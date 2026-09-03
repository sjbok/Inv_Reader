"""Summarize general documents and consolidate XLSX order forms."""

import argparse
import base64
import ctypes
from datetime import date, datetime
import json
import os
import re
import subprocess
import sys
import time
import urllib.error
import urllib.request
from contextlib import nullcontext, redirect_stderr, redirect_stdout
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Dict, List, Optional, Sequence
from urllib.parse import urlparse


DEFAULT_MODEL = "qwen3-vl:8b"
DEFAULT_OLLAMA_URL = "http://localhost:11434"
PORTABLE_OLLAMA_URL = "http://127.0.0.1:11435"
PORTABLE_CONFIG_NAME = "portable_config.json"
DEFAULT_CHUNK_SIZE = 24000
DEFAULT_MAX_OUTPUT_TOKENS = 1024
OLLAMA_STARTUP_TIMEOUT = 120
OUTPUT_WORKBOOK_NAME = "2026 한진양식.xlsx"
DATABASE_HEADERS = ("이름", "전화", "주소", "품목", "업체명", "특기사항/배송 메모")
INPUT_LOG_DATE_FORMAT = "%Y-%m-%d"

TEXT_EXTENSIONS = {
    ".csv",
    ".html",
    ".htm",
    ".json",
    ".log",
    ".markdown",
    ".md",
    ".rtf",
    ".txt",
    ".xml",
}
IMAGE_EXTENSIONS = {".bmp", ".gif", ".jpeg", ".jpg", ".png", ".tif", ".tiff", ".webp"}

SUMMARY_INSTRUCTIONS = """You summarize documents accurately and concisely.
Documents may be written in English, Korean, or a mixture of both. Read and understand
both languages, including Korean names, dates, numbers, and proper nouns. Preserve
important original Korean and English wording when listing specific details. For names,
proper nouns, addresses, item names, and company names, copy the wording exactly as it
appears in the document; do not translate, romanize, or normalize it.
Return plain text only, with no markdown code fences and no discussion of this prompt.
Include:
- a short title or description
- the document's purpose and main points
- important names, dates, identifiers, quantities, and monetary amounts
- specifically look for and explicitly list these fields when present: 이름, 전화, 주소,
  품목 (수량), 업체명. Preserve the original Korean or English text for each field's value.
- decisions, action items, or next steps when present
- uncertainties or missing information
For dates, understand common formats such as YYYY/MM/DD, YY/MM/DD, and
DD/MM/YYYY or MM/DD/YYYY. When reporting a date, use YYYY-MM-DD when the date
is unambiguous.
Never invent information. If a detail is not present, omit it."""


@dataclass
class DocumentContent:
    """Content prepared for the multimodal model."""

    text: str = ""
    images: Optional[List[str]] = None


@dataclass
class Item:
    """One item from an order form."""

    code: str = ""
    name: str = ""
    quantity: str = ""
    unit: str = ""


@dataclass
class PurchaseOrder:
    """Only the fields that are copied into the output database."""

    order_date: Optional[date] = None
    recipient: str = ""
    phone: str = ""
    address: str = ""
    company: str = ""
    memo: str = ""
    items: Optional[List[Item]] = None


@dataclass
class ConfidenceAnalysis:
    """Completeness score for one source file."""

    filename: str
    confidence: float


@dataclass
class ProcessingReport:
    """Issues found while processing one input batch."""

    duplicates: List[str]
    low_confidence: List[ConfidenceAnalysis]


class OllamaClient:
    """Minimal client for the local Ollama chat API."""

    def __init__(self, model: str = DEFAULT_MODEL, base_url: str = DEFAULT_OLLAMA_URL,
                 timeout: int = 600) -> None:
        self.model = model
        self.endpoint = base_url.rstrip("/") + "/api/chat"
        self.timeout = timeout

    def chat(self, prompt: str, images: Optional[Sequence[str]] = None) -> str:
        message = {"role": "user", "content": prompt}
        if images:
            message["images"] = list(images)

        payload = {
            "model": self.model,
            "messages": [message],
            "stream": False,
            "think": False,
            "options": {"temperature": 0.2},
        }
        payload["options"]["num_predict"] = DEFAULT_MAX_OUTPUT_TOKENS
        if not images:
            # qwen3-vl:8b has a bare template that ignores the API think flag.
            # Prefilling an empty thinking block keeps text requests finite.
            payload["raw"] = True
            payload["messages"].append({
                "role": "assistant",
                "content": "<think>\n\n</think>\n\n",
            })
        request = urllib.request.Request(
            self.endpoint,
            data=json.dumps(payload).encode("utf-8"),
            headers={"Content-Type": "application/json"},
            method="POST",
        )
        try:
            with urllib.request.urlopen(request, timeout=self.timeout) as response:
                result = json.loads(response.read().decode("utf-8"))
        except urllib.error.HTTPError as error:
            details = error.read().decode("utf-8", errors="replace").strip()
            raise RuntimeError("Ollama returned HTTP {}: {}".format(error.code, details)) from error
        except urllib.error.URLError as error:
            raise RuntimeError(
                "Could not connect to Ollama at {}. Start Ollama and make sure {} is available."
                .format(self.endpoint, self.model)
            ) from error

        try:
            response_message = result["message"]
            content = response_message["content"]
        except (KeyError, TypeError) as error:
            raise RuntimeError("Ollama returned an unexpected response: {}".format(result)) from error
        if not isinstance(content, str) or not content.strip():
            if isinstance(response_message, dict) and response_message.get("thinking"):
                raise RuntimeError(
                    "Ollama returned only thinking and no summary content; use an instruct "
                    "model or a model with a working thinking toggle"
                )
            raise RuntimeError("Ollama returned an empty summary")
        return content.strip()


def application_root() -> Path:
    """Return the directory that owns the executable, or the project root in development."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent.parent


def bundled_ollama_path(root: Path) -> Optional[Path]:
    """Find the Ollama binary included in a portable distribution."""
    candidates = [root / "ollama" / "ollama.exe", root / "ollama.exe"]
    if os.name != "nt":
        candidates.extend([root / "ollama" / "ollama", root / "ollama"])
    for candidate in candidates:
        if candidate.is_file():
            return candidate
    return None


def _runtime_path(path: Path, root: Path) -> Path:
    return path if path.is_absolute() else root / path


def configured_model(root: Path) -> str:
    """Read the model selected by the Windows build, if one was configured."""
    config_path = root / PORTABLE_CONFIG_NAME
    try:
        config = json.loads(config_path.read_text(encoding="utf-8-sig"))
    except (OSError, ValueError):
        return DEFAULT_MODEL
    model = config.get("model") if isinstance(config, dict) else None
    return model if isinstance(model, str) and model.strip() else DEFAULT_MODEL


class OllamaRuntime:
    """Own a private Ollama server for the lifetime of the application."""

    def __init__(self, executable: Path, models_dir: Path, base_url: str,
                 model: str, log_path: Path,
                 startup_timeout: int = OLLAMA_STARTUP_TIMEOUT) -> None:
        self.executable = executable
        self.models_dir = models_dir
        self.base_url = base_url.rstrip("/")
        self.model = model
        self.log_path = log_path
        self.startup_timeout = startup_timeout
        self.process = None
        self._log = None

    def __enter__(self) -> "OllamaRuntime":
        self.start()
        return self

    def __exit__(self, exc_type, exc_value, traceback) -> None:
        self.stop()

    def start(self) -> None:
        self.models_dir.mkdir(parents=True, exist_ok=True)
        self._log = self.log_path.open("a", encoding="utf-8")
        environment = os.environ.copy()
        environment["OLLAMA_HOST"] = self._ollama_host()
        environment["OLLAMA_MODELS"] = str(self.models_dir)
        creationflags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
        try:
            self.process = subprocess.Popen(
                [str(self.executable), "serve"],
                stdin=subprocess.DEVNULL,
                stdout=self._log,
                stderr=subprocess.STDOUT,
                env=environment,
                creationflags=creationflags,
            )
            self._wait_until_ready()
            if not self._model_available():
                self._pull_model(environment, creationflags)
                if not self._model_available():
                    raise RuntimeError(
                        "Ollama started, but model {} was not available after downloading it. "
                        "See {}.".format(self.model, self.log_path)
                    )
        except Exception:
            self.stop()
            raise

    def stop(self) -> None:
        process = self.process
        self.process = None
        if process is not None:
            if os.name == "nt":
                # Ollama can launch a model runner child. Kill the owned tree so
                # a runner cannot survive after the EXE exits.
                subprocess.run(
                    ["taskkill", "/PID", str(process.pid), "/T", "/F"],
                    check=False,
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL,
                    creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
                )
            elif process.poll() is None:
                process.terminate()
                try:
                    process.wait(timeout=10)
                except subprocess.TimeoutExpired:
                    process.kill()
                    process.wait()
        if self._log is not None:
            self._log.close()
            self._log = None

    def _ollama_host(self) -> str:
        parsed = urlparse(self.base_url)
        return "{}:{}".format(parsed.hostname or "127.0.0.1", parsed.port or 11434)

    def _wait_until_ready(self) -> None:
        deadline = time.monotonic() + self.startup_timeout
        tags_url = self.base_url + "/api/tags"
        last_error = "unknown error"
        while time.monotonic() < deadline:
            if self.process is not None and self.process.poll() is not None:
                raise RuntimeError(
                    "Ollama exited while starting. See {}.".format(self.log_path)
                )
            try:
                with urllib.request.urlopen(tags_url, timeout=2) as response:
                    json.loads(response.read().decode("utf-8"))
                return
            except (OSError, ValueError) as error:
                last_error = str(error)
                time.sleep(0.5)
        raise RuntimeError(
            "Ollama did not become ready within {} seconds: {}. See {}.".format(
                self.startup_timeout, last_error, self.log_path
            )
        )

    def _model_available(self) -> bool:
        tags_url = self.base_url + "/api/tags"
        try:
            with urllib.request.urlopen(tags_url, timeout=10) as response:
                result = json.loads(response.read().decode("utf-8"))
        except (OSError, ValueError) as error:
            raise RuntimeError("Could not inspect Ollama models: {}".format(error)) from error
        models = result.get("models", []) if isinstance(result, dict) else []
        return any(isinstance(item, dict) and item.get("name") == self.model for item in models)

    def _pull_model(self, environment: Dict[str, str], creationflags: int) -> None:
        if self._log is None:
            raise RuntimeError("Ollama log is not open")
        self._log.write("Model {} is missing; downloading it.\n".format(self.model))
        self._log.flush()
        pull = subprocess.Popen(
            [str(self.executable), "pull", self.model],
            stdin=subprocess.DEVNULL,
            stdout=self._log,
            stderr=subprocess.STDOUT,
            env=environment,
            creationflags=creationflags,
        )
        if pull.wait() != 0:
            raise RuntimeError(
                "Could not download Ollama model {}. See {}.".format(self.model, self.log_path)
            )


def _read_text(path: Path) -> str:
    return path.read_bytes().decode("utf-8-sig", errors="replace").strip()


def _read_docx(path: Path) -> str:
    try:
        from docx import Document
    except ImportError as error:
        raise RuntimeError("DOCX support requires python-docx; install requirements.txt") from error

    document = Document(str(path))
    sections = [paragraph.text.strip() for paragraph in document.paragraphs if paragraph.text.strip()]
    for table in document.tables:
        rows = []
        for row in table.rows:
            rows.append(" | ".join(cell.text.strip() for cell in row.cells))
        if rows:
            sections.append("\n".join(rows))
    return "\n\n".join(sections).strip()


def _read_xlsx(path: Path) -> str:
    try:
        from openpyxl import load_workbook
    except ImportError as error:
        raise RuntimeError("XLSX support requires openpyxl; install requirements.txt") from error

    workbook = load_workbook(str(path), read_only=True, data_only=True)
    sheets = []
    try:
        for worksheet in workbook.worksheets:
            rows = []
            for values in worksheet.iter_rows(values_only=True):
                raw_values = list(values)
                cells = [_cell_text(value) for value in values]
                for column, value in enumerate(cells):
                    if _normalise_label(value) != _normalise_label("발주일자"):
                        continue
                    for date_column in range(column + 1, len(cells)):
                        if not cells[date_column]:
                            continue
                        parsed_date = _parse_order_date(raw_values[date_column])
                        if parsed_date is not None:
                            cells[date_column] = parsed_date.isoformat()
                        break
                while cells and not cells[-1]:
                    cells.pop()
                if any(cells):
                    rows.append("\t".join(cells))
            if rows:
                sheets.append("[Sheet: {}]\n{}".format(worksheet.title, "\n".join(rows)))
    finally:
        workbook.close()
    return "\n\n".join(sheets).strip()


def _normalise_label(value: Any) -> str:
    """Normalize labels while keeping the actual extracted values unchanged."""
    if value is None:
        return ""
    value = getattr(value, "value", value)
    return re.sub(r"[\s:：•·/]", "", str(value)).casefold()


def _cell_text(value: Any) -> str:
    value = getattr(value, "value", value)
    if value is None:
        return ""
    if isinstance(value, datetime):
        return value.isoformat(sep=" ")
    if isinstance(value, date):
        return value.isoformat()
    if isinstance(value, float) and value.is_integer():
        return str(int(value))
    return str(value).strip()


def _is_grey_italic_cell(value: Any) -> bool:
    """Return whether a cell uses the form's grey italic example style."""
    font = getattr(value, "font", None)
    if font is None or not font.italic:
        return False
    color = font.color
    if color is None:
        return False
    if color.type == "rgb" and color.rgb:
        rgb = color.rgb[-6:]
        return rgb[0:2] == rgb[2:4] == rgb[4:6] and rgb[0:2] not in {"00", "FF"}
    # Some workbooks use a custom indexed palette for the same grey style.
    return color.type == "indexed" and color.indexed not in {8, 9}


def _is_example_item_row(row: tuple) -> bool:
    return any(_is_grey_italic_cell(cell) and _cell_text(cell) for cell in row)


def _row_field(rows: List[tuple], label: str) -> str:
    """Return the first nonempty cell to the right of a labeled cell."""
    wanted = _normalise_label(label)
    for row in rows:
        for column, value in enumerate(row):
            if _normalise_label(value) != wanted:
                continue
            for candidate in row[column + 1:]:
                text = _cell_text(candidate)
                if text:
                    return text
    return ""


def _memo_field(rows: List[tuple]) -> str:
    wanted = _normalise_label("특기사항/배송 메모")
    for row_number, row in enumerate(rows):
        for column, value in enumerate(row):
            if _normalise_label(value) != wanted:
                continue
            for candidate in row[column + 1:]:
                text = _cell_text(candidate)
                if text:
                    return text
            for following_row in rows[row_number + 1:]:
                for candidate in following_row:
                    text = _cell_text(candidate)
                    if text:
                        return text
    return ""


def _parse_order_date(value: Any) -> Optional[date]:
    if isinstance(value, datetime):
        return value.date()
    if isinstance(value, date):
        return value

    if isinstance(value, (int, float)) and not isinstance(value, bool):
        if 1 <= value <= 2958465:
            try:
                from openpyxl.utils.datetime import from_excel
                converted = from_excel(value)
                if isinstance(converted, datetime):
                    return converted.date()
                if isinstance(converted, date):
                    return converted
            except (ImportError, TypeError, ValueError, OverflowError):
                pass

    text = _cell_text(value)
    compact_match = re.search(r"(?<!\d)(\d{8})(?!\d)", text)
    if compact_match:
        parts = [int(compact_match.group(1)[:4]), int(compact_match.group(1)[4:6]),
                 int(compact_match.group(1)[6:])]
        return _valid_date_parts(*parts)

    serial_match = re.fullmatch(r"\d{5}", text)
    if serial_match:
        try:
            from openpyxl.utils.datetime import from_excel
            converted = from_excel(int(text))
            if isinstance(converted, datetime):
                return converted.date()
            if isinstance(converted, date):
                return converted
        except (ImportError, TypeError, ValueError, OverflowError):
            pass

    parts = re.findall(r"\d{1,4}", text)
    if len(parts) < 3:
        return None
    first, second, third = parts[:3]
    first_number, second_number, third_number = map(int, (first, second, third))

    if len(first) == 4:
        return _valid_date_parts(first_number, second_number, third_number)
    if len(third) == 4:
        if first_number > 12 and second_number <= 12:
            return _valid_date_parts(third_number, second_number, first_number)
        if second_number > 12 and first_number <= 12:
            return _valid_date_parts(third_number, first_number, second_number)
        # Korean order is the least surprising interpretation when both are valid months.
        return _valid_date_parts(third_number, second_number, first_number)

    # Two-digit years are normally written first in Korean order (26/09/02).
    # If that interpretation is not plausible, support day/month/year and month/day/year.
    if (20 <= first_number <= 69 and 1 <= second_number <= 12 and
            1 <= third_number <= 31):
        return _valid_date_parts(2000 + first_number, second_number, third_number)
    if first_number > 12 and second_number <= 12:
        return _valid_date_parts(2000 + third_number, second_number, first_number)
    if second_number > 12 and first_number <= 12:
        return _valid_date_parts(2000 + third_number, first_number, second_number)
    return _valid_date_parts(2000 + third_number, second_number, first_number)


def _valid_date_parts(year: int, month: int, day: int) -> Optional[date]:
    try:
        return date(year, month, day)
    except ValueError:
        return None


def _item_columns(rows: List[tuple]) -> Optional[Dict[str, int]]:
    aliases = {
        "code": _normalise_label("품목코드"),
        "name": _normalise_label("품목명[규격]"),
        "quantity": _normalise_label("수량"),
        "unit": _normalise_label("단위"),
    }
    for row_number, row in enumerate(rows):
        columns = {}
        for column, value in enumerate(row):
            normalized = _normalise_label(value)
            for field, wanted in aliases.items():
                if normalized == wanted:
                    columns[field] = column
        if len(columns) == len(aliases):
            return {"header": row_number, **columns}
    return None


def _extract_order_from_rows(rows: List[tuple]) -> PurchaseOrder:
    columns = _item_columns(rows)
    items = []
    section_labels = {
        _normalise_label("특기사항/배송 메모"),
        _normalise_label("발주자 정보"),
        _normalise_label("납품 • 배송정보"),
    }
    if columns is not None:
        for row in rows[columns["header"] + 1:]:
            if any(_normalise_label(value) in section_labels for value in row):
                continue
            if _is_example_item_row(row):
                continue
            values = {
                field: _cell_text(row[column]) if column < len(row) else ""
                for field, column in columns.items()
                if field != "header"
            }
            if not any(values.values()):
                continue
            # Footer and delivery-note rows do not have an item-column value.
            items.append(Item(**values))

    return PurchaseOrder(
        order_date=_parse_order_date(_row_field(rows, "발주일자")),
        recipient=_row_field(rows, "수령인"),
        phone=_row_field(rows, "수령인 연락처"),
        address=_row_field(rows, "배송지 주소"),
        company=_row_field(rows, "발주처"),
        memo=_memo_field(rows),
        items=items,
    )


def extract_purchase_order(path: Path) -> PurchaseOrder:
    """Extract the structured order fields from an XLSX workbook."""
    try:
        from openpyxl import load_workbook
    except ImportError as error:
        raise RuntimeError("XLSX support requires openpyxl; install requirements.txt") from error

    workbook = load_workbook(str(path), read_only=False, data_only=True)
    try:
        for worksheet in workbook.worksheets:
            rows = list(worksheet.iter_rows())
            order = _extract_order_from_rows(rows)
            if any((order.order_date, order.recipient, order.phone, order.address,
                    order.company, order.memo, order.items)):
                return order
    finally:
        workbook.close()
    return PurchaseOrder(items=[])


def confidence_level(order: PurchaseOrder) -> float:
    """Score the completeness of the fields needed by the output database."""
    checks = [
        order.order_date is not None,
        bool(order.recipient),
        bool(order.phone),
        bool(order.address),
        bool(order.company),
        bool(order.memo),
        bool(order.items),
    ]
    for item in order.items or []:
        checks.append(all((item.code, item.name, item.quantity, item.unit)))
    if not checks:
        return 0.0
    return round(100.0 * sum(checks) / len(checks), 2)


def _item_text(item: Item) -> str:
    parts = []
    if item.code:
        parts.append("({})".format(item.code))
    if item.name:
        parts.append(item.name)
    quantity = "{}{}".format(item.quantity, item.unit)
    if quantity:
        parts.append(quantity)
    return " ".join(parts)


def _database_row(order: PurchaseOrder, item: Item) -> List[str]:
    return [
        order.recipient,
        order.phone,
        order.address,
        _item_text(item),
        order.company,
        order.memo,
    ]


def _sheet_name(order: PurchaseOrder) -> str:
    if order.order_date is None:
        raise ValueError("Order is missing a valid 발주일자")
    return order.order_date.strftime("%m.%d")


def _new_database_workbook(output_path: Path):
    from openpyxl import Workbook, load_workbook

    if output_path.exists():
        return load_workbook(str(output_path))
    workbook = Workbook()
    workbook.remove(workbook.active)
    return workbook


def _input_log_path(output_dir: Path, order_date: date) -> Path:
    return output_dir / (order_date.strftime(INPUT_LOG_DATE_FORMAT) + ".txt")


def _read_input_log(path: Path) -> set:
    if not path.exists():
        return set()
    return {
        line.strip()
        for line in path.read_text(encoding="utf-8").splitlines()
        if line.strip()
    }


def _append_input_log(path: Path, filename: str) -> None:
    with path.open("a", encoding="utf-8") as log:
        log.write(filename + "\n")


def _prepare_database_sheet(workbook: Any, name: str) -> Any:
    from openpyxl.styles import Alignment, Font, PatternFill

    if name in workbook.sheetnames:
        worksheet = workbook[name]
    else:
        worksheet = workbook.create_sheet(name)

    first_row = [worksheet.cell(1, column).value for column in range(1, len(DATABASE_HEADERS) + 1)]
    if not any(value is not None for value in first_row):
        for column, header in enumerate(DATABASE_HEADERS, 1):
            cell = worksheet.cell(1, column, header)
            cell.font = Font(bold=True)
            cell.fill = PatternFill("solid", fgColor="D9EAF7")
            cell.alignment = Alignment(horizontal="center", vertical="center")
        worksheet.freeze_panes = "A2"
    elif tuple(first_row) != DATABASE_HEADERS:
        raise ValueError("Existing sheet '{}' does not use the expected database columns".format(name))

    widths = (18, 18, 45, 45, 28, 45)
    for column, width in enumerate(widths, 1):
        worksheet.column_dimensions[chr(64 + column)].width = width
    worksheet.auto_filter.ref = "A1:F{}".format(max(worksheet.max_row, 1))
    return worksheet


def append_order_to_workbook(workbook: Any, order: PurchaseOrder) -> int:
    """Append every item in an order and return the number of rows written."""
    worksheet = _prepare_database_sheet(workbook, _sheet_name(order))
    rows_written = 0
    for item in order.items or []:
        next_row = max(worksheet.max_row + 1, 2)
        worksheet.cell(next_row, 1).value = order.recipient
        for column, value in enumerate(_database_row(order, item), 1):
            worksheet.cell(next_row, column).value = value
        rows_written += 1
    worksheet.auto_filter.ref = "A1:F{}".format(max(worksheet.max_row, 1))
    return rows_written


def _read_pdf(path: Path) -> DocumentContent:
    try:
        import fitz
    except ImportError:
        fitz = None

    if fitz is not None:
        with fitz.open(str(path)) as document:
            pages = [page.get_text("text", sort=True).strip() for page in document]
        text = "\n\n".join("[Page {}]\n{}".format(index, page)
                              for index, page in enumerate(pages, 1) if page).strip()
        if text:
            return DocumentContent(text=text)

    try:
        from pypdf import PdfReader
    except ImportError as error:
        raise RuntimeError("PDF support requires pypdf; install requirements.txt") from error

    reader = PdfReader(str(path))
    pages = [page.extract_text() or "" for page in reader.pages]
    text = "\n\n".join("[Page {}]\n{}".format(index, page.strip())
                          for index, page in enumerate(pages, 1) if page.strip()).strip()
    if text:
        return DocumentContent(text=text)

    if fitz is None:
        raise RuntimeError(
            "This PDF has no extractable text and scanned-PDF support requires PyMuPDF; "
            "install requirements.txt"
        )

    images = []
    with fitz.open(str(path)) as document:
        for page in document:
            pixmap = page.get_pixmap(matrix=fitz.Matrix(1.5, 1.5), alpha=False)
            images.append(base64.b64encode(pixmap.tobytes("png")).decode("ascii"))
    return DocumentContent(images=images)


def read_document(path: Path) -> DocumentContent:
    """Read a supported document into text and/or base64 encoded images."""
    suffix = path.suffix.lower()
    if suffix in TEXT_EXTENSIONS:
        return DocumentContent(text=_read_text(path))
    if suffix == ".docx":
        return DocumentContent(text=_read_docx(path))
    if suffix == ".xlsx":
        return DocumentContent(text=_read_xlsx(path))
    if suffix == ".pdf":
        return _read_pdf(path)
    if suffix in IMAGE_EXTENSIONS:
        image = base64.b64encode(path.read_bytes()).decode("ascii")
        return DocumentContent(images=[image])
    raise ValueError("Unsupported document type: {}".format(path.suffix or "no extension"))


def _chunks(text: str, chunk_size: int) -> List[str]:
    paragraphs = text.split("\n\n")
    chunks = []
    current = ""
    for paragraph in paragraphs:
        paragraph = paragraph.strip()
        if not paragraph:
            continue
        if current and len(current) + len(paragraph) + 2 > chunk_size:
            chunks.append(current)
            current = ""
        if len(paragraph) > chunk_size:
            if current:
                chunks.append(current)
                current = ""
            chunks.extend(paragraph[index:index + chunk_size]
                          for index in range(0, len(paragraph), chunk_size))
        else:
            current = paragraph if not current else current + "\n\n" + paragraph
    if current:
        chunks.append(current)
    return chunks


def summarize_content(content: DocumentContent, client: OllamaClient,
                      chunk_size: int = DEFAULT_CHUNK_SIZE) -> str:
    """Summarize extracted document content, consolidating long text when needed."""
    if content.images:
        prompt = SUMMARY_INSTRUCTIONS + "\n\nRead all supplied pages/images as one document and summarize them."
        if content.text:
            prompt += "\n\nExtracted text:\n" + content.text
        return client.chat(prompt, content.images)

    if not content.text.strip():
        raise ValueError("Document contains no readable text or images")
    parts = _chunks(content.text, chunk_size)
    if len(parts) == 1:
        return client.chat(SUMMARY_INSTRUCTIONS + "\n\nDOCUMENT:\n" + parts[0])

    partial_summaries = []
    for index, part in enumerate(parts, 1):
        partial_summaries.append(client.chat(
            SUMMARY_INSTRUCTIONS +
            "\nThis is part {}/{} of a longer document. Summarize only this part, "
            "preserving concrete details.\n\nDOCUMENT PART:\n{}".format(index, len(parts), part)
        ))
    combined = "\n\n".join("PART {} SUMMARY:\n{}".format(index, summary)
                            for index, summary in enumerate(partial_summaries, 1))
    return client.chat(
        SUMMARY_INSTRUCTIONS +
        "\n\nCombine the following partial summaries into one coherent summary for the full document. "
        "Remove repetition and retain the most important concrete details.\n\n{}".format(combined)
    )


def _is_relative_to(path: Path, parent: Path) -> bool:
    try:
        path.relative_to(parent)
        return True
    except ValueError:
        return False


def iter_documents(input_dir: Path, output_dir: Optional[Path] = None) -> List[Path]:
    """Return supported files recursively, excluding the output directory."""
    resolved_output = output_dir.resolve() if output_dir else None
    documents = []
    for path in sorted(input_dir.rglob("*")):
        if not path.is_file() or path.name.startswith("."):
            continue
        if resolved_output and _is_relative_to(path.resolve(), resolved_output):
            continue
        if path.suffix.lower() in TEXT_EXTENSIONS | IMAGE_EXTENSIONS | {".docx", ".pdf", ".xlsx"}:
            documents.append(path)
    return documents


def output_name(path: Path, input_dir: Path) -> str:
    """Create a collision-resistant flat output name from an input relative path."""
    relative = path.relative_to(input_dir)
    parts = [part.replace("__", "____") for part in relative.with_suffix("").parts]
    return "__".join(parts) + ".txt"


def process_documents(input_dir: Path, output_dir: Path, client: Optional[OllamaClient] = None,
                      chunk_size: int = DEFAULT_CHUNK_SIZE,
                      report: Optional[ProcessingReport] = None) -> int:
    input_dir = input_dir.resolve()
    output_dir = output_dir.resolve()
    if not input_dir.is_dir():
        raise FileNotFoundError("Input directory does not exist: {}".format(input_dir))
    output_dir.mkdir(parents=True, exist_ok=True)
    documents = iter_documents(input_dir, output_dir)
    if not documents:
        raise FileNotFoundError("No supported documents found in {}".format(input_dir))

    failures = 0
    xlsx_documents = [path for path in documents if path.suffix.lower() == ".xlsx"]
    other_documents = [path for path in documents if path.suffix.lower() != ".xlsx"]
    confidence_results = []

    if xlsx_documents:
        database_path = output_dir / OUTPUT_WORKBOOK_NAME
        database_existed = database_path.exists()
        workbook = _new_database_workbook(database_path)
        extracted_orders = []
        recorded_filenames = {}
        pending_log_entries = []
        for path in xlsx_documents:
            print("Reading {} -> {}".format(path, database_path))
            try:
                order = extract_purchase_order(path)
                confidence_results.append(
                    ConfidenceAnalysis(path.name, confidence_level(order))
                )
                extracted_orders.append((path, order))
            except Exception as error:
                failures += 1
                confidence_results.append(ConfidenceAnalysis(path.name, 0.0))
                print("ERROR: {}: {}".format(path, error), file=sys.stderr)

        for path, order in sorted(
            extracted_orders,
            key=lambda result: (_sheet_name(result[1]) if result[1].order_date else "99.99",
                                result[0].name),
        ):
            if order.order_date is None:
                try:
                    _sheet_name(order)
                except Exception as error:
                    failures += 1
                    print("ERROR: {}: {}".format(path, error), file=sys.stderr)
                continue

            log_path = _input_log_path(output_dir, order.order_date)
            if log_path not in recorded_filenames:
                recorded_filenames[log_path] = _read_input_log(log_path)
            if path.name in recorded_filenames[log_path]:
                if report is not None:
                    report.duplicates.append(str(path.relative_to(input_dir)))
                print("DUPLICATE: {} is already recorded in {}; skipping.".format(
                    path.name, log_path.name
                ))
                continue
            try:
                append_order_to_workbook(workbook, order)
                recorded_filenames[log_path].add(path.name)
                pending_log_entries.append((log_path, path.name))
            except Exception as error:
                failures += 1
                print("ERROR: {}: {}".format(path, error), file=sys.stderr)
        workbook_saved = False
        try:
            if workbook.sheetnames or database_existed:
                workbook.save(str(database_path))
                workbook_saved = True
        finally:
            workbook.close()
        if workbook_saved:
            for log_path, filename in pending_log_entries:
                try:
                    _append_input_log(log_path, filename)
                except Exception as error:
                    failures += 1
                    print("ERROR: {}: could not record {}: {}".format(
                        log_path, filename, error
                    ), file=sys.stderr)

    if other_documents:
        client = client or OllamaClient()
    for path in other_documents:
        destination = output_dir / output_name(path, input_dir)
        print("Summarizing {} -> {}".format(path, destination))
        try:
            summary = summarize_content(read_document(path), client, chunk_size)
            destination.write_text(summary.rstrip() + "\n", encoding="utf-8")
        except Exception as error:
            failures += 1
            print("ERROR: {}: {}".format(path, error), file=sys.stderr)

    for result in confidence_results:
        print("Confidence: {} - {:.2f}%".format(result.filename, result.confidence))
    for result in confidence_results:
        if result.confidence < 90.0:
            print("HUMAN REVIEW REQUIRED: {} - {:.2f}%".format(
                result.filename, result.confidence
            ))
    if report is not None:
        report.low_confidence.extend(
            result for result in confidence_results if result.confidence < 90.0
        )
    return failures


def parse_args(argv: Optional[Sequence[str]] = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Read input documents and append XLSX orders to an Excel database."
    )
    parser.add_argument("--input", type=Path, default=Path("input"), help="Input folder (default: input)")
    parser.add_argument("--output", type=Path, default=Path("output"), help="Output folder (default: output)")
    parser.add_argument("--model", default=None,
                        help="Ollama model (default: portable config or {})".format(DEFAULT_MODEL))
    parser.add_argument("--ollama-url", default=None,
                        help="Ollama base URL (default: bundled server or {})".format(
                            DEFAULT_OLLAMA_URL
                        ))
    parser.add_argument("--ollama-executable", type=Path, default=None,
                        help="Ollama executable to manage (normally found beside the EXE)")
    parser.add_argument("--ollama-models", type=Path, default=None,
                        help="Ollama model directory (default: models beside the EXE)")
    parser.add_argument("--ollama-startup-timeout", type=int, default=OLLAMA_STARTUP_TIMEOUT,
                        help="Seconds to wait for bundled Ollama (default: %(default)s)")
    parser.add_argument("--no-ollama-management", action="store_true",
                        help="Do not start or stop Ollama automatically")
    parser.add_argument("--timeout", type=int, default=600, help="Per-request timeout in seconds")
    parser.add_argument("--chunk-size", type=int, default=DEFAULT_CHUNK_SIZE,
                        help="Maximum extracted text per model request")
    args = parser.parse_args(argv)
    if args.timeout <= 0 or args.chunk_size <= 0 or args.ollama_startup_timeout <= 0:
        parser.error("--timeout, --chunk-size, and --ollama-startup-timeout must be positive")
    return args


def _run_application(argv: Optional[Sequence[str]] = None) -> int:
    args = parse_args(argv)
    root = application_root()
    model = args.model or configured_model(root)
    input_dir = _runtime_path(args.input, root)
    output_dir = _runtime_path(args.output, root)
    input_dir.mkdir(parents=True, exist_ok=True)
    output_dir.mkdir(parents=True, exist_ok=True)
    documents = iter_documents(input_dir, output_dir)
    needs_ollama = any(path.suffix.lower() != ".xlsx" for path in documents)

    executable = _runtime_path(args.ollama_executable, root) if args.ollama_executable else (
        bundled_ollama_path(root) if not args.no_ollama_management else None
    )
    models_dir = _runtime_path(args.ollama_models, root) if args.ollama_models else root / "models"
    if executable and needs_ollama:
        ollama_url = args.ollama_url or PORTABLE_OLLAMA_URL
        runtime = OllamaRuntime(
            executable,
            models_dir,
            ollama_url,
            model,
            root / "run.log",
            args.ollama_startup_timeout,
        )
    else:
        ollama_url = args.ollama_url or DEFAULT_OLLAMA_URL
        runtime = nullcontext()
    report = ProcessingReport([], [])

    try:
        if needs_ollama and getattr(sys, "frozen", False) and not executable:
            raise FileNotFoundError(
                "Bundled ollama.exe was not found beside the application. "
                "Copy the complete portable application directory."
            )
        with runtime:
            failures = process_documents(
                input_dir,
                output_dir,
                OllamaClient(model, ollama_url, args.timeout),
                args.chunk_size,
                report,
            )
        show_issue_popups(report)
    except (FileNotFoundError, OSError, RuntimeError) as error:
        print("ERROR: {}".format(error), file=sys.stderr)
        return 1
    if failures:
        print("{} document(s) failed.".format(failures), file=sys.stderr)
        return 1
    return 0


def show_issue_popups(report: ProcessingReport) -> None:
    """Show grouped processing issues in the windowed Windows application."""
    if not getattr(sys, "frozen", False) or os.name != "nt":
        return
    try:
        if report.duplicates:
            duplicate_list = "\n".join("- {}".format(filename) for filename in report.duplicates)
            ctypes.windll.user32.MessageBoxW(
                None,
                "These files were already imported and were skipped:\n\n{}".format(duplicate_list),
                "Duplicate files",
                0x30,
            )
        if report.low_confidence:
            review_list = "\n".join(
                "- {}: {:.2f}%".format(result.filename, result.confidence)
                for result in report.low_confidence
            )
            ctypes.windll.user32.MessageBoxW(
                None,
                "Human review is required for these files:\n\n{}".format(review_list),
                "Human review required",
                0x30,
            )
    except (AttributeError, OSError) as error:
        print("ERROR: could not show issue popup: {}".format(error), file=sys.stderr)


def main(argv: Optional[Sequence[str]] = None) -> int:
    # Windowed PyInstaller builds do not provide stdout/stderr. Keep a diagnostic
    # log beside the EXE so a failed background run is still explainable.
    if getattr(sys, "frozen", False):
        log_path = application_root() / "run.log"
        with log_path.open("a", encoding="utf-8") as log:
            with redirect_stdout(log), redirect_stderr(log):
                return _run_application(argv)
    return _run_application(argv)


if __name__ == "__main__":
    raise SystemExit(main())

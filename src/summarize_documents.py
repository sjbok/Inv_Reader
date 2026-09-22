"""Summarize general documents and consolidate XLSX order forms."""

import argparse
import base64
import ctypes
from datetime import date, datetime
import json
import os
import queue
import re
import subprocess
import sys
import threading
import time
import urllib.error
import urllib.request
from contextlib import nullcontext, redirect_stderr, redirect_stdout
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Callable, Dict, List, Optional, Sequence
from urllib.parse import urlparse


DEFAULT_MODEL = "qwen3-vl:8b"
DEFAULT_OLLAMA_URL = "http://localhost:11434"
PORTABLE_OLLAMA_URL = "http://127.0.0.1:11435"
PORTABLE_CONFIG_NAME = "portable_config.json"
DEFAULT_CHUNK_SIZE = 24000
DEFAULT_MAX_OUTPUT_TOKENS = 1024
OLLAMA_STARTUP_TIMEOUT = 120
OUTPUT_WORKBOOK_NAME = "2026 한진양식.xlsx"
DATABASE_SHEET_NAME = "Orders"
DATABASE_HEADERS = (
    "이름", "전화", "우편번호", "주소", "수량", "품목", "운임타입", "지불조건",
    "특기사항", "업체명", "비고",
)
PREVIOUS_DATABASE_HEADERS = DATABASE_HEADERS[:-1]
LEGACY_DATABASE_HEADERS = ("이름", "전화", "주소", "품목", "업체명", "특기사항/배송 메모")
INPUT_LOG_NAME = "processed_files.txt"

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
    manager: str = ""
    email: str = ""
    business_registration_number: str = ""
    business_address: str = ""
    contact: str = ""
    memo: str = ""
    remarks: str = ""
    items: Optional[List[Item]] = None


@dataclass
class ConfidenceAnalysis:
    """Completeness score for one source file."""

    filename: str
    confidence: float


@dataclass
class FileProcessingStatus:
    """Current display status for one input file."""

    filename: str
    status: str
    confidence: Optional[float] = None
    message: str = ""


@dataclass
class MissingData:
    """Required input columns that were empty for one source file."""

    filename: str
    fields: List[str]


@dataclass
class InvalidData:
    """Input data that is present but fails a required format check."""

    filename: str
    field: str
    reason: str


@dataclass
class ProcessingReport:
    """Issues found while processing one input batch."""

    duplicates: List[str]
    low_confidence: List[ConfidenceAnalysis]
    missing_data: List[MissingData] = field(default_factory=list)
    invalid_data: List[InvalidData] = field(default_factory=list)


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
    field_labels = {
        _normalise_label("발주일자"),
        _normalise_label("발주처"),
        _normalise_label("담당자"),
        _normalise_label("이메일"),
        _normalise_label("발주사업자등록증번호"),
        _normalise_label("사업자 주소"),
        _normalise_label("연락처"),
        _normalise_label("수령인"),
        _normalise_label("수령인 연락처"),
        _normalise_label("배송지 주소"),
        _normalise_label("납품처"),
        _normalise_label("특기사항/배송 메모"),
        _normalise_label("비고"),
    }
    for row in rows:
        for column, value in enumerate(row):
            if _normalise_label(value) != wanted:
                continue
            for candidate in row[column + 1:]:
                if _normalise_label(candidate) in field_labels:
                    break
                text = _cell_text(candidate)
                if text:
                    return text
    return ""


def _memo_field(rows: List[tuple]) -> str:
    messages = _memo_messages(rows)
    return messages[0] if messages else ""


def _memo_messages(rows: List[tuple]) -> List[str]:
    wanted = {
        _normalise_label("특기사항/배송 메모"),
        _normalise_label("비고"),
    }
    section_labels = wanted | {
        _normalise_label("발주자 정보"),
        _normalise_label("납품 • 배송정보"),
        _normalise_label("발주일자"),
        _normalise_label("발주처"),
        _normalise_label("담당자"),
        _normalise_label("이메일"),
        _normalise_label("발주사업자등록증번호"),
        _normalise_label("사업자 주소"),
        _normalise_label("연락처"),
        _normalise_label("수령인"),
        _normalise_label("수령인 연락처"),
        _normalise_label("배송지 주소"),
        _normalise_label("납품처"),
    }
    messages = []
    for row_number, row in enumerate(rows):
        for column, value in enumerate(row):
            if _normalise_label(value) not in wanted:
                continue
            found_message = False
            for candidate in row[column + 1:]:
                if _normalise_label(candidate) in wanted:
                    break
                text = _cell_text(candidate)
                if text:
                    messages.append(text)
                    found_message = True
            if not found_message:
                for following_row in rows[row_number + 1:]:
                    following_messages = []
                    for candidate in following_row:
                        normalized = _normalise_label(candidate)
                        if normalized in section_labels:
                            break
                        text = _cell_text(candidate)
                        if text:
                            following_messages.append(text)
                    if following_messages:
                        messages.extend(following_messages)
                        break
    return messages


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
        _normalise_label("비고"),
        _normalise_label("발주자 정보"),
        _normalise_label("납품 • 배송정보"),
        _normalise_label("발주일자"),
        _normalise_label("발주처"),
        _normalise_label("담당자"),
        _normalise_label("이메일"),
        _normalise_label("발주사업자등록증번호"),
        _normalise_label("사업자 주소"),
        _normalise_label("연락처"),
        _normalise_label("수령인"),
        _normalise_label("수령인 연락처"),
        _normalise_label("배송지 주소"),
        _normalise_label("납품처"),
    }
    note_section_labels = {
        _normalise_label("특기사항/배송 메모"),
        _normalise_label("비고"),
    }
    if columns is not None:
        in_note_section = False
        for row in rows[columns["header"] + 1:]:
            if any(_normalise_label(value) in note_section_labels for value in row):
                in_note_section = True
                continue
            if in_note_section:
                continue
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
            values["unit"] = "개"
            # Footer and delivery-note rows do not have an item-column value.
            items.append(Item(**values))

    return PurchaseOrder(
        order_date=_parse_order_date(_row_field(rows, "발주일자")),
        recipient=_row_field(rows, "수령인"),
        phone=_row_field(rows, "수령인 연락처"),
        address=_row_field(rows, "배송지 주소"),
        company=_row_field(rows, "발주처"),
        manager=_row_field(rows, "담당자"),
        email=_row_field(rows, "이메일"),
        business_registration_number=_row_field(rows, "발주사업자등록증번호"),
        business_address=_row_field(rows, "사업자 주소"),
        contact=_row_field(rows, "연락처"),
        memo=_memo_field(rows),
        remarks="\n".join(_memo_messages(rows)),
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
                    order.company, order.manager, order.email,
                    order.business_registration_number, order.business_address,
                    order.contact, order.memo, order.remarks, order.items)):
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
        bool(order.items),
    ]
    for item in order.items or []:
        checks.append(all((item.name, item.quantity)))
    if not checks:
        return 0.0
    return round(100.0 * sum(checks) / len(checks), 2)


def confidence_reasons(order: PurchaseOrder) -> List[str]:
    """Return short field-level reasons for a confidence score below 100%."""
    reasons = []
    if order.order_date is None:
        reasons.append("발주일자")
    if not order.recipient:
        reasons.append("수령인")
    if not order.phone:
        reasons.append("수령인 연락처")
    if not order.address:
        reasons.append("배송지 주소")
    if not order.company:
        reasons.append("발주처")
    if not order.items:
        reasons.append("품목")
    for index, item in enumerate(order.items or [], 1):
        if not all((item.name, item.quantity)):
            reasons.append("품목 {}의 명칭/수량".format(index))
    return reasons


def missing_required_fields(order: PurchaseOrder) -> List[str]:
    """Return the input columns that must be populated before import."""
    missing = []
    recipient_is_phone = False
    if order.recipient and order.phone:
        recipient_digits = order.recipient.replace("-", "")
        phone_digits = order.phone.replace("-", "")
        recipient_is_phone = recipient_digits.isdigit() and recipient_digits == phone_digits
    required_fields = (
        ("발주일자", order.order_date),
        ("발주처", order.company),
        ("담당자", order.manager),
        ("이메일", order.email),
        ("발주사업자등록증번호", order.business_registration_number),
        ("사업자 주소", order.business_address),
        ("연락처", order.contact),
        ("수령인", "" if recipient_is_phone else order.recipient),
        ("수령인 연락처", order.phone),
        ("배송지 주소", order.address),
    )
    missing.extend(label for label, value in required_fields if not value)
    if not order.items:
        missing.append("품목")
    else:
        item_fields = (
            ("품목명", "name"),
            ("수량", "quantity"),
        )
        for index, item in enumerate(order.items, 1):
            for label, attribute in item_fields:
                if not getattr(item, attribute):
                    missing.append("{} (품목 {})".format(label, index))
    return missing


def _item_text(item: Item) -> str:
    description = []
    if item.code:
        description.append("({})".format(item.code))
    if item.name:
        description.append(item.name)
    quantity = "{}개".format(item.quantity) if item.quantity else ""
    item_description = " ".join(description)
    if item_description and quantity:
        return "{}-{}".format(item_description, quantity)
    return item_description or quantity


def _display_recipient(value: Any) -> Any:
    if not isinstance(value, str):
        return value
    if not value:
        return None
    return value if value.endswith("님") else value + "님"


def format_phone_number(phone: str) -> str:
    """Validate and normalize a Korean phone number for the output workbook."""
    value = phone.strip()
    digits = value.replace("-", "")
    if not digits.isdigit():
        raise ValueError("전화번호에는 숫자와 '-'만 사용할 수 있습니다")
    if not digits.startswith("0"):
        raise ValueError("전화번호는 0으로 시작해야 합니다")

    formats = {
        "010": {11: (3, 4, 4)},
        "02": {9: (2, 3, 4), 10: (2, 4, 4)},
        "03": {10: (3, 3, 4)},
        "04": {10: (3, 3, 4)},
        "05": {10: (3, 3, 4), 12: (4, 4, 4)},
        "06": {10: (3, 3, 4)},
        "07": {11: (3, 4, 4)},
    }
    prefix = next((candidate for candidate in formats if digits.startswith(candidate)), None)
    if prefix is None:
        raise ValueError("전화번호 국번은 010 또는 02~07이어야 합니다")
    if not 9 <= len(digits) <= 12:
        raise ValueError("하이픈을 제외한 전화번호 길이는 9~12자리여야 합니다")
    groups = formats[prefix].get(len(digits))
    if groups is None:
        raise ValueError("전화번호 형식에 맞지 않는 자리수입니다")
    parts = []
    position = 0
    for length in groups:
        parts.append(digits[position:position + length])
        position += length
    return "-".join(parts)


def _database_row(order: PurchaseOrder, item: Item) -> List[Any]:
    return [
        _display_recipient(order.recipient),
        order.phone,
        "",
        order.address,
        1,
        _item_text(item),
        "a",
        "신용",
        "빠른배송바랍니다",
        order.company,
        order.remarks,
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


def _input_log_path(output_dir: Path) -> Path:
    return output_dir / INPUT_LOG_NAME


def _read_input_log(path: Path) -> set:
    if not path.exists():
        return set()
    return {
        line.strip()
        for line in path.read_text(encoding="utf-8").splitlines()
        if line.strip()
    }


def _read_recorded_filenames(output_dir: Path) -> set:
    """Read the shared log and legacy date logs used by older versions."""
    filenames = _read_input_log(_input_log_path(output_dir))
    for path in output_dir.glob("*.txt"):
        if re.fullmatch(r"\d{4}-\d{2}-\d{2}\.txt", path.name):
            filenames.update(_read_input_log(path))
    return filenames


def _append_input_log(path: Path, filename: str) -> None:
    with path.open("a", encoding="utf-8") as log:
        log.write(filename + "\n")


def _prepare_database_sheet_schema(worksheet: Any) -> Any:
    from openpyxl.styles import Alignment, Font, PatternFill

    first_row = [worksheet.cell(1, column).value for column in range(1, len(DATABASE_HEADERS) + 1)]
    if not any(value is not None for value in first_row):
        for column, header in enumerate(DATABASE_HEADERS, 1):
            cell = worksheet.cell(1, column, header)
            cell.font = Font(bold=True)
            cell.fill = PatternFill("solid", fgColor="D9EAF7")
            cell.alignment = Alignment(horizontal="center", vertical="center")
        worksheet.freeze_panes = "A2"
    elif tuple(first_row[:len(LEGACY_DATABASE_HEADERS)]) == LEGACY_DATABASE_HEADERS:
        for row in range(2, worksheet.max_row + 1):
            legacy_values = [worksheet.cell(row, column).value
                             for column in range(1, len(LEGACY_DATABASE_HEADERS) + 1)]
            new_values = [
                legacy_values[0],
                legacy_values[1],
                "",
                legacy_values[2],
                1,
                legacy_values[3],
                "a",
                "신용",
                "빠른배송바랍니다",
                legacy_values[4],
                legacy_values[5],
            ]
            for column, value in enumerate(new_values, 1):
                worksheet.cell(row, column).value = value
        for column, header in enumerate(DATABASE_HEADERS, 1):
            worksheet.cell(1, column).value = header
        if worksheet.max_column > len(DATABASE_HEADERS):
            worksheet.delete_cols(len(DATABASE_HEADERS) + 1,
                                  worksheet.max_column - len(DATABASE_HEADERS))
    elif (first_row[-1] is None
          and tuple(first_row[:len(PREVIOUS_DATABASE_HEADERS)]) == PREVIOUS_DATABASE_HEADERS):
        for row in range(2, worksheet.max_row + 1):
            previous_note = worksheet.cell(row, 9).value
            worksheet.cell(row, len(DATABASE_HEADERS)).value = previous_note
            worksheet.cell(row, 9).value = "빠른배송바랍니다"
        worksheet.cell(1, len(DATABASE_HEADERS)).value = "비고"
    elif tuple(first_row) != DATABASE_HEADERS:
        raise ValueError("Existing sheet '{}' does not use the expected database columns".format(
            worksheet.title
        ))

    for row in range(2, worksheet.max_row + 1):
        worksheet.cell(row, 1).value = _display_recipient(worksheet.cell(row, 1).value)
        phone_cell = worksheet.cell(row, 2)
        if phone_cell.value:
            try:
                phone_cell.value = format_phone_number(str(phone_cell.value))
            except ValueError:
                pass

    widths = (18, 18, 12, 45, 10, 45, 12, 12, 24, 28, 30)
    for column, width in enumerate(widths, 1):
        worksheet.column_dimensions[chr(64 + column)].width = width
    worksheet.auto_filter.ref = "A1:K{}".format(max(worksheet.max_row, 1))
    return worksheet


def _prepare_database_sheet(workbook: Any) -> Any:
    if DATABASE_SHEET_NAME in workbook.sheetnames:
        worksheet = workbook[DATABASE_SHEET_NAME]
    elif workbook.sheetnames:
        worksheet = workbook.worksheets[0]
        worksheet.title = DATABASE_SHEET_NAME
    else:
        worksheet = workbook.create_sheet(DATABASE_SHEET_NAME)

    _prepare_database_sheet_schema(worksheet)
    for other in list(workbook.worksheets):
        if other is worksheet:
            continue
        _prepare_database_sheet_schema(other)
        for values in other.iter_rows(
            min_row=2, max_col=len(DATABASE_HEADERS), values_only=True
        ):
            if any(value is not None for value in values):
                worksheet.append(values)
        workbook.remove(other)
    _prepare_database_sheet_schema(worksheet)
    return worksheet


def append_order_to_workbook(workbook: Any, order: PurchaseOrder) -> int:
    """Append every item in an order and return the number of rows written."""
    worksheet = _prepare_database_sheet(workbook)
    rows_written = 0
    for item in order.items or []:
        next_row = max(worksheet.max_row + 1, 2)
        for column, value in enumerate(_database_row(order, item), 1):
            worksheet.cell(next_row, column).value = value
        rows_written += 1
    worksheet.auto_filter.ref = "A1:K{}".format(max(worksheet.max_row, 1))
    return rows_written


def _database_identity(order: PurchaseOrder, item: Item) -> tuple:
    return (
        _display_recipient(order.recipient),
        order.phone,
        order.address,
        _item_text(item),
        order.company,
    )


def _database_identities(worksheet: Any) -> set:
    return {
        (row[0], row[1], row[3], row[5], row[9])
        for row in worksheet.iter_rows(min_row=2, max_col=len(DATABASE_HEADERS), values_only=True)
        if any(value is not None for value in row)
    }


def _order_identities(order: PurchaseOrder) -> set:
    return {_database_identity(order, item) for item in order.items or []}


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


def _file_status_callback(callback: Optional[Callable[[FileProcessingStatus], None]],
                          path: Path, input_dir: Path, status: str,
                          confidence: Optional[float] = None,
                          message: str = "") -> None:
    if callback is not None:
        callback(FileProcessingStatus(
            str(path.relative_to(input_dir)), status, confidence, message
        ))


def process_documents(input_dir: Path, output_dir: Path, client: Optional[OllamaClient] = None,
                      chunk_size: int = DEFAULT_CHUNK_SIZE,
                      report: Optional[ProcessingReport] = None,
                      on_file_status: Optional[Callable[[FileProcessingStatus], None]] = None
                      ) -> int:
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
    missing_filenames = set()

    if xlsx_documents:
        database_path = output_dir / OUTPUT_WORKBOOK_NAME
        database_existed = database_path.exists()
        workbook = _new_database_workbook(database_path)
        worksheet = _prepare_database_sheet(workbook)
        existing_identities = _database_identities(worksheet)
        extracted_orders = []
        recorded_filenames = _read_recorded_filenames(output_dir)
        log_path = _input_log_path(output_dir)
        pending_log_entries = []
        for path in xlsx_documents:
            print("Reading {} -> {}".format(path, database_path))
            _file_status_callback(on_file_status, path, input_dir, "reading")
            try:
                order = extract_purchase_order(path)
                confidence = confidence_level(order)
                confidence_results.append(ConfidenceAnalysis(path.name, confidence))
                missing_fields = missing_required_fields(order)
                if missing_fields:
                    failures += 1
                    missing_filenames.add(path.name)
                    if report is not None:
                        report.missing_data.append(MissingData(
                            str(path.relative_to(input_dir)), missing_fields
                        ))
                    _file_status_callback(
                        on_file_status,
                        path,
                        input_dir,
                        "failed",
                        confidence,
                        "Missing required data: {}".format(", ".join(missing_fields)),
                    )
                    print("ERROR: {}: missing required data: {}".format(
                        path, ", ".join(missing_fields)
                    ), file=sys.stderr)
                    continue
                if any(character.isdigit() for character in order.recipient):
                    failures += 1
                    reason = "수령인에는 숫자를 사용할 수 없습니다"
                    if report is not None:
                        report.invalid_data.append(InvalidData(
                            str(path.relative_to(input_dir)), "수령인", reason
                        ))
                    _file_status_callback(
                        on_file_status, path, input_dir, "failed", confidence, reason
                    )
                    print("ERROR: {}: invalid 수령인: {}".format(path, reason), file=sys.stderr)
                    continue
                try:
                    order.phone = format_phone_number(order.phone)
                except ValueError as error:
                    failures += 1
                    if report is not None:
                        report.invalid_data.append(InvalidData(
                            str(path.relative_to(input_dir)), "전화", str(error)
                        ))
                    _file_status_callback(
                        on_file_status, path, input_dir, "failed", confidence, str(error)
                    )
                    print("ERROR: {}: invalid 전화: {}".format(path, error), file=sys.stderr)
                    continue
                _file_status_callback(
                    on_file_status,
                    path,
                    input_dir,
                    "low_confidence" if confidence < 90.0 else "success",
                    confidence,
                    ", ".join(confidence_reasons(order)) if confidence < 90.0 else "",
                )
                extracted_orders.append((path, order))
            except Exception as error:
                failures += 1
                confidence_results.append(ConfidenceAnalysis(path.name, 0.0))
                _file_status_callback(on_file_status, path, input_dir, "error", 0.0, str(error))
                print("ERROR: {}: {}".format(path, error), file=sys.stderr)

        for path, order in sorted(
            extracted_orders,
            key=lambda result: (_sheet_name(result[1]) if result[1].order_date else "99.99",
                                result[0].name),
        ):
            order_identities = _order_identities(order)
            if (
                path.name in recorded_filenames
                and order_identities
                and order_identities.issubset(existing_identities)
            ):
                if report is not None:
                    report.duplicates.append(str(path.relative_to(input_dir)))
                _file_status_callback(on_file_status, path, input_dir, "duplicate")
                print("DUPLICATE: {} is already recorded in {}; skipping.".format(
                    path.name, log_path.name
                ))
                continue
            try:
                append_order_to_workbook(workbook, order)
                existing_identities.update(order_identities)
                recorded_filenames.add(path.name)
                pending_log_entries.append(path.name)
            except Exception as error:
                failures += 1
                _file_status_callback(on_file_status, path, input_dir, "error", message=str(error))
                print("ERROR: {}: {}".format(path, error), file=sys.stderr)
        workbook_saved = False
        try:
            if workbook.sheetnames or database_existed:
                workbook.save(str(database_path))
                workbook_saved = True
        finally:
            workbook.close()
        if workbook_saved:
            for filename in pending_log_entries:
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
        _file_status_callback(on_file_status, path, input_dir, "reading")
        try:
            summary = summarize_content(read_document(path), client, chunk_size)
            destination.write_text(summary.rstrip() + "\n", encoding="utf-8")
            _file_status_callback(on_file_status, path, input_dir, "success")
        except Exception as error:
            failures += 1
            _file_status_callback(on_file_status, path, input_dir, "error", message=str(error))
            print("ERROR: {}: {}".format(path, error), file=sys.stderr)

    for result in confidence_results:
        print("Confidence: {} - {:.2f}%".format(result.filename, result.confidence))
    for result in confidence_results:
        if result.confidence < 90.0 and result.filename not in missing_filenames:
            print("HUMAN REVIEW REQUIRED: {} - {:.2f}%".format(
                result.filename, result.confidence
            ))
    if report is not None:
        report.low_confidence.extend(
            result for result in confidence_results
            if result.confidence < 90.0 and result.filename not in missing_filenames
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


def is_windowed_application() -> bool:
    """Use the live status window only for the packaged Windows application."""
    return bool(getattr(sys, "frozen", False) and os.name == "nt")


class ProcessingWindow:
    """Small Tk window that renders file processing updates without blocking the scan."""

    def __init__(self, input_dir: Path, documents: Sequence[Path]) -> None:
        import tkinter as tk
        from tkinter import messagebox, ttk

        self._tk = tk
        self._ttk = ttk
        self._messagebox = messagebox
        self._status_queue = queue.Queue()
        self._done = threading.Event()
        self._failures = 0
        self._error: Optional[Exception] = None
        self._on_complete = None
        self._rows = {}
        self._row_status = {}
        self._row_confidence = {}

        self.root = tk.Tk()
        self.root.title("Inv Reader")
        self.root.geometry("760x520")
        self.root.minsize(560, 320)
        self.root.configure(bg="#f4f1eb")
        self.root.protocol("WM_DELETE_WINDOW", self._close)

        style = ttk.Style(self.root)
        try:
            style.theme_use("vista")
        except tk.TclError:
            pass
        style.configure("Status.TLabel", background="#f4f1eb", foreground="#263238")
        style.configure("Status.Header.TLabel", background="#f4f1eb", foreground="#6d746e",
                        font=("Segoe UI", 9, "bold"))
        style.configure("Status.Title.TLabel", background="#f4f1eb", foreground="#16251f",
                        font=("Segoe UI", 18, "bold"))
        style.configure("Status.Close.TButton", padding=(16, 7))

        outer = ttk.Frame(self.root, padding=22)
        outer.pack(fill="both", expand=True)
        ttk.Label(outer, text="Inv Reader", style="Status.Title.TLabel").pack(anchor="w")
        self.summary = tk.StringVar(value="Reading input folder...")
        ttk.Label(outer, textvariable=self.summary, style="Status.TLabel").pack(
            anchor="w", pady=(4, 14)
        )

        table_frame = ttk.Frame(outer)
        table_frame.pack(fill="both", expand=True)
        canvas = tk.Canvas(table_frame, background="#f4f1eb", highlightthickness=0)
        rows_frame = ttk.Frame(canvas)
        canvas_window = canvas.create_window((0, 0), window=rows_frame, anchor="nw")
        rows_frame.bind(
            "<Configure>",
            lambda event: canvas.configure(scrollregion=canvas.bbox("all")),
        )
        canvas.bind(
            "<Configure>",
            lambda event: canvas.itemconfigure(canvas_window, width=event.width),
        )
        scrollbar = ttk.Scrollbar(table_frame, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        rows_frame.columnconfigure(1, weight=1)
        ttk.Label(rows_frame, text="", style="Status.Header.TLabel").grid(
            row=0, column=0, padx=(8, 10), pady=(4, 8), sticky="w"
        )
        ttk.Label(rows_frame, text="File", style="Status.Header.TLabel").grid(
            row=0, column=1, padx=(0, 12), pady=(4, 8), sticky="w"
        )
        ttk.Label(rows_frame, text="Confidence", style="Status.Header.TLabel").grid(
            row=0, column=2, padx=(0, 16), pady=(4, 8), sticky="e"
        )
        ttk.Label(rows_frame, text="Result", style="Status.Header.TLabel").grid(
            row=0, column=3, padx=(0, 8), pady=(4, 8), sticky="e"
        )

        for index, path in enumerate(documents):
            filename = str(path.relative_to(input_dir))
            row_frame = ttk.Frame(rows_frame)
            row_frame.grid(row=index + 1, column=0, columnspan=4, sticky="ew")
            row_frame.columnconfigure(1, weight=1)
            indicator = ttk.Label(row_frame, text="...", width=4, anchor="center")
            indicator.grid(row=0, column=0, padx=(8, 10), pady=5, sticky="w")
            filename_label = ttk.Label(row_frame, text=filename, style="Status.TLabel")
            filename_label.grid(row=0, column=1, padx=(0, 12), pady=5, sticky="w")
            confidence_label = ttk.Label(row_frame, text="—", style="Status.TLabel")
            confidence_label.grid(row=0, column=2, padx=(0, 16), pady=5, sticky="e")
            result_frame = ttk.Frame(row_frame)
            result_frame.grid(row=0, column=3, padx=(0, 8), pady=5, sticky="e")
            self._rows[filename] = (
                indicator, filename_label, confidence_label, result_frame
            )
            self._row_status[filename] = "waiting"
            self._row_confidence[filename] = None
            ttk.Label(result_frame, text="Waiting", style="Status.TLabel").pack(
                anchor="e"
            )
        self._total = len(self._rows)

        self.close_button = ttk.Button(
            outer, text="Close", command=self._close, style="Status.Close.TButton", state="disabled"
        )
        self.close_button.pack(anchor="e", pady=(14, 0))

    def _close(self) -> None:
        if self._done.is_set():
            self.root.destroy()

    def _post_status(self, status: FileProcessingStatus) -> None:
        self._status_queue.put(status)

    def _render_status(self, status: FileProcessingStatus) -> None:
        row_widgets = self._rows.get(status.filename)
        if row_widgets is None:
            return
        indicator_label, filename_label, confidence_label, result_frame = row_widgets
        if status.confidence is not None:
            self._row_confidence[status.filename] = status.confidence
        if status.status == "success":
            indicator, details = "✓", ""
        elif status.status == "low_confidence":
            indicator = "✕"
            details = "{:.2f}%".format(status.confidence or 0.0)
        elif status.status == "duplicate":
            indicator, details = "✕", "duplicate"
        elif status.status == "failed":
            indicator, details = "✕", "failed"
        elif status.status == "error":
            indicator, details = "✕", "error"
        elif status.status == "reading":
            indicator, details = "...", "Reading"
        else:
            indicator, details = "...", "Waiting"
        colors = {
            "success": "#207a45",
            "low_confidence": "#bd302b",
            "duplicate": "#bd302b",
            "failed": "#bd302b",
            "error": "#bd302b",
            "reading": "#7b817d",
            "waiting": "#7b817d",
        }
        color = colors.get(status.status, colors["waiting"])
        indicator_label.configure(text=indicator, foreground=color)
        filename_label.configure(foreground=color)
        confidence = self._row_confidence[status.filename]
        confidence_label.configure(
            text="—" if confidence is None else "{:.2f}%".format(confidence),
            foreground=color,
        )
        for child in result_frame.winfo_children():
            child.destroy()
        if status.status == "low_confidence":
            self._ttk.Button(
                result_frame,
                text=details,
                command=lambda: self._show_confidence_details(status),
            ).pack(anchor="e")
        else:
            label = self._tk.Label(
                result_frame,
                text=details,
                background="#f4f1eb",
                foreground=color,
                font=("Segoe UI", 10),
            )
            label.pack(anchor="e")
        self._row_status[status.filename] = status.status
        completed = sum(
            current_status in ("success", "low_confidence", "duplicate", "failed", "error")
            for current_status in self._row_status.values()
        )
        self.summary.set("Processed {} of {} files".format(completed, self._total))

    def _show_confidence_details(self, status: FileProcessingStatus) -> None:
        reasons = status.message or "The extracted fields were incomplete."
        reason_lines = "\n".join("- {}".format(reason) for reason in reasons.split(", "))
        self._messagebox.showinfo(
            "Confidence details",
            "File: {}\nConfidence: {:.2f}%\n\nLowered by:\n{}".format(
                status.filename, status.confidence or 0.0, reason_lines
            ),
            parent=self.root,
        )

    def _poll(self) -> None:
        while True:
            try:
                status = self._status_queue.get_nowait()
            except queue.Empty:
                break
            self._render_status(status)

        if self._done.is_set():
            if self._error is not None:
                self.summary.set("Scan stopped: {}".format(self._error))
            elif self._failures:
                self.summary.set("Scan complete with {} error(s)".format(self._failures))
            else:
                self.summary.set("Scan complete. Review the file results above.")
            self.close_button.configure(state="normal")
            if self._on_complete is not None:
                callback = self._on_complete
                self._on_complete = None
                callback(self._error)
            return
        self.root.after(50, self._poll)

    def run(self, processor: Callable[[Callable[[FileProcessingStatus], None]], int],
            on_complete: Optional[Callable[[Optional[Exception]], None]] = None):
        self._on_complete = on_complete

        def worker() -> None:
            try:
                self._failures = processor(self._post_status)
            except Exception as error:
                self._error = error
            finally:
                self._done.set()

        self.root.after(50, self._poll)
        threading.Thread(target=worker, name="document-processing", daemon=True).start()
        self.root.mainloop()
        return self._failures, self._error


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

    def process_batch(on_file_status: Optional[Callable[[FileProcessingStatus], None]] = None) -> int:
        with runtime:
            return process_documents(
                input_dir,
                output_dir,
                OllamaClient(model, ollama_url, args.timeout),
                args.chunk_size,
                report,
                on_file_status,
            )

    def show_completed_popups(processing_error: Optional[Exception]) -> None:
        if processing_error is None:
            show_issue_popups(report)

    try:
        if needs_ollama and getattr(sys, "frozen", False) and not executable:
            raise FileNotFoundError(
                "Bundled ollama.exe was not found beside the application. "
                "Copy the complete portable application directory."
            )
        if is_windowed_application():
            window = ProcessingWindow(input_dir, documents)
            failures, error = window.run(process_batch, show_completed_popups)
            if error is not None:
                raise error
        else:
            failures = process_batch()
            show_issue_popups(report)
    except Exception as error:
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
        if report.missing_data:
            missing_list = "\n".join(
                "- {}: {}".format(result.filename, ", ".join(result.fields))
                for result in report.missing_data
            )
            ctypes.windll.user32.MessageBoxW(
                None,
                "These files were skipped because required data was empty:\n\n{}".format(
                    missing_list
                ),
                "Missing required data",
                0x10,
            )
        if report.invalid_data:
            invalid_list = "\n".join(
                "- {}: {} - {}".format(result.filename, result.field, result.reason)
                for result in report.invalid_data
            )
            ctypes.windll.user32.MessageBoxW(
                None,
                "These files were skipped because input data was invalid:\n\n{}".format(
                    invalid_list
                ),
                "Invalid input data",
                0x10,
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

# Inv_Reader

Summarize documents offline with the local Ollama `qwen3-vl:8b` model. The
summarizer can read documents containing English, Korean, or both languages.

## Setup

Install Ollama and pull the model once:

```bash
ollama pull qwen3-vl:8b
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

Put documents in `input/`. The program supports text files, Markdown, CSV, JSON, HTML, XML, DOCX, XLSX, PDF, and common image formats. XLSX order forms are read directly and appended to `output/2026 한진양식.xlsx` in one `Orders` worksheet. Recognized date formats are normalized before use, including `YYYY/M/D`, `YY/M/D` such as `26/9/3`, and common year-last forms. The worksheet uses the columns 이름, 전화, 우편번호, 주소, 수량, 품목, 운임타입, 지불조건, 특기사항, and 업체명. 이름 receives the `님` suffix; 전화 must use the 010 or 02~07 prefix and is formatted with hyphens; every 품목 quantity uses the unit `개` and is joined to the item description with `-`; 우편번호 is blank; 수량 is `1`; 운임타입 is `a`; 지불조건 is `신용`; and 특기사항 copies the input `특기사항 • 배송 메모` or `비고` value, remaining blank when the input is blank. Other supported document types continue to use the Ollama summarizer. PDFs with selectable text use layout-aware extraction; scanned PDFs are rendered as images for Qwen3-VL. Text requests use a bounded, non-thinking API path because the `qwen3-vl:8b` Ollama template can otherwise consume the entire response on internal reasoning.

Start Ollama, then run:

```bash
python3 -m src.summarize_documents
```

XLSX files are appended to the shared workbook. `output/processed_files.txt` records input filenames already processed. A file is reported as `DUPLICATE` and skipped only when its filename is logged and all five identity fields (이름, 전화, 주소, 품목, 업체명) match an existing row. Files missing required import fields (수령인, 수령인 연락처, 배송지 주소, 발주처, or item data) are skipped, marked failed in the GUI, and listed with the empty columns in a popup. The program prints a confidence percentage for every XLSX file and marks results below 90% with `HUMAN REVIEW REQUIRED`. Non-XLSX documents produce one UTF-8 text file in `output/`; input directories are searched recursively and nested paths are flattened with `__` in the output filename.

Useful options:

```bash
python3 -m src.summarize_documents --input input --output output \
  --model qwen3-vl:8b --ollama-url http://localhost:11434
```

Run tests without Ollama:

```bash
python3 -m unittest discover -s tests
```

## Portable Windows Application

This creates a portable Windows application. The build must be performed on a
Windows computer because PyInstaller and Ollama use platform-specific binaries. A
Mac or Linux build does not produce a Windows EXE.

### 1. Build the Application

Use a Windows build computer. This can be the same computer that will eventually run
the application, or a separate Windows computer.

1. Copy the entire repository to the Windows build computer. It must include
   `src`, `tests`, `requirements.txt`, and `build_windows.ps1`. The Mac `venv` folder
   is not needed.
2. Install 64-bit Python for Windows from
   [python.org](https://www.python.org/downloads/windows/). Enable **Add Python to
   PATH** during installation.
3. Open a new PowerShell window and verify Python:

   ```powershell
   python --version
   ```

4. Install Ollama for Windows from [ollama.com](https://ollama.com/download/windows).
5. Open a new PowerShell window and verify Ollama:

   ```powershell
   ollama --version
   ```

6. Open the repository folder in File Explorer. Right-click inside the folder and
   choose **Open in Terminal** or **Open PowerShell window here**.
7. Run the build script:

   ```powershell
   Set-ExecutionPolicy -Scope Process Bypass
   .\build_windows.ps1
   ```

   The script installs the Python dependencies, downloads `qwen3-vl:8b`, starts a
   temporary Ollama server for the download, builds the PyInstaller application, and
   copies Ollama and the model into the portable package. Internet access is required
   on the build computer.

8. To build with a different Ollama model, pass the model name to the script:

   ```powershell
   .\build_windows.ps1 -Model "qwen3-vl:8b"
   ```

### 2. Check the Build Output

When the script finishes, the portable application is in:

```text
dist\Inv_Reader\
```

The folder should contain all of the following:

```text
Inv_Reader\
  Inv_Reader.exe
  ollama\
  models\
  input\
  output\
  portable_config.json
  _internal\
```

Do not distribute only `Inv_Reader.exe`. The `ollama`, `models`, and `_internal`
folders are required. The `models` folder may be several gigabytes.

Before distributing the application, test it on the build computer:

1. Put a supported document into `dist\Inv_Reader\input\`.
2. Double-click `dist\Inv_Reader\Inv_Reader.exe`.
3. Check that the result appears in `dist\Inv_Reader\output\`.
4. Check `dist\Inv_Reader\run.log` if anything does not work.

### 3. Create the Distribution ZIP

From the repository folder, run:

```powershell
Compress-Archive `
  -Path .\dist\Inv_Reader `
  -DestinationPath .\Inv_Reader_Windows.zip
```

Give the other computer the complete `Inv_Reader_Windows.zip`. Do not give it only
the EXE or only the build script.

### 4. Run on Another Windows Computer

1. Extract `Inv_Reader_Windows.zip` completely to a writable location, such as
   Documents or the Desktop. Do not extract it into `C:\Program Files`, where the
   application may not be allowed to write its folders and log.
2. Confirm that `Inv_Reader.exe`, `ollama`, `models`, and `_internal` are all still
   present in the extracted folder.
3. Put documents to process into the `input` folder.
4. Double-click `Inv_Reader.exe`.

The application creates `input`, `output`, and `models` beside the EXE when needed.
It starts the bundled Ollama server invisibly on a private local port, processes all
supported files in `input`, writes results to `output`, and stops its Ollama server
when it exits.

The target computer does not need Python, Ollama installed separately, a terminal,
or internet access when the `models` folder was bundled successfully. Diagnostic
output from the EXE and Ollama is written to `run.log`.

The portable EXE opens a live status window listing every input file as it is read.
Successfully read files show a green check mark; low-confidence XLSX files show a
red X and confidence percentage; duplicate files show a red X and `duplicate`.
After scanning, the existing popups still list duplicate and low-confidence files.
These issues are also recorded in `run.log`.

### 5. Portable Application Limitations

This package is intended for compatible 64-bit Windows computers. It does not
directly run on macOS or Linux. The target computer needs enough disk space and
memory for the bundled `qwen3-vl:8b` model. CPU-only computers can run the model but
may process documents more slowly. Keep the entire `Inv_Reader` folder together when
moving or copying the application.

For development or an existing Ollama server, the original command still works:

```bash
python3 -m src.summarize_documents --no-ollama-management
```

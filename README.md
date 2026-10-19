# pdf2docx-ocr (100% Offline Scanned PDF to Rich DOCX Converter)

A fast, **100% local and offline** CLI tool that converts scanned/image-based PDFs and images into **editable Microsoft Word (`.docx`) documents** while preserving:

- **Editable Text & Spatial Layout:** Proper vertical order, indentation, centered headings, and Table of Contents dot-leaders (`......`).
- **Font Color, Size & Bold Weight:** Samples the true RGB ink color of every word directly from the page, calculates font size in points (`pt`), and detects bold text.
- **Editable Boxes & Table of Contents Frames:** Detects rectangular boxed sections (like a Table of Contents box or callout panel) and builds **real bordered & shaded Word boxes** containing editable text—not a flat image.
- **Non-Text Figures & Photos:** Automatically crops diagrams, charts, photos, and illustrations that are not text and inserts them as images in their exact position.
- **10 Output Formats:** `docx` (Rich Layout), `docx-plain`, `docx-image`, `txt`, `md`, `html`, `pdf` (Searchable PDF), `json`, `png`, and `jpg`.

---

## Does It Use the Internet?

**No.** Once installed, `pdf2docx_ocr.py` runs completely offline on your computer. No files are uploaded anywhere, and no API keys are required.

---

## Installation Instructions

### Step 1: Install Python Packages

Open your terminal (or Command Prompt / PowerShell on Windows) and install the required libraries:

```bash
pip install pillow numpy opencv-python pymupdf python-docx pytesseract pypdf pdf2image
```

*(Note: `pymupdf` lets the script read PDF files directly without needing to install Poppler!)*

### Step 2: Install Tesseract OCR Engine (Offline Engine)

- **Windows:**
  1. Download the official installer from [UB Mannheim Tesseract Releases](https://github.com/UB-Mannheim/tesseract/wiki).
  2. Run the installer and note the installation path (usually `C:\Program Files\Tesseract-OCR\tesseract.exe`).
  3. Add `C:\Program Files\Tesseract-OCR` to your Windows `PATH`, or pass `--tesseract-cmd "C:\Program Files\Tesseract-OCR\tesseract.exe"` when running the script.
- **macOS (Homebrew):**
  ```bash
  brew install tesseract
  ```
- **Linux (Ubuntu / Debian):**
  ```bash
  sudo apt update && sudo apt install tesseract-ocr
  ```

### Step 3: Verify Your Setup

Run the built-in health check to verify everything is detected:

```bash
python pdf2docx_ocr.py --check
```

---

## Usage & Examples

### 1. Convert a Scanned PDF to a Rich Editable DOCX (Default)
Preserves font colors, font sizes, editable bordered boxes (such as Table of Contents boxes), and cropped non-text images:

```bash
python pdf2docx_ocr.py scan.pdf
```

### 2. Specify a Custom Font Family (e.g., Times New Roman or Arial)
```bash
python pdf2docx_ocr.py scan.pdf --font-name "Times New Roman"
```

### 3. Convert to Multiple Formats at Once
Generate a rich `.docx`, a styled `.html`, a Markdown `.md`, and a searchable `.pdf` in a single run:

```bash
python pdf2docx_ocr.py scan.pdf -f docx,html,md,pdf
```

### 4. Convert Specific Pages Only
```bash
python pdf2docx_ocr.py scan.pdf --pages 1-5,8,10-12
```

### 5. Improve OCR on Noisy Scans (While Keeping Original Font Colors!)
Even when you enable `--grayscale` or `--threshold` to help Tesseract read faint text, the script still samples the **original RGB colors** from the un-grayscaled page:

```bash
python pdf2docx_ocr.py scan.pdf -f docx --grayscale --autocontrast
```

### 6. Batch Convert an Entire Folder
```bash
python pdf2docx_ocr.py ./my_scans --batch -f docx -d ./converted_output
```

### 7. Interactive Menu Mode (No Flags Needed)
Simply run the script without arguments to get an interactive step-by-step prompt:

```bash
python pdf2docx_ocr.py
```

---

## All Output Formats (`-f` / `--format`)

| Format | Description |
| :--- | :--- |
| `docx` | **Rich Layout Word document** with editable text, sampled RGB font colors, font sizes, bordered boxes/TOC, and cropped non-text images |
| `docx-plain` | Simple Word document containing only plain unformatted paragraphs |
| `docx-image` | Word document with each full page embedded as a picture |
| `txt` | Plain UTF-8 text file |
| `md` | Markdown file with detected headings and blockquoted boxes |
| `html` | Styled HTML file preserving font colors, bordered boxes, and embedded images |
| `pdf` | Searchable PDF (original page image + invisible selectable text layer) |
| `json` | Structured JSON containing every word, bounding box, font size, and RGB color |
| `png` / `jpg` | Exports each PDF page as a high-resolution image |

---

## CLI Options Reference

| Flag | Description |
| :--- | :--- |
| `-f, --format` | Comma-separated formats, e.g., `docx,txt,html,pdf` (default: `docx`) |
| `-o, --output` | Custom output file name |
| `-d, --outdir` | Output directory path |
| `--pages` | Page range to process, e.g., `1-3,5,8-` |
| `--split` | Save one output file per page (`scan_p001.docx`, `scan_p002.docx`, ...) |
| `--batch` | Convert all PDFs and images inside a folder |
| `-i, --interactive` | Launch the interactive terminal menu |
| `--font-name` | Font family to use in Word/HTML output (default: `Calibri`) |
| `--no-colors` | Force all output text to black instead of sampling RGB colors |
| `--no-boxes` | Disable bordered box / Table of Contents box detection |
| `--no-images` | Disable cropping and embedding non-text figures/diagrams |
| `--dpi` | Rendering resolution for PDF pages (default: `300`) |
| `--lang` | Tesseract OCR language code, e.g., `eng` or `eng+hin` |
| `--grayscale` | Preprocess image as grayscale for OCR |
| `--autocontrast` | Maximize image contrast before OCR |
| `--sharpen` | Apply sharpen filter before OCR |
| `--threshold` | Binarize image at threshold `0-255` (e.g., `150`) |
| `--auto-rotate` | Automatically fix upside-down or sideways pages |
| `--tesseract-cmd` | Full path to `tesseract` binary (useful on Windows) |
| `--check` | Verify installed Python packages and system tools |

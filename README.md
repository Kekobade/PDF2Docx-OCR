# pdf2docx-ocr

This folder has **two tools**:

| Tool | Use it when |
|------|-------------|
| `layout_ocr_to_docx.py` | You want the Word file to **look like the original page**: same positions, text colours, pictures kept as pictures, boxes and tables redrawn (see "Layout mode" below) |
| `pdf2docx_ocr.py` | You just want the **text** (or other formats like TXT, Markdown, HTML, searchable PDF, PNG/JPG) |


A command-line tool that converts **image-based PDFs** (scans, screenshots, photos saved as PDF) and **image files** into editable documents or cleaned-up images.

**It runs 100% offline.** All processing happens on your computer using Tesseract OCR and Poppler. Your files are never uploaded anywhere. Internet is only needed once, to download and install the tools.

## Layout mode: `layout_ocr_to_docx.py`

Rebuilds each page of an image-only PDF (or photos of pages) as a Word page where:

- **Text is real text.** Every block is a Word text box at its original position. You can click into it, select, copy and edit.
- **Colour and size are matched.** Font size, text colour and bold are estimated from the pixels.
- **Pictures stay pictures.** Photos, charts and logos are cut out and inserted as images, not turned into garbage text.
- **Boxes and lines are redrawn.** Box outlines, table borders and rules become Word shapes. Coloured bars become filled shapes, and a coloured page background is kept.
- **Right-aligned parts keep their place.** Page numbers at the right of a table of contents sit where they were; lone numbers and labels that normal OCR skips are found by an extra pass.

Everything runs offline.

### Quick start

```bash
python layout_ocr_to_docx.py scan.pdf
python layout_ocr_to_docx.py scan.pdf -o result.docx --font "Times New Roman"
python layout_ocr_to_docx.py scan.pdf --pages 1-3 --debug
python layout_ocr_to_docx.py photo_of_page.jpg -o result.docx
```

Output defaults to `<input>_layout.docx`. One input file per run (a PDF can have many pages).

### What it can and cannot do

| Works well | Limits |
|------------|--------|
| Positions of text, pictures, boxes, lines | **The exact font cannot be identified** from an image. One font is used for the whole document: set it with `--font` to the closest match (for example `Times New Roman`, `Calibri`, `Georgia`) |
| Text colour, relative text size | Sizes are estimates; use `--font-scale 1.05` etc. to adjust all of them |
| Photos and diagrams as pictures | Text inside pictures stays part of the picture (use `--keep-image-text` to also read it) |
| Box outlines and table grids as shapes | Tables are drawn as lines and separate text, **not** as a real Word table you can add rows to |
| Bold, italic, and underline detection | Bold is detected via relative stroke weight and tuned with `--bold-factor` (or `--no-bold`). Italic is detected via projection profile shear deskewing (or `--no-italic`). Underlines are matched from line shapes directly to text runs (or `--no-underline`) |
| Paragraph and text alignment | Left, center, right, and justified alignment are automatically detected and emitted as Word alignment properties (or `--no-align`) |
| Coloured header bars and page backgrounds | Gradients, shadows, rounded corners and dashed lines are not reproduced |

Because the page is rebuilt from fixed positions, it is a **look-alike you can copy text from**, not a flowing document that reflows when you edit it. If your PDF contains real selectable text (not only images), `pip install pdf2docx` gives much better results for it.

The file is tested with LibreOffice. Open it in Word 2013 or newer. If Word reports a problem, tell me the message.

### Options

| Option | Description | Default |
|--------|-------------|---------|
| `-o, --output` | Output `.docx` | `<input>_layout.docx` |
| `--pages` | PDF pages, e.g. `1-3,5` | all |
| `--dpi` | Render resolution for PDFs | `300` |
| `--image-page-width` | Assumed page width in inches for photo inputs | `8.27` (A4) |
| `--lang` | OCR language(s), e.g. `eng+hin` | `eng` |
| `--psm` | Tesseract page segmentation mode | Tesseract default |
| `--min-conf` | Drop OCR words below this confidence | `40` |
| `--font` | Font for all text | `Arial` |
| `--font-scale` | Multiply all font sizes | `1.0` |
| `--no-bold` | Turn off bold detection | off |
| `--bold-factor` | Higher means less bold | `1.2` |
| `--no-italic` | Turn off italic slant detection | off |
| `--no-underline` | Turn off underline detection (retains separate line shapes) | off |
| `--no-align` | Turn off paragraph alignment detection (keep all left-aligned) | off |
| `--y-offset` | Move all text down by N points (negative = up) | `0` |
| `--box-per-line` | One text box per line: most exact positions, but each line is separate to copy | off |
| `--page-batch` | Number of PDF pages to render per Poppler batch | `5` |
| `--min-image-pct` | Smallest picture to keep, as % of page area | `0.3` |
| `--bg-tol` | How different from paper colour counts as content | `45` |
| `--no-images` | Do not extract pictures | off |
| `--no-lines` | Do not redraw boxes and lines | off |
| `--keep-image-text` | Also keep OCR text found inside pictures | off |
| `--no-sparse-pass` | Skip the extra OCR pass for lone numbers/labels (faster, may miss page numbers) | off |
| `--debug` | Save a PNG per page showing what was detected | off |
| `--poppler-path`, `--tesseract-cmd` | Tool locations on Windows | |

### Tuning the result

Run with `--debug` first. It saves `<output>_p001_debug.png` per page with coloured rectangles: **red** = pictures, **blue** = filled areas, **green** = lines, **orange** = text boxes.

| Problem | Try |
|---------|-----|
| A picture is missing or cut into pieces | Lower `--min-image-pct` (e.g. `0.1`) |
| Text became a picture, or noise became a picture | Raise `--min-image-pct`, or use `--no-images` |
| Whole page treated as content (scan with grey paper) | Raise `--bg-tol` (e.g. `70`) |
| Text too big or too small | `--font-scale 0.9` or `1.1` |
| Text sits slightly too high or low | `--y-offset 2` or `-2` |
| Text wraps or is cut off at box edges | Pick a narrower font with `--font`, or use `--box-per-line` |
| Wrong words in places | Check OCR quality: higher `--dpi`, right `--lang`, cleaner scan |
| Bold on the wrong lines | `--bold-factor 1.5` (less) or `1.1` (more), or `--no-bold` |
| Lines or box borders missing | They may be dashed or very faint; lower `--bg-tol` (e.g. `30`) |

---

## Installation (needed for both tools)

See section 1 below: install Tesseract, Poppler, then `pip install -r requirements.txt`.

---

## Text and format conversion: `pdf2docx_ocr.py`

## Output formats

| Format | What you get | Needs OCR? |
|--------|--------------|:---------:|
| `docx` | Word document with editable text | yes |
| `docx-image` | Word document with each page as a picture | no |
| `txt` | Plain text | yes |
| `md` | Markdown text | yes |
| `html` | Simple HTML page with the text | yes |
| `pdf` | Searchable PDF (page images plus a hidden text layer) | yes |
| `png` | Every page saved as a PNG image | no |
| `jpg` | Every page saved as a JPG image | no |

You can request several at once: `-f docx,txt,pdf`.

Formats that don't need OCR (`docx-image`, `png`, `jpg`) work even without Tesseract installed.

---

## 1. Installation

### Step 1: Tesseract OCR

**Windows:** download the installer from the UB Mannheim Tesseract builds (search "Tesseract UB Mannheim"). Tick the extra languages you need during install. Default path: `C:\Program Files\Tesseract-OCR`.

**macOS:**
```bash
brew install tesseract
```

**Linux (Debian/Ubuntu):**
```bash
sudo apt update && sudo apt install tesseract-ocr
```

### Step 2: Poppler (turns PDF pages into images)

**Windows:** download a Poppler for Windows build, extract it (for example to `C:\poppler`), then either add `C:\poppler\Library\bin` to PATH or use `--poppler-path` each time.

**macOS:**
```bash
brew install poppler
```

**Linux:**
```bash
sudo apt install poppler-utils
```

### Step 3: Python packages

Python 3.8+ is required. A virtual environment is recommended:

```bash
python -m venv venv

# Windows
venv\Scripts\activate
# macOS / Linux
source venv/bin/activate

pip install -r requirements.txt
```

### Step 4: Verify

```bash
python pdf2docx_ocr.py --check
```

This lists which Python packages, Tesseract languages and Poppler were found, and tells you what is missing.

---

## 2. Three ways to run it

### A) Interactive menu (easiest)

Run it with no arguments, or with `-i`:

```bash
python pdf2docx_ocr.py
```

It asks for the input file or folder, output formats, pages, language, DPI, image cleanup and output folder, then converts.

### B) One-line commands

```bash
python pdf2docx_ocr.py scan.pdf
```
Converts to `scan.docx` (editable text).

### C) Batch a whole folder

```bash
python pdf2docx_ocr.py ./my_scans --batch -f docx
```
Converts every PDF and image in the folder, one output set per file.

---

## 3. Recipes

```bash
# Editable Word + plain text + Markdown in one go
python pdf2docx_ocr.py scan.pdf -f docx,txt,md

# Only pages 1 to 5 and page 9
python pdf2docx_ocr.py scan.pdf --pages 1-5,9

# From page 3 to the end
python pdf2docx_ocr.py scan.pdf --pages 3-

# Make a scanned PDF searchable (you can Ctrl+F it)
python pdf2docx_ocr.py scan.pdf -f pdf

# Keep pages looking identical inside a Word file
python pdf2docx_ocr.py scan.pdf -f docx-image

# Export every PDF page as a PNG
python pdf2docx_ocr.py scan.pdf -f png

# One Word file per page
python pdf2docx_ocr.py scan.pdf -f docx --split

# Combine several photos into one document, in the order given
python pdf2docx_ocr.py p1.jpg p2.jpg p3.jpg -f docx -o combined

# Hindi + English text, higher resolution
python pdf2docx_ocr.py scan.pdf --lang eng+hin --dpi 400

# Poor scan: clean it up first
python pdf2docx_ocr.py scan.pdf --grayscale --autocontrast --sharpen

# Very faint or shadowed scan: force black and white
python pdf2docx_ocr.py scan.pdf --threshold 150

# Sideways pages
python pdf2docx_ocr.py scan.pdf --rotate 90
python pdf2docx_ocr.py scan.pdf --auto-rotate

# Put results in a specific folder
python pdf2docx_ocr.py scan.pdf -f docx,pdf -d ./converted
```

### Windows with tools not on PATH

```bat
python pdf2docx_ocr.py scan.pdf ^
  --tesseract-cmd "C:\Program Files\Tesseract-OCR\tesseract.exe" ^
  --poppler-path "C:\poppler\Library\bin"
```

---

## 4. All options

| Option | Description | Default |
|--------|-------------|---------|
| `INPUT ...` | PDF/image files (or a folder with `--batch`) | interactive if empty |
| `-f, --format` | Output format(s), comma-separated | `docx` |
| `-o, --output` | Output name or base path (extension optional) | input name |
| `-d, --outdir` | Folder to write outputs into | next to input |
| `--pages` | Pages to convert for PDFs: `1-3,5,8-` | all |
| `--split` | One output file per page | off |
| `--batch` | Input is a folder, convert each file separately | off |
| `-i, --interactive` | Menu-driven mode | off |
| `--lang` | Tesseract language(s), e.g. `eng+hin` | `eng` |
| `--psm` | Tesseract page segmentation mode (0 to 13) | Tesseract default |
| `--keep-linebreaks` | Keep original line breaks instead of joining lines | off |
| `--list-langs` | Show installed OCR languages | |
| `--dpi` | Render resolution for PDFs | `300` |
| `--grayscale` | Convert to grayscale | off |
| `--autocontrast` | Stretch contrast | off |
| `--sharpen` | Sharpen the image | off |
| `--threshold N` | Pure black/white at level N (0 to 255, try 150) | off |
| `--rotate N` | Rotate clockwise by N degrees | `0` |
| `--auto-rotate` | Detect orientation with Tesseract (best effort) | off |
| `--page-batch` | Number of PDF pages to render per Poppler batch | `5` |
| `--check` | Verify your installation | |
| `--poppler-path` | Folder containing Poppler binaries (Windows) | |
| `--tesseract-cmd` | Full path to the Tesseract executable | |

**Output names:** with several formats or `--split`, files are named like `scan.docx`, `scan_images.docx`, `scan_p001.docx`, `scan_p001.png`. Existing files with the same name are overwritten.

**Page selection** applies to PDF inputs. Image files are always converted whole.

---

## 5. Tips for better results

- **DPI:** 300 is a good default. Use 400 for small print. Beyond that it is slow and rarely helps.
- **Bad scans:** try `--grayscale --autocontrast`, then `--sharpen`, then `--threshold 150` (adjust between 120 and 200).
- **Language:** use the right code. Combine with `+`, such as `eng+hin`. Run `--list-langs` to see what's installed.
- **Layouts:** for single columns of text, `--psm 6` can help. For sparse text or labels, try `--psm 11`.
- **Handwriting:** Tesseract reads it poorly. Use `docx-image` or `png` to keep the pages as pictures.
- **Always proofread** OCR output, especially numbers and names.

### Installing extra OCR languages

- **Windows:** re-run the Tesseract installer and select more languages.
- **macOS:** `brew install tesseract-lang`
- **Linux:** `sudo apt install tesseract-ocr-hin` (replace `hin` with the language code)

---

## 6. Troubleshooting

| Problem | Fix |
|---------|-----|
| `TesseractNotFoundError` | Install Tesseract, or pass `--tesseract-cmd` with the full path |
| `Unable to get page count` / Poppler error | Install Poppler, or pass `--poppler-path` |
| `Error loading language 'xyz'` | Install that language pack (see above) |
| `ModuleNotFoundError` | Activate your virtual environment and run `pip install -r requirements.txt` |
| Searchable PDF fails with a pypdf message | `pip install pypdf` |
| Output text is gibberish | Raise `--dpi`, check `--lang`, use the image cleanup options, or choose an image format |
| Pages come out sideways | Use `--rotate 90` (or 180, 270) or `--auto-rotate` |
| Slow or high memory use on big files | Lower `--dpi` to 200, or convert in chunks with `--pages` |

Run `python pdf2docx_ocr.py --check` first: it catches most setup problems.

---

## 7. How it works

1. Each PDF page is rendered to an image with Poppler, one page at a time to keep memory low. Image files are loaded directly (phone photo rotation is respected).
2. Optional cleanup is applied: rotate, grayscale, contrast, sharpen, threshold.
3. For text formats, Tesseract reads the text. For `pdf`, Tesseract builds a page with a hidden text layer, and the pages are merged.
4. The results are written in each format you asked for.

## 8. Limitations

- OCR output is plain paragraphs. Tables, columns, fonts and exact layout are not rebuilt.
- Handwriting recognition is weak.
- Image-based outputs (`docx-image`, `png`, `jpg`) are larger and not editable as text.

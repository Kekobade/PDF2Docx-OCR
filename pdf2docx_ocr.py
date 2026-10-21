#!/usr/bin/env python3
"""pdf2docx-ocr: convert image-based PDFs and images to DOCX, TXT, Markdown,
HTML, searchable PDF, or PNG/JPG pages. Runs 100% offline.

Quick examples:
  python pdf2docx_ocr.py scan.pdf
  python pdf2docx_ocr.py scan.pdf -f docx,txt,md --pages 1-5
  python pdf2docx_ocr.py scan.pdf -f pdf --grayscale --autocontrast
  python pdf2docx_ocr.py scan.pdf -f docx-image
  python pdf2docx_ocr.py ./folder --batch -f docx
  python pdf2docx_ocr.py                      # interactive menu
  python pdf2docx_ocr.py --check              # verify your setup
"""
import argparse
import html
import io
import re
import shutil
import subprocess
import sys
from pathlib import Path

from PIL import Image, ImageFilter, ImageOps

IMAGE_EXTS = {".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff", ".webp"}

FORMAT_HELP = [
    ("docx", "Word document with OCR text (editable)"),
    ("docx-image", "Word document with each page as a picture"),
    ("txt", "Plain text from OCR"),
    ("md", "Markdown text from OCR"),
    ("html", "HTML page with OCR text"),
    ("pdf", "Searchable PDF (image + hidden text layer, needs pypdf)"),
    ("png", "Each page saved as a PNG image"),
    ("jpg", "Each page saved as a JPG image"),
]
FORMATS = [k for k, _ in FORMAT_HELP]
EXT = {"docx": ".docx", "docx-image": ".docx", "txt": ".txt", "md": ".md",
       "html": ".html", "pdf": ".pdf", "png": ".png", "jpg": ".jpg"}
ALIASES = {"word": "docx", "text": "txt", "markdown": "md", "jpeg": "jpg",
           "searchable-pdf": "pdf", "image-docx": "docx-image"}
TEXT_FORMATS = {"docx", "txt", "md", "html"}


# --------------------------------------------------------------------------
# Helpers
# --------------------------------------------------------------------------
def parse_formats(spec):
    out = []
    for item in spec.split(","):
        item = item.strip().lower()
        if not item:
            continue
        item = ALIASES.get(item, item)
        if item not in FORMATS:
            sys.exit(f"Unknown format '{item}'. Choose from: {', '.join(FORMATS)}")
        if item not in out:
            out.append(item)
    if not out:
        sys.exit("No output format given.")
    return out


def parse_pages(spec, total):
    """'1-3,5,8-' -> [1,2,3,5,8,...,total]"""
    if not spec:
        return list(range(1, total + 1))
    pages = []
    for part in spec.split(","):
        part = part.strip()
        if not part:
            continue
        if "-" in part:
            a, b = part.split("-", 1)
            start = int(a) if a.strip() else 1
            end = int(b) if b.strip() else total
            pages.extend(range(start, end + 1))
        else:
            pages.append(int(part))
    seen, out = set(), []
    for p in pages:
        if 1 <= p <= total and p not in seen:
            seen.add(p)
            out.append(p)
    if not out:
        sys.exit(f"No valid pages in '{spec}' (document has {total} pages).")
    return out


def get_tesseract(a):
    try:
        import pytesseract
    except ImportError:
        sys.exit("pytesseract is not installed. Run: pip install pytesseract")
    if a.tesseract_cmd:
        pytesseract.pytesseract.tesseract_cmd = a.tesseract_cmd
    return pytesseract


def iter_pages(inputs, a):
    """Yield one PIL image per page, loading PDF pages in bounded batches to maximize speed."""
    for path in inputs:
        suffix = path.suffix.lower()
        if suffix == ".pdf":
            from pdf2image import convert_from_path, pdfinfo_from_path
            total = pdfinfo_from_path(str(path), poppler_path=a.poppler_path)["Pages"]
            pages = parse_pages(a.pages, total)
            batch_size = max(1, getattr(a, "page_batch", 5))
            for i in range(0, len(pages), batch_size):
                chunk = pages[i:i + batch_size]
                if chunk == list(range(chunk[0], chunk[-1] + 1)):
                    imgs = convert_from_path(str(path), dpi=a.dpi, first_page=chunk[0],
                                             last_page=chunk[-1], poppler_path=a.poppler_path)
                    for img in imgs:
                        yield img
                else:
                    for n in chunk:
                        imgs = convert_from_path(str(path), dpi=a.dpi, first_page=n,
                                                 last_page=n, poppler_path=a.poppler_path)
                        yield imgs[0]
        elif suffix in IMAGE_EXTS:
            img = Image.open(path)
            img = ImageOps.exif_transpose(img)  # respect phone-photo rotation
            yield img.convert("RGB")
        else:
            sys.exit(f"Unsupported file type: {path}")


def preprocess(img, a, tess):
    if a.auto_rotate and tess is not None:
        try:
            osd = tess.image_to_osd(img, output_type=tess.Output.DICT)
            if osd.get("rotate"):
                img = img.rotate(-osd["rotate"], expand=True)
        except Exception:
            pass  # OSD can fail on pages with little text; just continue
    if a.rotate:
        img = img.rotate(-a.rotate, expand=True)  # clockwise degrees
    if a.grayscale or a.threshold:
        img = img.convert("L")
    if a.autocontrast:
        img = ImageOps.autocontrast(img)
    if a.sharpen:
        img = img.filter(ImageFilter.SHARPEN)
    if a.threshold:
        t = a.threshold
        img = img.point(lambda p: 255 if p > t else 0)
    return img


def paragraphs(text, keep_breaks):
    for block in re.split(r"\n\s*\n", text.strip()):
        block = block.strip()
        if block:
            yield block if keep_breaks else " ".join(block.split())


# --------------------------------------------------------------------------
# Writers (idx = list of page positions to include in this file)
# --------------------------------------------------------------------------
def write_docx_text(idx, texts, path, keep):
    from docx import Document
    doc = Document()
    for k, i in enumerate(idx):
        for p in paragraphs(texts[i], keep):
            doc.add_paragraph(p)
        if k < len(idx) - 1:
            doc.add_page_break()
    doc.save(str(path))


def write_docx_images(idx, imgs, path):
    from docx import Document
    from docx.shared import Inches
    doc = Document()
    for k, i in enumerate(idx):
        buf = io.BytesIO()
        imgs[i].save(buf, format="PNG")
        buf.seek(0)
        doc.add_picture(buf, width=Inches(6.5))
        if k < len(idx) - 1:
            doc.add_page_break()
    doc.save(str(path))


def write_txt(idx, texts, path, keep):
    parts = []
    for i in idx:
        body = "\n\n".join(paragraphs(texts[i], keep))
        parts.append(f"--- Page {i + 1} ---\n\n{body}" if len(idx) > 1 else body)
    Path(path).write_text("\n\n".join(parts) + "\n", encoding="utf-8")


def write_md(idx, texts, path, keep):
    parts = []
    for i in idx:
        body = "\n\n".join(paragraphs(texts[i], keep))
        parts.append(body)
    Path(path).write_text("\n\n---\n\n".join(parts) + "\n", encoding="utf-8")


def write_html(idx, texts, path, keep):
    sections = []
    for i in idx:
        paras = "\n".join(f"<p>{html.escape(p)}</p>" for p in paragraphs(texts[i], keep))
        sections.append(f'<section id="page-{i + 1}">\n{paras}\n</section>')
    doc = ("<!DOCTYPE html>\n<html><head><meta charset=\"utf-8\">"
           f"<title>{html.escape(Path(path).stem)}</title>"
           "<style>body{font-family:sans-serif;max-width:800px;margin:2rem auto;"
           "line-height:1.6;padding:0 1rem}section{border-bottom:1px solid #ccc;"
           "padding-bottom:1rem;margin-bottom:1rem;white-space:pre-wrap}</style>"
           "</head><body>\n" + "\n".join(sections) + "\n</body></html>\n")
    Path(path).write_text(doc, encoding="utf-8")


def write_pdf(idx, pdfs, path):
    try:
        from pypdf import PdfReader, PdfWriter
    except ImportError:
        sys.exit("Searchable PDF output needs pypdf. Run: pip install pypdf")
    writer = PdfWriter()
    for i in idx:
        for page in PdfReader(io.BytesIO(pdfs[i])).pages:
            writer.add_page(page)
    with open(path, "wb") as f:
        writer.write(f)


def save_image(img, path, fmt):
    if fmt == "jpg":
        img.convert("RGB").save(path, quality=92)
    else:
        img.save(path)


# --------------------------------------------------------------------------
# Job runner
# --------------------------------------------------------------------------
def out_base(a, inputs):
    known = set(EXT.values())
    if a.output:
        base = Path(a.output)
        if base.suffix.lower() in known:
            base = base.with_suffix("")
    else:
        base = inputs[0].with_suffix("")
        if len(inputs) > 1:
            base = base.with_name(base.name + "_combined")
    if a.outdir:
        Path(a.outdir).mkdir(parents=True, exist_ok=True)
        base = Path(a.outdir) / base.name
    return base


def run_job(inputs, a):
    fmts = a.formats
    need_text = bool(set(fmts) & TEXT_FORMATS)
    need_pdf = "pdf" in fmts
    need_img = bool(set(fmts) & {"docx-image", "png", "jpg"})
    tess = get_tesseract(a) if (need_text or need_pdf or a.auto_rotate) else None
    cfg = f"--psm {a.psm}" if a.psm else ""

    print(f"Reading: {', '.join(str(p) for p in inputs)}")
    texts, pdfs, imgs = [], [], []
    count = 0
    for raw in iter_pages(inputs, a):
        count += 1
        print(f"  Processing page {count} ...", end="\r", flush=True)
        img = preprocess(raw, a, tess)
        if need_text:
            texts.append(tess.image_to_string(img, lang=a.lang, config=cfg))
        if need_pdf:
            pdfs.append(tess.image_to_pdf_or_hocr(img, extension="pdf",
                                                   lang=a.lang, config=cfg))
        if need_img:
            imgs.append(img)
    print(" " * 40, end="\r")
    if count == 0:
        sys.exit("No pages were read.")

    base = out_base(a, inputs)
    for fmt in fmts:
        ext = EXT[fmt]
        tag = "_images" if fmt == "docx-image" else ""

        if fmt in ("png", "jpg"):
            for n, img in enumerate(imgs, 1):
                p = Path(f"{base}_p{n:03d}{ext}")
                save_image(img, p, fmt)
            print(f"Saved {count} {fmt.upper()} image(s) next to {base}_p001{ext}")
            continue

        groups = [[i] for i in range(count)] if a.split else [list(range(count))]
        for g in groups:
            part = f"_p{g[0] + 1:03d}" if a.split else ""
            path = Path(f"{base}{tag}{part}{ext}")
            if fmt == "docx":
                write_docx_text(g, texts, path, a.keep_linebreaks)
            elif fmt == "docx-image":
                write_docx_images(g, imgs, path)
            elif fmt == "txt":
                write_txt(g, texts, path, a.keep_linebreaks)
            elif fmt == "md":
                write_md(g, texts, path, a.keep_linebreaks)
            elif fmt == "html":
                write_html(g, texts, path, a.keep_linebreaks)
            elif fmt == "pdf":
                write_pdf(g, pdfs, path)
            if not a.split:
                print(f"Saved: {path}")
        if a.split:
            print(f"Saved {len(groups)} {fmt} file(s) as {base}{tag}_p001{ext}, ...")


# --------------------------------------------------------------------------
# Check / interactive / CLI
# --------------------------------------------------------------------------
def check(a):
    print(f"Python: {sys.version.split()[0]}")
    for mod, pkg in [("PIL", "pillow"), ("pdf2image", "pdf2image"),
                     ("docx", "python-docx"), ("pytesseract", "pytesseract"),
                     ("pypdf", "pypdf (only for searchable PDF)")]:
        try:
            __import__(mod)
            print(f"  [ok]      {pkg}")
        except ImportError:
            print(f"  [missing] {pkg}  ->  pip install {pkg.split()[0]}")

    cmd = a.tesseract_cmd or shutil.which("tesseract")
    if cmd:
        try:
            out = subprocess.run([cmd, "--version"], capture_output=True, text=True)
            print(f"  [ok]      tesseract: {(out.stdout or out.stderr).splitlines()[0]}")
            langs = subprocess.run([cmd, "--list-langs"], capture_output=True, text=True)
            names = [x for x in langs.stdout.splitlines()[1:] if x.strip()]
            print(f"            languages: {', '.join(names)}")
        except Exception as e:
            print(f"  [error]   tesseract found but failed to run: {e}")
    else:
        print("  [missing] tesseract (install it or use --tesseract-cmd)")

    poppler = shutil.which("pdftoppm")
    if not poppler and a.poppler_path:
        for name in ("pdftoppm", "pdftoppm.exe"):
            if (Path(a.poppler_path) / name).exists():
                poppler = str(Path(a.poppler_path) / name)
    print(f"  [ok]      poppler: {poppler}" if poppler else
          "  [missing] poppler (install it or use --poppler-path)")


def ask(prompt, default=""):
    suffix = f" [{default}]" if default else ""
    s = input(f"{prompt}{suffix}: ").strip().strip('"').strip("'")
    return s or default


def yes(prompt, default="n"):
    return ask(prompt + " (y/n)", default).lower().startswith("y")


def interactive(a):
    print("\n=== PDF / image converter (interactive) ===\n")
    path = ask("Input file or folder")
    if not path:
        sys.exit("No input given.")
    a.inputs = [Path(path)]
    if a.inputs[0].is_dir():
        a.batch = True

    print("\nOutput formats:")
    for n, (key, desc) in enumerate(FORMAT_HELP, 1):
        print(f"  {n}. {key:<11} {desc}")
    raw = ask("Choose formats (numbers or names, comma-separated)", "1")
    chosen = []
    for tok in raw.split(","):
        tok = tok.strip()
        if tok.isdigit() and 1 <= int(tok) <= len(FORMAT_HELP):
            chosen.append(FORMAT_HELP[int(tok) - 1][0])
        elif tok:
            chosen.append(tok)
    a.format = ",".join(chosen)

    a.pages = ask("Pages, e.g. 1-3,5 (blank = all)") or None
    a.lang = ask("OCR language", a.lang)
    a.dpi = int(ask("DPI", str(a.dpi)))
    if yes("Clean up images first (grayscale + autocontrast)?"):
        a.grayscale = a.autocontrast = True
    a.split = yes("One output file per page?")
    a.outdir = ask("Output folder (blank = next to input)") or None
    print()


def build_parser():
    p = argparse.ArgumentParser(
        description="Offline converter: image-based PDFs/images -> DOCX, TXT, MD, "
                    "HTML, searchable PDF, PNG/JPG.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="Formats: " + ", ".join(FORMATS))
    p.add_argument("inputs", nargs="*", type=Path,
                   help="PDF/image files, or a folder with --batch")
    p.add_argument("-f", "--format", default="docx",
                   help="Output format(s), comma-separated (default: docx)")
    p.add_argument("-o", "--output", help="Output name/base path (extension optional)")
    p.add_argument("-d", "--outdir", help="Folder to write outputs into")
    p.add_argument("--pages", help="Pages to convert for PDFs, e.g. 1-3,5,8-")
    p.add_argument("--split", action="store_true", help="One output file per page")
    p.add_argument("--batch", action="store_true",
                   help="Input is a folder: convert each PDF/image inside separately")
    p.add_argument("-i", "--interactive", action="store_true", help="Menu-driven mode")

    g = p.add_argument_group("OCR")
    g.add_argument("--lang", default="eng", help="Tesseract language(s), e.g. eng+hin")
    g.add_argument("--psm", type=int, help="Tesseract page segmentation mode (0-13)")
    g.add_argument("--keep-linebreaks", action="store_true",
                   help="Keep original line breaks instead of joining lines")
    g.add_argument("--list-langs", action="store_true", help="List installed languages")

    g = p.add_argument_group("Image options")
    g.add_argument("--dpi", type=int, default=300, help="Render DPI for PDFs (default 300)")
    g.add_argument("--grayscale", action="store_true")
    g.add_argument("--autocontrast", action="store_true")
    g.add_argument("--sharpen", action="store_true")
    g.add_argument("--threshold", type=int, metavar="0-255",
                   help="Convert to pure black/white at this level (try 150)")
    g.add_argument("--rotate", type=int, default=0, help="Rotate clockwise by degrees")
    g.add_argument("--auto-rotate", action="store_true",
                   help="Detect page orientation with Tesseract (best effort)")
    g.add_argument("--page-batch", type=int, default=5,
                   help="Number of PDF pages to render per Poppler batch (default 5)")

    g = p.add_argument_group("Tools")
    g.add_argument("--check", action="store_true", help="Check your installation")
    g.add_argument("--poppler-path", help="Poppler bin folder (Windows)")
    g.add_argument("--tesseract-cmd", help="Full path to tesseract executable")
    return p


def main():
    a = build_parser().parse_args()
    try:
        if a.check:
            return check(a)
        if a.list_langs:
            print("\n".join(get_tesseract(a).get_languages()))
            return
        if a.interactive or not a.inputs:
            if not sys.stdin.isatty() and not a.inputs:
                sys.exit("No input given. Run with --help for usage.")
            interactive(a)

        a.formats = parse_formats(a.format)
        if a.threshold is not None and not 0 <= a.threshold <= 255:
            sys.exit("--threshold must be between 0 and 255.")
        for p in a.inputs:
            if not p.exists():
                sys.exit(f"Not found: {p}")

        if a.batch:
            a.output = None
            folder = a.inputs[0]
            if not folder.is_dir():
                sys.exit("--batch needs a folder as input.")
            files = sorted(f for f in folder.iterdir()
                           if f.suffix.lower() == ".pdf" or f.suffix.lower() in IMAGE_EXTS)
            if not files:
                sys.exit("No PDFs or images found in that folder.")
            for f in files:
                run_job([f], a)
        else:
            run_job(a.inputs, a)
    except KeyboardInterrupt:
        sys.exit("\nCancelled.")
    except SystemExit:
        raise
    except Exception as e:
        sys.exit(f"Error: {e}")


if __name__ == "__main__":
    main()

# Changelog

All notable changes to `pdf2docx-ocr` and `layout_ocr_to_docx` are documented in this file.

## [Unreleased] - 2026-10-20

### Added
- **Italic font detection via projection profile shear variance (`layout_ocr_to_docx.py`)**:
  - *What*: Added `detect_word_italic` and integrated it into `word_style` and run generation. Text runs now emit `<w:i/>` when words exhibit characteristic rightward shear. Added `--no-italic` CLI flag to disable when needed.
  - *Why*: Previously italic text was completely undetected and rendered as regular text.
  - *How tested*: Verified against synthetic italic sentences in `test_data/case1_styles.pdf`, achieving 100% precision and recall on Liberation Sans italic and bold-italic test words. Verified OpenXML output in `word/document.xml` and visual rendering via LibreOffice PDF/pdftoppm round-trip.

- **Underline detection and text run integration (`layout_ocr_to_docx.py`)**:
  - *What*: Added `match_underlines` to match horizontal line shapes spanning immediately beneath text runs, converting them into true Word `<w:u w:val="single"/>` run attributes and suppressing redundant overlapping rectangle shapes. Added `--no-underline` CLI flag.
  - *Why*: Previously underlines were drawn as separate free-floating shapes that often overlapped or misaligned with text in Word.
  - *How tested*: Verified on underlined section heading in `case1_styles.pdf` and table cells in `case2_table.pdf`, confirming `<w:u>` generation and elimination of duplicate line geometries.

- **Paragraph and text alignment detection (`layout_ocr_to_docx.py`)**:
  - *What*: Added automatic alignment detection (center, right, justified, left) in `make_text_boxes`. Positioned page-centered and right-aligned text boxes appropriately and emitted `<w:jc w:val="..."/>` in `<w:pPr>`. Added `--no-align` CLI flag.
  - *Why*: Previously all text was left-aligned with static pixel indents (`w:ind w:left`), causing centered titles and right-aligned metadata to misalign upon editing.
  - *How tested*: Verified on centered titles, right-aligned dates/references, and justified paragraphs across `case1_styles.pdf` and `case4_noisy.pdf`.

- **Table grid line snapping and continuity (`layout_ocr_to_docx.py`)**:
  - *What*: Added `snap_table_grid` to snap nearly-touching horizontal and vertical line endpoints within 3 pixels.
  - *Why*: Scanning discretization and morphological operations previously created 1-2 pixel corner gaps in table borders.
  - *How tested*: Verified on 4x4 data table in `case2_table.pdf` via LibreOffice visual round-trip rendering.

- **Bounded-batch PDF page streaming (`layout_ocr_to_docx.py`, `pdf2docx_ocr.py`)**:
  - *What*: Added `--page-batch N` (default: 5) to batch consecutive PDF page conversions via Poppler while strictly bounding memory footprint.
  - *Why*: Calling `convert_from_path(first_page=n, last_page=n)` individually per page in a loop incurred massive process startup and PDF re-parsing overhead on multi-page documents.
  - *How tested*: Verified on `case5_multipage.pdf` (3 pages) and benchmarked rendering throughput.

- **Automated synthetic test generator and verification harness**:
  - *What*: Created `generate_test_cases.py` and `run_tests.py` testing noise, blur, tilt, colored paper, inverted text, styles, alignments, tables, charts, multi-page PDFs, OpenXML schema inspection, and headless LibreOffice visual rendering.
  - *Why*: To ensure comprehensive automated regression testing and measurable quality validation.

### Changed
- **Bold detection accuracy and orphan-noise isolation (`layout_ocr_to_docx.py`)**:
  - *What*: Dilated text mask during graphics separation to eliminate character edge residues. Added sanity filtering to `orphan_words` and excluded orphan fragments from the global stroke width distribution (`page_med`). Calibrated bold threshold to relative ratio (`max(bold_factor * page_med, 0.105)`).
  - *Why*: Hallucinated orphan characters had previously inflated `page_med` stroke width, causing genuine bold headings and bold body text to be missed.
  - *How tested*: Verified on `case1_styles.pdf`, `case2_table.pdf`, and `case3_banner.pdf`, correctly detecting bold headings and paragraphs without false positives on regular body text.

- **Color fidelity preservation (`layout_ocr_to_docx.py`)**:
  - *What*: Refined `snap_color` to only snap near-neutral dark gray to pure black (`max(c) < 36 and spread < 16`) and near-neutral light gray to pure white (`min(c) > 244 and spread < 12`).
  - *Why*: Previous threshold (`max(c) < 60`) aggressively crushed dark blue and dark green elements to pure black, and washed out cream paper backgrounds.
  - *How tested*: Verified dark blue banner (`#123466`) and warm cream paper (`#FCF8EE`) in `case3_banner.pdf`.

- **Font size estimation refinement (`layout_ocr_to_docx.py`)**:
  - *What*: Differentiated cap-height/ascender lines (divisor 0.76) from lowercase-only lines (divisor 0.54) instead of a fixed 0.85 divisor.
  - *Why*: Standard typography cap-height is ~70-75% of body font size; dividing by 0.85 consistently underestimated true point sizes by 10-15%.
  - *How tested*: Verified on 48pt, 32pt, 26pt, and 24pt ground-truth fonts in `case1_styles.pdf`.

- **Bordered fill shape classification (`layout_ocr_to_docx.py`)**:
  - *What*: Added interior erosion prior to color standard deviation evaluation in candidate shape fills.
  - *Why*: Dark borders around light-colored headers previously caused them to be misclassified as raster images instead of filled vector shapes.
  - *How tested*: Verified table headers in `case2_table.pdf` and alert boxes in `case3_banner.pdf` are rendered as native Word vector shapes.

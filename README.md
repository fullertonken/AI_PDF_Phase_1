# PDF Invoice Processor

A fully local, privacy-first invoice processing system. No documents ever leave your machine.

## Overview

### Phase 1: Document Preparation (this release)
- **Split** multi-page PDFs into single pages
- **Detect & correct orientation** using Tesseract OSD
- **Apply OCR** to scanned pages that lack embedded text
- **Output**: Clean, upright, text-searchable single-page PDFs + images + manifest

### Phase 2: Classification & Extraction (planned)
- Classify pages as invoice vs. supporting document
- Group pages belonging to the same invoice
- Extract structured fields (invoice #, date, totals, GST/PST, etc.)
- Uses Ollama (local LLM) — no cloud APIs

### Phase 3: Annotation & Human-in-the-Loop (planned)
- Annotate PDFs with highlighted extracted fields
- Manual review interface for low-confidence extractions
- Training data pipeline for Donut fine-tuning on your RTX 2060

---

## Setup

### 1. System Dependencies (Ubuntu/Debian)

```bash
sudo bash install_system_deps.sh
```

Or manually:
```bash
sudo apt update
sudo apt install -y tesseract-ocr tesseract-ocr-eng poppler-utils ghostscript unpaper
```

### 2. Python Dependencies

```bash
pip install -r requirements.txt
```

### 3. Verify Installation

```bash
# Check Tesseract
tesseract --version

# Check poppler
pdftoppm -v

# Check Ghostscript
gs --version
```

---

## Usage

### GUI Mode (recommended)

```bash
python phase1_processor.py
```

1. Click **Browse** to select the folder containing your PDFs
2. Optionally change the output folder (defaults to `processed_output/` inside the input folder)
3. Check **Include subfolders** if your PDFs are in nested directories
4. Click **Start Processing**
5. Watch the log for per-page status

### What It Does Per Page

```
Source PDF
  ↓
Split into single page
  ↓
Check for embedded text (pdfplumber)
  ↓
Render to image at 300 DPI
  ↓
Tesseract OSD → detect rotation needed
  ↓
Rotate if confidence > threshold
  ↓
If no text → OCR (ocrmypdf or Tesseract fallback)
  ↓
Save: corrected PDF + PNG image
```

---

## Output Structure

```
processed_output/
├── manifest.json          # Full metadata for every processed page
├── pages/                 # Single-page PDFs (oriented, OCR'd)
│   ├── invoice_batch_page_0001.pdf
│   ├── invoice_batch_page_0002.pdf
│   └── ...
└── images/                # Page images (PNG, oriented)
    ├── invoice_batch_page_0001.png
    ├── invoice_batch_page_0002.png
    └── ...
```

### Manifest Format

Each page record in `manifest.json`:

```json
{
  "page_id": "invoice_batch_page_0001",
  "source_pdf": "/path/to/invoice_batch.pdf",
  "source_pdf_basename": "invoice_batch.pdf",
  "original_page": 1,
  "total_source_pages": 47,
  "orientation_detected": 90,
  "orientation_correction": 90,
  "orientation_confidence": 12.5,
  "had_text": false,
  "ocr_applied": true,
  "text_length": 1823,
  "output_pdf": "processed_output/pages/invoice_batch_page_0001.pdf",
  "output_image": "processed_output/images/invoice_batch_page_0001.png",
  "processing_notes": [
    "No usable embedded text (12 chars)",
    "OSD: rotate=90°, conf=12.5",
    "Applying 90° rotation",
    "OCR via ocrmypdf"
  ]
}
```

---

## Configuration

Key constants in `phase1_processor.py` that you can adjust:

| Constant | Default | Description |
|----------|---------|-------------|
| `TEXT_THRESHOLD` | 40 | Minimum characters to consider a page as having usable embedded text |
| `OCR_DPI` | 300 | DPI for rendering pages to images (higher = better OCR, slower) |
| `ORIENTATION_CONFIDENCE` | 2.0 | Minimum Tesseract OSD confidence to trust rotation detection |

---

## Hardware Notes

- **RTX 2060 12GB**: Phase 1 uses CPU (Tesseract). The GPU will be utilized in Phase 2/3 for Ollama inference and Donut fine-tuning.
- **Large batches**: The tool processes pages sequentially. For a 500-page PDF, expect roughly 5-15 seconds per page depending on whether OCR is needed.
- **Disk space**: Each page produces a PDF + PNG image. Budget ~1-3 MB per page for the output.

---

## Troubleshooting

**"No PDF files found"** — Make sure the selected folder directly contains `.pdf` files, or check "Include subfolders".

**"Missing: tesseract-ocr"** — Run `sudo apt install tesseract-ocr tesseract-ocr-eng`.

**Orientation detection wrong** — Increase `ORIENTATION_CONFIDENCE` (e.g., to 5.0) to be more conservative, or decrease it to trust Tesseract more often. Pages with very little text may get unreliable OSD results.

**OCR quality poor** — Ensure `OCR_DPI` is at least 300. For very small text, try 400. Install additional Tesseract language packs if documents aren't English.

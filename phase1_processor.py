"""
PDF Invoice Processor - Phase 1: Document Preparation Pipeline

Reads PDFs from a selected folder, splits multi-page PDFs into single pages,
detects and corrects orientation, and applies OCR where needed.

All processing is done locally — no data leaves the machine.

Requirements:
    System: tesseract-ocr, poppler-utils, ghostscript
    Python: pypdf, pytesseract, pdf2image, Pillow, ocrmypdf, pdfplumber, tqdm, deskew, numpy
"""

import os
import sys
import json
import shutil
import logging
import subprocess
import threading
import tempfile
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from pathlib import Path
from datetime import datetime
from dataclasses import dataclass, asdict, field
from typing import Optional

# ---------------------------------------------------------------------------
# PDF / Image / OCR imports
# ---------------------------------------------------------------------------
try:
    from pypdf import PdfReader, PdfWriter
except ImportError:
    sys.exit("Missing 'pypdf'. Run: pip install pypdf")

try:
    from PIL import Image
except ImportError:
    sys.exit("Missing 'Pillow'. Run: pip install Pillow")

try:
    import pytesseract
except ImportError:
    sys.exit("Missing 'pytesseract'. Run: pip install pytesseract")

try:
    from pdf2image import convert_from_path
except ImportError:
    sys.exit("Missing 'pdf2image'. Run: pip install pdf2image")

try:
    import pdfplumber
except ImportError:
    sys.exit("Missing 'pdfplumber'. Run: pip install pdfplumber")

try:
    from tqdm import tqdm
except ImportError:
    tqdm = None  # non-critical, we handle progress in the GUI

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------
TEXT_THRESHOLD = 40          # min chars to consider a page as having usable text
OCR_DPI = 300                # DPI for rendering pages to images
ORIENTATION_CONFIDENCE = 2.0 # min confidence from Tesseract OSD to trust rotation
LOG_FORMAT = "%(asctime)s [%(levelname)s] %(message)s"

# ---------------------------------------------------------------------------
# Data classes
# ---------------------------------------------------------------------------
@dataclass
class PageRecord:
    """Metadata for a single processed page."""
    page_id: str
    source_pdf: str
    source_pdf_basename: str
    original_page: int           # 1-based
    total_source_pages: int
    orientation_detected: int    # degrees detected by OSD
    orientation_correction: int  # degrees actually rotated
    orientation_confidence: float
    had_text: bool               # did the original page have embedded text?
    ocr_applied: bool
    text_length: int             # length of final extracted text
    output_pdf: str              # path to the processed single-page PDF
    output_image: str            # path to the corrected page image (PNG)
    processing_notes: list = field(default_factory=list)


# ---------------------------------------------------------------------------
# Utility helpers
# ---------------------------------------------------------------------------
def check_system_dependencies() -> list[str]:
    """Return a list of missing system tools."""
    missing = []
    # On Windows, Ghostscript is installed as gswin64c.exe or gswin64.exe
    gs_candidates = ("gswin64c", "gswin64", "gs") if sys.platform == "win32" else ("gs",)
    for tool in ("tesseract", "pdftoppm"):
        if shutil.which(tool) is None:
            missing.append(tool)
    if not any(shutil.which(g) for g in gs_candidates):
        missing.append("gs")
    return missing


def extract_text_from_pdf_page(pdf_path: str, page_num: int = 0) -> str:
    """Extract embedded text from a single page of a PDF using pdfplumber."""
    try:
        with pdfplumber.open(pdf_path) as pdf:
            if page_num < len(pdf.pages):
                text = pdf.pages[page_num].extract_text() or ""
                return text.strip()
    except Exception:
        pass
    return ""


def render_page_to_image(pdf_path: str, page_num: int = 0, dpi: int = OCR_DPI) -> Image.Image:
    """Render a single PDF page to a PIL Image."""
    images = convert_from_path(
        pdf_path,
        first_page=page_num + 1,
        last_page=page_num + 1,
        dpi=dpi,
    )
    return images[0]


def detect_orientation(image: Image.Image) -> tuple[int, float]:
    """
    Use Tesseract OSD to detect text orientation.
    Returns (rotation_degrees, confidence).
    rotation_degrees is the amount the image needs to be rotated clockwise
    to make text upright.
    """
    try:
        osd = pytesseract.image_to_osd(image, output_type=pytesseract.Output.DICT)
        rotate = osd.get("rotate", 0)
        confidence = float(osd.get("orientation_conf", 0.0))
        return rotate, confidence
    except pytesseract.TesseractError:
        return 0, 0.0


def rotate_image(image: Image.Image, degrees: int) -> Image.Image:
    """Rotate image by given degrees (counter-clockwise in PIL convention).
    Tesseract 'rotate' value = clockwise rotation needed, so we negate for PIL."""
    if degrees == 0:
        return image
    # PIL expand=True keeps the full image visible after rotation
    return image.rotate(-degrees, expand=True)


def deskew_image_pil(image: Image.Image) -> tuple[Image.Image, float]:
    """
    Detect and correct small-angle scan skew using the 'deskew' package.
    Returns (corrected_image, angle_applied).  angle_applied is 0.0 when no
    correction was needed or the library is unavailable.
    Requires: pip install deskew numpy
    """
    try:
        import numpy as np
        from deskew import determine_skew
        gray = np.array(image.convert("L"))
        angle = determine_skew(gray)
        if angle is not None and abs(angle) > 0.3:  # ignore sub-0.3° noise
            corrected = image.rotate(angle, expand=True, fillcolor="white")
            return corrected, round(angle, 2)
    except ImportError:
        pass
    except Exception as e:
        logging.debug(f"Deskew failed: {e}")
    return image, 0.0


def ocr_image_to_pdf_bytes(image: Image.Image) -> bytes:
    """Run Tesseract on an image and produce a single-page PDF with text layer."""
    pdf_bytes = pytesseract.image_to_pdf_or_hocr(image, extension="pdf")
    return pdf_bytes


def create_single_page_pdf(reader: PdfReader, page_index: int, output_path: str):
    """Extract one page from a PdfReader and write it to output_path."""
    writer = PdfWriter()
    writer.add_page(reader.pages[page_index])
    with open(output_path, "wb") as f:
        writer.write(f)


def apply_rotation_to_pdf(input_pdf: str, degrees: int, output_pdf: str):
    """Rotate a single-page PDF by the given degrees clockwise and save."""
    reader = PdfReader(input_pdf)
    writer = PdfWriter()
    page = reader.pages[0]
    page.rotate(degrees)
    writer.add_page(page)
    with open(output_pdf, "wb") as f:
        writer.write(f)


def apply_ocr_to_pdf(input_pdf: str, output_pdf: str, force_ocr: bool = True) -> bool:
    """
    Use ocrmypdf to add an OCR text layer to a PDF.

    force_ocr=True  → OCR all pages (use for pages without embedded text).
    force_ocr=False → only OCR pages that lack a text layer; others pass through.
    Small-angle deskew is handled at the image level (deskew_image_pil) rather
    than here, because unpaper (needed by ocrmypdf deskew) is not available on Windows.
    Falls back gracefully if ocrmypdf is not installed.
    """
    try:
        import ocrmypdf
        ocrmypdf.ocr(
            input_pdf,
            output_pdf,
            language="eng",
            force_ocr=force_ocr,
            optimize=1,
        )
        return True
    except ImportError:
        return False
    except Exception as e:
        logging.warning(f"ocrmypdf failed: {e}, falling back to manual OCR")
        return False


# ---------------------------------------------------------------------------
# Main processing pipeline
# ---------------------------------------------------------------------------
def process_page(
    source_pdf_path: str,
    page_index: int,
    total_pages: int,
    output_dir: str,
    images_dir: str,
    page_counter: int,
) -> PageRecord:
    """
    Process a single page from a source PDF:
      1. Extract to single-page PDF
      2. Check for embedded text
      3. Render to image
      4. Detect orientation via OSD
      5. Correct orientation
      6. Apply OCR if needed
      7. Save processed PDF + image
    """
    source_basename = os.path.basename(source_pdf_path)
    stem = Path(source_pdf_path).stem
    page_id = f"{stem}_page_{page_index + 1:04d}"

    notes = []
    temp_files = []

    try:
        # Step 1: Extract single page
        reader = PdfReader(source_pdf_path)
        temp_single = os.path.join(output_dir, f"_temp_{page_id}.pdf")
        create_single_page_pdf(reader, page_index, temp_single)
        temp_files.append(temp_single)

        # Step 2: Check embedded text
        embedded_text = extract_text_from_pdf_page(temp_single, 0)
        has_text = len(embedded_text) >= TEXT_THRESHOLD

        if has_text:
            notes.append(f"Embedded text found ({len(embedded_text)} chars)")
        else:
            notes.append(f"No usable embedded text ({len(embedded_text)} chars)")

        # Step 3: Render to image for orientation detection
        page_image = render_page_to_image(temp_single, 0, dpi=OCR_DPI)

        # Step 4: Detect orientation
        rotation_needed, osd_confidence = detect_orientation(page_image)
        notes.append(f"OSD: rotate={rotation_needed}°, conf={osd_confidence:.1f}")

        # Step 5: Decide on rotation
        actual_rotation = 0
        if osd_confidence >= ORIENTATION_CONFIDENCE and rotation_needed != 0:
            actual_rotation = rotation_needed
            notes.append(f"Applying {actual_rotation}° rotation")
        elif rotation_needed != 0:
            notes.append(f"Low confidence ({osd_confidence:.1f}), skipping rotation")

        # Step 6: Apply rotation to image
        corrected_image = rotate_image(page_image, actual_rotation)

        # Step 6b: Detect and correct small-angle scan skew at the image level.
        # This is done in Python (no external binary needed) and runs on every page.
        corrected_image, deskew_angle = deskew_image_pil(corrected_image)
        if deskew_angle:
            notes.append(f"Image deskew: {deskew_angle:.2f}°")

        # Save the corrected image (orientation + deskew applied)
        image_path = os.path.join(images_dir, f"{page_id}.png")
        corrected_image.save(image_path, "PNG")

        # Step 7: Build the final PDF
        output_pdf_path = os.path.join(output_dir, f"{page_id}.pdf")

        if deskew_angle:
            # The image was geometrically corrected — the output PDF must be rebuilt
            # from that corrected image so the deskew is baked in.
            # Tesseract is used directly because ocrmypdf can't consume our pre-deskewed image.
            pdf_bytes = ocr_image_to_pdf_or_hocr_fallback(corrected_image)
            with open(output_pdf_path, "wb") as f:
                f.write(pdf_bytes)
            notes.append(f"PDF rebuilt from deskewed image via Tesseract")
            ocr_applied = True

        elif has_text and actual_rotation == 0:
            # No corrections needed — copy the original page as-is.
            shutil.copy2(temp_single, output_pdf_path)
            ocr_applied = False
            notes.append("Kept original PDF (no corrections needed)")

        elif has_text and actual_rotation != 0:
            # Only a 90°/180°/270° rotation needed — apply it directly to the PDF.
            apply_rotation_to_pdf(temp_single, actual_rotation, output_pdf_path)
            ocr_applied = False
            notes.append("Rotated existing PDF")

        else:
            # No embedded text, no deskew — standard OCR path.
            if actual_rotation != 0:
                rotated_temp = os.path.join(output_dir, f"_temp_rot_{page_id}.pdf")
                apply_rotation_to_pdf(temp_single, actual_rotation, rotated_temp)
                temp_files.append(rotated_temp)
                ocr_input = rotated_temp
            else:
                ocr_input = temp_single

            ocr_success = apply_ocr_to_pdf(ocr_input, output_pdf_path, force_ocr=True)
            if not ocr_success:
                pdf_bytes = ocr_image_to_pdf_or_hocr_fallback(corrected_image)
                with open(output_pdf_path, "wb") as f:
                    f.write(pdf_bytes)
                notes.append("OCR via Tesseract direct (ocrmypdf unavailable)")
            else:
                notes.append("OCR via ocrmypdf")
            ocr_applied = True

        # Get final text content
        final_text = extract_text_from_pdf_page(output_pdf_path, 0)

        return PageRecord(
            page_id=page_id,
            source_pdf=source_pdf_path,
            source_pdf_basename=source_basename,
            original_page=page_index + 1,
            total_source_pages=total_pages,
            orientation_detected=rotation_needed,
            orientation_correction=actual_rotation,
            orientation_confidence=osd_confidence,
            had_text=has_text,
            ocr_applied=ocr_applied,
            text_length=len(final_text),
            output_pdf=output_pdf_path,
            output_image=image_path,
            processing_notes=notes,
        )

    finally:
        # Cleanup temp files
        for tf in temp_files:
            try:
                if os.path.exists(tf):
                    os.remove(tf)
            except OSError:
                pass


def ocr_image_to_pdf_or_hocr_fallback(image: Image.Image) -> bytes:
    """Tesseract direct: image → PDF bytes with text layer."""
    return pytesseract.image_to_pdf_or_hocr(image, extension="pdf")


def discover_pdfs(folder: str) -> list[str]:
    """Find all PDF files in the given folder (non-recursive by default)."""
    pdfs = []
    for entry in sorted(os.listdir(folder)):
        if entry.lower().endswith(".pdf"):
            pdfs.append(os.path.join(folder, entry))
    return pdfs


def count_total_pages(pdf_paths: list[str]) -> int:
    """Count total pages across all PDFs."""
    total = 0
    for p in pdf_paths:
        try:
            reader = PdfReader(p)
            total += len(reader.pages)
        except Exception:
            pass
    return total


def run_pipeline(
    input_folder: str,
    output_folder: str,
    recursive: bool = False,
    progress_callback=None,
    log_callback=None,
    cancel_event: threading.Event = None,
) -> list[PageRecord]:
    """
    Main Phase 1 pipeline.

    Args:
        input_folder: path to folder containing source PDFs
        output_folder: path where processed output goes
        recursive: whether to scan subfolders
        progress_callback: fn(current, total, message) for GUI progress
        log_callback: fn(message) for GUI log
        cancel_event: threading.Event to signal cancellation

    Returns:
        List of PageRecord for every processed page.
    """
    def log(msg):
        logging.info(msg)
        if log_callback:
            log_callback(msg)

    def progress(current, total, msg=""):
        if progress_callback:
            progress_callback(current, total, msg)

    # Discover PDFs
    if recursive:
        pdf_paths = []
        for root, _, files in os.walk(input_folder):
            for f in sorted(files):
                if f.lower().endswith(".pdf"):
                    pdf_paths.append(os.path.join(root, f))
    else:
        pdf_paths = discover_pdfs(input_folder)

    if not pdf_paths:
        log("No PDF files found in the selected folder.")
        return []

    log(f"Found {len(pdf_paths)} PDF file(s)")

    # Count total pages for progress
    total_pages = count_total_pages(pdf_paths)
    log(f"Total pages to process: {total_pages}")

    # Create output directories
    pages_dir = os.path.join(output_folder, "pages")
    images_dir = os.path.join(output_folder, "images")
    os.makedirs(pages_dir, exist_ok=True)
    os.makedirs(images_dir, exist_ok=True)

    records = []
    page_counter = 0

    for pdf_path in pdf_paths:
        if cancel_event and cancel_event.is_set():
            log("Processing cancelled by user.")
            break

        try:
            reader = PdfReader(pdf_path)
            num_pages = len(reader.pages)
        except Exception as e:
            log(f"ERROR: Cannot read '{os.path.basename(pdf_path)}': {e}")
            continue

        log(f"Processing: {os.path.basename(pdf_path)} ({num_pages} pages)")

        for page_idx in range(num_pages):
            if cancel_event and cancel_event.is_set():
                break

            page_counter += 1
            progress(page_counter, total_pages,
                     f"{os.path.basename(pdf_path)} p.{page_idx + 1}/{num_pages}")

            try:
                record = process_page(
                    source_pdf_path=pdf_path,
                    page_index=page_idx,
                    total_pages=num_pages,
                    output_dir=pages_dir,
                    images_dir=images_dir,
                    page_counter=page_counter,
                )
                records.append(record)

                status = "OCR" if record.ocr_applied else "TEXT"
                rot = f" rot:{record.orientation_correction}°" if record.orientation_correction else ""
                log(f"  Page {page_idx + 1}: [{status}]{rot} → {record.page_id}.pdf "
                    f"({record.text_length} chars)")

            except Exception as e:
                log(f"  ERROR on page {page_idx + 1}: {e}")
                logging.exception(f"Failed processing {pdf_path} page {page_idx}")

    # Save manifest
    manifest_path = os.path.join(output_folder, "manifest.json")
    manifest_data = {
        "created": datetime.now().isoformat(),
        "input_folder": input_folder,
        "output_folder": output_folder,
        "total_source_files": len(pdf_paths),
        "total_pages_processed": len(records),
        "pages": [asdict(r) for r in records],
    }
    with open(manifest_path, "w") as f:
        json.dump(manifest_data, f, indent=2)

    log(f"\nDone! Processed {len(records)} pages from {len(pdf_paths)} file(s)")
    log(f"Output: {output_folder}")
    log(f"Manifest: {manifest_path}")

    return records


# ===========================================================================
# GUI Application
# ===========================================================================
class InvoiceProcessorGUI:
    """Tkinter GUI for the Phase 1 PDF preparation pipeline."""

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("PDF Invoice Processor — Phase 1: Document Preparation")
        self.root.geometry("820x680")
        self.root.minsize(700, 550)

        self.input_folder = tk.StringVar()
        self.output_folder = tk.StringVar()
        self.recursive = tk.BooleanVar(value=False)
        self.processing = False
        self.cancel_event = threading.Event()

        self._build_ui()
        self._check_dependencies()

    # -----------------------------------------------------------------------
    # UI construction
    # -----------------------------------------------------------------------
    def _build_ui(self):
        # Main container
        main = ttk.Frame(self.root, padding=10)
        main.pack(fill=tk.BOTH, expand=True)

        # --- Header ---
        header = ttk.Label(main, text="PDF Invoice Processor — Phase 1",
                           font=("Segoe UI", 16, "bold"))
        header.pack(pady=(0, 5))
        subtitle = ttk.Label(main,
                             text="Split · Orient · OCR — All processing done locally",
                             font=("Segoe UI", 10))
        subtitle.pack(pady=(0, 10))

        # --- Dependency status ---
        self.dep_frame = ttk.LabelFrame(main, text="System Dependencies", padding=5)
        self.dep_frame.pack(fill=tk.X, pady=(0, 10))
        self.dep_label = ttk.Label(self.dep_frame, text="Checking...")
        self.dep_label.pack(anchor=tk.W)

        # --- Input folder ---
        input_frame = ttk.LabelFrame(main, text="Input Folder (containing PDFs)", padding=5)
        input_frame.pack(fill=tk.X, pady=(0, 5))

        input_row = ttk.Frame(input_frame)
        input_row.pack(fill=tk.X)
        ttk.Entry(input_row, textvariable=self.input_folder).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        ttk.Button(input_row, text="Browse…", command=self._browse_input).pack(side=tk.RIGHT)

        ttk.Checkbutton(input_frame, text="Include subfolders (recursive)",
                        variable=self.recursive).pack(anchor=tk.W, pady=(3, 0))

        # --- Output folder ---
        output_frame = ttk.LabelFrame(main, text="Output Folder", padding=5)
        output_frame.pack(fill=tk.X, pady=(0, 5))

        output_row = ttk.Frame(output_frame)
        output_row.pack(fill=tk.X)
        ttk.Entry(output_row, textvariable=self.output_folder).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        ttk.Button(output_row, text="Browse…", command=self._browse_output).pack(side=tk.RIGHT)

        # --- Progress ---
        progress_frame = ttk.LabelFrame(main, text="Progress", padding=5)
        progress_frame.pack(fill=tk.X, pady=(0, 5))

        self.progress_label = ttk.Label(progress_frame, text="Ready")
        self.progress_label.pack(anchor=tk.W)
        self.progress_bar = ttk.Progressbar(progress_frame, mode="determinate")
        self.progress_bar.pack(fill=tk.X, pady=(3, 0))

        # --- Buttons ---
        btn_frame = ttk.Frame(main)
        btn_frame.pack(fill=tk.X, pady=(5, 5))

        self.start_btn = ttk.Button(btn_frame, text="▶  Start Processing",
                                    command=self._start_processing)
        self.start_btn.pack(side=tk.LEFT, padx=(0, 5))

        self.cancel_btn = ttk.Button(btn_frame, text="■  Cancel",
                                     command=self._cancel_processing, state=tk.DISABLED)
        self.cancel_btn.pack(side=tk.LEFT, padx=(0, 5))

        self.open_btn = ttk.Button(btn_frame, text="📂 Open Output Folder",
                                   command=self._open_output, state=tk.DISABLED)
        self.open_btn.pack(side=tk.RIGHT)

        # --- Log ---
        log_frame = ttk.LabelFrame(main, text="Processing Log", padding=5)
        log_frame.pack(fill=tk.BOTH, expand=True)

        self.log_text = scrolledtext.ScrolledText(log_frame, height=12, wrap=tk.WORD,
                                                   font=("Consolas", 9))
        self.log_text.pack(fill=tk.BOTH, expand=True)

    # -----------------------------------------------------------------------
    # Dependency check
    # -----------------------------------------------------------------------
    def _check_dependencies(self):
        missing = check_system_dependencies()
        if missing:
            tool_names = {"tesseract": "tesseract-ocr", "pdftoppm": "poppler-utils", "gs": "ghostscript"}
            packages = [tool_names.get(t, t) for t in missing]
            if sys.platform == "win32":
                install_hint = "Download from: https://github.com/UB-Mannheim/tesseract/wiki (tesseract), https://github.com/oschwartz10612/poppler-windows (poppler), https://www.ghostscript.com/download.html (ghostscript)"
            else:
                install_hint = f"Install with: sudo apt install {' '.join(packages)}"
            self.dep_label.config(
                text=f"⚠ Missing: {', '.join(packages)}. {install_hint}",
                foreground="red",
            )
            self.start_btn.config(state=tk.DISABLED)
        else:
            self.dep_label.config(text="✓ All dependencies found (tesseract, poppler, ghostscript)",
                                  foreground="green")

    # -----------------------------------------------------------------------
    # Folder browsing
    # -----------------------------------------------------------------------
    def _browse_input(self):
        folder = filedialog.askdirectory(title="Select folder containing PDFs")
        if folder:
            self.input_folder.set(folder)
            # Auto-set output folder
            if not self.output_folder.get():
                default_out = os.path.join(folder, "processed_output")
                self.output_folder.set(default_out)

    def _browse_output(self):
        folder = filedialog.askdirectory(title="Select output folder")
        if folder:
            self.output_folder.set(folder)

    # -----------------------------------------------------------------------
    # Processing
    # -----------------------------------------------------------------------
    def _start_processing(self):
        input_dir = self.input_folder.get().strip()
        output_dir = self.output_folder.get().strip()

        if not input_dir or not os.path.isdir(input_dir):
            messagebox.showerror("Error", "Please select a valid input folder.")
            return
        if not output_dir:
            messagebox.showerror("Error", "Please select an output folder.")
            return

        # Create output dir
        os.makedirs(output_dir, exist_ok=True)

        # Reset
        self.log_text.delete("1.0", tk.END)
        self.progress_bar["value"] = 0
        self.cancel_event.clear()
        self.processing = True

        self.start_btn.config(state=tk.DISABLED)
        self.cancel_btn.config(state=tk.NORMAL)
        self.open_btn.config(state=tk.DISABLED)

        # Run in background thread
        thread = threading.Thread(target=self._run_pipeline_thread,
                                  args=(input_dir, output_dir), daemon=True)
        thread.start()

    def _run_pipeline_thread(self, input_dir, output_dir):
        try:
            records = run_pipeline(
                input_folder=input_dir,
                output_folder=output_dir,
                recursive=self.recursive.get(),
                progress_callback=self._update_progress,
                log_callback=self._append_log,
                cancel_event=self.cancel_event,
            )
            self.root.after(0, self._processing_done, len(records))
        except Exception as e:
            self.root.after(0, self._processing_error, str(e))

    def _cancel_processing(self):
        self.cancel_event.set()
        self._append_log("Cancelling... (will stop after current page)")

    def _processing_done(self, count):
        self.processing = False
        self.start_btn.config(state=tk.NORMAL)
        self.cancel_btn.config(state=tk.DISABLED)
        self.open_btn.config(state=tk.NORMAL)
        self.progress_label.config(text=f"Complete — {count} pages processed")
        self.progress_bar["value"] = 100

    def _processing_error(self, error_msg):
        self.processing = False
        self.start_btn.config(state=tk.NORMAL)
        self.cancel_btn.config(state=tk.DISABLED)
        messagebox.showerror("Processing Error", f"An error occurred:\n{error_msg}")

    # -----------------------------------------------------------------------
    # UI update helpers (thread-safe)
    # -----------------------------------------------------------------------
    def _update_progress(self, current, total, message):
        def _update():
            pct = (current / total * 100) if total > 0 else 0
            self.progress_bar["value"] = pct
            self.progress_label.config(text=f"[{current}/{total}] {message}")
        self.root.after(0, _update)

    def _append_log(self, message):
        def _append():
            self.log_text.insert(tk.END, message + "\n")
            self.log_text.see(tk.END)
        self.root.after(0, _append)

    def _open_output(self):
        folder = self.output_folder.get()
        if folder and os.path.isdir(folder):
            if sys.platform == "win32":
                os.startfile(folder)
            elif sys.platform == "darwin":
                subprocess.Popen(["open", folder])
            else:
                subprocess.Popen(["xdg-open", folder])


# ===========================================================================
# Entry point
# ===========================================================================
def main():
    logging.basicConfig(level=logging.INFO, format=LOG_FORMAT)

    root = tk.Tk()

    # Try to apply a modern theme
    try:
        style = ttk.Style()
        available = style.theme_names()
        for preferred in ("clam", "alt", "vista", "xpnative"):
            if preferred in available:
                style.theme_use(preferred)
                break
    except Exception:
        pass

    app = InvoiceProcessorGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()

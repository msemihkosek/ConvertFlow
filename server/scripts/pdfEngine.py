# ConvertFlow - High-Fidelity Hybrid PDF Engine (PyMuPDF + pdf2docx + python-docx)
import sys
import os
import io
import warnings

warnings.filterwarnings("ignore")

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
if hasattr(sys.stderr, "reconfigure"):
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")

import pymupdf
from pdf2docx import Converter
from docx import Document
from docx.shared import Pt, Inches


def analyze_pdf_type(pdf_path: str) -> tuple[str, int]:
    """Inspects PDF pages to determine whether it has a real text layer or is scanned/image-only."""
    doc = pymupdf.open(pdf_path)
    total_pages = len(doc)
    sample_count = min(total_pages, 15)
    total_chars = 0

    for i in range(sample_count):
        text = doc[i].get_text("text") or ""
        total_chars += len(text.strip())

    doc.close()
    avg_chars = total_chars / max(1, sample_count)
    pdf_type = "native" if avg_chars >= 25 else "scanned"
    return pdf_type, total_pages


def convert_scanned_pdf_to_docx(pdf_path: str, output_path: str, dpi: int = 150, jpg_quality: int = 85) -> None:
    """Renders scanned PDF pages at crisp DPI with PyMuPDF and embeds each page cleanly in Word."""
    pdf_doc = pymupdf.open(pdf_path)
    word_doc = Document()

    for idx, page in enumerate(pdf_doc):
        rect = page.rect
        width_pt = rect.width
        height_pt = rect.height

        if idx == 0:
            section = word_doc.sections[0]
        else:
            section = word_doc.add_section()

        section.page_width = Pt(width_pt)
        section.page_height = Pt(height_pt)
        section.top_margin = Inches(0.1)
        section.bottom_margin = Inches(0.1)
        section.left_margin = Inches(0.1)
        section.right_margin = Inches(0.1)

        pix = page.get_pixmap(dpi=dpi, alpha=False)
        img_bytes = pix.tobytes("jpeg", jpg_quality=jpg_quality)
        img_stream = io.BytesIO(img_bytes)

        paragraph = word_doc.add_paragraph()
        paragraph.paragraph_format.space_before = Pt(0)
        paragraph.paragraph_format.space_after = Pt(0)
        run = paragraph.add_run()
        # Fit slightly inside 0.1in margin on all sides
        usable_w = max(72.0, width_pt - 16.0)
        run.add_picture(img_stream, width=Pt(usable_w))

    pdf_doc.close()
    word_doc.save(output_path)


def convert_pdf_to_docx(pdf_path: str, output_path: str) -> str:
    pdf_type, total_pages = analyze_pdf_type(pdf_path)

    if pdf_type == "native":
        try:
            cv = Converter(pdf_path)
            cv.convert(output_path, start=0, end=None)
            cv.close()
            if os.path.exists(output_path) and os.path.getsize(output_path) > 0:
                return f"native:{total_pages}"
        except Exception as err:
            sys.stderr.write(f"pdf2docx fallback triggered: {err}\n")

    # Scanned PDF or pdf2docx fallback -> adaptive DPI & quality for fast, compact DOCX
    if total_pages > 80:
        dpi, jpg_quality = 110, 75
    elif total_pages > 20:
        dpi, jpg_quality = 140, 82
    else:
        dpi, jpg_quality = 170, 85

    convert_scanned_pdf_to_docx(pdf_path, output_path, dpi=dpi, jpg_quality=jpg_quality)
    return f"scanned:{total_pages}"


def render_pdf_to_images(pdf_path: str, output_dir: str, ext: str = "png", dpi: int = 180) -> int:
    os.makedirs(output_dir, exist_ok=True)
    doc = pymupdf.open(pdf_path)
    total = len(doc)
    pad = max(3, len(str(total)))
    clean_ext = ext.lower().lstrip(".")
    if clean_ext == "jpeg":
        clean_ext = "jpg"

    for i, page in enumerate(doc):
        pix = page.get_pixmap(dpi=dpi, alpha=False)
        filename = f"sayfa_{str(i + 1).zfill(pad)}.{clean_ext}"
        out_file = os.path.join(output_dir, filename)
        if clean_ext == "jpg":
            pix.save(out_file, jpg_quality=90)
        else:
            pix.save(out_file)

    doc.close()
    return total


def extract_pdf_text(pdf_path: str, output_path: str, mode: str = "txt") -> int:
    doc = pymupdf.open(pdf_path)
    total = len(doc)
    base_name = os.path.splitext(os.path.basename(pdf_path))[0]

    if mode == "txt":
        parts = [f"=== {base_name}.pdf ({total} Sayfa) ===\n\n"]
        for i, page in enumerate(doc):
            txt = (page.get_text("text") or "").strip()
            parts.append(f"--- Sayfa {i + 1} ---\n{txt}\n\n")
        with open(output_path, "w", encoding="utf-8-sig") as f:
            f.write("".join(parts))

    elif mode == "md":
        parts = [f"# {base_name}\n\n*ConvertFlow ile dönüştürüldü ({total} sayfa)*\n\n---\n\n"]
        for i, page in enumerate(doc):
            txt = (page.get_text("text") or "").strip()
            parts.append(f"## Sayfa {i + 1}\n\n{txt or '*(Metin katmanı yok)*'}\n\n---\n\n")
        with open(output_path, "w", encoding="utf-8-sig") as f:
            f.write("".join(parts))

    elif mode == "html":
        page_blocks = []
        for i, page in enumerate(doc):
            page_html = page.get_text("html") or "<p><em>(İçerik yok)</em></p>"
            page_blocks.append(f'<article class="cf-page"><div class="cf-tag">Sayfa {i + 1} / {total}</div>{page_html}</article>')

        full_html = f"""<!DOCTYPE html>
<html lang="tr">
<head>
  <meta charset="UTF-8">
  <title>{base_name}</title>
  <style>
    body {{ font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; background: #f1f5f9; color: #0f172a; max-width: 900px; margin: 2rem auto; padding: 0 1rem; }}
    .cf-page {{ background: #fff; border: 1px solid #e2e8f0; border-radius: 10px; padding: 2rem; margin-bottom: 1.5rem; overflow-x: auto; box-shadow: 0 2px 6px rgba(0,0,0,0.04); }}
    .cf-tag {{ font-size: 12px; font-weight: 600; color: #64748b; margin-bottom: 1rem; border-bottom: 1px solid #f1f5f9; padding-bottom: 0.5rem; }}
  </style>
</head>
<body>
  {''.join(page_blocks)}
</body>
</html>"""
        with open(output_path, "w", encoding="utf-8") as f:
            f.write(full_html)

    doc.close()
    return total


def main():
    if len(sys.argv) < 4:
        sys.stderr.write("Usage: pdfEngine.py <mode> <input_pdf> <output_path> [extra]\n")
        sys.exit(1)

    mode = sys.argv[1].lower()
    input_pdf = sys.argv[2]
    output_path = sys.argv[3]

    if mode == "docx":
        info = convert_pdf_to_docx(input_pdf, output_path)
        print(f"OK:{info}")
    elif mode in ("png", "jpg", "jpeg"):
        count = render_pdf_to_images(input_pdf, output_path, ext=mode)
        print(f"OK:images:{count}")
    elif mode in ("txt", "md", "html"):
        count = extract_pdf_text(input_pdf, output_path, mode=mode)
        print(f"OK:{mode}:{count}")
    else:
        sys.stderr.write(f"Unsupported mode: {mode}\n")
        sys.exit(1)


if __name__ == "__main__":
    main()


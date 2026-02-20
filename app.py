from __future__ import annotations

import io
import zipfile
from typing import List, Tuple

from flask import Flask, render_template, request, send_file
from PIL import Image
import fitz  # PyMuPDF
from pypdf import PdfReader, PdfWriter, PageObject
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter

app = Flask(__name__)


def _get_uploaded_pdf(field: str = "pdf") -> PdfReader:
    file = request.files.get(field)
    if not file or file.filename == "":
        raise ValueError("PDF file missing")
    return PdfReader(file.stream)


def _send_pdf_bytes(data: bytes, filename: str):
    return send_file(io.BytesIO(data), as_attachment=True, download_name=filename, mimetype="application/pdf")


def _parse_pages_spec(spec: str, total_pages: int) -> List[int]:
    pages: List[int] = []
    if not spec.strip():
        return list(range(total_pages))

    for chunk in spec.split(","):
        chunk = chunk.strip()
        if not chunk:
            continue
        if "-" in chunk:
            start, end = chunk.split("-", 1)
            s, e = int(start), int(end)
            for p in range(min(s, e), max(s, e) + 1):
                if 1 <= p <= total_pages:
                    pages.append(p - 1)
        else:
            p = int(chunk)
            if 1 <= p <= total_pages:
                pages.append(p - 1)
    deduped = []
    seen = set()
    for p in pages:
        if p not in seen:
            seen.add(p)
            deduped.append(p)
    return deduped


def _build_text_overlay(width: float, height: float, text: str, x: float, y: float, font_size: int = 16):
    packet = io.BytesIO()
    c = canvas.Canvas(packet, pagesize=(width, height))
    c.setFont("Helvetica", font_size)
    c.drawString(x, y, text)
    c.save()
    packet.seek(0)
    return PdfReader(packet)


def _build_image_overlay(width: float, height: float, image_bytes: bytes, x: float, y: float, scale: float = 0.25):
    packet = io.BytesIO()
    c = canvas.Canvas(packet, pagesize=(width, height))
    img = Image.open(io.BytesIO(image_bytes))
    iw, ih = img.size
    draw_w = width * scale
    draw_h = (ih / iw) * draw_w
    c.drawInlineImage(img, x, y, width=draw_w, height=draw_h)
    c.save()
    packet.seek(0)
    return PdfReader(packet)


def _pdfwriter_bytes(writer: PdfWriter) -> bytes:
    out = io.BytesIO()
    writer.write(out)
    out.seek(0)
    return out.read()


def _compress_pdf_to_target(pdf_bytes: bytes, target_kb: int) -> bytes:
    # Rasterize pages and rebuild; binary search DPI/quality to approach target size.
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    low, high = 40, 170
    best = None

    def render(dpi: int) -> bytes:
        images: List[bytes] = []
        for page in doc:
            pix = page.get_pixmap(dpi=dpi, alpha=False)
            images.append(pix.tobytes("jpg", jpg_quality=70))
        writer = PdfWriter()
        for jpg in images:
            img = Image.open(io.BytesIO(jpg)).convert("RGB")
            bio = io.BytesIO()
            img.save(bio, format="PDF")
            part = PdfReader(io.BytesIO(bio.getvalue()))
            writer.add_page(part.pages[0])
        return _pdfwriter_bytes(writer)

    for _ in range(7):
        mid = (low + high) // 2
        candidate = render(mid)
        size_kb = len(candidate) / 1024
        if size_kb <= target_kb:
            best = candidate
            low = mid + 1
        else:
            high = mid - 1

    if best is None:
        best = render(max(35, high))
    doc.close()
    return best


@app.route("/")
def index():
    return render_template("index.html")


@app.post("/merge")
def merge_pdf():
    writer = PdfWriter()
    files = request.files.getlist("pdfs")
    if not files:
        return "At least one PDF required", 400
    for f in files:
        reader = PdfReader(f.stream)
        for p in reader.pages:
            writer.add_page(p)
    return _send_pdf_bytes(_pdfwriter_bytes(writer), "merged.pdf")


@app.post("/split")
def split_pdf():
    reader = _get_uploaded_pdf()
    ranges = request.form.get("pages", "")
    selected = _parse_pages_spec(ranges, len(reader.pages))
    if not selected:
        return "No valid pages selected", 400

    mem = io.BytesIO()
    with zipfile.ZipFile(mem, "w", zipfile.ZIP_DEFLATED) as zf:
        for idx in selected:
            writer = PdfWriter()
            writer.add_page(reader.pages[idx])
            zf.writestr(f"page_{idx + 1}.pdf", _pdfwriter_bytes(writer))
    mem.seek(0)
    return send_file(mem, as_attachment=True, download_name="split_pages.zip", mimetype="application/zip")


@app.post("/rotate")
def rotate_pdf():
    reader = _get_uploaded_pdf()
    angle = int(request.form.get("angle", "90"))
    writer = PdfWriter()
    for page in reader.pages:
        page.rotate(angle)
        writer.add_page(page)
    return _send_pdf_bytes(_pdfwriter_bytes(writer), "rotated.pdf")


@app.post("/page_numbers")
def page_numbers():
    reader = _get_uploaded_pdf()
    writer = PdfWriter()
    position = request.form.get("position", "bottom")
    for i, page in enumerate(reader.pages, start=1):
        w = float(page.mediabox.width)
        h = float(page.mediabox.height)
        y = 20 if position == "bottom" else h - 30
        overlay = _build_text_overlay(w, h, str(i), w - 40, y)
        page.merge_page(overlay.pages[0])
        writer.add_page(page)
    return _send_pdf_bytes(_pdfwriter_bytes(writer), "page_numbered.pdf")


@app.post("/watermark")
def watermark_pdf():
    reader = _get_uploaded_pdf()
    mode = request.form.get("mode", "text")
    writer = PdfWriter()
    image = request.files.get("image")
    text = request.form.get("text", "WATERMARK")

    image_bytes = image.read() if image and image.filename else None

    for page in reader.pages:
        w = float(page.mediabox.width)
        h = float(page.mediabox.height)
        if mode == "image" and image_bytes:
            overlay = _build_image_overlay(w, h, image_bytes, x=20, y=20)
        else:
            overlay = _build_text_overlay(w, h, text, x=30, y=h / 2, font_size=28)
        page.merge_page(overlay.pages[0])
        writer.add_page(page)
    return _send_pdf_bytes(_pdfwriter_bytes(writer), "watermarked.pdf")


@app.post("/protect")
def protect_pdf():
    reader = _get_uploaded_pdf()
    password = request.form.get("password", "")
    if not password:
        return "Password required", 400
    writer = PdfWriter()
    for p in reader.pages:
        writer.add_page(p)
    writer.encrypt(password)
    return _send_pdf_bytes(_pdfwriter_bytes(writer), "protected.pdf")


@app.post("/unlock")
def unlock_pdf():
    file = request.files.get("pdf")
    password = request.form.get("password", "")
    if not file:
        return "PDF missing", 400
    reader = PdfReader(file.stream)
    if reader.is_encrypted:
        if reader.decrypt(password) == 0:
            return "Wrong password", 400
    writer = PdfWriter()
    for p in reader.pages:
        writer.add_page(p)
    return _send_pdf_bytes(_pdfwriter_bytes(writer), "unlocked.pdf")


@app.post("/organize")
def organize_pdf():
    reader = _get_uploaded_pdf()
    order = request.form.get("order", "")
    selected = _parse_pages_spec(order, len(reader.pages))
    if not selected:
        return "Invalid page order", 400
    writer = PdfWriter()
    for idx in selected:
        writer.add_page(reader.pages[idx])
    return _send_pdf_bytes(_pdfwriter_bytes(writer), "organized.pdf")


@app.post("/image_to_pdf")
def image_to_pdf():
    files = request.files.getlist("images")
    if not files:
        return "Select images", 400
    pil_images = [Image.open(f.stream).convert("RGB") for f in files]
    out = io.BytesIO()
    first, rest = pil_images[0], pil_images[1:]
    first.save(out, format="PDF", save_all=True, append_images=rest)
    out.seek(0)
    return send_file(out, as_attachment=True, download_name="images_to_pdf.pdf", mimetype="application/pdf")


@app.post("/pdf_to_image")
def pdf_to_image():
    file = request.files.get("pdf")
    if not file:
        return "PDF missing", 400
    doc = fitz.open(stream=file.read(), filetype="pdf")
    mem = io.BytesIO()
    with zipfile.ZipFile(mem, "w", zipfile.ZIP_DEFLATED) as zf:
        for i, page in enumerate(doc, start=1):
            pix = page.get_pixmap(dpi=200)
            zf.writestr(f"page_{i}.png", pix.tobytes("png"))
    mem.seek(0)
    return send_file(mem, as_attachment=True, download_name="pdf_images.zip", mimetype="application/zip")


@app.post("/crop")
def crop_pdf():
    reader = _get_uploaded_pdf()
    margin = float(request.form.get("margin", "10"))
    writer = PdfWriter()
    for page in reader.pages:
        left = float(page.mediabox.left) + margin
        bottom = float(page.mediabox.bottom) + margin
        right = float(page.mediabox.right) - margin
        top = float(page.mediabox.top) - margin
        page.cropbox.lower_left = (left, bottom)
        page.cropbox.upper_right = (right, top)
        writer.add_page(page)
    return _send_pdf_bytes(_pdfwriter_bytes(writer), "cropped.pdf")


@app.post("/compress")
def compress_pdf():
    file = request.files.get("pdf")
    if not file:
        return "PDF missing", 400
    preset = request.form.get("target_kb", "100")
    manual = request.form.get("manual_kb", "").strip()
    target_kb = int(manual) if manual else int(preset)
    data = file.read()
    compressed = _compress_pdf_to_target(data, target_kb)
    return _send_pdf_bytes(compressed, f"compressed_{target_kb}kb.pdf")


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)

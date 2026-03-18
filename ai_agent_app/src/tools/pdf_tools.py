"""
PDF automation tools.
"""
import os


def _workspace_output_dir() -> str:
    output_dir = os.path.join(os.getcwd(), "output")
    os.makedirs(output_dir, exist_ok=True)
    return output_dir


def _normalize_output_folder(output_folder: str, suffix: str) -> str:
    if not output_folder:
        output_folder = os.path.join(_workspace_output_dir(), suffix)
    os.makedirs(output_folder, exist_ok=True)
    return output_folder


def _normalize_output_file(path: str, default_name: str, extension: str) -> str:
    if not path:
        return os.path.join(_workspace_output_dir(), default_name)

    abs_path = os.path.abspath(path)
    if os.path.isdir(abs_path):
        return os.path.join(abs_path, default_name)
    if not abs_path.lower().endswith(extension.lower()):
        return abs_path + extension
    return abs_path


def _ocr_page_image(page, matrix) -> str:
    try:
        import pytesseract
        from PIL import Image
    except ImportError:
        return ""

    pix = page.get_pixmap(matrix=matrix, alpha=False)
    image = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    return pytesseract.image_to_string(image).strip()


def pdf_extract_images(pdf_path: str, output_folder: str, page_number: int = 0) -> str:
    """
    Extract embedded images from a PDF.

    If a page has no embedded images, save a full-page render as fallback so scanned
    PDFs still produce a usable image output.
    """
    try:
        import fitz  # PyMuPDF

        if not pdf_path:
            return "ERROR: Please select a PDF file first."
        if not os.path.exists(pdf_path):
            return f"ERROR: File not found: {pdf_path}"

        output_folder = _normalize_output_folder(output_folder, "pdf_images")
        doc = fitz.open(pdf_path)
        matrix = fitz.Matrix(2, 2)

        if page_number > 0:
            page_indexes = [page_number - 1]
        else:
            page_indexes = list(range(len(doc)))

        saved = []
        fallback_pages = []

        for page_index in page_indexes:
            if page_index < 0 or page_index >= len(doc):
                return f"ERROR: Page {page_number} out of range. PDF has {len(doc)} page(s)."

            page = doc[page_index]
            image_list = page.get_images(full=True)

            if image_list:
                for img_index, img in enumerate(image_list, start=1):
                    xref = img[0]
                    extracted = doc.extract_image(xref)
                    image_bytes = extracted["image"]
                    image_ext = extracted.get("ext", "png")
                    out_path = os.path.join(
                        output_folder,
                        f"page{page_index + 1}_img{img_index}.{image_ext}",
                    )
                    with open(out_path, "wb") as handle:
                        handle.write(image_bytes)
                    saved.append(out_path)
                continue

            pixmap = page.get_pixmap(matrix=matrix, alpha=False)
            fallback_path = os.path.join(output_folder, f"page{page_index + 1}_render.png")
            pixmap.save(fallback_path)
            saved.append(fallback_path)
            fallback_pages.append(page_index + 1)

        if not saved:
            return f"WARNING: No images found in PDF ({'page ' + str(page_number) if page_number else 'all pages'})"

        message = (
            f"SUCCESS: Extracted {len(saved)} image file(s) to: {output_folder}\n"
            f"Files: {', '.join(os.path.basename(path) for path in saved[:12])}"
        )
        if len(saved) > 12:
            message += f" ... and {len(saved) - 12} more"
        if fallback_pages:
            message += (
                f"\nNote: Pages {', '.join(str(page) for page in fallback_pages)} "
                "had no embedded images, so full-page render(s) were saved."
            )
        return message
    except ImportError:
        return "ERROR: PyMuPDF not installed. Run: pip install pymupdf"
    except Exception as exc:
        return f"ERROR: Error extracting PDF images: {exc}"


def pdf_extract_text(pdf_path: str, page_number: int = 0) -> str:
    """Extract text from PDF, with OCR fallback for scanned pages."""
    try:
        import fitz  # PyMuPDF

        if not pdf_path:
            return "ERROR: Please select a PDF file first."
        if not os.path.exists(pdf_path):
            return f"ERROR: File not found: {pdf_path}"

        doc = fitz.open(pdf_path)
        matrix = fitz.Matrix(2, 2)

        if page_number > 0:
            page_indexes = [page_number - 1]
        else:
            page_indexes = list(range(len(doc)))

        result = []
        ocr_pages = []

        for page_index in page_indexes:
            if page_index < 0 or page_index >= len(doc):
                return f"ERROR: Page {page_number} out of range. PDF has {len(doc)} page(s)."

            page = doc[page_index]
            text = page.get_text("text").strip()
            if not text:
                text = _ocr_page_image(page, matrix)
                if text:
                    ocr_pages.append(page_index + 1)
            result.append(f"--- Page {page_index + 1} ---\n{text}")

        full = "\n\n".join(chunk for chunk in result if chunk.strip())
        if not full.strip():
            return "WARNING: No readable text found in the PDF."

        preview = full[:3000] + ("..." if len(full) > 3000 else "")
        message = f"SUCCESS: Extracted text from {pdf_path}\n\n{preview}"
        if ocr_pages:
            message += f"\n\nNote: OCR fallback was used for page(s): {', '.join(str(page) for page in ocr_pages)}"
        return message
    except ImportError:
        return "ERROR: PyMuPDF not installed. Run: pip install pymupdf"
    except Exception as exc:
        return f"ERROR: Error extracting PDF text: {exc}"


def pdf_word_to_pdf(docx_path: str, pdf_path: str) -> str:
    """Convert a Word file to PDF using docx2pdf first, then LibreOffice fallback."""
    try:
        if not docx_path:
            return "ERROR: Please select a Word file first."
        if not os.path.exists(docx_path):
            return f"ERROR: File not found: {docx_path}"

        output_pdf = _normalize_output_file(
            pdf_path,
            f"{os.path.splitext(os.path.basename(docx_path))[0]}.pdf",
            ".pdf",
        )
        os.makedirs(os.path.dirname(output_pdf), exist_ok=True)

        try:
            from docx2pdf import convert

            convert(docx_path, output_pdf)
            if os.path.exists(output_pdf):
                return f"SUCCESS: Converted {docx_path} to {output_pdf}"
        except Exception:
            pass

        import subprocess

        out_dir = os.path.dirname(output_pdf)
        result = subprocess.run(
            ["soffice", "--headless", "--convert-to", "pdf", "--outdir", out_dir, docx_path],
            capture_output=True,
            text=True,
            timeout=30,
        )
        if result.returncode == 0:
            produced_pdf = os.path.join(out_dir, f"{os.path.splitext(os.path.basename(docx_path))[0]}.pdf")
            if produced_pdf != output_pdf and os.path.exists(produced_pdf):
                os.replace(produced_pdf, output_pdf)
            return f"SUCCESS: Converted {docx_path} to {output_pdf}"
        return "ERROR: Could not convert Word to PDF. docx2pdf/LibreOffice both failed."
    except Exception as exc:
        return f"ERROR: Error converting to PDF: {exc}"

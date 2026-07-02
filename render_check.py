"""
Convert a PPTX to PDF via headless LibreOffice, then render each PDF page to
a PNG for visual inspection. Usage: python3 render_check.py <pptx_path>
"""
import subprocess
import sys
import os

def convert_to_pdf(pptx_path, out_dir):
    subprocess.run(
        ["soffice", "--headless", "--convert-to", "pdf", "--outdir", out_dir, pptx_path],
        check=True, timeout=120,
    )
    base = os.path.splitext(os.path.basename(pptx_path))[0]
    return os.path.join(out_dir, base + ".pdf")


def render_pages(pdf_path, out_dir, prefix):
    import fitz
    doc = fitz.open(pdf_path)
    paths = []
    for i, page in enumerate(doc):
        pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
        p = os.path.join(out_dir, f"{prefix}_slide{i+1}.png")
        pix.save(p)
        paths.append(p)
    return paths


if __name__ == "__main__":
    pptx_path = sys.argv[1]
    out_dir = sys.argv[2] if len(sys.argv) > 2 else os.path.dirname(pptx_path)
    pdf_path = convert_to_pdf(pptx_path, out_dir)
    print(f"PDF: {pdf_path}")
    prefix = os.path.splitext(os.path.basename(pptx_path))[0]
    imgs = render_pages(pdf_path, out_dir, prefix)
    for p in imgs:
        print(f"IMG: {p}")

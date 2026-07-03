"""Render each PDF page to a PNG for visual QC (Section 6.6)."""
import sys
import fitz  # pymupdf

def render(pdf_path, out_prefix, zoom=2.0):
    doc = fitz.open(pdf_path)
    mat = fitz.Matrix(zoom, zoom)
    paths = []
    for i, page in enumerate(doc):
        pix = page.get_pixmap(matrix=mat)
        out_path = f"{out_prefix}_slide{i+1}.png"
        pix.save(out_path)
        paths.append(out_path)
        print(out_path)
    return paths

if __name__ == "__main__":
    render(sys.argv[1], sys.argv[2])

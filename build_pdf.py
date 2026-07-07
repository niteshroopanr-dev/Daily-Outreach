"""
ProfitPulse Brief PDF export.

Converts the just built PPTX to PDF via headless LibreOffice, so the PDF is a
direct render of the same deck rather than a hand maintained reimplementation
that can drift from it (Section 6.5 requires the PDF to match the PPTX).
Usage: python3 build_pdf.py "Out-reach efforts/Brief_CompanyName_DDMonYYYY.pptx"
"""

import os
import subprocess
import sys

os.environ.setdefault("HOME", "/tmp/lo_home")
os.makedirs(os.environ["HOME"], exist_ok=True)


def convert(pptx_path):
    out_dir = os.path.dirname(os.path.abspath(pptx_path))
    result = subprocess.run(
        ["soffice", "--headless", "--convert-to", "pdf", "--outdir", out_dir, pptx_path],
        capture_output=True, text=True,
    )
    print(result.stdout)
    if result.returncode != 0:
        print(result.stderr, file=sys.stderr)
        raise SystemExit(result.returncode)
    pdf_path = os.path.splitext(pptx_path)[0] + ".pdf"
    print(f"PDF saved: {pdf_path}")
    return pdf_path


if __name__ == "__main__":
    if len(sys.argv) != 2:
        raise SystemExit("Usage: python3 build_pdf.py <path-to-pptx>")
    convert(sys.argv[1])

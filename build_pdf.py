"""
ProfitPulse Brief PDF Exporter
Exports the PPTX built by build_brief.py straight to PDF via LibreOffice, so the
PDF is a direct export of the same deck rather than a separately drawn document.
Usage: python3 build_pdf.py <path/to/Brief_X.pptx>
"""

import subprocess
import sys
import tempfile
from pathlib import Path


def export_pdf(pptx_path: str) -> str:
    pptx_path = Path(pptx_path).resolve()
    out_dir = pptx_path.parent
    with tempfile.TemporaryDirectory() as profile_dir:
        subprocess.run(
            [
                "soffice", "--headless",
                f"-env:UserInstallation=file://{profile_dir}",
                "--convert-to", "pdf",
                "--outdir", str(out_dir),
                str(pptx_path),
            ],
            check=True,
        )
    pdf_path = out_dir / (pptx_path.stem + ".pdf")
    print(f"PDF saved: {pdf_path}")
    return str(pdf_path)


if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python3 build_pdf.py <path/to/Brief_X.pptx>")
        sys.exit(1)
    export_pdf(sys.argv[1])

"""Render the presenter guide (handover/GIAI-THICH-QUY-TRINH.md) to PDF for sharing with the group.

Usage:  python -m pipeline.build_guide_pdf

Markdown → HTML (Python-Markdown) → PDF printed by Chrome (Playwright): A4, the dashboard's own fonts
(full Vietnamese support), clickable links, page numbers. Run it after every edit of the .md so the
two stay identical.
"""
import re
import sys
import tempfile
from pathlib import Path

import markdown

ROOT = Path(__file__).resolve().parent.parent
SRC = ROOT / "handover" / "GIAI-THICH-QUY-TRINH.md"
OUT = SRC.with_suffix(".pdf")
ASSETS = ROOT / "site" / "assets"

PAGE_CSS = """
@page { size: A4; margin: 16mm 16mm 18mm; }
body { font-family: "Source Sans 3", Arial, sans-serif; font-size: 10.5pt; line-height: 1.5; color: #15191e; }
h1 { font-family: "Source Serif 4", Georgia, serif; font-weight: 600; font-size: 20pt; line-height: 1.25; margin: 0 0 10pt; }
h2 { font-size: 13.5pt; font-weight: 700; margin: 18pt 0 6pt; padding-top: 6pt; border-top: 1px solid #d5dae1; break-after: avoid; }
p, li { margin: 0 0 6pt; }
ul, ol { padding-left: 16pt; margin: 0 0 8pt; }
table { border-collapse: collapse; width: 100%; margin: 6pt 0 10pt; font-size: 9.5pt; }
th, td { border: 1px solid #d5dae1; padding: 4pt 6pt; text-align: left; vertical-align: top; }
th { background: #f3f5f7; font-weight: 600; }
tr { break-inside: avoid; }
code { font-family: Consolas, "Courier New", monospace; font-size: 9pt; background: #f3f5f7; padding: 0 2pt; border-radius: 2pt; }
pre { background: #f3f5f7; padding: 6pt 8pt; border-radius: 3pt; white-space: pre-wrap; font-size: 9pt; }
pre code { background: none; padding: 0; }
a { color: #1d4e89; }
"""
FOOTER = ('<div style="width:100%;text-align:center;font:8pt Arial,sans-serif;color:#5b6472">'
          'Trang <span class="pageNumber"></span>/<span class="totalPages"></span></div>')


def to_html(md_text):
    # Bare URLs become autolinks so they stay clickable in the PDF.
    md_text = re.sub(r"(?<![<(\"])(https?://[^\s<>()|]+?)([.,;]?)(?=\s|$)", r"<\1>\2", md_text)
    body = markdown.markdown(md_text, extensions=["tables", "fenced_code", "sane_lists"])
    title = re.search(r"^# (.+)$", md_text, re.M).group(1)
    # Same @font-face rules as the dashboard; <base> makes their relative font paths resolve.
    fonts = "\n".join(line for line in (ASSETS / "styles.css").read_text(encoding="utf-8").splitlines()
                      if line.startswith("@font-face"))
    return (f'<!doctype html><html lang="vi"><head><meta charset="utf-8"><title>{title}</title>'
            f'<base href="{ASSETS.as_uri()}/"><style>{fonts}\n{PAGE_CSS}</style></head><body>{body}</body></html>')


def main():
    from playwright.sync_api import sync_playwright
    html = to_html(SRC.read_text(encoding="utf-8"))
    with tempfile.TemporaryDirectory() as tmp, sync_playwright() as p:
        page_file = Path(tmp) / "guide.html"
        page_file.write_text(html, encoding="utf-8")
        try:
            browser = p.chromium.launch(channel="chrome")
        except Exception:
            browser = p.chromium.launch()
        page = browser.new_page()
        page.goto(page_file.as_uri())
        page.evaluate("document.fonts.ready")
        page.pdf(path=str(OUT), prefer_css_page_size=True, print_background=True, display_header_footer=True,
                 header_template="<div></div>", footer_template=FOOTER)
        browser.close()
    print(f"Guide PDF written: {OUT.relative_to(ROOT)}")
    return 0


if __name__ == "__main__":
    sys.exit(main())

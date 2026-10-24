"""Re-read the MD&A-only figures directly on the rendered 10-K pages, a second way.

Usage:  python -m pipeline.recheck_pages                # read the live pages
        python -m pipeline.recheck_pages --from-saved   # re-evaluate the saved readings only

The currency-neutral growth rates and NIKE Brand Digital sales have no XBRL tag and no
second official source, so extraction alone leaves them unconfirmed. This step opens each
10-K in a browser (Playwright + Chrome), finds the printed page by its footer, and reads:
  * the four "% change excluding currency changes" values by column POSITION under the
    first header of that name (the extractor reads cells by index, so a wrong column or a
    shifted cell would show up as a mismatch here), and
  * the Digital amount from the sentence that follows the words "Digital sales".
It highlights what it read, saves screenshots and an evidence page in
handover/bang-chung-doi-chieu/, and writes data/extracted/page_check.csv, which the
workbook builder uses to mark matching figures as confirmed. Exits 1 on any mismatch.
"""
import csv
import datetime
import html
import re
import sys
from decimal import Decimal
from pathlib import Path

from . import dataset
from .sources import FILINGS, filing, ir_document_url, sec_document_url

ROOT = Path(__file__).resolve().parent.parent
OUT_CSV = ROOT / "data" / "extracted" / "page_check.csv"
EVIDENCE = ROOT / "handover" / "bang-chung-doi-chieu"
CN_ROWS = [("revenues_pct_change_cn", "TOTAL NIKE, INC. REVENUES", "% thay đổi doanh thu NIKE, Inc. loại trừ tỷ giá"),
           ("nike_brand_pct_change_cn", "TOTAL NIKE BRAND REVENUES", "% thay đổi NIKE Brand loại trừ tỷ giá"),
           ("nike_direct_pct_change_cn", "SALES THROUGH NIKE DIRECT", "% thay đổi NIKE Direct loại trừ tỷ giá"),
           ("nike_brand_wholesale_pct_change_cn", "SALES TO WHOLESALE CUSTOMERS", "% thay đổi bán buôn loại trừ tỷ giá")]
DIGITAL = ("nike_brand_digital_sales", "Doanh số NIKE Brand Digital (tỷ USD)")

# Shared by both page scripts: the printed page `page` runs from the footer of page-1 to its own footer.
JS_PAGE = r"""
  const norm = s => s.replace(/\s+/g, ' ').trim();
  const FOOT = /^\d{4} FORM 10-K\s*\d+$/;
  const footers = Array.from(document.querySelectorAll('div, p, td, span'))
    .filter(e => FOOT.test(norm(e.textContent)) && !Array.from(e.children).some(c => FOOT.test(norm(c.textContent))));
  const num = e => +norm(e.textContent).match(/(\d+)$/)[1];
  const end = footers.find(e => num(e) === page), start = footers.find(e => num(e) === page - 1);
  if (!end) return {error: 'page footer not found'};
  const inPage = e => (e.compareDocumentPosition(end) & Node.DOCUMENT_POSITION_FOLLOWING) &&
                      (!start || (start.compareDocumentPosition(e) & Node.DOCUMENT_POSITION_FOLLOWING));
"""
JS_TABLE = r"""([page, labels]) => {""" + JS_PAGE + r"""
  const t = Array.from(document.querySelectorAll('table')).filter(inPage)
    .find(t => /SALES THROUGH NIKE DIRECT/i.test(t.textContent) && /EXCLUDING\s+CURRENCY/i.test(t.textContent));
  if (!t) return {error: 'revenue table not found on the page'};
  const head = Array.from(t.querySelectorAll('td, th')).filter(c => /EXCLUDING\s+CURRENCY/i.test(norm(c.textContent)))
    .map(c => c.getBoundingClientRect()).filter(r => r.width > 0).sort((a, b) => a.left - b.left)[0];
  const values = {};
  for (const [id, label] of labels) {
    const row = Array.from(t.querySelectorAll('tr')).find(r => {
      const first = Array.from(r.querySelectorAll('td')).map(c => norm(c.textContent)).find(x => x);
      return first && first.replace(/\(\d+\)/g, '').trim().toUpperCase() === label;
    });
    if (!row) { values[id] = null; continue; }
    const cells = Array.from(row.querySelectorAll('td')).filter(c => {
      const r = c.getBoundingClientRect();
      return r.width > 0 && r.right > head.left + 2 && r.left < head.right - 2;
    });
    cells.forEach(c => { c.style.background = '#ffe066'; });
    row.style.outline = '2px solid #d92d20';
    values[id] = norm(cells.map(c => c.textContent).join(' '));
  }
  t.setAttribute('data-recheck', 'table');
  return {values};
}"""
JS_DIGITAL = r"""([page]) => {""" + JS_PAGE + r"""
  const b = Array.from(document.querySelectorAll('div, p')).filter(inPage)
    .find(e => /Digital (commerce )?sales/i.test(e.textContent) && /billion/.test(e.textContent) && !e.querySelector('div, p, table'));
  if (!b) return {error: 'Digital sentence not found on the page'};
  b.style.background = '#ffe066';
  b.style.outline = '2px solid #d92d20';
  b.setAttribute('data-recheck', 'digital');
  return {text: norm(b.textContent)};
}"""


def read_pct(text):
    # Some filings put the "%" sign in a cell of its own, outside the header's column span.
    m = re.search(r"(\(?-?\d+\)?)\s*%?", text or "")
    if not m:
        return None
    s = m.group(1)
    return Decimal(-int(s.strip("()-")) if s.startswith(("(", "-")) else int(s)) / 100


def read_digital(text, fy):
    """Amount for fiscal `fy` in the clause that follows the words 'Digital sales'."""
    for m in re.finditer(r"Digital (?:commerce )?sales", text):
        clause = text[m.end():m.end() + 220]
        for amount, year in re.findall(r"\$(\d+\.\d) billion (?:for|in) fiscal (\d{4})", clause):
            if int(year) == fy:
                return Decimal(amount)
    return None


def open_filing(pg, f):
    pg.goto(ir_document_url(f), wait_until="domcontentloaded", timeout=180000)
    # Poll until the printed page footers are in the document (the IR site keeps the network busy,
    # so neither "networkidle" nor animation-frame polling is reliable here).
    for _ in range(60):
        if pg.evaluate("() => /FORM 10-K\\s*\\d+/.test(document.body ? document.body.innerText : '')"):
            return
        pg.wait_for_timeout(2000)
    raise RuntimeError(f"{ir_document_url(f)}: page footers never appeared")


def check_year(pg, fy, facts):
    f = filing(fy)
    rows = []
    get = lambda mid, method: facts.get((fy, mid, fy, method))
    tpage = get("nike_direct_pct_change_cn", "mdna_table")["page"]
    dpage = get(DIGITAL[0], "mdna_text")["page"]
    open_filing(pg, f)
    r = pg.evaluate(JS_TABLE, [tpage, [[mid, label] for mid, label, _ in CN_ROWS]])
    shot = f"FY{fy}-bang-doanh-thu-tr{tpage}.png"
    if "values" in r:
        pg.locator("[data-recheck=table]").screenshot(path=str(EVIDENCE / shot))
    for mid, _label, vi in CN_ROWS:
        want = get(mid, "mdna_table")["value"]
        text = (r.get("values") or {}).get(mid)
        got = read_pct(text)
        rows.append(dict(filing_fy=fy, metric_id=mid, label=vi, page=tpage, read_text=text or r.get("error", ""),
                         read_value=got, extracted_value=want, result="match" if got == want else "MISMATCH",
                         screenshot=shot if "values" in r else ""))
    r = pg.evaluate(JS_DIGITAL, [dpage])
    shot = f"FY{fy}-digital-tr{dpage}.png"
    if "text" in r:
        pg.locator("[data-recheck=digital]").screenshot(path=str(EVIDENCE / shot))
    want = get(DIGITAL[0], "mdna_text")["value"]
    got = read_digital(r.get("text", ""), fy)
    rows.append(dict(filing_fy=fy, metric_id=DIGITAL[0], label=DIGITAL[1], page=dpage, read_text=r.get("text", r.get("error", ""))[:300],
                     read_value=got, extracted_value=want, result="match" if got == want else "MISMATCH",
                     screenshot=shot if "text" in r else ""))
    return rows


def show(mid, v):
    if v is None:
        return "—"
    if mid == DIGITAL[0]:
        return f"{v:.1f}".replace(".", ",")
    p = int(v * 100)
    return f"{'+' if p > 0 else '−' if p < 0 else ''}{abs(p)}%"


def write_evidence_page(rows):
    ok = sum(r["result"] == "match" for r in rows)
    y, m, d = max(r["checked_on"] for r in rows).split("-")
    out = ["<!doctype html><html lang='vi'><meta charset='utf-8'><title>Đối chiếu 35 số trên trang 10-K</title>",
           "<style>body{font:15px/1.55 Segoe UI,Arial,sans-serif;margin:28px auto;max-width:1400px;padding:0 16px;color:#15191e}"
           "table{border-collapse:collapse;margin:8px 0 14px}td,th{border-bottom:1px solid #ddd;padding:5px 12px;text-align:left}"
           "th{background:#f4f5f7}.ok{color:#0a7a3d;font-weight:600}.bad{color:#b42318;font-weight:700}"
           "img{max-width:100%;border:1px solid #ccc;margin:6px 0 20px}h2{margin-top:40px;border-top:2px solid #15191e;padding-top:12px}</style>",
           "<h1>Đối chiếu 35 số chỉ có trong phần MD&amp;A với trang 10-K</h1>",
           f"<p><b>{ok}/{len(rows)} số khớp.</b> Kiểm tra ngày {d}/{m}/{y} bằng trình duyệt Chrome điều khiển tự động (Playwright). "
           "Cách làm: mở từng báo cáo 10-K, tìm đúng trang theo chân trang “FORM 10-K số trang”, rồi đọc lại số trên trang đã hiển thị. "
           "Bốn tỷ lệ % đọc theo vị trí cột “% CHANGE EXCLUDING CURRENCY CHANGES” <i>đầu tiên</i> (bảng có hai cột cùng tên; cột thứ "
           "hai là của năm trước). Doanh số Digital đọc trong câu ngay sau cụm “Digital sales” (FY2020–2021: “Digital commerce sales”). "
           "Ô và đoạn văn được đọc tô nền vàng, dòng được đọc viền đỏ.</p>",
           "<p>Ảnh chụp từ bản sao 10-K trên NIKE Investor Relations, cùng số accession và cùng nội dung với bản trên SEC EDGAR "
           "(liên kết SEC ở từng năm).</p>"]
    for fy, *_ in FILINGS:
        yr = [r for r in rows if r["filing_fy"] == fy]
        out.append(f"<h2>FY{fy}</h2><p><a href='{sec_document_url(filing(fy))}'>Mở 10-K FY{fy} trên SEC EDGAR</a></p>"
                   "<table><tr><th>Chỉ tiêu</th><th>Trang</th><th>Số trong bộ dữ liệu</th><th>Đọc lại trên trang</th><th>Kết quả</th></tr>")
        for r in yr:
            good = r["result"] == "match"
            out.append(f"<tr><td>{html.escape(r['label'])}</td><td>{r['page']}</td><td>{show(r['metric_id'], r['extracted_value'])}</td>"
                       f"<td>{show(r['metric_id'], r['read_value'])}</td><td class='{'ok' if good else 'bad'}'>{'Khớp' if good else 'KHÔNG KHỚP'}</td></tr>")
        out.append("</table>")
        for shot in dict.fromkeys(r["screenshot"] for r in yr if r["screenshot"]):
            out.append(f"<img src='{html.escape(shot)}' alt='{html.escape(shot)}'>")
    (EVIDENCE / "doi-chieu-35-so.html").write_text("\n".join(out) + "\n", encoding="utf-8")


def read_live(facts):
    from playwright.sync_api import sync_playwright
    rows = []
    with sync_playwright() as p:
        try:
            browser = p.chromium.launch(channel="chrome")
        except Exception:
            browser = p.chromium.launch()
        pg = browser.new_page(viewport={"width": 1400, "height": 1100})
        for fy, *_ in FILINGS:
            year_rows = check_year(pg, fy, facts)
            if any(r["result"] != "match" for r in year_rows):   # one retry: the viewer occasionally renders late
                year_rows = check_year(pg, fy, facts)
            rows += year_rows
            print(f"FY{fy}: " + ", ".join(f"{r['metric_id']}={r['read_value']} {r['result']}" for r in year_rows))
            pg.wait_for_timeout(8000)   # be gentle: the IR site blocks rapid repeated requests
        browser.close()
    return rows


def read_saved(facts):
    """Re-evaluate the readings stored in page_check.csv (no download), e.g. after fixing a parser."""
    labels = {mid: vi for mid, _, vi in CN_ROWS} | {DIGITAL[0]: DIGITAL[1]}
    rows = []
    with open(OUT_CSV, encoding="utf-8") as fh:
        for r in csv.DictReader(fh):
            fy, mid = int(r["filing_fy"]), r["metric_id"]
            method = "mdna_text" if mid == DIGITAL[0] else "mdna_table"
            want = facts[(fy, mid, fy, method)]["value"]
            if mid == DIGITAL[0]:
                got = Decimal(r["read_value"]) if r["read_value"] else read_digital(r["read_text"], fy)
            else:
                got = read_pct(r["read_text"])
            rows.append(dict(filing_fy=fy, metric_id=mid, label=labels[mid], page=int(r["page"]), read_text=r["read_text"],
                             read_value=got, extracted_value=want, result="match" if got == want else "MISMATCH",
                             screenshot=r["screenshot"], checked_on=r.get("checked_on") or datetime.date.today().isoformat()))
    return rows


def main():
    facts = dataset.load_facts()
    EVIDENCE.mkdir(parents=True, exist_ok=True)
    rows = read_saved(facts) if "--from-saved" in sys.argv else read_live(facts)
    for r in rows:
        r.setdefault("checked_on", datetime.date.today().isoformat())
    fields = ["filing_fy", "metric_id", "page", "read_value", "extracted_value", "result", "checked_on", "screenshot", "read_text"]
    with open(OUT_CSV, "w", newline="", encoding="utf-8") as fh:
        w = csv.DictWriter(fh, fieldnames=fields, extrasaction="ignore")
        w.writeheader()
        for r in rows:
            w.writerow({k: ("" if r[k] is None else r[k]) for k in fields})
    write_evidence_page(rows)
    bad = [r for r in rows if r["result"] != "match"]
    print(f"{len(rows) - len(bad)}/{len(rows)} figures match the rendered 10-K pages -> {OUT_CSV.relative_to(ROOT)}, "
          f"{EVIDENCE.relative_to(ROOT)}")
    return 1 if bad else 0


if __name__ == "__main__":
    sys.exit(main())

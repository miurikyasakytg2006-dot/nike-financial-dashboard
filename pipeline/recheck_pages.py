"""Re-read the MD&A-only figures that no other source confirms, a second way, on the rendered report pages.

Usage:  python -m pipeline.recheck_pages                # read the pages
        python -m pipeline.recheck_pages --from-saved   # re-evaluate the saved readings only

The quarterly "% change excluding currency changes" figures have no XBRL tag, and a later report does not print
them again, so extraction alone leaves them unconfirmed. This step opens each quarterly report (the copy downloaded
from NIKE Investor Relations, rendered by Chrome through Playwright), finds the printed page by its page number,
and reads the four values of the three-month column by POSITION under the first "% change excluding currency
changes" header of the revenue table. The extractor reads cells by index, so a wrong column or a shifted cell would
show up here as a mismatch. It highlights what it read, saves a screenshot per report and an evidence page in
handover/bang-chung-doi-chieu/, and writes data/extracted/page_check.csv, which the dataset uses to mark matching
figures as confirmed. Exits 1 on any mismatch.
"""
import csv
import datetime
import html
import re
import sys
from decimal import Decimal
from pathlib import Path

from . import dataset
from .extract import filing_paths
from .sources import QUARTERS, document_title, quarterly_filing, sec_document_url

ROOT = Path(__file__).resolve().parent.parent
OUT_CSV = ROOT / "data" / "extracted" / "page_check.csv"
EVIDENCE = ROOT / "handover" / "bang-chung-doi-chieu"
EVIDENCE_PAGE = "doi-chieu-36-so.html"
CN_ROWS = [("revenues_pct_change_cn", "TOTAL NIKE, INC. REVENUES", "% thay đổi doanh thu NIKE, Inc. loại trừ tỷ giá"),
           ("nike_brand_pct_change_cn", "TOTAL NIKE BRAND REVENUES", "% thay đổi NIKE Brand loại trừ tỷ giá"),
           ("nike_direct_pct_change_cn", "SALES THROUGH NIKE DIRECT", "% thay đổi NIKE Direct loại trừ tỷ giá"),
           ("nike_brand_wholesale_pct_change_cn", "SALES TO WHOLESALE CUSTOMERS", "% thay đổi bán buôn loại trừ tỷ giá")]

# The printed page `page` of a 10-Q runs from the page break after page-1 to the page break after its own number.
JS_TABLE = r"""([page, labels]) => {
  const norm = s => s.replace(/\s+/g, ' ').trim();
  const breaks = Array.from(document.querySelectorAll('hr')).filter(h => /page-break/i.test(h.getAttribute('style') || ''));
  const walker = document.createTreeWalker(document.body, NodeFilter.SHOW_TEXT);
  const texts = [];
  let n;
  while ((n = walker.nextNode())) { const t = norm(n.nodeValue); if (t) texts.push(n); }
  const numberBefore = hr => {
    let last = null;
    for (const t of texts) {
      if (hr.compareDocumentPosition(t) & Node.DOCUMENT_POSITION_FOLLOWING) break;
      last = t;
    }
    const v = last ? norm(last.nodeValue) : '';
    return /^\d{1,3}$/.test(v) ? +v : null;
  };
  const numbered = breaks.map(h => ({hr: h, page: numberBefore(h)}));
  const endIdx = numbered.findIndex(b => b.page === page);
  if (endIdx < 0) return {error: 'page number not found'};
  const end = numbered[endIdx].hr, start = endIdx > 0 ? numbered[endIdx - 1].hr : null;
  const inPage = e => (e.compareDocumentPosition(end) & Node.DOCUMENT_POSITION_FOLLOWING) &&
                      (!start || (start.compareDocumentPosition(e) & Node.DOCUMENT_POSITION_FOLLOWING));
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
  t.scrollIntoView();
  return {values};
}"""


def read_pct(text):
    # Some reports put the "%" sign in a cell of its own, outside the header's column span.
    m = re.search(r"(\(?-?\d+\)?)\s*%?", text or "")
    if not m:
        return None
    s = m.group(1)
    return Decimal(-int(s.strip("()-")) if s.startswith(("(", "-")) else int(s)) / 100


def check_quarter(pg, fy, q, facts):
    f = quarterly_filing(fy, q)
    period = f"FY{fy}Q{q}"
    get = lambda mid: facts[(f["id"], mid, period, "mdna_table")]
    page = get(CN_ROWS[0][0])["page"]
    doc, _ = filing_paths(f)
    pg.goto(doc.resolve().as_uri(), wait_until="load", timeout=180000)
    r = pg.evaluate(JS_TABLE, [page, [[mid, label] for mid, label, _ in CN_ROWS]])
    shot = f"Q{q}-FY{fy}-bang-doanh-thu-tr{page}.png"
    if "values" in r:
        pg.locator("[data-recheck=table]").screenshot(path=str(EVIDENCE / shot))
    rows = []
    for mid, _label, vi in CN_ROWS:
        want = get(mid)["value"]
        text = (r.get("values") or {}).get(mid)
        got = read_pct(text)
        rows.append(dict(filing=f["id"], metric_id=mid, period=period, label=vi, page=page,
                         read_text=text or r.get("error", ""), read_value=got, extracted_value=want,
                         result="match" if got == want else "MISMATCH", screenshot=shot if "values" in r else ""))
    return rows


def show(v):
    if v is None:
        return "—"
    p = int(v * 100)
    return f"{'+' if p > 0 else '−' if p < 0 else ''}{abs(p)}%"


def write_evidence_page(rows):
    ok = sum(r["result"] == "match" for r in rows)
    y, m, d = max(r["checked_on"] for r in rows).split("-")
    out = ["<!doctype html><html lang='vi'><meta charset='utf-8'><meta name='viewport' content='width=device-width, initial-scale=1'>"
           f"<title>Đối chiếu {len(rows)} số trên trang 10-Q</title>",
           "<style>body{font:15px/1.55 Segoe UI,Arial,sans-serif;margin:28px auto;max-width:1400px;padding:0 16px;color:#15191e}"
           "table{border-collapse:collapse;margin:8px 0 14px}td,th{border-bottom:1px solid #ddd;padding:5px 12px;text-align:left}"
           "td+td{white-space:nowrap}"
           "th{background:#f4f5f7}.ok{color:#0a7a3d;font-weight:600}.bad{color:#b42318;font-weight:700}"
           "img{max-width:100%;border:1px solid #ccc;margin:6px 0 20px}h2{margin-top:40px;border-top:2px solid #15191e;padding-top:12px}"
           "@media (max-width:640px){body{font-size:14px;margin:16px auto}table{width:100%}td,th{padding:5px 6px}}</style>",
           f"<h1>Đối chiếu {len(rows)} tỷ lệ tăng trưởng loại trừ tỷ giá với trang 10-Q</h1>",
           f"<p><b>{ok}/{len(rows)} số khớp.</b> Kiểm tra ngày {d}/{m}/{y} bằng trình duyệt Chrome điều khiển tự động (Playwright). "
           "Đây là những số chỉ có trong phần MD&amp;A của báo cáo quý, không gắn thẻ XBRL và không được in lại ở báo cáo sau. "
           "Cách làm: mở từng báo cáo 10-Q, tìm đúng trang theo số trang in ở cuối trang, rồi đọc lại số trên trang đã hiển thị, "
           "theo vị trí cột “% CHANGE EXCLUDING CURRENCY CHANGES” <i>đầu tiên</i> (cột của ba tháng; báo cáo quý 2 và quý 3 còn "
           "một cột cùng tên cho số lũy kế). Ô số đã đọc được tô nền vàng, dòng tương ứng có viền đỏ.</p>",
           "<p>Ảnh chụp từ bản báo cáo tải về từ NIKE Investor Relations, cùng số accession và cùng nội dung với bản trên SEC "
           "EDGAR (liên kết SEC ở từng quý).</p>"]
    for fy, q in QUARTERS:
        yr = [r for r in rows if r["period"] == f"FY{fy}Q{q}"]
        if not yr:
            continue
        f = quarterly_filing(fy, q)
        out.append(f"<h2>Quý {q} FY{fy}</h2><p><a href='{sec_document_url(f)}'>Mở {html.escape(document_title(f))} trên SEC EDGAR</a></p>"
                   "<table><tr><th>Chỉ tiêu</th><th>Trang</th><th>Số trong bộ dữ liệu</th><th>Đọc lại trên trang</th><th>Kết quả</th></tr>")
        for r in yr:
            good = r["result"] == "match"
            out.append(f"<tr><td>{html.escape(r['label'])}</td><td>{r['page']}</td><td>{show(r['extracted_value'])}</td>"
                       f"<td>{show(r['read_value'])}</td><td class='{'ok' if good else 'bad'}'>{'Khớp' if good else 'KHÔNG KHỚP'}</td></tr>")
        out.append("</table>")
        for shot in dict.fromkeys(r["screenshot"] for r in yr if r["screenshot"]):
            out.append(f"<img src='{html.escape(shot)}' alt='{html.escape(shot)}'>")
    (EVIDENCE / EVIDENCE_PAGE).write_text("\n".join(out) + "\n", encoding="utf-8")


def read_live(facts):
    from playwright.sync_api import sync_playwright
    rows = []
    with sync_playwright() as p:
        try:
            browser = p.chromium.launch(channel="chrome")
        except Exception:
            browser = p.chromium.launch()
        pg = browser.new_page(viewport={"width": 1400, "height": 1100})
        # The saved copies are read as they are: nothing is fetched from the network.
        pg.route("**/*", lambda route: route.continue_() if route.request.url.startswith("file:") else route.abort())
        for fy, q in QUARTERS:
            if q == 4:
                continue
            rows_q = check_quarter(pg, fy, q, facts)
            rows += rows_q
            print(f"Q{q} FY{fy}: " + ", ".join(f"{r['metric_id']}={r['read_value']} {r['result']}" for r in rows_q))
        browser.close()
    return rows


def read_saved(facts):
    """Re-evaluate the readings stored in page_check.csv (no browser), e.g. after fixing a parser."""
    labels = {mid: vi for mid, _, vi in CN_ROWS}
    rows = []
    with open(OUT_CSV, encoding="utf-8") as fh:
        for r in csv.DictReader(fh):
            want = facts[(r["filing"], r["metric_id"], r["period"], "mdna_table")]["value"]
            got = read_pct(r["read_text"])
            rows.append(dict(filing=r["filing"], metric_id=r["metric_id"], period=r["period"], label=labels[r["metric_id"]],
                             page=int(r["page"]), read_text=r["read_text"], read_value=got, extracted_value=want,
                             result="match" if got == want else "MISMATCH", screenshot=r["screenshot"],
                             checked_on=r.get("checked_on") or datetime.date.today().isoformat()))
    return rows


def main():
    facts = dataset.load_facts()
    EVIDENCE.mkdir(parents=True, exist_ok=True)
    rows = read_saved(facts) if "--from-saved" in sys.argv else read_live(facts)
    for r in rows:
        r.setdefault("checked_on", datetime.date.today().isoformat())
    fields = ["filing", "metric_id", "period", "page", "read_value", "extracted_value", "result", "checked_on", "screenshot",
              "read_text"]
    with open(OUT_CSV, "w", newline="", encoding="utf-8") as fh:
        w = csv.DictWriter(fh, fieldnames=fields, extrasaction="ignore")
        w.writeheader()
        for r in rows:
            w.writerow({k: ("" if r[k] is None else r[k]) for k in fields})
    write_evidence_page(rows)
    bad = [r for r in rows if r["result"] != "match"]
    print(f"{len(rows) - len(bad)}/{len(rows)} figures match the rendered 10-Q pages -> {OUT_CSV.relative_to(ROOT)}, "
          f"{EVIDENCE.relative_to(ROOT)}")
    return 1 if bad else 0


if __name__ == "__main__":
    sys.exit(main())

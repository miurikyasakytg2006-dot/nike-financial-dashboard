"""Reader for NIKE 10-K and 10-Q documents.

Two formats occur in the period covered:
  * inline XBRL (XHTML), from the first quarter of fiscal 2020 on: the numbers are tagged in the document itself;
  * HTML with a separate XBRL instance file, for fiscal 2019 and earlier: the facts come from the instance, and
    each fact used is located in the printed document by its row label and value (see `find_row`).

One document-order pass turns the filing into blocks (text lines, tables, tagged numeric facts). Each block
carries the printed page number and the nearest preceding section heading, so every extracted value can be
traced to a place a reader can find in the filing. Page numbers come from the 10-K footer ("2020 FORM 10-K 41",
or "42 NIKE, INC." on even pages of the fiscal 2019 report) or, in a 10-Q, from the number printed at the
bottom of the page, just before the page break.
"""
import re
from decimal import Decimal

from lxml import etree, html

NS = {
    "ix": "http://www.xbrl.org/2013/inlineXBRL",
    "xbrli": "http://www.xbrl.org/2003/instance",
    "xbrldi": "http://xbrl.org/2006/xbrldi",
}
XH = "{http://www.w3.org/1999/xhtml}"
IX = "{%s}" % NS["ix"]

FOOTER_RE = re.compile(r"^(?:(\d{4}) FORM 10-K\s*(\d+)|(\d+) NIKE, INC\.)$")
PAGE_NUMBER_RE = re.compile(r"^\d{1,3}$")
HEADING_RE = re.compile(
    r"^((?:UNAUDITED )?(?:CONDENSED )?CONSOLIDATED (?:STATEMENTS? OF [A-Z' ,()]+|BALANCE SHEETS)"
    r"|NOTE \d+ ?[—–-] ?[A-Z][A-Z ,&'’()\-]+"
    r"|ITEM \d+[A-Z]?\.?\s+[A-Z][A-Za-z ,'’&]+"
    r"|SELECTED QUARTERLY FINANCIAL DATA(?: \(UNAUDITED\))?)$", re.IGNORECASE)
NUMBER_WORDS = {"no": 0, "none": 0, "one": 1, "two": 2, "three": 3}
SKIP_LINES = {"table of contents"}


def load(path):
    parser = etree.XMLParser(recover=False, huge_tree=True)
    return etree.parse(str(path), parser).getroot()


def _text(el):
    return " ".join("".join(el.itertext()).split())


def _local(el):
    return el.tag.rsplit("}", 1)[-1] if isinstance(el.tag, str) else ""


def contexts(root):
    out = {}
    for c in root.iter("{%s}context" % NS["xbrli"]):
        p = c.find("xbrli:period", NS)
        dims = {m.get("dimension"): m.text.strip() for m in c.iter("{%s}explicitMember" % NS["xbrldi"])}
        out[c.get("id")] = {
            "instant": p.findtext("xbrli:instant", namespaces=NS),
            "start": p.findtext("xbrli:startDate", namespaces=NS),
            "end": p.findtext("xbrli:endDate", namespaces=NS),
            "dims": dims,
        }
    return out


def parse_value(el):
    """Numeric value of an ix:nonFraction element in base units (USD), or None."""
    raw = _text(el)
    fmt = el.get("format") or ""
    if el.get("{http://www.w3.org/2001/XMLSchema-instance}nil") == "true":
        return None
    if "numwords" in fmt:
        if raw.lower() not in NUMBER_WORDS:
            return None
        val = Decimal(NUMBER_WORDS[raw.lower()])
    elif "zero" in fmt or raw in ("—", "–", "-"):
        val = Decimal(0)
    else:
        s = raw.replace(",", "").replace("$", "").strip()
        if not s:
            return None
        val = Decimal(s)
    val = val.scaleb(int(el.get("scale") or 0))
    if el.get("sign") == "-":
        val = -val
    return val


def displayed(value):
    """How the filing prints a USD amount in millions: 39,117 or (50)."""
    s = f"{abs(value) / 10 ** 6:,.0f}"
    return f"({s})" if value < 0 else s


def instance_facts(path):
    """Numeric facts of a traditional XBRL instance file, in the same shape as the inline ones (without location)."""
    root = load(path)
    ctx = contexts(root)
    prefix = {uri: p for p, uri in root.nsmap.items() if p}
    out = []
    for el in root:
        if not isinstance(el.tag, str) or el.get("contextRef") is None or el.get("unitRef") is None:
            continue
        uri, name = el.tag[1:].split("}")
        c = ctx[el.get("contextRef")]
        text = (el.text or "").strip()
        nil = el.get("{http://www.w3.org/2001/XMLSchema-instance}nil") == "true"
        out.append({
            "concept": f"{prefix.get(uri, uri)}:{name}",
            "start": c["start"], "end": c["end"], "instant": c["instant"], "dims": c["dims"],
            "value": None if nil or not text else Decimal(text), "decimals": el.get("decimals"), "displayed": "",
            "section": "", "table": None, "page": None,
        })
    return out, ctx


class Document:
    """A parsed 10-K or 10-Q: facts, tables and text lines, each located by page and section.

    `instance` is the XBRL instance file of a traditional (non-inline) filing; its facts carry no location until
    the extractor finds them in the document's tables.
    """

    def __init__(self, path, instance=None):
        self.inline = instance is None
        if self.inline:
            self.root = load(path)
            self.contexts = contexts(self.root)
            self.facts = []
        else:
            page = html.fromstring(open(path, "rb").read().decode("utf-8"))
            # The investor-relations copy wraps the filing in its own page; only the filing itself is read.
            main = page.find(".//main")
            self.root = main if main is not None else page
            self.facts, self.contexts = instance_facts(instance)
        self.tables, self.lines = [], []
        self._walk()

    def _walk(self):
        pending, section, sub = [], "", ""
        for el in self.root.iter():
            tag = _local(el)
            if not tag or tag == "header" and el.tag == IX + "header":
                continue
            if tag == "table":
                rows = self._table_rows(el)
                # Workiva lays out some headings (e.g. "NOTE 14 — REVENUES") as one-cell tables.
                for row in rows:
                    if len(row) == 1 and HEADING_RE.match(row[0]):
                        section, sub = row[0], ""
                before = [ln["text"] for ln in self.lines[-3:]]
                block = {"rows": rows, "section": section, "sub": sub, "page": None, "line": len(self.lines),
                         "before": before}
                self.tables.append(block)
                pending.append(block)
                for row in rows:
                    self.lines.append({"text": " | ".join(row), "section": section, "sub": sub, "page": None})
                    pending.append(self.lines[-1])
            elif el.tag == IX + "nonFraction":
                c = self.contexts[el.get("contextRef")]
                fact = {
                    "concept": el.get("name"),
                    "start": c["start"], "end": c["end"], "instant": c["instant"], "dims": c["dims"],
                    "value": parse_value(el), "decimals": el.get("decimals"), "displayed": _text(el),
                    "section": section, "table": len(self.tables) - 1 if self._in_table(el) else None,
                    "page": None,
                }
                self.facts.append(fact)
                pending.append(fact)
            elif tag == "hr" and "page-break" in (el.get("style") or ""):
                # A 10-Q prints the page number as the last line of the page.
                last = next((p for p in reversed(pending) if "text" in p and "rows" not in p), None)
                if last is not None and PAGE_NUMBER_RE.match(last["text"]) and last is self.lines[-1]:
                    self.lines.pop()
                    number = int(last["text"])
                    for item in pending:
                        if item is not last:
                            item["page"] = number
                pending = []
            elif tag in ("div", "p") and self._is_leaf_text(el):
                txt = _text(el)
                if not txt or txt.lower() in SKIP_LINES:
                    continue
                m = FOOTER_RE.match(txt)
                if m:
                    for item in pending:
                        item["page"] = int(m.group(2) or m.group(3))
                    pending = []
                    continue
                if HEADING_RE.match(txt):
                    section, sub = txt, ""
                elif txt.isupper() and 4 < len(txt) < 70 and not re.search(r"\d", txt):
                    sub = txt
                self.lines.append({"text": txt, "section": section, "sub": sub, "page": None})
                pending.append(self.lines[-1])

    def _table_rows(self, table):
        rows = []
        for tr in table.iter(*self._tags("tr")):
            cells = [_text(td) for td in tr.iter(*self._tags("td", "th"))]
            cells = [c for c in cells if c and c not in ("$", "%", ")")]
            if cells:
                rows.append(cells)
        return rows

    def _tags(self, *names):
        return [XH + n for n in names] if self.inline else list(names)

    def _in_table(self, el):
        return any(_local(a) == "table" for a in el.iterancestors())

    def _is_leaf_text(self, el):
        for d in el.iterdescendants():
            if _local(d) in ("div", "p", "table"):
                return False
        return not self._in_table(el)

    def find_row(self, section, label, shown):
        """The first table row (in a section matching `section`) whose label matches `label` and that shows the
        amount `shown`, with its table; None if the document has no such row."""
        sec, lab = re.compile(section, re.IGNORECASE), re.compile(label, re.IGNORECASE)
        for t in self.tables:
            if not sec.search(t["section"]):
                continue
            for r in t["rows"]:
                if lab.fullmatch(re.sub(r"(\(\d+\))+:?$", "", r[0]).strip()) and shown in r[1:]:
                    return t, r
        return None

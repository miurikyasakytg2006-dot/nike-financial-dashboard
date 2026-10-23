"""Minimal Inline XBRL reader for NIKE 10-K documents.

One document-order pass turns the filing into blocks (text lines, tables, tagged numeric
facts). Each block carries the printed page number (from the "YYYY FORM 10-K NN" page
footer that closes its page) and the nearest preceding section heading, so every
extracted value can be traced to a location a reader can find in the filing.
"""
import re
from decimal import Decimal

from lxml import etree

NS = {
    "ix": "http://www.xbrl.org/2013/inlineXBRL",
    "xbrli": "http://www.xbrl.org/2003/instance",
    "xbrldi": "http://xbrl.org/2006/xbrldi",
}
XH = "{http://www.w3.org/1999/xhtml}"
IX = "{%s}" % NS["ix"]

FOOTER_RE = re.compile(r"^(\d{4}) FORM 10-K\s*(\d+)$")
HEADING_RE = re.compile(
    r"^(CONSOLIDATED STATEMENTS? OF [A-Z' ,()]+|CONSOLIDATED BALANCE SHEETS|NOTE \d+ ?[—–-] ?[A-Z][A-Z ,&'()\-]+"
    r"|ITEM \d+[A-Z]?\.?\s+[A-Z][A-Za-z ,'’&]+)$"
)
NUMBER_WORDS = {"no": 0, "none": 0, "one": 1, "two": 2, "three": 3}


def load(path):
    parser = etree.XMLParser(recover=False, huge_tree=True)
    return etree.parse(str(path), parser).getroot()


def _text(el):
    return " ".join("".join(el.itertext()).split())


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


def _table_rows(table):
    rows = []
    for tr in table.iter(XH + "tr"):
        cells = [_text(td) for td in tr.iter(XH + "td")]
        cells = [c for c in cells if c and c not in ("$", "%", ")")]
        if cells:
            rows.append(cells)
    return rows


class Document:
    """A parsed 10-K: facts, tables and text lines, each located by page and section."""

    def __init__(self, path):
        self.root = load(path)
        self.contexts = contexts(self.root)
        self.facts, self.tables, self.lines = [], [], []
        self._walk()

    def _walk(self):
        pending, section = [], ""
        for el in self.root.iter():
            if not isinstance(el.tag, str) or el.tag == IX + "header":
                continue
            if el.tag == XH + "table":
                rows = _table_rows(el)
                # Workiva lays out some headings (e.g. "NOTE 14 — REVENUES") as one-cell tables.
                for row in rows:
                    if len(row) == 1 and HEADING_RE.match(row[0]):
                        section = row[0]
                block = {"rows": rows, "section": section, "page": None, "line": len(self.lines)}
                self.tables.append(block)
                pending.append(block)
                for row in rows:
                    self.lines.append({"text": " | ".join(row), "section": section, "page": None})
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
            elif el.tag == XH + "div" and self._is_leaf_text(el):
                txt = _text(el)
                if not txt:
                    continue
                m = FOOTER_RE.match(txt)
                if m:
                    for item in pending:
                        item["page"] = int(m.group(2))
                    pending = []
                    continue
                if HEADING_RE.match(txt):
                    section = txt
                self.lines.append({"text": txt, "section": section, "page": None})
                pending.append(self.lines[-1])

    @staticmethod
    def _in_table(el):
        return any(a.tag == XH + "table" for a in el.iterancestors())

    def _is_leaf_text(self, el):
        if el.find(".//" + XH + "div") is not None or el.find(".//" + XH + "table") is not None:
            return False
        return not self._in_table(el)

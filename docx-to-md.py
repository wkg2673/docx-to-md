#!/usr/bin/env python3
"""
DOCX -> Markdown converter.

Run the program and it asks for:
  1. the full path of the .docx file
  2. the full path of the .md output file

Requires:  pip install python-docx

Supported: headings, bold / italic / strikethrough / inline code, hyperlinks,
bullet and numbered lists (nested), block quotes, tables, and images
(images are extracted into a folder next to the output file).
"""

import os
import re
import sys

from docx import Document
from docx.oxml.ns import qn
from docx.table import Table
from docx.text.paragraph import Paragraph

CODE_FONTS = {"courier new", "courier", "consolas", "monaco", "menlo", "lucida console"}


# --------------------------------------------------------------------------- #
# Helpers
# --------------------------------------------------------------------------- #
def clean_path(raw: str) -> str:
    """Strip whitespace and surrounding quotes (e.g. from 'Copy as path')."""
    raw = raw.strip()
    if len(raw) >= 2 and raw[0] == raw[-1] and raw[0] in ("'", '"'):
        raw = raw[1:-1]
    return os.path.expanduser(raw.strip())


def escape_md(text: str) -> str:
    """Escape characters that would otherwise be read as Markdown syntax."""
    return re.sub(r"([\\`*_\[\]<>|])", r"\\\1", text)


def wrap(text: str, marker: str) -> str:
    """Wrap text in a Markdown marker, keeping edge whitespace outside it."""
    if not text.strip():
        return text
    lead = text[: len(text) - len(text.lstrip())]
    trail = text[len(text.rstrip()):]
    return f"{lead}{marker}{text.strip()}{marker}{trail}"


# --------------------------------------------------------------------------- #
# Converter
# --------------------------------------------------------------------------- #
class DocxToMarkdown:
    def __init__(self, docx_path: str, md_path: str):
        self.doc = Document(docx_path)
        self.md_path = md_path
        self.image_dir_name = os.path.splitext(os.path.basename(md_path))[0] + "_images"
        self.image_dir = os.path.join(os.path.dirname(os.path.abspath(md_path)), self.image_dir_name)
        self.image_count = 0
        self.image_cache = {}          # rId -> relative markdown path
        self.counters = {}             # (numId, ilvl) -> current number
        self.num_formats = self._load_numbering()

    # ---- numbering (list type detection) ---------------------------------- #
    def _load_numbering(self):
        """Map (numId, ilvl) -> numFmt (e.g. 'bullet', 'decimal')."""
        formats = {}
        try:
            numbering = self.doc.part.numbering_part.element
        except Exception:
            return formats

        abstract = {}
        for an in numbering.findall(qn("w:abstractNum")):
            aid = an.get(qn("w:abstractNumId"))
            for lvl in an.findall(qn("w:lvl")):
                fmt = lvl.find(qn("w:numFmt"))
                abstract[(aid, lvl.get(qn("w:ilvl")))] = fmt.get(qn("w:val")) if fmt is not None else "decimal"

        for num in numbering.findall(qn("w:num")):
            nid = num.get(qn("w:numId"))
            ref = num.find(qn("w:abstractNumId"))
            if ref is None:
                continue
            aid = ref.get(qn("w:val"))
            for (a, lvl), fmt in abstract.items():
                if a == aid:
                    formats[(nid, lvl)] = fmt
        return formats

    # ---- images ------------------------------------------------------------ #
    def _save_image(self, rid: str):
        if rid in self.image_cache:
            return self.image_cache[rid]
        try:
            part = self.doc.part.related_parts[rid]
        except KeyError:
            return None
        os.makedirs(self.image_dir, exist_ok=True)
        self.image_count += 1
        ext = os.path.splitext(part.partname)[1] or ".png"
        filename = f"image{self.image_count}{ext}"
        with open(os.path.join(self.image_dir, filename), "wb") as f:
            f.write(part.blob)
        rel = f"{self.image_dir_name}/{filename}"
        self.image_cache[rid] = rel
        return rel

    # ---- inline content ---------------------------------------------------- #
    def _run_format(self, run_el):
        """Return (bold, italic, strike, code) for a w:r element."""
        rpr = run_el.find(qn("w:rPr"))
        bold = italic = strike = code = False
        if rpr is not None:
            def on(tag):
                el = rpr.find(qn(tag))
                return el is not None and el.get(qn("w:val"), "true") not in ("0", "false", "off")
            bold, italic = on("w:b"), on("w:i")
            strike = on("w:strike") or on("w:dstrike")
            fonts = rpr.find(qn("w:rFonts"))
            if fonts is not None:
                name = (fonts.get(qn("w:ascii")) or fonts.get(qn("w:hAnsi")) or "").lower()
                code = name in CODE_FONTS
        return bold, italic, strike, code

    def _run_pieces(self, run_el):
        """Yield ('text', str) / ('image', path) pieces from a w:r element."""
        for child in run_el:
            tag = child.tag
            if tag == qn("w:t"):
                yield "text", child.text or ""
            elif tag == qn("w:tab"):
                yield "text", "    "
            elif tag in (qn("w:br"), qn("w:cr")):
                if child.get(qn("w:type")) != "page":
                    yield "text", "\n"
            elif tag in (qn("w:drawing"), qn("w:pict")):
                for blip in child.iter(qn("a:blip")):
                    rid = blip.get(qn("r:embed"))
                    if rid:
                        path = self._save_image(rid)
                        if path:
                            yield "image", path

    def _runs_to_md(self, runs):
        """
        Convert a list of w:r elements to Markdown, merging adjacent runs
        that share formatting so we don't produce '**a****b**'.
        """
        segments = []  # [fmt, text]
        for run_el in runs:
            fmt = self._run_format(run_el)
            for kind, value in self._run_pieces(run_el):
                if kind == "image":
                    segments.append(["image", f"![image]({value})"])
                    continue
                if segments and segments[-1][0] == fmt:
                    segments[-1][1] += value
                else:
                    segments.append([fmt, value])

        out = []
        for fmt, text in segments:
            if fmt == "image":
                out.append(text)
                continue
            bold, italic, strike, code = fmt
            if code:
                text = wrap(text, "`")
            else:
                text = escape_md(text)
                if bold and italic:
                    text = wrap(text, "***")
                elif bold:
                    text = wrap(text, "**")
                elif italic:
                    text = wrap(text, "*")
                if strike:
                    text = wrap(text, "~~")
            out.append(text)
        return "".join(out)

    def _paragraph_inline(self, p_el):
        """Convert all inline content of a w:p (runs + hyperlinks) to Markdown."""
        parts = []
        pending_runs = []

        def flush():
            if pending_runs:
                parts.append(self._runs_to_md(pending_runs))
                pending_runs.clear()

        for child in p_el:
            if child.tag == qn("w:r"):
                pending_runs.append(child)
            elif child.tag == qn("w:hyperlink"):
                flush()
                label = self._runs_to_md(child.findall(qn("w:r"))).strip()
                rid = child.get(qn("r:id"))
                anchor = child.get(qn("w:anchor"))
                url = None
                if rid and rid in self.doc.part.rels:
                    url = self.doc.part.rels[rid].target_ref
                elif anchor:
                    url = f"#{anchor}"
                parts.append(f"[{label}]({url})" if url and label else label)
            elif child.tag in (qn("w:ins"), qn("w:smartTag")):
                # Tracked insertions / smart tags wrap normal runs
                pending_runs.extend(child.findall(qn("w:r")))
        flush()
        return "".join(parts)

    # ---- block content ----------------------------------------------------- #
    def _list_info(self, paragraph):
        """Return (ordered: bool, level: int, numId) or None if not a list item."""
        ppr = paragraph._p.pPr
        num_pr = ppr.numPr if ppr is not None else None
        style = (paragraph.style.name if paragraph.style is not None else "").lower()

        if num_pr is not None and num_pr.numId is not None and num_pr.numId.val != 0:
            nid = str(num_pr.numId.val)
            lvl = str(num_pr.ilvl.val) if num_pr.ilvl is not None else "0"
            fmt = self.num_formats.get((nid, lvl), "decimal")
            return fmt != "bullet", int(lvl), nid

        if "list bullet" in style:
            m = re.search(r"(\d+)$", style)
            return False, (int(m.group(1)) - 1 if m else 0), "style-bullet"
        if "list number" in style:
            m = re.search(r"(\d+)$", style)
            return True, (int(m.group(1)) - 1 if m else 0), "style-number"
        return None

    def _paragraph_to_md(self, paragraph):
        text = self._paragraph_inline(paragraph._p)
        style = (paragraph.style.name if paragraph.style is not None else "").lower()

        # Headings
        if style == "title":
            self.counters.clear()
            return f"# {text.strip()}" if text.strip() else ""
        if style == "subtitle":
            return f"*{text.strip()}*" if text.strip() else ""
        m = re.match(r"heading (\d)", style)
        if m and text.strip():
            self.counters.clear()
            level = min(int(m.group(1)), 6)
            return f"{'#' * level} {re.sub(r'^\*+|\*+$', '', text.strip())}"

        # Lists
        info = self._list_info(paragraph)
        if info and text.strip():
            ordered, level, nid = info
            if ordered:
                self.counters[(nid, level)] = self.counters.get((nid, level), 0) + 1
                for key in [k for k in self.counters if k[0] == nid and k[1] > level]:
                    del self.counters[key]
                marker = f"{self.counters[(nid, level)]}."
            else:
                marker = "-"
            return f"{'    ' * level}{marker} {text.strip()}"
        self.counters.clear()

        # Quotes
        if style in ("quote", "intense quote", "block text"):
            return "> " + text.strip().replace("\n", "\n> ") if text.strip() else ""

        # Normal text: a manual line break becomes a Markdown hard break
        return text.strip().replace("\n", "  \n")

    def _table_to_md(self, table):
        rows = []
        for row in table.rows:
            cells = []
            seen = set()
            for cell in row.cells:
                if id(cell._tc) in seen:  # merged cells repeat
                    continue
                seen.add(id(cell._tc))
                content = "<br>".join(
                    self._paragraph_inline(p._p).strip() for p in cell.paragraphs if p.text.strip()
                )
                cells.append(re.sub(r"(?<!\\)\|", r"\\|", content))
            rows.append(cells)
        if not rows:
            return ""

        width = max(len(r) for r in rows)
        rows = [r + [""] * (width - len(r)) for r in rows]
        lines = ["| " + " | ".join(rows[0]) + " |", "| " + " | ".join(["---"] * width) + " |"]
        lines += ["| " + " | ".join(r) + " |" for r in rows[1:]]
        return "\n".join(lines)

    # ---- main -------------------------------------------------------------- #
    def convert(self) -> str:
        blocks = []
        previous_list_id = None  # numId of the list the previous block belonged to

        for child in self.doc.element.body.iterchildren():
            if child.tag == qn("w:p"):
                paragraph = Paragraph(child, self.doc)
                md = self._paragraph_to_md(paragraph)
                if not md.strip():
                    previous_list_id = None
                    continue
                info = self._list_info(paragraph)
                list_id = info[2] if info else None
                # Consecutive items of the same list are joined by a single newline
                if list_id is not None and list_id == previous_list_id and blocks:
                    blocks[-1] += "\n" + md
                else:
                    blocks.append(md)
                previous_list_id = list_id
            elif child.tag == qn("w:tbl"):
                md = self._table_to_md(Table(child, self.doc))
                if md:
                    blocks.append(md)
                previous_list_id = None

        return "\n\n".join(blocks) + "\n"


# --------------------------------------------------------------------------- #
# Command-line interaction
# --------------------------------------------------------------------------- #
def main():
    print("=== DOCX to Markdown Converter ===\n")

    # Input path
    while True:
        in_path = clean_path(input("Full path of the .docx file (including extension): "))
        if not in_path.lower().endswith(".docx"):
            print("  -> The input file must have the .docx extension.\n")
        elif not os.path.isfile(in_path):
            print("  -> File not found. Please check the path.\n")
        else:
            break

    # Output path
    while True:
        out_path = clean_path(input("Full output path (including extension, e.g. C:\\out\\result.md): "))
        if not out_path.lower().endswith(".md"):
            print("  -> The output file must have the .md extension.\n")
            continue
        out_dir = os.path.dirname(os.path.abspath(out_path))
        if not os.path.isdir(out_dir):
            create = input(f"  Folder '{out_dir}' does not exist. Create it? (y/n): ").strip().lower()
            if create != "y":
                continue
            os.makedirs(out_dir, exist_ok=True)
        if os.path.exists(out_path):
            overwrite = input("  Output file already exists. Overwrite? (y/n): ").strip().lower()
            if overwrite != "y":
                continue
        break

    try:
        converter = DocxToMarkdown(in_path, out_path)
        markdown = converter.convert()
        with open(out_path, "w", encoding="utf-8") as f:
            f.write(markdown)
    except Exception as exc:
        print(f"\nConversion failed: {exc}")
        sys.exit(1)

    print(f"\nDone! Markdown saved to: {out_path}")
    if converter.image_count:
        print(f"{converter.image_count} image(s) saved to: {converter.image_dir}")


if __name__ == "__main__":
    main()

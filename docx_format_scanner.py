"""
docx_format_scanner.py

Untersucht eine .docx-Datei ausschliesslich lesend auf Formatierungen,
die auf Hintergruende, Hervorhebungen und Textfarben zurueckgehen.

Arbeitet direkt auf dem XML (word/document.xml, word/styles.xml),
da die Bibliothek python-docx Zeichen-Schattierung (w:shd) und
einige Detailattribute nicht zuverlaessig offenlegt.
"""

from __future__ import annotations

import argparse
import csv
import sys
import zipfile
from collections.abc import Callable
from dataclasses import dataclass, field
from pathlib import Path

from lxml import etree

from i18n import category_key, get_translator, translate_location

NS = {
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
}


def qn(tag: str) -> str:
    """Wandelt 'w:xyz' in das vollqualifizierte Clark-Notation-Tag um."""
    prefix, local = tag.split(":")
    return f"{{{NS[prefix]}}}{local}"


@dataclass
class Finding:
    kategorie: str        # z.B. "Hervorhebung", "Zeichen-Schattierung", "Textfarbe", "Schriftart"
    wert: str              # z.B. "yellow", "FFFF00", "FF0000"
    ort: str               # z.B. "Absatz 12, Lauf 3" oder "Tabelle 1, Zelle (2,1)"
    textausschnitt: str    # kurzer Kontext, max. ~60 Zeichen
    style_name: str = ""   # falls die Formatierung ueber eine benannte Formatvorlage kommt
    xml_snippet: str = "" # XML-Fragment mit Text und zugehoeriger Formatierung


@dataclass
class ScanResult:
    findings: list[Finding] = field(default_factory=list)
    fonts_verwendet: set[str] = field(default_factory=set)
    quellen_dateien: list[str] = field(default_factory=list)


def _text_snippet(run_or_parent: etree._Element, max_len: int = 60) -> str:
    texts = run_or_parent.findall(".//" + qn("w:t"))
    joined = "".join(t.text or "" for t in texts)
    joined = joined.strip()
    if len(joined) > max_len:
        joined = joined[:max_len] + "..."
    return joined or "(kein Text)"


def _style_name(styles_root: etree._Element | None, style_id: str | None) -> str:
    if styles_root is None or not style_id:
        return ""
    style = styles_root.find(f".//{qn('w:style')}[@{qn('w:styleId')}='{style_id}']")
    if style is None:
        return style_id
    name_el = style.find(qn("w:name"))
    if name_el is not None:
        return name_el.get(qn("w:val")) or style_id
    return style_id


def _xml_fragment(element: etree._Element, max_chars: int = 800) -> str:
    xml = etree.tostring(element, encoding="unicode", pretty_print=True)
    xml = xml.strip()
    if max_chars > 0 and len(xml) > max_chars:
        xml = xml[:max_chars].rstrip() + "..."
    return xml


def _truncate_xml_snippet(snippet: str, max_chars: int = 800) -> str:
    if max_chars <= 0:
        return snippet
    if len(snippet) > max_chars:
        return snippet[:max_chars].rstrip() + "..."
    return snippet


def _scan_run_props(
    rpr: etree._Element,
    ort: str,
    text_ctx: str,
    styles_root: etree._Element | None,
    findings: list[Finding],
    fonts: set[str],
    xml_element: etree._Element,
) -> None:
    snippet = _xml_fragment(xml_element)

    highlight = rpr.find(qn("w:highlight"))
    if highlight is not None:
        val = highlight.get(qn("w:val"))
        findings.append(Finding("Hervorhebung", val or "?", ort, text_ctx, xml_snippet=snippet))

    shd = rpr.find(qn("w:shd"))
    if shd is not None:
        fill = shd.get(qn("w:fill"))
        val = shd.get(qn("w:val"))
        if fill and fill.lower() != "auto":
            findings.append(Finding("Zeichen-Schattierung", f"fill={fill}", ort, text_ctx, xml_snippet=snippet))
        elif val and val.lower() != "clear":
            findings.append(Finding("Zeichen-Schattierung", f"val={val}", ort, text_ctx, xml_snippet=snippet))

    color = rpr.find(qn("w:color"))
    if color is not None:
        val = color.get(qn("w:val"))
        if val and val.lower() not in ("auto",):
            findings.append(Finding("Textfarbe", val, ort, text_ctx, xml_snippet=snippet))

    rfonts = rpr.find(qn("w:rFonts"))
    if rfonts is not None:
        for attr in ("w:ascii", "w:hAnsi", "w:eastAsia", "w:cs"):
            v = rfonts.get(qn(attr))
            if v:
                fonts.add(v)

    rstyle = rpr.find(qn("w:rStyle"))
    if rstyle is not None:
        style_id = rstyle.get(qn("w:val"))
        name = _style_name(styles_root, style_id)
        findings.append(Finding("Zeichenformatvorlage", name, ort, text_ctx, style_name=name, xml_snippet=snippet))


def _scan_paragraph_shading(ppr: etree._Element, ort: str, text_ctx: str, findings: list[Finding], xml_element: etree._Element) -> None:
    shd = ppr.find(qn("w:shd"))
    if shd is not None:
        fill = shd.get(qn("w:fill"))
        val = shd.get(qn("w:val"))
        snippet = _xml_fragment(xml_element)
        if fill and fill.lower() != "auto":
            findings.append(Finding("Absatz-Schattierung", f"fill={fill}", ort, text_ctx, xml_snippet=snippet))
        elif val and val.lower() != "clear":
            findings.append(Finding("Absatz-Schattierung", f"val={val}", ort, text_ctx, xml_snippet=snippet))


def _scan_table_cell_shading(tc: etree._Element, ort: str, findings: list[Finding]) -> None:
    tcpr = tc.find(qn("w:tcPr"))
    if tcpr is None:
        return
    shd = tcpr.find(qn("w:shd"))
    if shd is not None:
        fill = shd.get(qn("w:fill"))
        val = shd.get(qn("w:val"))
        text_ctx = _text_snippet(tc)
        snippet = _xml_fragment(tc)
        if fill and fill.lower() != "auto":
            findings.append(Finding("Tabellenzellen-Schattierung", f"fill={fill}", ort, text_ctx, xml_snippet=snippet))
        elif val and val.lower() != "clear":
            findings.append(Finding("Tabellenzellen-Schattierung", f"val={val}", ort, text_ctx, xml_snippet=snippet))


def scan_document_xml(
    xml_bytes: bytes,
    quelle: str,
    styles_root: etree._Element | None,
) -> tuple[list[Finding], set[str]]:
    findings: list[Finding] = []
    fonts: set[str] = set()

    root = etree.fromstring(xml_bytes)

    absatz_index = 0
    for p in root.iter(qn("w:p")):
        absatz_index += 1
        text_ctx = _text_snippet(p)

        ppr = p.find(qn("w:pPr"))
        if ppr is not None:
            _scan_paragraph_shading(ppr, f"{quelle}: Absatz {absatz_index}", text_ctx, findings, p)

        lauf_index = 0
        for r in p.findall(qn("w:r")):
            lauf_index += 1
            rpr = r.find(qn("w:rPr"))
            if rpr is not None:
                ort = f"{quelle}: Absatz {absatz_index}, Lauf {lauf_index}"
                run_text_ctx = _text_snippet(r)
                _scan_run_props(rpr, ort, run_text_ctx, styles_root, findings, fonts, r)

    tabelle_index = 0
    for tbl in root.iter(qn("w:tbl")):
        tabelle_index += 1
        zeile_index = 0
        for tr in tbl.findall(qn("w:tr")):
            zeile_index += 1
            spalte_index = 0
            for tc in tr.findall(qn("w:tc")):
                spalte_index += 1
                ort = f"{quelle}: Tabelle {tabelle_index}, Zelle ({zeile_index},{spalte_index})"
                _scan_table_cell_shading(tc, ort, findings)

    return findings, fonts


def scan_docx(pfad: Path) -> ScanResult:
    result = ScanResult()

    with zipfile.ZipFile(pfad, "r") as zf:
        namen = zf.namelist()

        styles_root = None
        if "word/styles.xml" in namen:
            styles_root = etree.fromstring(zf.read("word/styles.xml"))

        ziel_dateien = [n for n in namen if n == "word/document.xml"]
        ziel_dateien += [n for n in namen if n.startswith("word/header") and n.endswith(".xml")]
        ziel_dateien += [n for n in namen if n.startswith("word/footer") and n.endswith(".xml")]
        ziel_dateien += [n for n in namen if n in ("word/footnotes.xml", "word/endnotes.xml")]

        for datei in ziel_dateien:
            xml_bytes = zf.read(datei)
            findings, fonts = scan_document_xml(xml_bytes, datei, styles_root)
            result.findings.extend(findings)
            result.fonts_verwendet.update(fonts)
            result.quellen_dateien.append(datei)

    return result


def _format_bericht(
    result: ScanResult,
    dateiname: str,
    out_xml: bool = False,
    xml_max_chars: int = 800,
    translate: Callable[[str], str] | None = None,
) -> str:
    translate = translate or get_translator("en")
    lines: list[str] = []
    lines.append(translate("report_title").format(filename=dateiname))
    lines.append("=" * 60)
    lines.append(translate("report_parts").format(parts=", ".join(result.quellen_dateien)))
    lines.append("")

    if not result.findings:
        lines.append(translate("no_findings"))
    else:
        by_kategorie: dict[str, list[Finding]] = {}
        for f in result.findings:
            by_kategorie.setdefault(f.kategorie, []).append(f)

        for kategorie, items in by_kategorie.items():
            category = translate(category_key(kategorie))
            lines.append(translate("finding_count").format(category=category, count=len(items)))
            werte: dict[str, list[Finding]] = {}
            for it in items:
                werte.setdefault(it.wert, []).append(it)
            for wert, treffer in werte.items():
                lines.append(translate("value_count").format(value=wert, count=len(treffer)))
                for t in treffer:
                    location = translate_location(t.ort, translate)
                    lines.append(f"    - {location}: \"{t.textausschnitt}\"")
                    if out_xml and t.xml_snippet:
                        snippet = _truncate_xml_snippet(t.xml_snippet, xml_max_chars)
                        if snippet:
                            lines.append("      XML:")
                            for xml_line in snippet.splitlines():
                                lines.append(f"        {xml_line}")
            lines.append("")

    if result.fonts_verwendet:
        lines.append(translate("fonts").format(fonts=", ".join(sorted(result.fonts_verwendet))))
        lines.append(translate("font_hint_1"))
        lines.append(translate("font_hint_2"))
        lines.append("")

    lines.append(translate("note_1"))
    lines.append(translate("note_2"))
    lines.append(translate("note_3"))
    lines.append(translate("note_4"))

    return "\n".join(lines)


def _write_csv(result: ScanResult, ausgabe_pfad: Path, xml_max_chars: int = 800) -> None:
    with open(ausgabe_pfad, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerow(["Kategorie", "Wert", "Ort", "Textausschnitt", "Formatvorlage", "XML"])
        for finding in result.findings:
            xml_value = _truncate_xml_snippet(finding.xml_snippet, xml_max_chars)
            writer.writerow(
                [
                    finding.kategorie,
                    finding.wert,
                    finding.ort,
                    finding.textausschnitt,
                    finding.style_name,
                    xml_value,
                ]
            )


def main() -> int:
    parser = argparse.ArgumentParser(
        description="Read-only inspection of a .docx file for background, "
        "highlight, and text-color formatting."
    )
    parser.add_argument("datei", type=Path, help="Path to the .docx file to inspect")
    parser.add_argument(
        "--csv",
        type=Path,
        default=None,
        help="Optional: path for additionally saving the results as CSV",
    )
    parser.add_argument(
        "--out-xml",
        action="store_true",
        help="Optional: also show the related XML fragment for findings",
    )
    parser.add_argument(
        "--xml-max-chars",
        type=int,
        default=800,
        help="Maximum XML snippet length per finding. 0 = complete snippet, -1 = alias for 0",
    )
    parser.add_argument(
        "--annotate",
        type=Path,
        default=None,
        metavar="AUSGABE.docx",
        help="Optional: create a copy at this path with findings marked as Word "
        "comments. The original file remains unchanged.",
    )
    args = parser.parse_args()

    if not args.datei.exists():
        print(f"File not found: {args.datei}", file=sys.stderr)
        return 1

    if args.datei.suffix.lower() != ".docx":
        print("Warning: file does not have a .docx extension. Continuing anyway.", file=sys.stderr)

    try:
        result = scan_docx(args.datei)
    except zipfile.BadZipFile:
        print("Error: file is not a valid .docx (not a ZIP archive).", file=sys.stderr)
        return 1

    xml_max_chars = args.xml_max_chars
    if xml_max_chars < 0:
        xml_max_chars = 0

    print(_format_bericht(result, args.datei.name, out_xml=args.out_xml, xml_max_chars=xml_max_chars))

    if args.csv:
        _write_csv(result, args.csv, xml_max_chars=xml_max_chars)
        print(f"\nCSV export saved to: {args.csv}")

    if args.annotate:
        from comment_writer import annotate_docx  # lokaler Import, um Zirkelbezuege zu vermeiden

        if args.annotate.resolve() == args.datei.resolve():
            print("Error: --annotate destination must not be identical to the original file.", file=sys.stderr)
            return 1
        try:
            anzahl = annotate_docx(args.datei, args.annotate)
        except zipfile.BadZipFile:
            print("Error: file is not a valid .docx (not a ZIP archive).", file=sys.stderr)
            return 1
        print(f"\n{anzahl} comment(s) written to: {args.annotate}")

    return 0


if __name__ == "__main__":
    sys.exit(main())
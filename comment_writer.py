"""
comment_writer.py

Erstellt aus einer .docx-Datei eine KOPIE, in der alle von
docx_format_scanner gefundenen Formatierungen (Hervorhebung,
Zeichen-Schattierung, Textfarbe, Zeichenformatvorlage,
Absatz-Schattierung, Tabellenzellen-Schattierung) als native
Word-Kommentare an der jeweiligen Fundstelle markiert werden.

Wichtig:
- Die Original-Datei wird ausschliesslich lesend geoeffnet und NICHT
  veraendert. Es wird immer eine neue Datei unter dem angegebenen
  Zielpfad geschrieben.
- Kommentiert wird word/document.xml (Haupttext inkl. Tabellen).
  Kopf-/Fusszeilen sowie Fuss-/Endnoten werden von
  docx_format_scanner.scan_docx() weiterhin mitgescannt und im Bericht
  aufgefuehrt, hier aber nicht zusaetzlich mit Kommentaren versehen.
"""

from __future__ import annotations

import zipfile
from dataclasses import dataclass
from datetime import datetime, timezone
from pathlib import Path

from lxml import etree

from docx_format_scanner import NS, _style_name, _text_snippet, qn

COMMENT_AUTHOR = "WordHighLightFormatInspect"
COMMENT_INITIALS = "WFI"

CONTENT_TYPES_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
COMMENTS_CONTENT_TYPE = (
    "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"
)
COMMENTS_REL_TYPE = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
)
XML_SPACE_ATTR = "{http://www.w3.org/XML/1998/namespace}space"


@dataclass
class CommentEntry:
    comment_id: int
    text: str


def _comment_text(kategorie: str, wert: str, style_name: str, textausschnitt: str) -> str:
    teile = [f"{kategorie}: {wert}"]
    if style_name:
        teile.append(f"(Formatvorlage: {style_name})")
    teile.append(f'- Text: "{textausschnitt}"')
    return " ".join(teile)


def _make_marker(tag: str, comment_id: int) -> etree._Element:
    el = etree.Element(qn(tag))
    el.set(qn("w:id"), str(comment_id))
    return el


def _make_reference_run(comment_id: int) -> etree._Element:
    r = etree.Element(qn("w:r"))
    rpr = etree.SubElement(r, qn("w:rPr"))
    rstyle = etree.SubElement(rpr, qn("w:rStyle"))
    rstyle.set(qn("w:val"), "CommentReference")
    ref = etree.SubElement(r, qn("w:commentReference"))
    ref.set(qn("w:id"), str(comment_id))
    return r


def _wrap_run_with_comment(p: etree._Element, r: etree._Element, comment_id: int) -> None:
    """Setzt commentRangeStart/-End und Referenz-Lauf direkt um 'r' herum."""
    kinder = list(p)
    idx = kinder.index(r)
    p.insert(idx, _make_marker("w:commentRangeStart", comment_id))
    idx_r = list(p).index(r)
    p.insert(idx_r + 1, _make_marker("w:commentRangeEnd", comment_id))
    p.insert(idx_r + 2, _make_reference_run(comment_id))


def _wrap_paragraph_with_comment(p: etree._Element, comment_id: int) -> None:
    """Umschliesst den gesamten Absatzinhalt (nach pPr) mit Kommentarmarken."""
    ppr = p.find(qn("w:pPr"))
    insert_pos = 1 if ppr is not None else 0
    p.insert(insert_pos, _make_marker("w:commentRangeStart", comment_id))
    p.append(_make_marker("w:commentRangeEnd", comment_id))
    p.append(_make_reference_run(comment_id))


def _wrap_cell_with_comment(tc: etree._Element, comment_id: int) -> None:
    """Umschliesst den gesamten Zelleninhalt (erster bis letzter Absatz) mit Kommentarmarken."""
    absaetze = tc.findall(qn("w:p"))
    if not absaetze:
        return
    erster, letzter = absaetze[0], absaetze[-1]
    ppr = erster.find(qn("w:pPr"))
    insert_pos = 1 if ppr is not None else 0
    erster.insert(insert_pos, _make_marker("w:commentRangeStart", comment_id))
    letzter.append(_make_marker("w:commentRangeEnd", comment_id))
    letzter.append(_make_reference_run(comment_id))


def _annotate_document_xml(
    xml_bytes: bytes,
    styles_root: etree._Element | None,
    comments: list[CommentEntry],
    next_id,
) -> bytes:
    """
    Zwei Phasen, bewusst getrennt:

    1. Lesephase: der Baum wird NICHT veraendert. Dabei werden alle
       Fundstellen samt betroffenem Original-Element eingesammelt.
    2. Schreibphase: erst danach werden die Kommentarmarken eingefuegt.

    Das ist noetig, weil die neu eingefuegten Kommentar-Referenzlaeufe
    selbst ein rPr/rStyle enthalten (rStyle="CommentReference"). Wuerde
    man waehrend des Durchlaufs sofort mutieren, wuerden diese neuen
    Laeufe bei der naechsten findall()-Abfrage faelschlich erneut als
    Fundstelle (Zeichenformatvorlage) erkannt.
    """
    root = etree.fromstring(xml_bytes)

    paragraph_targets: list[tuple[etree._Element, int]] = []
    run_targets: list[tuple[etree._Element, etree._Element, list[int]]] = []
    cell_targets: list[tuple[etree._Element, int]] = []

    # --- Lesephase ---
    for p in root.iter(qn("w:p")):
        text_ctx = _text_snippet(p)

        ppr = p.find(qn("w:pPr"))
        if ppr is not None:
            shd = ppr.find(qn("w:shd"))
            if shd is not None:
                wert = _schattierungswert(shd)
                if wert:
                    cid = next_id()
                    comments.append(
                        CommentEntry(cid, _comment_text("Absatz-Schattierung", wert, "", text_ctx))
                    )
                    paragraph_targets.append((p, cid))

        for r in p.findall(qn("w:r")):
            rpr = r.find(qn("w:rPr"))
            if rpr is None:
                continue
            run_text_ctx = _text_snippet(r)
            comment_ids: list[int] = []

            highlight = rpr.find(qn("w:highlight"))
            if highlight is not None:
                val = highlight.get(qn("w:val")) or "?"
                cid = next_id()
                comments.append(
                    CommentEntry(cid, _comment_text("Hervorhebung", val, "", run_text_ctx))
                )
                comment_ids.append(cid)

            shd = rpr.find(qn("w:shd"))
            if shd is not None:
                wert = _schattierungswert(shd)
                if wert:
                    cid = next_id()
                    comments.append(
                        CommentEntry(cid, _comment_text("Zeichen-Schattierung", wert, "", run_text_ctx))
                    )
                    comment_ids.append(cid)

            color = rpr.find(qn("w:color"))
            if color is not None:
                val = color.get(qn("w:val"))
                if val and val.lower() != "auto":
                    cid = next_id()
                    comments.append(
                        CommentEntry(cid, _comment_text("Textfarbe", val, "", run_text_ctx))
                    )
                    comment_ids.append(cid)

            rstyle = rpr.find(qn("w:rStyle"))
            if rstyle is not None:
                style_id = rstyle.get(qn("w:val"))
                name = _style_name(styles_root, style_id)
                cid = next_id()
                comments.append(
                    CommentEntry(
                        cid, _comment_text("Zeichenformatvorlage", name, name, run_text_ctx)
                    )
                )
                comment_ids.append(cid)

            if comment_ids:
                run_targets.append((p, r, comment_ids))

    for tbl in root.iter(qn("w:tbl")):
        for tr in tbl.findall(qn("w:tr")):
            for tc in tr.findall(qn("w:tc")):
                tcpr = tc.find(qn("w:tcPr"))
                if tcpr is None:
                    continue
                shd = tcpr.find(qn("w:shd"))
                if shd is None:
                    continue
                wert = _schattierungswert(shd)
                if wert:
                    text_ctx = _text_snippet(tc)
                    cid = next_id()
                    comments.append(
                        CommentEntry(
                            cid, _comment_text("Tabellenzellen-Schattierung", wert, "", text_ctx)
                        )
                    )
                    cell_targets.append((tc, cid))

    # --- Schreibphase: Reihenfolge so gewaehlt, dass sich die Einfuegungen
    # nicht gegenseitig stoeren (Absatz-/Zellrand zuerst, dann Laeufe innen). ---
    for p, cid in paragraph_targets:
        _wrap_paragraph_with_comment(p, cid)

    for p, r, ids_for_run in run_targets:
        for cid in ids_for_run:
            _wrap_run_with_comment(p, r, cid)

    for tc, cid in cell_targets:
        _wrap_cell_with_comment(tc, cid)

    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def _schattierungswert(shd: etree._Element) -> str | None:
    fill = shd.get(qn("w:fill"))
    val = shd.get(qn("w:val"))
    if fill and fill.lower() != "auto":
        return f"fill={fill}"
    if val and val.lower() != "clear":
        return f"val={val}"
    return None


def _build_comments_xml(comments: list[CommentEntry]) -> bytes:
    root = etree.Element(qn("w:comments"), nsmap={"w": NS["w"]})
    now = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    for entry in comments:
        comment = etree.SubElement(root, qn("w:comment"))
        comment.set(qn("w:id"), str(entry.comment_id))
        comment.set(qn("w:author"), COMMENT_AUTHOR)
        comment.set(qn("w:date"), now)
        comment.set(qn("w:initials"), COMMENT_INITIALS)
        p = etree.SubElement(comment, qn("w:p"))
        r = etree.SubElement(p, qn("w:r"))
        t = etree.SubElement(r, qn("w:t"))
        t.set(XML_SPACE_ATTR, "preserve")
        t.text = entry.text
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def _ensure_comments_content_type(xml_bytes: bytes | None) -> bytes:
    if xml_bytes is None:
        root = etree.Element(f"{{{CONTENT_TYPES_NS}}}Types", nsmap={None: CONTENT_TYPES_NS})
    else:
        root = etree.fromstring(xml_bytes)
    vorhanden = root.find(f"{{{CONTENT_TYPES_NS}}}Override[@PartName='/word/comments.xml']")
    if vorhanden is None:
        override = etree.SubElement(root, f"{{{CONTENT_TYPES_NS}}}Override")
        override.set("PartName", "/word/comments.xml")
        override.set("ContentType", COMMENTS_CONTENT_TYPE)
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def _ensure_comments_relationship(xml_bytes: bytes | None) -> bytes:
    if xml_bytes is None:
        root = etree.Element(f"{{{RELS_NS}}}Relationships", nsmap={None: RELS_NS})
    else:
        root = etree.fromstring(xml_bytes)

    for rel in root.findall(f"{{{RELS_NS}}}Relationship"):
        if rel.get("Type") == COMMENTS_REL_TYPE:
            return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)

    max_id = 0
    for rel in root.findall(f"{{{RELS_NS}}}Relationship"):
        rid = rel.get("Id", "")
        if rid.startswith("rId") and rid[3:].isdigit():
            max_id = max(max_id, int(rid[3:]))

    neue_rel = etree.SubElement(root, f"{{{RELS_NS}}}Relationship")
    neue_rel.set("Id", f"rId{max_id + 1}")
    neue_rel.set("Type", COMMENTS_REL_TYPE)
    neue_rel.set("Target", "comments.xml")
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)


def annotate_docx(quelle: Path, ziel: Path) -> int:
    """
    Erstellt unter 'ziel' eine Kopie von 'quelle', in der alle in
    word/document.xml gefundenen Formatierungen als Word-Kommentare
    markiert sind.

    'quelle' wird nur lesend geoeffnet und bleibt unveraendert.

    Gibt die Anzahl der eingefuegten Kommentare zurueck.
    """
    with zipfile.ZipFile(quelle, "r") as zin:
        namen = list(zin.namelist())
        inhalte = {name: zin.read(name) for name in namen}

    if "word/document.xml" not in inhalte:
        raise ValueError(
            "Kein word/document.xml gefunden - ist dies eine gueltige .docx-Datei?"
        )

    styles_root = None
    if "word/styles.xml" in inhalte:
        styles_root = etree.fromstring(inhalte["word/styles.xml"])

    comments: list[CommentEntry] = []
    zaehler = {"n": 0}

    def next_id() -> int:
        cid = zaehler["n"]
        zaehler["n"] += 1
        return cid

    inhalte["word/document.xml"] = _annotate_document_xml(
        inhalte["word/document.xml"], styles_root, comments, next_id
    )

    if comments:
        inhalte["word/comments.xml"] = _build_comments_xml(comments)
        if "word/comments.xml" not in namen:
            namen.append("word/comments.xml")

        inhalte["[Content_Types].xml"] = _ensure_comments_content_type(
            inhalte.get("[Content_Types].xml")
        )

        rels_name = "word/_rels/document.xml.rels"
        inhalte[rels_name] = _ensure_comments_relationship(inhalte.get(rels_name))
        if rels_name not in namen:
            namen.append(rels_name)

    ziel = Path(ziel)
    ziel.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(ziel, "w", zipfile.ZIP_DEFLATED) as zout:
        for name in namen:
            zout.writestr(name, inhalte[name])

    return len(comments)

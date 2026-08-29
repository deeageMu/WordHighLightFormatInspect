from docx_format_scanner import Finding, ScanResult, _format_bericht, scan_document_xml


def test_scan_document_xml_includes_xml_snippet_for_formatting():
    xml = b'''<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p>
    <w:r>
      <w:rPr>
        <w:highlight w:val="yellow"/>
        <w:color w:val="FF0000"/>
      </w:rPr>
      <w:t>Hallo</w:t>
    </w:r>
  </w:p>
</w:document>'''

    findings, _ = scan_document_xml(xml, "word/document.xml", None)

    assert len(findings) >= 2
    assert any("<w:rPr>" in finding.xml_snippet for finding in findings)
    assert any("<w:t>Hallo</w:t>" in finding.xml_snippet for finding in findings)


def test_format_bericht_only_shows_xml_when_flag_is_enabled():
    result = ScanResult(
        findings=[
            Finding(
                kategorie="Hervorhebung",
                wert="yellow",
                ort="word/document.xml: Absatz 1, Lauf 1",
                textausschnitt="Hallo",
                xml_snippet="<w:r><w:rPr><w:highlight w:val=\"yellow\"/></w:rPr><w:t>Hallo</w:t></w:r>",
            )
        ]
    )

    report_default = _format_bericht(result, "sample.docx")
    report_with_xml = _format_bericht(result, "sample.docx", out_xml=True)
    report_with_full_xml = _format_bericht(result, "sample.docx", out_xml=True, xml_max_chars=0)

    assert "XML:" not in report_default
    assert "XML:" in report_with_xml
    assert "<w:rPr>" in report_with_xml
    assert "<w:t>Hallo</w:t>" in report_with_full_xml
    assert "..." not in report_with_full_xml

from __future__ import annotations

import tempfile
import unittest
import zipfile
from unittest.mock import patch
from pathlib import Path
from xml.etree import ElementTree as ET

import stamp_reports
from stamp_reports import MEDIA_CACHET, MEDIA_SIGNATURE, process_source, stamp_docx


REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"


class StampReportsTests(unittest.TestCase):
    def test_stamp_docx_inserts_cachet_and_signature_images(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "report.docx"
            _make_stamp_docx(path)

            cachet = Path(tmp) / "cachet.png"
            cachet.write_bytes(b"cachet-bytes")
            signature = Path(tmp) / "signature.png"
            signature.write_bytes(b"signature-bytes")

            updated, message = stamp_docx(path, cachet.read_bytes(), signature.read_bytes())
            self.assertTrue(updated)
            self.assertEqual(message, "stamped")

            with zipfile.ZipFile(path) as docx:
                names = set(docx.namelist())
                self.assertIn(MEDIA_CACHET, names)
                self.assertIn(MEDIA_SIGNATURE, names)

                rels_xml = ET.fromstring(docx.read("word/_rels/document.xml.rels"))
                targets = {
                    rel.attrib.get("Target", "")
                    for rel in rels_xml.findall(f".//{{{REL_NS}}}Relationship")
                }
                self.assertIn("media/cachet_gthconsult.png", targets)
                self.assertIn("media/signature_gthconsult.png", targets)

    def test_process_source_uses_default_signature_from_assets(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            source = Path(tmp) / "source"
            source.mkdir()
            _make_stamp_docx(source / "report.docx")

            cachet = Path(tmp) / "cachet.png"
            cachet.write_bytes(b"cachet-bytes")
            signature = Path(tmp) / "signature.png"
            signature.write_bytes(b"signature-bytes")
            output = Path(tmp) / "out"

            with patch.object(stamp_reports, "DEFAULT_CACHET", cachet), patch.object(
                stamp_reports, "DEFAULT_SIGNATURE", signature
            ):
                results = process_source(source, output=output)

            self.assertTrue(any(line.startswith("OK\t") for line in results))
            with zipfile.ZipFile(output / "report.docx") as docx:
                names = set(docx.namelist())
                self.assertIn(MEDIA_SIGNATURE, names)


def _make_stamp_docx(path: Path) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    document = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:tbl>
      <w:tr>
        <w:tc><w:p><w:r><w:t>Inspecteur</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>Administration GTHCONSULT</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:p/></w:tc>
        <w:tc><w:p/></w:tc>
      </w:tr>
    </w:tbl>
    <w:sectPr/>
  </w:body>
</w:document>"""
    rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""
    content_types = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"""
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as docx:
        docx.writestr("[Content_Types].xml", content_types)
        docx.writestr("_rels/.rels", rels)
        docx.writestr("word/document.xml", document)
        docx.writestr("word/_rels/document.xml.rels", rels)


if __name__ == "__main__":
    unittest.main()

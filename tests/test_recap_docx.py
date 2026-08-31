from __future__ import annotations

import tempfile
import unittest
import zipfile
from pathlib import Path
from xml.etree import ElementTree as ET

from recap_docx import ReportSummary, _xls_summary_cells, build_workbook, discover_docx_directories


MAIN_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"


class RecapDocxTests(unittest.TestCase):
    def test_recursive_scan_creates_only_direct_docx_directory_sheets(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp) / "O2X 2026"
            _make_docx(root / "Pont ciseaux 2026" / "a.docx", ["Rapport A", "Controle du pont ciseaux."])
            _make_docx(root / "Pont ciseaux 2026" / "b.docx", ["Rapport B"])
            _make_docx(root / "Atelier" / "Zone A" / "Palan a chaine 2026" / "c.docx", ["Rapport C"])
            _make_docx(root / "Atelier" / "Zone A" / "Palan a chaine 2026" / "d.docx", ["Rapport D"])
            _make_docx(root / "Ponts 2 colonnes 2026" / "e.docx", ["Rapport E"])

            matches = discover_docx_directories(root)
            discovered_names = [path.name for path, _files in matches]
            self.assertCountEqual(
                discovered_names,
                ["Pont ciseaux 2026", "Palan a chaine 2026", "Ponts 2 colonnes 2026"],
            )
            self.assertNotIn("O2X 2026", discovered_names)
            self.assertNotIn("Atelier", discovered_names)
            self.assertNotIn("Zone A", discovered_names)

            output = Path(tmp) / "recap.xlsx"
            created_sheets = build_workbook(root, output)

            self.assertCountEqual(
                created_sheets,
                ["Pont ciseaux 2026", "Palan a chaine 2026", "Ponts 2 colonnes 2026"],
            )
            self.assertCountEqual(_sheet_names(output), created_sheets)

    def test_scan_continues_below_directory_that_also_contains_docx(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            _make_docx(root / "Parent" / "parent.docx", ["Parent doc"])
            _make_docx(root / "Parent" / "Child" / "child.docx", ["Child doc"])

            output = root / "recap.xlsx"
            created_sheets = build_workbook(root, output)

            self.assertEqual(created_sheets, ["Parent", "Child"])

    def test_reference_column_is_first_in_output(self) -> None:
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            _make_docx(root / "Folder" / "report.docx", ["Doc", "More text"])
            output = root / "recap.xls"
            build_workbook(root, output)

            headers = _sheet_row_values(output, 3)
            self.assertEqual(headers[0], "R\u00e9f\u00e9rence du rapport")
            self.assertEqual(headers[1], "n\u00b0")

    def test_criticite_column_is_colored_and_placed_correctly(self) -> None:
        blocked_report = ReportSummary(
            file_name="blocked.docx",
            relative_path="Folder/blocked.docx",
            department="Dept",
            equipment="Equip",
            title="Title",
            report_reference="REF-1",
            verification_date="",
            manufacturer="Maker",
            manufacturer_type="",
            service_year="",
            serial_number="SN-1",
            internal_number="REP-1",
            location="Site",
            equipment_type="Type",
            capacity_kg="1000",
            verification_type="",
            blocking_observations="Observation bloquante.",
            nonblocking_observations="",
            complementary_observations="",
            conclusion="",
        )
        blocked_cells = _xls_summary_cells(1, blocked_report)
        self.assertEqual(blocked_cells[11][0], "Bloqu\u00e9")
        self.assertEqual(blocked_cells[11][1], "Blocked")
        self.assertEqual(blocked_cells[10][1], "Cell")

        clear_report = ReportSummary(
            file_name="clear.docx",
            relative_path="Folder/clear.docx",
            department="Dept",
            equipment="Equip",
            title="Title",
            report_reference="REF-2",
            verification_date="",
            manufacturer="Maker",
            manufacturer_type="",
            service_year="",
            serial_number="SN-2",
            internal_number="REP-2",
            location="Site",
            equipment_type="Type",
            capacity_kg="1000",
            verification_type="",
            blocking_observations="",
            nonblocking_observations="Observation non bloquante.",
            complementary_observations="",
            conclusion="",
        )
        clear_cells = _xls_summary_cells(2, clear_report)
        self.assertEqual(clear_cells[11][0], "Non bloqu\u00e9")
        self.assertEqual(clear_cells[11][1], "NotBlocked")


def _make_docx(path: Path, paragraphs: list[str]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    body = "".join(
        f"<w:p><w:r><w:t>{text}</w:t></w:r></w:p>"
        for text in paragraphs
    )
    document = f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>{body}<w:sectPr/></w:body>
</w:document>"""
    core = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/">
  <dc:title>Fixture</dc:title>
</cp:coreProperties>"""
    rels = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""
    content_types = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
</Types>"""
    with zipfile.ZipFile(path, "w", compression=zipfile.ZIP_DEFLATED) as docx:
        docx.writestr("[Content_Types].xml", content_types)
        docx.writestr("_rels/.rels", rels)
        docx.writestr("word/document.xml", document)
        docx.writestr("docProps/core.xml", core)


def _sheet_names(path: Path) -> list[str]:
    with zipfile.ZipFile(path) as workbook:
        xml = workbook.read("xl/workbook.xml")
    root = ET.fromstring(xml)
    return [sheet.attrib["name"] for sheet in root.findall(f".//{{{MAIN_NS}}}sheet")]


def _sheet_row_values(path: Path, row_index: int) -> list[str]:
    ns = {"ss": "urn:schemas-microsoft-com:office:spreadsheet"}
    root = ET.fromstring(Path(path).read_text(encoding="utf-8"))
    row = root.findall(".//ss:Worksheet/ss:Table/ss:Row", ns)[row_index - 1]
    values = []
    for cell in row.findall("ss:Cell", ns):
        data = cell.find("ss:Data", ns)
        values.append("" if data is None or data.text is None else data.text)
    return values


if __name__ == "__main__":
    unittest.main()

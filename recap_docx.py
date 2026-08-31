from __future__ import annotations

import argparse
import html
import os
import re
import shutil
import subprocess
import tempfile
import textwrap
import zipfile
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, NamedTuple
from xml.etree import ElementTree as ET


EXCEL_MAX_SHEET_NAME = 31
INVALID_SHEET_CHARS = re.compile(r"[\[\]\:\*\?\/\\]")
REPORT_TITLE = "récapitulatif des observations relevées en 2026"
DOC_NS = {
    "cp": "http://schemas.openxmlformats.org/package/2006/metadata/core-properties",
    "dc": "http://purl.org/dc/elements/1.1/",
    "dcterms": "http://purl.org/dc/terms/",
    "w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main",
}


class LogoAsset(NamedTuple):
    extension: str
    content_type: str
    data: bytes


@dataclass(frozen=True)
class ReportSummary:
    file_name: str
    relative_path: str
    department: str
    equipment: str
    title: str
    report_reference: str
    verification_date: str
    manufacturer: str
    manufacturer_type: str
    service_year: str
    serial_number: str
    internal_number: str
    location: str
    equipment_type: str
    capacity_kg: str
    verification_type: str
    blocking_observations: str
    nonblocking_observations: str
    complementary_observations: str
    conclusion: str


def parse_docx(path: Path, root: Path, department: str = "") -> ReportSummary:
    with zipfile.ZipFile(path) as package:
        paragraphs = []
        for name in sorted(package.namelist()):
            if name.startswith("word/") and name.endswith(".xml"):
                paragraphs.extend(_extract_paragraphs(package.read(name)))

    fields = _extract_report_fields(paragraphs)
    equipment_type = fields.get("Type d'appareil", "")
    serial_number = _clean_value(fields.get("Numero(s) de serie (plaque constructeur)", ""))
    title = _make_report_title(equipment_type, serial_number, fields.get("Equipement", ""), path)

    return ReportSummary(
        file_name=path.name,
        relative_path=str(path.relative_to(root)),
        department=department,
        equipment=fields.get("Equipement", ""),
        title=title,
        report_reference=fields.get("Reference du rapport", ""),
        verification_date=fields.get("Date(s) de(s) verification(s)", ""),
        manufacturer=fields.get("Constructeur", ""),
        manufacturer_type=fields.get("Type constructeur (plaque)", ""),
        service_year=fields.get("Annee de mise en service (plaque constructeur)", ""),
        serial_number=serial_number,
        internal_number=fields.get("Numero(s) interne(s)", ""),
        location=fields.get("Localisation de(s) l'appareil (s) lors de la visite", ""),
        equipment_type=equipment_type,
        capacity_kg=fields.get("Capacite Kg", ""),
        verification_type=fields.get("Type de verification", ""),
        blocking_observations=fields.get("Observations bloquantes", ""),
        nonblocking_observations=fields.get("Observations non bloquantes", ""),
        complementary_observations=fields.get("Observations complementaires", ""),
        conclusion=fields.get("Conclusion", ""),
    )


def discover_docx_directories(root: Path) -> list[tuple[Path, list[Path]]]:
    matches: list[tuple[Path, list[Path]]] = []
    for current, dirs, files in os.walk(root):
        dirs[:] = [directory for directory in dirs if directory not in {"__MACOSX", ".git"}]
        dirs.sort(key=str.casefold)
        current_path = Path(current)
        docx_files = sorted(
            current_path / name
            for name in files
            if name.lower().endswith(".docx") and not name.startswith("~$")
        )
        if docx_files:
            matches.append((current_path, docx_files))
    return matches


def build_workbook(source: Path, output: Path) -> list[str]:
    source = source.resolve()
    output = output.resolve()
    with _source_root(source) as root:
        directories = discover_docx_directories(root)
        sheets: list[tuple[str, list[ReportSummary]]] = []
        used_sheet_names: set[str] = set()

        for directory, docx_files in directories:
            sheet_name = make_unique_sheet_name(directory.name, used_sheet_names)
            used_sheet_names.add(sheet_name)
            department = _department_from_path(directory, root)
            reports = [parse_docx(docx, root, department) for docx in docx_files]
            sheets.append((sheet_name, reports))

        if not sheets:
            raise ValueError(f"Aucun fichier .docx trouve dans: {root}")

        if output.suffix.lower() in {".xls", ".xml"}:
            write_xls_xml(output, sheets)
        else:
            write_xlsx(output, sheets)
        return [sheet_name for sheet_name, _ in sheets]


def _department_from_path(directory: Path, root: Path) -> str:
    parts = directory.relative_to(root).parts
    if len(parts) >= 2:
        return re.sub(r"\s*2026\s*$", "", parts[1], flags=re.IGNORECASE).strip()
    if parts:
        return re.sub(r"\s*2026\s*$", "", parts[0], flags=re.IGNORECASE).strip()
    return ""


def make_unique_sheet_name(directory_name: str, used: set[str]) -> str:
    base = INVALID_SHEET_CHARS.sub(" ", _sanitize_excel_text(directory_name)).strip().strip("'")
    base = re.sub(r"\s+", " ", base) or "Feuille"
    base = base[:EXCEL_MAX_SHEET_NAME]
    if base not in used:
        return base

    index = 2
    while True:
        suffix = f" ({index})"
        candidate = f"{base[: EXCEL_MAX_SHEET_NAME - len(suffix)]}{suffix}"
        if candidate not in used:
            return candidate
        index += 1


def write_xlsx(output: Path, sheets: list[tuple[str, list[ReportSummary]]], logo: LogoAsset | None = None) -> None:
    output.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(output, "w", compression=zipfile.ZIP_DEFLATED) as xlsx:
        _writestr(xlsx, "[Content_Types].xml", _content_types(len(sheets), logo))
        _writestr(xlsx, "_rels/.rels", _root_rels())
        _writestr(xlsx, "xl/workbook.xml", _workbook_xml([name for name, _ in sheets]))
        _writestr(xlsx, "xl/_rels/workbook.xml.rels", _workbook_rels(len(sheets)))
        _writestr(xlsx, "xl/styles.xml", _styles_xml())
        _writestr(xlsx, "docProps/core.xml", _package_core_xml())
        _writestr(xlsx, "docProps/app.xml", _package_app_xml(len(sheets)))

        if logo:
            xlsx.writestr(f"xl/media/logo.{logo.extension}", logo.data)

        for index, (sheet_name, rows) in enumerate(sheets, start=1):
            _writestr(xlsx, f"xl/worksheets/sheet{index}.xml", _worksheet_xml(rows, bool(logo)))
            if logo:
                _writestr(xlsx, f"xl/worksheets/_rels/sheet{index}.xml.rels", _worksheet_rels(index))
                _writestr(xlsx, f"xl/drawings/drawing{index}.xml", _drawing_xml())
                _writestr(xlsx, f"xl/drawings/_rels/drawing{index}.xml.rels", _drawing_rels(logo.extension))


def write_xls_xml(output: Path, sheets: list[tuple[str, list[ReportSummary]]]) -> None:
    output.parent.mkdir(parents=True, exist_ok=True)
    worksheets = "\n".join(_xls_worksheet_xml(sheet_name, reports) for sheet_name, reports in sheets)
    content = f"""<?xml version="1.0" encoding="UTF-8"?>
<?mso-application progid="Excel.Sheet"?>
<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:o="urn:schemas-microsoft-com:office:office"
 xmlns:x="urn:schemas-microsoft-com:office:excel"
 xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:html="http://www.w3.org/TR/REC-html40">
 <Styles>
  <Style ss:ID="Default" ss:Name="Normal">
   <Alignment ss:Vertical="Top" ss:WrapText="1"/>
   <Font ss:FontName="Calibri" ss:Size="11"/>
  </Style>
  <Style ss:ID="Title">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
   <Font ss:FontName="Calibri" ss:Size="16" ss:Bold="1" ss:Color="#FFFFFF"/>
   <Interior ss:Color="#1F4E78" ss:Pattern="Solid"/>
  </Style>
  <Style ss:ID="Logo">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center"/>
   <Font ss:FontName="Calibri" ss:Size="14" ss:Bold="1" ss:Color="#1F4E78"/>
  </Style>
  <Style ss:ID="Header">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
   <Font ss:FontName="Calibri" ss:Size="11" ss:Bold="1" ss:Color="#FFFFFF"/>
   <Interior ss:Color="#4472C4" ss:Pattern="Solid"/>
   <Borders>{_xls_border_xml()}</Borders>
  </Style>
  <Style ss:ID="Cell">
   <Alignment ss:Vertical="Top" ss:WrapText="1"/>
   <Borders>{_xls_border_xml()}</Borders>
  </Style>
  <Style ss:ID="Blocked">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
   <Font ss:FontName="Calibri" ss:Size="11" ss:Bold="1"/>
   <Interior ss:Color="#FFC7CE" ss:Pattern="Solid"/>
   <Borders>{_xls_border_xml()}</Borders>
  </Style>
  <Style ss:ID="NotBlocked">
   <Alignment ss:Horizontal="Center" ss:Vertical="Center" ss:WrapText="1"/>
   <Font ss:FontName="Calibri" ss:Size="11" ss:Bold="1"/>
   <Interior ss:Color="#C6EFCE" ss:Pattern="Solid"/>
   <Borders>{_xls_border_xml()}</Borders>
  </Style>
 </Styles>
{worksheets}
</Workbook>
"""
    output.write_text(content, encoding="utf-8")


def _xls_worksheet_xml(sheet_name: str, reports: list[ReportSummary]) -> str:
    widths = [88, 42, 95, 155, 105, 125, 105, 70, 205, 115, 260, 85, 205, 240]
    columns = "\n".join(f'  <Column ss:Width="{width}"/>' for width in widths)
    rows = [
        _xls_row([
            ("GTHCONSULT", "Logo", 1),
            (REPORT_TITLE, "Title", 12),
        ], height=42),
        _xls_row([("", "Default", 13)], height=8),
        _xls_row([(header, "Header", 1) for header in _summary_headers()], height=34),
    ]
    rows.extend(
        _xls_row(_xls_summary_cells(index, report), height=_xls_report_row_height(report))
        for index, report in enumerate(reports, start=1)
    )
    return f""" <Worksheet ss:Name="{_xml_escape(sheet_name)}">
  <Table>
{columns}
{''.join(rows)}
  </Table>
  <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">
   <FreezePanes/>
   <FrozenNoSplit/>
   <SplitHorizontal>3</SplitHorizontal>
   <TopRowBottomPane>3</TopRowBottomPane>
   <ActivePane>2</ActivePane>
  </WorksheetOptions>
 </Worksheet>"""


def _summary_headers() -> list[str]:
    return [
        "Référence du rapport",
        "n°",
        "Département",
        "Equipement",
        "Repère utilisateur",
        "Constructeur",
        "N° de fabrication",
        "Capacité Kg",
        "Lieu d'implantation",
        "Etat des lieux après CR",
        "Actions recommandées",
        "Criticité",
        "Observations supplémentaires",
        "Conclusion",
    ]


def _xls_summary_cells(index: int, report: ReportSummary) -> list[tuple[object, str, int]]:
    values = _summary_row(index, report)
    cells: list[tuple[object, str, int]] = []
    for column_index, value in enumerate(values, start=1):
        style = "Cell"
        if column_index == 12:
            style = "Blocked" if value == "Bloqué" else "NotBlocked"
        cells.append((value, style, 1))
    return cells


def _xls_report_row_height(report: ReportSummary) -> int:
    action = report.blocking_observations if _has_real_observation(report.blocking_observations) else report.nonblocking_observations
    long_text = "\n".join(
        [
            _format_observation_list(action),
            _format_observation_list(report.complementary_observations),
            _format_conclusion(report.conclusion),
        ]
    )
    line_count = max(1, long_text.count("\n") + 1)
    return min(150, max(48, 22 + line_count * 14))


def _xls_row(cells: list[tuple[object, str, int]], height: int | None = None) -> str:
    height_attr = f' ss:Height="{height}"' if height else ""
    content = "".join(_xls_cell(value, style, merge) for value, style, merge in cells)
    return f'   <Row{height_attr}>{content}</Row>\n'


def _xls_cell(value: object, style: str, merge_across: int = 1) -> str:
    merge_attr = f' ss:MergeAcross="{merge_across - 1}"' if merge_across > 1 else ""
    data_type = "Number" if isinstance(value, int) else "String"
    return f'<Cell ss:StyleID="{style}"{merge_attr}><Data ss:Type="{data_type}">{_xml_escape(str(value))}</Data></Cell>'


def _xls_border_xml() -> str:
    return (
        '<Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#D9E2F3"/>'
        '<Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#D9E2F3"/>'
        '<Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#D9E2F3"/>'
        '<Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="1" ss:Color="#D9E2F3"/>'
    )


def _extract_paragraphs(document_xml: bytes) -> list[str]:
    root = ET.fromstring(document_xml)
    paragraphs: list[str] = []
    for paragraph in root.findall(".//w:p", DOC_NS):
        parts: list[str] = []
        for node in paragraph.iter():
            tag = _local_name(node.tag)
            if tag == "t" and node.text:
                parts.append(node.text)
            elif tag == "tab":
                parts.append(" ")
            elif tag == "br":
                parts.append(" ")
        text = re.sub(r"\s+", " ", "".join(parts)).strip()
        if text:
            paragraphs.append(text)
    return paragraphs


def _extract_report_fields(paragraphs: list[str]) -> dict[str, str]:
    normalized = [_normalize_label(text) for text in paragraphs]
    fields: dict[str, str] = {}

    label_map = {
        "equipement": "Equipement",
        "reference du rapport": "Reference du rapport",
        "n rapport": "Reference du rapport",
        "date(s) de(s) verification(s)": "Date(s) de(s) verification(s)",
        "constructeur": "Constructeur",
        "identification constructeur": "Constructeur",
        "type constructeur (plaque)": "Type constructeur (plaque)",
        "annee de mise en service (plaque constructeur)": "Annee de mise en service (plaque constructeur)",
        "numero(s) de serie (plaque constructeur)": "Numero(s) de serie (plaque constructeur)",
        "n de serie": "Numero(s) de serie (plaque constructeur)",
        "numero(s) interne(s)": "Numero(s) interne(s)",
        "localisation de(s) l'appareil (s) lors de la visite": "Localisation de(s) l'appareil (s) lors de la visite",
        "lieu de verification": "Localisation de(s) l'appareil (s) lors de la visite",
        "type d'appareil": "Type d'appareil",
        "type de verification": "Type de verification",
    }
    stop_labels = set(label_map) | {
        "marque",
        "page",
        "annee",
        "reference client",
        "n affaire",
        "verificateur(s) agree",
        "accompagnes de",
        "date d'emission du rapport",
        "le present rapport comporte",
        "documentation technique constructeur (notice d'instructions, de montage, d'utilisation)",
        "essais en charge",
        "examen de montage et d'installation",
        "modification(s) apportee(s) ou autre(s) remarque(s) eventuelle(s) concernant l'appareil examine",
    }

    for index, label in enumerate(normalized):
        target = label_map.get(label)
        if target and target not in fields:
            fields[target] = _next_value(paragraphs, normalized, index, stop_labels)

    fields["Observations bloquantes"] = _extract_observation_after(
        paragraphs,
        normalized,
        "observations ne permettant pas l'utilisation de l'appareil",
    )
    fields["Observations non bloquantes"] = _extract_observation_after(
        paragraphs,
        normalized,
        "observations ne s'opposant pas a l'utilisation de l'appareil",
    )
    fields["Observations complementaires"] = _extract_section_text(
        paragraphs,
        normalized,
        "observations complementaires",
        {"e. conclusion", "conclusion"},
    )
    fields["Conclusion"] = _extract_section_text(
        paragraphs,
        normalized,
        "e. conclusion",
        {"gth-maq", "gthconsult", "sarl", "patente", "fixe", "n affaire", "n rapport", "page", "annee", "reference client"},
    ) or _extract_section_text(
        paragraphs,
        normalized,
        "conclusion",
        {"gth-maq", "gthconsult", "sarl", "patente", "fixe", "n affaire", "n rapport", "page", "annee", "reference client"},
    )
    fields["Capacite Kg"] = _extract_capacity_kg(paragraphs)
    return fields


def _extract_capacity_kg(paragraphs: list[str]) -> str:
    patterns = [
        r"\bCMU\s*\(kg\)\s*:\s*([0-9]+(?:[,.][0-9]+)?)",
        r"Charge maximale utile\s*\(kg\)\s*:\s*([0-9]+(?:[,.][0-9]+)?)",
        r"charge maximale d'utilisation\s*\(CMU\).*?([0-9]+(?:[,.][0-9]+)?)\s*kg",
        r"charge de\s*([0-9]+(?:[,.][0-9]+)?)\s*kg",
    ]
    for paragraph in paragraphs:
        normalized = _clean_value(paragraph)
        for pattern in patterns:
            match = re.search(pattern, normalized, flags=re.IGNORECASE)
            if match:
                return match.group(1).replace(",", ".")
    return ""


def _find_logo(docx_files: Iterable[Path]) -> LogoAsset | None:
    content_types = {
        ".jpeg": "image/jpeg",
        ".jpg": "image/jpeg",
        ".png": "image/png",
    }
    for docx in docx_files:
        try:
            with zipfile.ZipFile(docx) as package:
                media = [
                    name for name in package.namelist()
                    if name.startswith("word/media/") and Path(name).suffix.lower() in content_types
                ]
                if not media:
                    continue
                logo_name = sorted(media, key=_logo_priority)[0]
                extension = Path(logo_name).suffix.lower().lstrip(".")
                return LogoAsset(extension, content_types[f".{extension}"], package.read(logo_name))
        except zipfile.BadZipFile:
            continue
    return None


def _logo_priority(name: str) -> tuple[int, str]:
    file_name = Path(name).name.lower()
    if file_name.startswith("image2."):
        return (0, file_name)
    if "gth" in file_name or "consult" in file_name:
        return (1, file_name)
    return (2, file_name)


def _next_value(paragraphs: list[str], normalized: list[str], index: int, labels: set[str]) -> str:
    values: list[str] = []
    for cursor in range(index + 1, min(len(paragraphs), index + 5)):
        value = _clean_value(paragraphs[cursor])
        if not value:
            continue
        if _is_stop_label(normalized[cursor], labels) or _looks_like_section_heading(normalized[cursor]):
            break
        values.append(value)
        if len(values) >= 2:
            break
    return " ".join(values)


def _is_stop_label(value: str, labels: set[str]) -> bool:
    return value in labels or any(value.startswith(f"{label} ") or value.startswith(f"{label} :") for label in labels)


def _extract_observation_after(paragraphs: list[str], normalized: list[str], heading: str) -> str:
    try:
        start = normalized.index(heading)
    except ValueError:
        return ""

    values: list[str] = []
    for cursor in range(start + 1, min(len(paragraphs), start + 12)):
        current = normalized[cursor]
        if current in {"suite donnee", "obs. n", "obs n"}:
            continue
        if current.startswith("observations ne ") or current in {"observations complementaires", "e. conclusion", "conclusion"}:
            break
        value = _clean_value(paragraphs[cursor])
        if value and value != "-":
            values.append(value)
    return " ".join(values)


def _extract_section_text(
    paragraphs: list[str],
    normalized: list[str],
    heading: str,
    stop_headings: set[str],
) -> str:
    try:
        start = normalized.index(heading)
    except ValueError:
        return ""

    values: list[str] = []
    for cursor in range(start + 1, len(paragraphs)):
        if normalized[cursor] in stop_headings or any(normalized[cursor].startswith(prefix) for prefix in stop_headings):
            break
        value = _clean_value(paragraphs[cursor])
        if value:
            values.append(value)
    return textwrap.shorten(" ".join(values), width=800, placeholder="...")


def _make_report_title(equipment_type: str, serial_number: str, equipment: str, path: Path) -> str:
    clean_type = _clean_value(equipment_type).rstrip(",.")
    clean_serial = _clean_serial(serial_number)
    if clean_type and clean_serial:
        return f"{clean_type} - {clean_serial}"
    if equipment:
        return _clean_value(equipment).rstrip(",.")
    return path.stem


def _clean_serial(value: str) -> str:
    value = _clean_value(value).rstrip(",.")
    value = re.sub(r"^(n[°o]\s*)", "", value, flags=re.IGNORECASE)
    return value


def _clean_value(value: str) -> str:
    value = value.replace("\u00a0", " ")
    value = re.sub(r"\s+", " ", value).strip()
    return value


def _normalize_label(value: str) -> str:
    value = _clean_value(value).lower().strip(" :;,")
    value = re.sub(r"\bn[°º\ufffd]\s*", "n ", value)
    replacements = {
        "�": "e",
        "é": "e",
        "è": "e",
        "ê": "e",
        "à": "a",
        "â": "a",
        "ù": "u",
        "û": "u",
        "î": "i",
        "ï": "i",
        "ô": "o",
        "ç": "c",
        "’": "'",
        "°": "",
    }
    for source, target in replacements.items():
        value = value.replace(source, target)
    value = value.replace("n ", "n ")
    value = re.sub(r"\s+", " ", value)
    return value.strip(" :;,")


def _looks_like_section_heading(value: str) -> bool:
    return value in {
        "a. renseignements generaux",
        "renseignements generaux",
        "description de l'appareil verifie",
        "c. examen et essais de l'appareil",
        "d. liste recapitulatif des observations",
    }


def _source_root(source: Path):
    if source.is_dir():
        return _NullContext(source)
    if source.is_file() and source.suffix.lower() == ".zip":
        temp_dir = Path(tempfile.mkdtemp(prefix="recap_docx_"))
        with zipfile.ZipFile(source) as archive:
            archive.extractall(temp_dir)
        return _TempDirContext(temp_dir)
    if source.is_file() and source.suffix.lower() == ".rar":
        temp_dir = Path(tempfile.mkdtemp(prefix="recap_docx_"))
        tar = shutil.which("tar")
        if not tar:
            raise ValueError("Archive .rar detectee, mais tar.exe est introuvable pour l'extraction.")
        result = subprocess.run(
            [tar, "-xf", str(source), "-C", str(temp_dir)],
            check=False,
            capture_output=True,
            text=True,
        )
        if result.returncode != 0:
            shutil.rmtree(temp_dir, ignore_errors=True)
            message = (result.stderr or result.stdout or "extraction impossible").strip()
            raise ValueError(f"Impossible d'extraire l'archive .rar: {message}")
        return _TempDirContext(temp_dir)
    raise ValueError(f"Source introuvable ou non prise en charge: {source}")


class _NullContext:
    def __init__(self, value: Path) -> None:
        self.value = value

    def __enter__(self) -> Path:
        return self.value

    def __exit__(self, *_exc) -> None:
        return None


class _TempDirContext:
    def __init__(self, value: Path) -> None:
        self.value = value

    def __enter__(self) -> Path:
        return self.value

    def __exit__(self, *_exc) -> None:
        shutil.rmtree(self.value, ignore_errors=True)


def _worksheet_xml(reports: list[ReportSummary], has_logo: bool = False) -> str:
    headers = [
        "n°",
        "Département",
        "Equipement",
        "Repère utilisateur",
        "Constructeur",
        "N° de fabrication",
        "Capacité Kg",
        "Lieu d'implantation",
        "Etat des lieux après CR",
        "Actions recommandées",
        "Criticité",
        "Observations supplémentaires",
        "Conclusion",
    ]
    title_row = ["GTHCONSULT", REPORT_TITLE, "", "", "", "", "", "", "", "", "", "", ""]
    matrix: list[list[object]] = [
        title_row,
        ["", "", "", "", "", "", "", "", "", "", "", "", ""],
        headers,
    ]
    matrix.extend(_summary_row(index, report) for index, report in enumerate(reports, start=1))

    sheet_data = "\n".join(_row_xml(row_index, values) for row_index, values in enumerate(matrix, start=1))
    title_merge = "B1:N2"
    drawing = '\n  <drawing r:id="rId1"/>' if has_logo else ""
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheetViews><sheetView workbookViewId="0"><pane ySplit="3" topLeftCell="A4" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>
  <cols>
    <col min="1" max="1" width="6" customWidth="1"/>
    <col min="2" max="2" width="18" customWidth="1"/>
    <col min="3" max="3" width="28" customWidth="1"/>
    <col min="4" max="4" width="18" customWidth="1"/>
    <col min="5" max="5" width="22" customWidth="1"/>
    <col min="6" max="6" width="18" customWidth="1"/>
    <col min="7" max="7" width="12" customWidth="1"/>
    <col min="8" max="8" width="34" customWidth="1"/>
    <col min="9" max="9" width="18" customWidth="1"/>
    <col min="10" max="10" width="46" customWidth="1"/>
    <col min="11" max="11" width="14" customWidth="1"/>
    <col min="12" max="13" width="34" customWidth="1"/>
  </cols>
  <sheetData>
{sheet_data}
  </sheetData>
  <autoFilter ref="A3:N{max(3, len(matrix))}"/>
  <mergeCells count="1"><mergeCell ref="{title_merge}"/></mergeCells>{drawing}
</worksheet>"""


def _summary_row(index: int, report: ReportSummary) -> list[object]:
    blocking = _has_real_observation(report.blocking_observations)
    action = report.blocking_observations if blocking else report.nonblocking_observations
    if not _has_real_observation(action):
        action = "R.A.S"
    return [
        report.report_reference,
        index,
        report.department,
        _clean_value(report.equipment_type).rstrip(",.") or report.equipment,
        report.internal_number or "--",
        report.manufacturer,
        report.serial_number,
        report.capacity_kg,
        report.location,
        "--",
        _format_observation_list(action),
        "Bloqué" if blocking else "Non bloqué",
        _format_observation_list(report.complementary_observations),
        _format_conclusion(report.conclusion),
    ]


def _has_real_observation(value: str) -> bool:
    normalized = _normalize_label(value)
    if not normalized:
        return False
    normalized = normalized.replace(".", "").strip(" -")
    return normalized not in {"sans objet", "ras", "r a s", "--"}


def _format_observation_list(value: str) -> str:
    value = _clean_value(value)
    if not _has_real_observation(value):
        return "R.A.S"

    matches = list(re.finditer(r"\bO\s*(\d+)\b", value, flags=re.IGNORECASE))
    if matches:
        lines = []
        for index, match in enumerate(matches):
            start = match.end()
            end = matches[index + 1].start() if index + 1 < len(matches) else len(value)
            detail = value[start:end].strip(" -:;.")
            if detail:
                lines.append(f"- O{match.group(1)} : {detail}.")
        if lines:
            return "\n".join(lines)

    parts = _split_readable_sentences(value)
    return "\n".join(f"- {part}" for part in parts) if parts else value


def _format_conclusion(value: str) -> str:
    value = _clean_value(value)
    if not value:
        return ""
    parts = _split_readable_sentences(value)
    return "\n".join(f"- {part}" for part in parts) if parts else value


def _split_readable_sentences(value: str) -> list[str]:
    sentences = [
        sentence.strip(" -;")
        for sentence in re.split(r"(?<=[.!?])\s+", value)
        if sentence.strip(" -;")
    ]
    if len(sentences) <= 1 and len(value) > 140:
        sentences = [
            part.strip(" -;")
            for part in re.split(r"\s+(?=[A-ZÉÈÀÂÎÔÛÇ][a-zéèàâîôûç])", value)
            if part.strip(" -;")
        ]
    return sentences


def _row_xml(row_index: int, values: list[object]) -> str:
    cells = []
    for column_index, value in enumerate(values, start=1):
        ref = f"{_column_name(column_index)}{row_index}"
        style = _cell_style(row_index, column_index, value)
        cells.append(_cell_xml(ref, value, f' s="{style}"' if style else ""))
    height = ' ht="42" customHeight="1"' if row_index == 1 else (' ht="34" customHeight="1"' if row_index >= 4 else "")
    return f'    <row r="{row_index}"{height}>' + "".join(cells) + "</row>"


def _cell_style(row_index: int, column_index: int, value: object) -> int:
    if row_index == 1:
        return 1
    if row_index == 3:
        return 2
    if column_index == 11 and value == "Bloqué":
        return 4
    if column_index == 11 and value == "Non bloqué":
        return 5
    if row_index >= 4:
        return 3
    return 0


def _cell_xml(ref: str, value: object, style: str = "") -> str:
    if isinstance(value, int):
        return f'<c r="{ref}"{style}><v>{value}</v></c>'
    return f'<c r="{ref}" t="inlineStr"{style}><is><t>{_xml_escape(str(value))}</t></is></c>'


def _column_name(index: int) -> str:
    name = ""
    while index:
        index, remainder = divmod(index - 1, 26)
        name = chr(65 + remainder) + name
    return name


def _content_types(sheet_count: int, logo: LogoAsset | None = None) -> str:
    sheet_overrides = "\n".join(
        f'  <Override PartName="/xl/worksheets/sheet{index}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
        for index in range(1, sheet_count + 1)
    )
    drawing_overrides = ""
    media_default = ""
    if logo:
        drawing_overrides = "\n".join(
            f'  <Override PartName="/xl/drawings/drawing{index}.xml" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>'
            for index in range(1, sheet_count + 1)
        )
        media_default = f'\n  <Default Extension="{logo.extension}" ContentType="{logo.content_type}"/>'
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>{media_default}
  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>
  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>
{sheet_overrides}
{drawing_overrides}
</Types>"""


def _worksheet_rels(index: int) -> str:
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing" Target="../drawings/drawing{index}.xml"/>
</Relationships>"""


def _drawing_rels(extension: str) -> str:
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/logo.{extension}"/>
</Relationships>"""


def _drawing_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <xdr:twoCellAnchor editAs="oneCell">
    <xdr:from><xdr:col>0</xdr:col><xdr:colOff>80000</xdr:colOff><xdr:row>0</xdr:row><xdr:rowOff>80000</xdr:rowOff></xdr:from>
    <xdr:to><xdr:col>1</xdr:col><xdr:colOff>500000</xdr:colOff><xdr:row>2</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>
    <xdr:pic>
      <xdr:nvPicPr><xdr:cNvPr id="2" name="GTHCONSULT logo"/><xdr:cNvPicPr/></xdr:nvPicPr>
      <xdr:blipFill><a:blip r:embed="rId1"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>
      <xdr:spPr><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr>
    </xdr:pic>
    <xdr:clientData/>
  </xdr:twoCellAnchor>
</xdr:wsDr>"""


def _root_rels() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
  <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
</Relationships>"""


def _workbook_xml(sheet_names: Iterable[str]) -> str:
    sheets = "\n".join(
        f'    <sheet name="{_xml_escape(name)}" sheetId="{index}" r:id="rId{index}"/>'
        for index, name in enumerate(sheet_names, start=1)
    )
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <sheets>
{sheets}
  </sheets>
</workbook>"""


def _workbook_rels(sheet_count: int) -> str:
    rels = [
        f'  <Relationship Id="rId{index}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet{index}.xml"/>'
        for index in range(1, sheet_count + 1)
    ]
    rels.append(
        f'  <Relationship Id="rId{sheet_count + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
    )
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
""" + "\n".join(rels) + "\n</Relationships>"


def _styles_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <fonts count="4">
    <font><sz val="11"/><name val="Calibri"/></font>
    <font><b/><color rgb="FFFFFFFF"/><sz val="16"/><name val="Calibri"/></font>
    <font><b/><color rgb="FFFFFFFF"/><sz val="11"/><name val="Calibri"/></font>
    <font><b/><color rgb="FF000000"/><sz val="11"/><name val="Calibri"/></font>
  </fonts>
  <fills count="6">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF1F4E78"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FF4472C4"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFFC7CE"/><bgColor indexed="64"/></patternFill></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFC6EFCE"/><bgColor indexed="64"/></patternFill></fill>
  </fills>
  <borders count="2">
    <border><left/><right/><top/><bottom/><diagonal/></border>
    <border><left style="thin"><color rgb="FFD9E2F3"/></left><right style="thin"><color rgb="FFD9E2F3"/></right><top style="thin"><color rgb="FFD9E2F3"/></top><bottom style="thin"><color rgb="FFD9E2F3"/></bottom><diagonal/></border>
  </borders>
  <cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>
  <cellXfs count="6">
    <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
    <xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="2" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment wrapText="1" vertical="top"/></xf>
    <xf numFmtId="0" fontId="3" fillId="4" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
    <xf numFmtId="0" fontId="3" fillId="5" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>
  </cellXfs>
  <cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>
</styleSheet>"""


def _package_core_xml() -> str:
    return """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
  <dc:creator>Recap DOCX</dc:creator>
  <dc:title>Recapitulatif DOCX</dc:title>
</cp:coreProperties>"""


def _package_app_xml(sheet_count: int) -> str:
    return f"""<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
  <Application>Recap DOCX</Application>
  <DocSecurity>0</DocSecurity>
  <ScaleCrop>false</ScaleCrop>
  <HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>{sheet_count}</vt:i4></vt:variant></vt:vector></HeadingPairs>
</Properties>"""


def _writestr(package: zipfile.ZipFile, name: str, content: str) -> None:
    package.writestr(name, content.encode("utf-8"))


def _xml_escape(value: str) -> str:
    return html.escape(_sanitize_excel_text(value), quote=True)


def _sanitize_excel_text(value: str) -> str:
    value = value.replace("\ufffd", "e")
    return "".join(
        char
        for char in value
        if char in "\t\n\r" or ord(char) >= 32
    )


def _local_name(tag: str) -> str:
    return tag.rsplit("}", 1)[-1]


def main() -> int:
    parser = argparse.ArgumentParser(description="Cree un recap Excel depuis des dossiers contenant des DOCX.")
    parser.add_argument("source", type=Path, help="Dossier extrait ou fichier .zip a scanner.")
    parser.add_argument("-o", "--output", type=Path, default=Path("outputs/recap.xlsx"), help="Fichier .xlsx de sortie.")
    args = parser.parse_args()

    try:
        sheet_names = build_workbook(args.source, args.output)
    except ValueError as exc:
        print(f"Erreur: {exc}")
        return 1

    print(f"Classeur cree: {args.output.resolve()}")
    print(f"Feuilles creees: {len(sheet_names)}")
    for name in sheet_names:
        print(f"- {name}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

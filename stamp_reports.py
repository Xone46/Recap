from __future__ import annotations

import argparse
import os
import re
import shutil
import stat
import tempfile
import zipfile
from pathlib import Path
from xml.dom import minidom

from recap_docx import _source_root


ROOT = Path(__file__).resolve().parent
DEFAULT_CACHET = Path(r"C:\Users\Test\Desktop\DH\tools\cachet_gthconsult_transparent.png")
DEFAULT_SIGNATURE = ROOT / "assets" / "signature.png"
MEDIA_CACHET = "word/media/cachet_gthconsult.png"
MEDIA_SIGNATURE = "word/media/signature_gthconsult.png"

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing"
A = "http://schemas.openxmlformats.org/drawingml/2006/main"
PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture"
CT = "http://schemas.openxmlformats.org/package/2006/content-types"
REL = "http://schemas.openxmlformats.org/package/2006/relationships"


def _safe_name(name: str) -> str:
    base = Path(name).stem
    base = re.sub(r"[^\w -]+", "_", base, flags=re.UNICODE).strip()
    base = re.sub(r"\s+", " ", base)
    return base or "reports"


def elements(parent, ns, local):
    return [
        node
        for node in parent.childNodes
        if node.nodeType == node.ELEMENT_NODE and node.namespaceURI == ns and node.localName == local
    ]


def descendants(parent, ns, local):
    return [
        node
        for node in parent.getElementsByTagNameNS(ns, local)
        if node.nodeType == node.ELEMENT_NODE
    ]


def text_of(node):
    return "".join(
        child.firstChild.nodeValue
        for child in node.getElementsByTagNameNS(W, "t")
        if child.firstChild and child.firstChild.nodeType == child.TEXT_NODE
    ).strip()


def parse_xml(data):
    return minidom.parseString(data)


def serialize(doc):
    return doc.toxml(encoding="utf-8")


def next_rid(rels_doc):
    ids = []
    for rel in rels_doc.getElementsByTagNameNS(REL, "Relationship"):
        rid = rel.getAttribute("Id")
        if rid.startswith("rId") and rid[3:].isdigit():
            ids.append(int(rid[3:]))
    return f"rId{(max(ids) if ids else 0) + 1}"


def next_docpr_id(doc):
    ids = []
    for node in doc.getElementsByTagNameNS(WP, "docPr"):
        try:
            ids.append(int(node.getAttribute("id")))
        except ValueError:
            continue
    return str((max(ids) if ids else 914000) + 1)


def ensure_namespace(root, prefix, uri):
    attr = f"xmlns:{prefix}"
    if not root.hasAttribute(attr):
        root.setAttribute(attr, uri)


def make_el(doc, ns, name, attrs=None):
    el = doc.createElementNS(ns, name)
    if attrs:
        for key, value in attrs.items():
            if ":" in key:
                prefix = key.split(":", 1)[0]
                uri = {"w": W, "r": R}.get(prefix)
                el.setAttributeNS(uri, key, value)
            else:
                el.setAttribute(key, value)
    return el


def append(parent, ns, name, attrs=None):
    el = make_el(parent.ownerDocument, ns, name, attrs)
    parent.appendChild(el)
    return el


def make_image_paragraph(doc, rid, name, cx, cy, y_offset):
    p = make_el(doc, W, "w:p")
    p_pr = append(p, W, "w:pPr")
    append(p_pr, W, "w:jc", {"w:val": "center"})
    append(p_pr, W, "w:spacing", {"w:before": "0", "w:after": "0", "w:line": "240", "w:lineRule": "auto"})
    r_el = append(p, W, "w:r")
    drawing = append(r_el, W, "w:drawing")
    anchor = append(
        drawing,
        WP,
        "wp:anchor",
        {
            "distT": "0",
            "distB": "0",
            "distL": "0",
            "distR": "0",
            "simplePos": "0",
            "relativeHeight": "251658240",
            "behindDoc": "0",
            "locked": "0",
            "layoutInCell": "1",
            "allowOverlap": "1",
        },
    )
    append(anchor, WP, "wp:simplePos", {"x": "0", "y": "0"})
    pos_h = append(anchor, WP, "wp:positionH", {"relativeFrom": "column"})
    append(pos_h, WP, "wp:align").appendChild(doc.createTextNode("center"))
    pos_v = append(anchor, WP, "wp:positionV", {"relativeFrom": "paragraph"})
    append(pos_v, WP, "wp:posOffset").appendChild(doc.createTextNode(str(y_offset)))
    append(anchor, WP, "wp:extent", {"cx": str(cx), "cy": str(cy)})
    append(anchor, WP, "wp:effectExtent", {"l": "0", "t": "0", "r": "0", "b": "0"})
    append(anchor, WP, "wp:wrapNone")
    append(anchor, WP, "wp:docPr", {"id": next_docpr_id(doc), "name": name})
    append(anchor, WP, "wp:cNvGraphicFramePr")
    graphic = append(anchor, A, "a:graphic")
    graphic_data = append(
        graphic,
        A,
        "a:graphicData",
        {"uri": "http://schemas.openxmlformats.org/drawingml/2006/picture"},
    )
    pic = append(graphic_data, PIC, "pic:pic")
    nv = append(pic, PIC, "pic:nvPicPr")
    append(nv, PIC, "pic:cNvPr", {"id": "0", "name": f"{name}.png"})
    append(nv, PIC, "pic:cNvPicPr")
    blip_fill = append(pic, PIC, "pic:blipFill")
    append(blip_fill, A, "a:blip", {"r:embed": rid})
    stretch = append(blip_fill, A, "a:stretch")
    append(stretch, A, "a:fillRect")
    sp_pr = append(pic, PIC, "pic:spPr")
    xfrm = append(sp_pr, A, "a:xfrm")
    append(xfrm, A, "a:off", {"x": "0", "y": "0"})
    append(xfrm, A, "a:ext", {"cx": str(cx), "cy": str(cy)})
    geom = append(sp_pr, A, "a:prstGeom", {"prst": "rect"})
    append(geom, A, "a:avLst")
    return p


def clear_cell_keep_props(tc):
    kept = None
    for child in elements(tc, W, "tcPr"):
        kept = child.cloneNode(deep=True)
        break
    while tc.firstChild:
        tc.removeChild(tc.firstChild)
    if kept is not None:
        tc.appendChild(kept)
        if not elements(kept, W, "vAlign"):
            append(kept, W, "w:vAlign", {"w:val": "center"})


def _find_stamp_cells(doc):
    for tbl in descendants(doc, W, "tbl"):
        rows = elements(tbl, W, "tr")
        if len(rows) < 2:
            continue
        first_cells = elements(rows[0], W, "tc")
        second_cells = elements(rows[1], W, "tc")
        if len(first_cells) >= 2 and len(second_cells) >= 2:
            left = text_of(first_cells[0]).upper()
            right = text_of(first_cells[1]).upper()
            if "INSPECTEUR" in left and "ADMINISTRATION" in right and "GTHCONSULT" in right:
                return second_cells[0], second_cells[1]
    return None


def _ensure_png_content_type(ct_doc):
    has_png = any(
        node.getAttribute("Extension").lower() == "png"
        for node in ct_doc.getElementsByTagNameNS(CT, "Default")
    )
    if not has_png:
        default = ct_doc.createElementNS(CT, "Default")
        default.setAttribute("Extension", "png")
        default.setAttribute("ContentType", "image/png")
        ct_doc.documentElement.appendChild(default)


def stamp_docx(path: Path, cachet_bytes: bytes, signature_bytes: bytes | None = None) -> tuple[bool, str]:
    with zipfile.ZipFile(path, "r") as zin:
        data = {name: zin.read(name) for name in zin.namelist()}

    if "word/document.xml" not in data or "word/_rels/document.xml.rels" not in data:
        return False, "missing document.xml"

    doc = parse_xml(data["word/document.xml"])
    root = doc.documentElement
    ensure_namespace(root, "wp", WP)
    ensure_namespace(root, "a", A)
    ensure_namespace(root, "pic", PIC)
    ensure_namespace(root, "r", R)

    stamp_cells = _find_stamp_cells(doc)
    if stamp_cells is None:
        return False, "signature table not found"

    signature_cell, cachet_cell = stamp_cells
    rels_doc = parse_xml(data["word/_rels/document.xml.rels"])
    ct_doc = parse_xml(data["[Content_Types].xml"])
    _ensure_png_content_type(ct_doc)

    rels_payloads: list[tuple[str, bytes]] = []

    if signature_bytes:
        sig_rid = next_rid(rels_doc)
        sig_rel = rels_doc.createElementNS(REL, "Relationship")
        sig_rel.setAttribute("Id", sig_rid)
        sig_rel.setAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image")
        sig_rel.setAttribute("Target", "media/signature_gthconsult.png")
        rels_doc.documentElement.appendChild(sig_rel)
        clear_cell_keep_props(signature_cell)
        signature_cell.appendChild(make_image_paragraph(doc, sig_rid, "Signature GTHCONSULT", 1120000, 520000, "-130000"))
        rels_payloads.append((MEDIA_SIGNATURE, signature_bytes))

    cachet_rid = next_rid(rels_doc)
    cachet_rel = rels_doc.createElementNS(REL, "Relationship")
    cachet_rel.setAttribute("Id", cachet_rid)
    cachet_rel.setAttribute("Type", "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image")
    cachet_rel.setAttribute("Target", "media/cachet_gthconsult.png")
    rels_doc.documentElement.appendChild(cachet_rel)
    clear_cell_keep_props(cachet_cell)
    cachet_cell.appendChild(make_image_paragraph(doc, cachet_rid, "Cachet GTHCONSULT", 1417320, 905256, "-155000"))
    rels_payloads.append((MEDIA_CACHET, cachet_bytes))

    data["word/document.xml"] = serialize(doc)
    data["word/_rels/document.xml.rels"] = serialize(rels_doc)
    data["[Content_Types].xml"] = serialize(ct_doc)
    for rel_path, payload in rels_payloads:
        data[rel_path] = payload

    tmp = path.with_suffix(path.suffix + ".tmp")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for name, payload in data.items():
            zout.writestr(name, payload)
    os.chmod(path, stat.S_IWRITE | stat.S_IREAD)
    os.replace(tmp, path)
    return True, "stamped"


def _copy_source_tree(source_root: Path, output_root: Path) -> None:
    if output_root.exists():
        shutil.rmtree(output_root)
    shutil.copytree(source_root, output_root)


def _resolve_output_root(source: Path, output: Path | None) -> Path:
    if output is not None:
        return output.resolve()
    return (ROOT / "outputs" / f"{_safe_name(source.name)}_cachet").resolve()


def process_source(source: Path, output: Path | None = None, cachet_path: Path | None = None, signature_path: Path | None = None) -> list[str]:
    source = source.resolve()
    output_root = _resolve_output_root(source, output)
    cachet_asset = cachet_path or DEFAULT_CACHET
    if not cachet_asset.exists():
        raise FileNotFoundError(f"Cachet introuvable: {cachet_asset}")
    signature_asset = signature_path or (DEFAULT_SIGNATURE if DEFAULT_SIGNATURE.exists() else None)
    if signature_asset is not None and not signature_asset.exists():
        raise FileNotFoundError(f"Signature introuvable: {signature_asset}")
    cachet_bytes = cachet_asset.read_bytes()
    signature_bytes = signature_asset.read_bytes() if signature_asset else None

    with _source_root(source) as root:
        _copy_source_tree(root, output_root)
        docs = sorted(
            p
            for p in output_root.rglob("*.docx")
            if not p.name.startswith("~$") and "_backup_before_cachet" not in p.parts
        )
        results: list[str] = []
        for path in docs:
            rel = path.relative_to(output_root)
            try:
                updated, msg = stamp_docx(path, cachet_bytes, signature_bytes)
                results.append(f"{'OK' if updated else 'SKIP'}\t{rel}\t{msg}")
            except Exception as exc:
                results.append(f"ERROR\t{rel}\t{exc}")
        return results


def main() -> None:
    parser = argparse.ArgumentParser(description="Ajoute le cachet et la signature dans tous les rapports DOCX.")
    parser.add_argument("source", help="Dossier ou archive .zip/.rar contenant les rapports DOCX")
    parser.add_argument("-o", "--output", help="Dossier de sortie pour les DOCX estampilles")
    parser.add_argument("--cachet", help="Chemin vers l'image du cachet")
    parser.add_argument("--signature", help="Chemin vers l'image de signature (par defaut: assets/signature.png)")
    args = parser.parse_args()

    source = Path(args.source)
    output = Path(args.output) if args.output else None
    cachet_path = Path(args.cachet) if args.cachet else None
    signature_path = Path(args.signature) if args.signature else None

    results = process_source(source, output=output, cachet_path=cachet_path, signature_path=signature_path)
    out_dir = _resolve_output_root(source.resolve(), output)
    processed = len(results)
    ok = sum(line.startswith("OK\t") for line in results)
    skipped = sum(line.startswith("SKIP\t") for line in results)
    errors = sum(line.startswith("ERROR\t") for line in results)

    report_path = ROOT / "cachet_reports.txt"
    report_path.write_text(
        "\n".join(
            [
                f"source={source}",
                f"output={out_dir}",
                f"processed={processed}",
                f"stamped={ok}",
                f"skipped={skipped}",
                f"errors={errors}",
                "",
                *results,
                "",
            ]
        ),
        encoding="utf-8",
    )
    print(f"source={source}")
    print(f"output={out_dir}")
    print(f"processed={processed} stamped={ok} skipped={skipped} errors={errors}")
    print(report_path)


if __name__ == "__main__":
    main()

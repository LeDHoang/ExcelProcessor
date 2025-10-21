import argparse
import json
import os
import posixpath
import re
import shutil
import sys
import zipfile
from collections import defaultdict, deque
from dataclasses import dataclass, asdict
from datetime import datetime
from typing import Any, Dict, List, Optional, Tuple

try:
    import openpyxl  # type: ignore
except Exception:
    openpyxl = None


NAMESPACES: Dict[str, str] = {
    "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
    "r": "http://schemas.openxmlformats.org/officeDocument/2006/relationships",
    "dgm": "http://schemas.openxmlformats.org/drawingml/2006/diagram",
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main",
    "pr": "http://schemas.openxmlformats.org/package/2006/relationships",
}


def ensure_dir(path: str) -> None:
    if not os.path.isdir(path):
        os.makedirs(path, exist_ok=True)


def load_xml(zipf: zipfile.ZipFile, path: str):
    from xml.etree import ElementTree as ET

    with zipf.open(path) as f:
        data = f.read()
    return ET.fromstring(data)


def read_text_nodes(elem) -> List[str]:
    # Collect all a:t descendants as text pieces
    texts: List[str] = []
    for t_node in elem.findall(".//a:t", NAMESPACES):
        if t_node.text is not None:
            texts.append(t_node.text)
    return texts


def normalize_rel_target(base_dir_in_zip: str, target: str) -> str:
    # Relationships targets can be relative like ../media/image1.png
    # base_dir_in_zip is e.g. xl/drawings
    joined = posixpath.normpath(posixpath.join(base_dir_in_zip, target))
    # Remove any leading './' artifacts
    while joined.startswith("./"):
        joined = joined[2:]
    return joined


def parse_workbook_sheets(zipf: zipfile.ZipFile) -> List[Tuple[str, str]]:
    # Returns list of (sheet_name, sheet_path) in the zip
    from xml.etree import ElementTree as ET

    workbook_xml_path = "xl/workbook.xml"
    workbook_rels_path = "xl/_rels/workbook.xml.rels"

    workbook_root = load_xml(zipf, workbook_xml_path)
    rels_root = load_xml(zipf, workbook_rels_path)

    rid_to_target: Dict[str, str] = {}
    for rel in rels_root.findall(".//pr:Relationship", NAMESPACES):
        r_type = rel.attrib.get("Type", "")
        if r_type.endswith("/worksheet"):
            target = rel.attrib.get("Target", "")
            # Targets are relative to xl/workbook.xml → base 'xl'
            target_path = normalize_rel_target("xl", target)
            rid_to_target[rel.attrib["Id"]] = target_path

    sheets: List[Tuple[str, str]] = []
    for sheet in workbook_root.findall(".//s:sheet", NAMESPACES):
        name = sheet.attrib.get("name", "Sheet")
        rid = sheet.attrib.get(f"{{{NAMESPACES['r']}}}id")
        if rid and rid in rid_to_target:
            # Normalize path to xl/worksheets/sheetX.xml
            target = rid_to_target[rid]
            sheet_path = posixpath.normpath(target)
            sheets.append((name, sheet_path))
    return sheets


def read_sheet_relationships(zipf: zipfile.ZipFile, sheet_path: str) -> Dict[str, Dict[str, str]]:
    # Returns map of rel_id -> {Type, Target}
    rels_dir = posixpath.dirname(sheet_path).replace("xl/worksheets", "xl/worksheets/_rels")
    rels_filename = posixpath.basename(sheet_path) + ".rels"
    rels_path = posixpath.join(rels_dir, rels_filename)
    rels: Dict[str, Dict[str, str]] = {}
    if rels_path in zipf.namelist():
        root = load_xml(zipf, rels_path)
        for rel in root.findall(".//pr:Relationship", NAMESPACES):
            rels[rel.attrib["Id"]] = {
                "Type": rel.attrib.get("Type", ""),
                "Target": rel.attrib.get("Target", ""),
            }
    return rels


def read_rels(zipf: zipfile.ZipFile, part_path: str) -> Dict[str, Dict[str, str]]:
    # Generic rels reader for a given part (e.g., xl/drawings/drawing1.xml)
    base_dir = posixpath.dirname(part_path)
    rels_dir = posixpath.join(base_dir, "_rels")
    rels_path = posixpath.join(rels_dir, posixpath.basename(part_path) + ".rels")
    rels: Dict[str, Dict[str, str]] = {}
    if rels_path in zipf.namelist():
        root = load_xml(zipf, rels_path)
        for rel in root.findall(".//pr:Relationship", NAMESPACES):
            rels[rel.attrib["Id"]] = {
                "Type": rel.attrib.get("Type", ""),
                "Target": rel.attrib.get("Target", ""),
            }
    return rels


@dataclass
class ImageInfo:
    sheet_name: str
    image_filename: str
    original_part: str
    anchor: Dict[str, int]


@dataclass
class ShapeText:
    sheet_name: str
    text: str
    anchor: Dict[str, int]


@dataclass
class SmartArtNode:
    id: str
    text: str
    children: List["SmartArtNode"]


def extract_anchor(anchor_elem) -> Dict[str, int]:
    # Supports xdr:twoCellAnchor and xdr:oneCellAnchor 'from' position
    def read_pos(from_elem) -> Dict[str, int]:
        def read_child(name: str) -> int:
            child = from_elem.find(f"xdr:{name}", NAMESPACES)
            if child is None or child.text is None:
                return 0
            try:
                return int(child.text)
            except Exception:
                return 0

        return {
            "col": read_child("col"),
            "colOff": read_child("colOff"),
            "row": read_child("row"),
            "rowOff": read_child("rowOff"),
        }

    from_pos = anchor_elem.find("xdr:from", NAMESPACES)
    if from_pos is not None:
        return read_pos(from_pos)
    # Fallback default
    return {"col": 0, "colOff": 0, "row": 0, "rowOff": 0}


def extract_images_and_shapes(
    zipf: zipfile.ZipFile,
    drawing_path: str,
    output_images_dir: str,
    sheet_name: str,
) -> Tuple[List[ImageInfo], List[ShapeText], List[Tuple[Dict[str, int], str, str]]]:
    # Returns (images, shape_texts, smartart_refs) where smartart_refs contains (anchor, rel_id, target_path)
    from xml.etree import ElementTree as ET

    images: List[ImageInfo] = []
    shapes: List[ShapeText] = []
    smartart_refs: List[Tuple[Dict[str, int], str, str]] = []

    drawing_root = load_xml(zipf, drawing_path)
    rels = read_rels(zipf, drawing_path)

    base_dir_in_zip = posixpath.dirname(drawing_path)

    def iter_anchors():
        for tag in ("twoCellAnchor", "oneCellAnchor"):  # order preserved
            for anchor in drawing_root.findall(f"xdr:{tag}", NAMESPACES):
                yield anchor

    image_counter = 0
    for anchor in iter_anchors():
        anchor_pos = extract_anchor(anchor)

        # Pictures
        pic = anchor.find("xdr:pic", NAMESPACES)
        if pic is not None:
            blip = pic.find(".//a:blip", NAMESPACES)
            if blip is not None:
                rid = blip.attrib.get(f"{{{NAMESPACES['r']}}}embed")
                if rid and rid in rels:
                    target_rel = rels[rid]["Target"]
                    part_in_zip = normalize_rel_target(base_dir_in_zip, target_rel)
                    if part_in_zip in zipf.namelist():
                        # Save image
                        image_counter += 1
                        ext = os.path.splitext(part_in_zip)[1] or ".bin"
                        safe_sheet = re.sub(r"[^A-Za-z0-9_-]+", "_", sheet_name)
                        out_name = f"{safe_sheet}_img_{image_counter}{ext}"
                        out_path = os.path.join(output_images_dir, out_name)
                        with zipf.open(part_in_zip) as src, open(out_path, "wb") as dst:
                            shutil.copyfileobj(src, dst)
                        images.append(
                            ImageInfo(
                                sheet_name=sheet_name,
                                image_filename=os.path.relpath(out_path, os.path.dirname(output_images_dir)),
                                original_part=part_in_zip,
                                anchor=anchor_pos,
                            )
                        )
            continue

        # Shape with text
        sp = anchor.find("xdr:sp", NAMESPACES)
        if sp is not None:
            tx_body = sp.find("xdr:txBody", NAMESPACES)
            if tx_body is not None:
                texts = read_text_nodes(tx_body)
                text_content = " ".join(t.strip() for t in texts if t is not None).strip()
                if text_content:
                    shapes.append(ShapeText(sheet_name=sheet_name, text=text_content, anchor=anchor_pos))
            continue

        # Graphic frame (charts, smartart diagrams)
        gframe = anchor.find("xdr:graphicFrame", NAMESPACES)
        if gframe is not None:
            gdata = gframe.find(".//a:graphicData", NAMESPACES)
            if gdata is not None:
                uri = gdata.attrib.get("uri", "")
                # SmartArt diagram
                if uri.endswith("/diagram"):
                    rel_ids = gdata.find("dgm:relIds", NAMESPACES)
                    if rel_ids is not None:
                        # Prefer data model (dm)
                        dm_rid = rel_ids.attrib.get(f"{{{NAMESPACES['r']}}}dm")
                        if dm_rid and dm_rid in rels:
                            target_rel = rels[dm_rid]["Target"]
                            part_in_zip = normalize_rel_target(base_dir_in_zip, target_rel)
                            smartart_refs.append((anchor_pos, dm_rid, part_in_zip))
                        # Some files use lo (layout) or qs (quick style) to reach data via nested rels
                        for key in ("lo", "qs", "cs"):
                            rid = rel_ids.attrib.get(f"{{{NAMESPACES['r']}}}{key}")
                            if rid and rid in rels:
                                target_rel = rels[rid]["Target"]
                                part_in_zip = normalize_rel_target(base_dir_in_zip, target_rel)
                                smartart_refs.append((anchor_pos, rid, part_in_zip))
            continue

    return images, shapes, smartart_refs


def build_smartart_hierarchy(zipf: zipfile.ZipFile, data_model_path: str) -> List[SmartArtNode]:
    # Parse xl/diagrams/dataX.xml and build a tree
    root = load_xml(zipf, data_model_path)

    # Points (nodes)
    id_to_text: Dict[str, str] = {}
    for pt in root.findall(".//dgm:pt", NAMESPACES):
        model_id = pt.attrib.get("modelId") or pt.attrib.get("modelId", "")
        text_chunks = read_text_nodes(pt)
        text_value = " ".join(t.strip() for t in text_chunks if t is not None).strip()
        if not text_value:
            # Try name on prSet if text is missing
            pr = pt.find("dgm:prSet", NAMESPACES)
            name_attr = pr.attrib.get("name") if pr is not None else None
            if name_attr:
                text_value = name_attr
        if model_id:
            id_to_text[model_id] = text_value

    # Connections (edges) define hierarchy/flow
    children_map: Dict[str, List[str]] = defaultdict(list)
    indegree: Dict[str, int] = defaultdict(int)
    for cxn in root.findall(".//dgm:cxn", NAMESPACES):
        src = cxn.attrib.get("srcId")
        dst = cxn.attrib.get("destId")
        if src and dst:
            children_map[src].append(dst)
            indegree[dst] += 1
            # ensure keys exist
            indegree.setdefault(src, indegree.get(src, 0))

    # Roots are nodes with indegree == 0
    roots = [node_id for node_id in id_to_text.keys() if indegree.get(node_id, 0) == 0]

    def build_node(node_id: str) -> SmartArtNode:
        text = id_to_text.get(node_id, "")
        child_ids = children_map.get(node_id, [])
        return SmartArtNode(
            id=node_id,
            text=text,
            children=[build_node(cid) for cid in child_ids],
        )

    nodes = [build_node(r) for r in roots] if roots else [build_node(nid) for nid in id_to_text.keys()]
    return nodes


def extract_cells_text(excel_path: str) -> Dict[str, List[Dict[str, Any]]]:
    results: Dict[str, List[Dict[str, Any]]] = defaultdict(list)
    if openpyxl is None:
        return results
    try:
        wb = openpyxl.load_workbook(excel_path, data_only=True)
    except Exception:
        return results
    for ws in wb.worksheets:
        for row in ws.iter_rows(values_only=False):
            for cell in row:
                value = cell.value
                if value is None:
                    continue
                text = str(value).strip()
                if text == "":
                    continue
                results[ws.title].append(
                    {
                        "cell": cell.coordinate,
                        "text": text,
                    }
                )
    return results


def generate_markdown(output_json: Dict[str, Any]) -> str:
    lines: List[str] = []
    lines.append("# Excel to Markdown Conversion")
    lines.append(f"Generated on: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    lines.append("")
    # TOC
    lines.append("## Table of Contents")
    for sheet in output_json.get("sheets", []):
        sheet_name = sheet.get("name", "Sheet")
        anchor = re.sub(r"[^a-z0-9]+", "-", sheet_name.lower()).strip("-")
        lines.append(f"- [{sheet_name}](#{anchor})")
    lines.append("")

    for sheet in output_json.get("sheets", []):
        sheet_name = sheet.get("name", "Sheet")
        anchor = re.sub(r"[^a-z0-9]+", "-", sheet_name.lower()).strip("-")
        lines.append(f"## {sheet_name}")

        # Cell text
        cell_texts = sheet.get("text_content", [])
        if cell_texts:
            lines.append("")
            lines.append("### Cell Text")
            for item in cell_texts:
                lines.append(f"- {item.get('cell')}: {item.get('text')}")

        # Shape text
        shapes = sheet.get("shapes_text", [])
        if shapes:
            lines.append("")
            lines.append("### Shapes Text")
            for st in shapes:
                pos = st.get("anchor", {})
                lines.append(f"- (r{pos.get('row')}, c{pos.get('col')}): {st.get('text')}")

        # Images
        images = sheet.get("images", [])
        if images:
            lines.append("")
            lines.append("### Images")
            for im in images:
                rel_path = im.get("image_filename")
                caption = f"Image at r{im.get('anchor',{}).get('row')} c{im.get('anchor',{}).get('col')}"
                lines.append(f"![{caption}]({rel_path})")

        # SmartArt
        smartarts = sheet.get("smartart", [])
        if smartarts:
            lines.append("")
            lines.append("### SmartArt")

            def walk(nodes: List[Dict[str, Any]], depth: int = 0) -> None:
                for node in nodes:
                    bullet = "  " * depth + "- "
                    text = node.get("text", "").strip() or f"(id: {node.get('id')})"
                    lines.append(f"{bullet}{text}")
                    walk(node.get("children", []), depth + 1)

            walk(smartarts)

        lines.append("")

    return "\n".join(lines).rstrip() + "\n"


def process_excel(input_path: str, output_dir: str) -> Dict[str, Any]:
    ensure_dir(output_dir)
    images_dir = os.path.join(output_dir, "images")
    ensure_dir(images_dir)

    # Base JSON structure
    result: Dict[str, Any] = {
        "text_content": [],
        "images": [],
        "smartart": [],
        "links": [],
        "metadata": {
            "filename": os.path.basename(input_path),
            "created": datetime.now().isoformat(),
        },
        "sheets": [],
    }

    # Extract cells text using openpyxl (best effort)
    cells_by_sheet = extract_cells_text(input_path)

    with zipfile.ZipFile(input_path, "r") as zipf:
        # Discover sheets and their parts
        try:
            sheets = parse_workbook_sheets(zipf)
        except Exception:
            # Fallback to default mapping if parsing fails
            sheets = [("Sheet1", "xl/worksheets/sheet1.xml")]

        result["metadata"]["sheets"] = [name for name, _ in sheets]

        for sheet_name, sheet_part in sheets:
            sheet_entry: Dict[str, Any] = {
                "name": sheet_name,
                "text_content": cells_by_sheet.get(sheet_name, []),
                "images": [],
                "shapes_text": [],
                "smartart": [],
                "structure": {
                    "sheet_part": sheet_part,
                    "hierarchy": {},
                },
            }

            # Resolve drawing relationships
            s_rels = read_sheet_relationships(zipf, sheet_part)
            for rel_id, rel in s_rels.items():
                if rel.get("Type", "").endswith("/drawing"):
                    drawing_part = normalize_rel_target(posixpath.dirname(sheet_part), rel["Target"]).replace("worksheets/../", "xl/")
                    drawing_part = posixpath.normpath(drawing_part)
                    if drawing_part in zipf.namelist():
                        images, shape_texts, smartart_refs = extract_images_and_shapes(
                            zipf, drawing_part, images_dir, sheet_name
                        )
                        # Images
                        for im in images:
                            im_dict = asdict(im)
                            sheet_entry["images"].append(im_dict)
                            result["images"].append(im_dict)
                        # Shapes text
                        for st in shape_texts:
                            st_dict = asdict(st)
                            sheet_entry["shapes_text"].append(st_dict)

                        # SmartArt data models
                        for anchor, rid, data_model_path in smartart_refs:
                            try:
                                nodes = build_smartart_hierarchy(zipf, data_model_path)
                                # Store as plain dicts for JSON
                                def node_to_dict(n: SmartArtNode) -> Dict[str, Any]:
                                    return {
                                        "id": n.id,
                                        "text": n.text,
                                        "children": [node_to_dict(c) for c in n.children],
                                    }

                                sheet_entry["smartart"].extend([node_to_dict(n) for n in nodes])
                                result["smartart"].extend([node_to_dict(n) for n in nodes])
                            except Exception:
                                # Best-effort; continue other content
                                continue

            result["sheets"].append(sheet_entry)

    return result


def main(argv: Optional[List[str]] = None) -> int:
    parser = argparse.ArgumentParser(description="Extract text, SmartArt, and images from Excel preserving hierarchy.")
    parser.add_argument("-i", "--input", required=True, help="Path to input .xlsx file")
    parser.add_argument("-o", "--output", default="output", help="Output directory")
    args = parser.parse_args(argv)

    input_path = os.path.abspath(args.input)
    output_dir = os.path.abspath(args.output)

    if not os.path.isfile(input_path):
        print(f"Input file not found: {input_path}", file=sys.stderr)
        return 2

    data = process_excel(input_path, output_dir)

    # Write JSON
    json_path = os.path.join(output_dir, "extracted_data.json")
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)

    # Write Markdown
    md = generate_markdown(data)
    md_path = os.path.join(output_dir, "converted.md")
    with open(md_path, "w", encoding="utf-8") as f:
        f.write(md)

    print(f"Wrote: {json_path}")
    print(f"Wrote: {md_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

   
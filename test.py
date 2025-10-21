import zipfile
import xml.etree.ElementTree as ET
from typing import List, Dict, Any

# Define the XML namespaces
namespaces = {
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'dgm': 'http://schemas.openxmlformats.org/drawingml/2006/diagram'
}

def extract_smartart_structure(xlsx_path: str) -> List[Dict[str, Any]]:
    """
    Extracts a structured representation of all SmartArt objects from an .xlsx file.
    """
    all_diagrams_data = []

    try:
        with zipfile.ZipFile(xlsx_path, 'r') as zf:
            diagram_files = [f for f in zf.namelist() if f.startswith('xl/diagrams/data')]
            
            for diagram_file in diagram_files:
                xml_content = zf.read(diagram_file)
                root = ET.fromstring(xml_content)
                
                # The ptLst contains the hierarchical points (shapes)
                point_list_element = root.find('dgm:ptLst', namespaces)
                if point_list_element is not None:
                    diagram_structure = []
                    # Use a recursive helper function to parse the hierarchy
                    _parse_points(point_list_element, diagram_structure, level=0)
                    all_diagrams_data.append({
                        "source_file": diagram_file,
                        "structure": diagram_structure
                    })

    except Exception as e:
        print(f"An error occurred: {e}")
    
    return all_diagrams_data

def _parse_points(parent_element: ET.Element, structure_list: List, level: int):
    """Recursively parses <dgm:pt> elements to capture hierarchy."""
    for pt_element in parent_element.findall('dgm:pt', namespaces):
        # Find the text associated with this point
        text_body = pt_element.find('.//a:t', namespaces)
        text = text_body.text if text_body is not None else ""
        
        point_data = {
            "id": pt_element.get('id'),
            "text": text,
            "level": level
        }
        structure_list.append(point_data)
        
        # Recursively process any child points
        _parse_points(pt_element, structure_list, level + 1)

# --- USAGE ---
file_path = 'Sheet1.xlsx'
smartart_data = extract_smartart_structure(file_path)

if smartart_data:
    print("Successfully extracted SmartArt structure:")
    for diagram in smartart_data:
        print(f"\n--- Diagram: {diagram['source_file']} ---")
        for node in diagram['structure']:
            indent = "  " * node['level']
            print(f"{indent}ID: {node['id']}, Level: {node['level']}, Text: '{node['text']}'")
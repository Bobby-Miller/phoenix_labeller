from pprint import pprint
import xml.etree.ElementTree as ET
import base64
import zipfile
import io
import pandas as pd
from datetime import datetime
import os
import xml.dom.minidom

# --- CONFIGURATION ---
SOURCE_DATA = "labels_to_print.xlsx"  # Can be .csv or .xlsx
OUTPUT_FILE = "AUTOMATED_PROJECT.mtp"
LABELS_PER_MATERIAL = 7

def extract_mtp_content(file_path, extracted_directory_name):
    # 1. Parse the outer XML file
    tree = ET.parse(file_path)
    root = tree.getroot()

    # 2. Find the Content tag (handling namespaces if necessary)
    # The provided XML uses a flat structure for Content
    content_element = root.find('Content')

    if content_element is None or not content_element.text:
        print("No content found in the MTP file.")
        return

    # 3. Decode the Base64 string into bytes
    encoded_data = content_element.text.strip()
    decoded_bytes = base64.b64decode(encoded_data)

    # 4. Open the decoded bytes as a Zip file
    output_dir = extracted_directory_name
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)

    with zipfile.ZipFile(io.BytesIO(decoded_bytes)) as z:
        z.extractall(output_dir)
        print(f"Successfully extracted {len(z.namelist())} files to '{output_dir}':")
        for name in z.namelist():
            print(f" - {name}")


def xml_data_from_file(source):
# Load and parse the file
    tree = ET.parse(source)
    root = tree.getroot()

# Convert the tree back into a string
    xml_string = ET.tostring(root, encoding='unicode', method='xml')
    return xml_string

def generate_mtp(output_filename, mat_file, xml_data):

    # 2. Create the ZIP archive in memory
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, 'a', zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("Data.xml", xml_data)
        # If you have a specific .mat file, you'd add it here:
        zf.write(mat_file, f"Materials/{mat_file.split('/')[-1]}")

    # 3. Encode the entire ZIP to Base64
    zip_bytes = zip_buffer.getvalue()
    encoded_content = base64.b64encode(zip_bytes).decode('utf-8')

    # 4. Create the outer .mtp (XML) wrapper 
    now = datetime.now().strftime("%y-%m-%d %H/%M")
    mtp_root = ET.Element("Project", {
        "DataCreation": now,
        "DataLastModyfication": now,
        "Version": "12",
        "IsComplex": "true"
    })

    content_element = ET.SubElement(mtp_root, "Content")
    content_element.text = encoded_content

    # 5. Save to file
    tree = ET.ElementTree(mtp_root)
    tree.write(output_filename, encoding='utf-8', xml_declaration=True)
    print(f"Successfully generated {output_filename}")


def encode_b64(text):
    if not text or pd.isna(text):
        return ""
    return base64.b64encode(str(text).encode('utf-8')).decode('utf-8')

def load_labels(file_path):
    """Loads labels from the 'Labels' column of a CSV or Excel file."""
    try:
        if file_path.endswith('.csv'):
            df = pd.read_csv(file_path)
        else:
            df = pd.read_excel(file_path)
        
        if 'Labels' not in df.columns:
            print(f"Error: Could not find a column named 'Labels' in {file_path}")
            return []
        
        return df['Labels'].tolist()
    except Exception as e:
        print(f"Failed to load file: {e}")
        return []

def create_data_xml(label_list):
    now = datetime.now().strftime("%y-%m-%d %H/%M")
    root = ET.Element("Project", {
        "DataCreation": now, 
        "DataLastModyfication": now, 
        "Version": "12",
        "OrderNumberOfCopy": "1"
    })

    # LogicNode with double-encoded "Root"
    logic = ET.SubElement(root, "LogicNode", {"RefId": "1"})
    ET.SubElement(logic, "ParentRefId").text = "-1"
    ET.SubElement(logic, "Name").text = encode_b64(encode_b64("Root"))
    ET.SubElement(logic, "Level").text = "-1"

    for i in range(0, len(label_list), LABELS_PER_MATERIAL):
        chunk = label_list[i : i + LABELS_PER_MATERIAL]
        
        material = ET.SubElement(root, "Material", {
            "Name": "TR_WML6(13X13)R",
            "Device": "D17",
            "Height": "35.75",
            "Width": "175",
            "OrderMaterialColor": "0816252:WH",
            "TransformedByAngle": "0",
        })


        for index in range(LABELS_PER_MATERIAL):
            label_text = chunk[index] if index < len(chunk) else "default"
            
            lbl = ET.SubElement(material, "Label", {
                "Index": str(index),
                "OrderIndex": str(index),
                "IsPrinted": "true",
                "TreeRefId": "1",
                "HorizontalAlign": "Center",
                "VerticalAlign": "Center",
                "Rotation": "0",
                "Width": "12.7",
                "Height": "12.7",
                "Mirrored": "false",
                "GroupID": "0",
                "LineSpacing": "1",
                "ContentProtection": "false",
                "IsPrinted": "false",
                "PrintedTime": "2026-01-29T13:55:49.5675461-05:00",
                "CurrentColor": "0, 0, 0",
                "Section": "0",
                "TreeRefId": "1",
                
            })

            tr = ET.SubElement(lbl, "TextRange")
            ET.SubElement(tr, "Color").text = "0, 0, 0"
            ET.SubElement(tr, "Text").text = encode_b64(label_text)
            
            ET.SubElement(lbl, "CutOrPerforations")
            # Formatting block
            for tag, el in {"Font": tr, "DefaultFont": lbl, "CurrentFont": lbl}.items():
                f = ET.SubElement(el, tag, {"Size": "2.5", "Type": "TTF", "LineThickness": "0", "Stretching": "1.5", "Proportional": "false", "Kerning": "false", "WidthFactor": "1", "FormattingType": "0"})
                ET.SubElement(f, "Name").text = "Arial Narrow" if "Default" not in tag else "Arial"

    wire_info_cols_dict = {
        "MaterialDisplayNameColumn": "false",
        "IndexOfWireColumn": "false",
        "CrossSectionColumn": "false",
        "CrossSectionUnitColumn": "false",
        "WireColorDescriptorColumn": "false",
        "PartNumberColumn": "false",
        "OrderNumberColumn": "false",
        "TypeDesignationColumn": "false",
        "WireLengthColumn": "false",
        "WireDiameterColumn": "false",
        "WireDiameterUnitColumn": "false",
        "EndSourceDirectionDescriptorColumn": "false",
        "EndTargetDirectionDescriptor": "false",
        "SourceColumn": "false",
        "RoutingTrackColumn": "false",
        "TargetColumn": "false",
        "NumberOfCopiesColumn": "false",
        "PrintStatusColumn": "false",
        "PrintStampColumn": "false",
        "FunctionalAssignmentSourceColumn": "true",
        "FunctionalAssignmentTargetColumn": "true",
        "HighLevelFunctionSourceColumn": "true",
        "HighLevelFunctionTargetColumn": "true",
        "InstallationSiteSourceColumn": "true",
        "InstallationSiteTargetColumn": "true",
        "MountingLocationSourceColumn": "true",
        "MountingLocationTargetColumn": "true",
        "DtSourceColumn": "true",
        "DtTargetColumn": "true",
        "ConnectionPointSourceColumn": "true",
        "ConnectionPointTargetColumn": "true",
        "PageSourceColumn": "true",
        "PageTargetColumn": "true",
        "WireTerminationProcessingSourceColumn": "true",
        "WireTerminationProcessingTargetColumn": "true",
        "StrippingLengthSourceColumn": "true",
        "StrippingLengthTargetColumn": "true",
        "ConnectionDimensionSourceColumn": "true",
        "ConnectionDimensionTargetColumn": "true",
        "ConnectionDesignationColumn": "true",
    }

    # Minimal required metadata columns
    for idx, (col_id, val) in enumerate(wire_info_cols_dict.items()):
        ET.SubElement(root, "WireInformationColumns", {"ColumnId": col_id, "Index": str(idx), "IsReadonly": "true", "AdditionalColumn": val, "HeaderText": "", "IsVisible": "false",})



    cad_db = ET.SubElement(root, "CadDataBase")
    ET.SubElement(cad_db, "Entrys")
    ET.SubElement(root, "Resources")

    return ET.tostring(root, encoding='utf-8', xml_declaration=True)

def pretty_xml_data(xml_data):


# 1. Parse the string into a DOM object
# We use .decode() because minidom.parseString expects a string, not bytes
    dom = xml.dom.minidom.parseString(xml_data.decode('utf-8'))

# 2. Use toprettyxml() to add indentation and newlines
    pretty_xml = dom.toprettyxml(indent="    ")

    return pretty_xml


def main():
    # --- Execution ---
    labels = load_labels(SOURCE_DATA)
    if labels:
        # xml_data = xml_data_from_file("extracted_mtp_data/Data.xml")
        xml_data = create_data_xml(labels)
        pretty_xml = pretty_xml_data(xml_data)
        # print(pretty_xml_data(xml_data))
        # print(xml_data)
        generate_mtp("AUTOGEN_CONFIG_J.mtp", "extracted_mtp_data/Materials/TR_WML6(13X13)R.mat", pretty_xml)
        extract_mtp_content('AUTOGEN_CONFIG_J.mtp', "extracted_mtp_data_J")


if __name__ == "__main__":
    main()

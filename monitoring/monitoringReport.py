import os
import json
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.shared import Inches
from PIL import Image
from docx.shared import Pt

# Constants file path
CONFIG_PATH = "monitoring/config/constants.json"


def load_constants():
    with open(CONFIG_PATH, 'r') as file:
        return json.load(file)


CONSTANTS = load_constants()


def add_page_number(doc):
    """Adds page numbers to the footer of the document."""
    for section in doc.sections:
        footer = section.footer
        paragraph = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
        paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
        run = paragraph.add_run()
        fldChar1 = OxmlElement("w:fldChar")
        fldChar1.set(qn("w:fldCharType"), "begin")
        instrText = OxmlElement("w:instrText")
        instrText.set(qn("xml:space"), "preserve")
        instrText.text = "PAGE"
        fldChar2 = OxmlElement("w:fldChar")
        fldChar2.set(qn("w:fldCharType"), "end")
        run._r.append(fldChar1)
        run._r.append(instrText)
        run._r.append(fldChar2)


def add_title_page(doc, report_frequency):
    """Adds the title page."""
    doc.add_paragraph(f"{report_frequency.capitalize()} Monitoring Report", "Title")
    doc.add_page_break()

# def add_doc_control(doc, report_frequency):
#     doc.add_paragraph()


def add_table_of_contents(doc):
    """Adds a TOC field that updates when `F9` is pressed in Word."""
    doc.add_paragraph("Contents", "TOC Heading")

    paragraph = doc.add_paragraph()
    paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    run = paragraph.add_run()

    fldChar1 = OxmlElement("w:fldChar")
    fldChar1.set(qn("w:fldCharType"), "begin")

    instrText = OxmlElement("w:instrText")
    instrText.set(qn("xml:space"), "preserve")
    instrText.text = "TOC \\o \"1-3\" \\h \\z \\u"

    fldChar2 = OxmlElement("w:fldChar")
    fldChar2.set(qn("w:fldCharType"), "separate")

    fldChar3 = OxmlElement("w:fldChar")
    fldChar3.set(qn("w:fldCharType"), "end")

    run._r.append(fldChar1)
    run._r.append(instrText)
    run._r.append(fldChar2)
    run._r.append(fldChar3)

    add_list_of_tables_and_figures(doc)

def add_list_of_tables_and_figures(doc):
    """Adds separate 'List of Tables' and 'List of Figures' sections to the document."""

    # 📌 List of Tables
    doc.add_page_break()
    doc.add_paragraph("Tables", "TOC Heading")

    paragraph = doc.add_paragraph()
    paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    run = paragraph.add_run()

    fldChar1 = OxmlElement("w:fldChar")
    fldChar1.set(qn("w:fldCharType"), "begin")

    instrText = OxmlElement("w:instrText")
    instrText.set(qn("xml:space"), "preserve")
    instrText.text = "TOC \\h \\z \\t \"Heading 4,1\""  # ✅ Extracts only "Heading 4" for Tables

    fldChar2 = OxmlElement("w:fldChar")
    fldChar2.set(qn("w:fldCharType"), "separate")

    fldChar3 = OxmlElement("w:fldChar")
    fldChar3.set(qn("w:fldCharType"), "end")

    run._r.append(fldChar1)
    run._r.append(instrText)
    run._r.append(fldChar2)
    run._r.append(fldChar3)

    # 📌 List of Figures
    doc.add_paragraph("")
    doc.add_paragraph("Figures", "TOC Heading")

    paragraph = doc.add_paragraph()
    paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
    run = paragraph.add_run()

    fldChar1_fig = OxmlElement("w:fldChar")
    fldChar1_fig.set(qn("w:fldCharType"), "begin")

    instrText_fig = OxmlElement("w:instrText")
    instrText_fig.set(qn("xml:space"), "preserve")
    instrText_fig.text = "TOC \\h \\z \\t \"Heading 5,1\""  # ✅ Extracts only "Heading 5" for Figures

    fldChar2_fig = OxmlElement("w:fldChar")
    fldChar2_fig.set(qn("w:fldCharType"), "separate")

    fldChar3_fig = OxmlElement("w:fldChar")
    fldChar3_fig.set(qn("w:fldCharType"), "end")

    run._r.append(fldChar1_fig)
    run._r.append(instrText_fig)
    run._r.append(fldChar2_fig)
    run._r.append(fldChar3_fig)


def add_list_of_tables(doc):
    """Extracts tables from the document using 'Heading 4' and adds a 'List of Tables' section."""
    doc.add_page_break()
    doc.add_paragraph("Tables", "Title")

    for para in doc.paragraphs:
        if para.style.name.startswith("Heading 4") and "Table" in para.text:
            doc.add_paragraph(para.text, style="Normal")


def add_list_of_figures(doc):
    """Extracts figures from the document using 'Heading 4' and adds a 'List of Figures' section."""
    doc.add_page_break()
    doc.add_paragraph("Figures", "Title")

    for para in doc.paragraphs:
        if para.style.name.startswith("Heading 4") and "Figure" in para.text:
            doc.add_paragraph(para.text, style="Normal")



def replace_placeholders(text, placeholders):
    """Replaces placeholders in text dynamically."""
    if not text:
        return ""

    for key, value in placeholders.items():
        text = text.replace(f"{{{key}}}", str(value))

    return text


def add_section(doc, section_title, section_data, section_number, placeholders, numbering_tracker):
    """Recursively adds sections, subsections, and sub-subsections, while managing numbering of tables and figures at the section level."""

    # Determine section heading level
    heading_level = section_number.count(".") + 1

    # Add a page break only for main sections (level 1 heading)
    if heading_level == 1:
        doc.add_page_break()

    # Fetch the correct section title from JSON (fallback to formatted key name if not found)
    json_title = section_data.get("title", section_title.replace("_", " ").title())

    # Add section heading
    doc.add_heading(f"{section_number}. {json_title}", level=heading_level)

    # Precompute table and figure numbers
    computed_table_numbers, computed_figure_numbers = precompute_numbers(section_data, section_number, numbering_tracker)

    # Replace placeholders and add section text
    process_section_text(doc, section_data, placeholders, computed_table_numbers, computed_figure_numbers)

    # Handle special sections: Scope of Work, Regulatory Standards
    filtered_subsections = process_special_sections(section_title, section_data, placeholders, doc)

    # Add bullet lists
    add_bullet_list(doc, section_data, placeholders)

    # Insert tables, images, and graphs
    insert_tables(doc, section_data, placeholders, computed_table_numbers)
    insert_images_and_graphs(doc, section_data, computed_figure_numbers, placeholders)

    # Recursively process subsections
    for idx, (sub_key, sub_data) in enumerate(filtered_subsections.items(), start=1):
        sub_section_number = f"{section_number}.{idx}"
        add_section(doc, sub_key, sub_data, sub_section_number, placeholders, numbering_tracker)


def precompute_numbers(section_data, section_number, numbering_tracker):
    """Precompute table and figure numbers before replacing placeholders."""
    main_section_number = section_number.split(".")[0]  # Extract main section (e.g., "4" from "4.1.2")

    computed_table_numbers = []
    computed_figure_numbers = []  # Includes both images & graphs

    # ✅ Correctly Precompute Table Numbers for Multiple Tables
    if "table" in section_data:
        numbering_tracker["table"][main_section_number] = numbering_tracker["table"].get(main_section_number, 0) + 1
        computed_table_numbers.append(f"{main_section_number}.{numbering_tracker['table'][main_section_number]}")
    elif "tables" in section_data:
        num_tables = len(section_data["tables"])
        for _ in range(num_tables):
            numbering_tracker["table"][main_section_number] = numbering_tracker["table"].get(main_section_number, 0) + 1
            computed_table_numbers.append(f"{main_section_number}.{numbering_tracker['table'][main_section_number]}")

    # ✅ Precompute Figure Numbers for Multiple Images and Graphs
    num_figures = 0
    if "image" in section_data:
        num_figures += 1
    if "images" in section_data:
        num_figures += len(section_data["images"])  # ✅ Count all images
    if "graph" in section_data:
        num_figures += 1

    for _ in range(num_figures):
        numbering_tracker["figure"][main_section_number] = numbering_tracker["figure"].get(main_section_number, 0) + 1
        computed_figure_numbers.append(f"{main_section_number}.{numbering_tracker['figure'][main_section_number]}")

    return computed_table_numbers, computed_figure_numbers


def process_section_text(doc, section_data, placeholders, computed_table_numbers, computed_figure_numbers):
    """Replace placeholders and add section text."""
    text = section_data.get("text", "")

    # Replace `{table_number}` placeholders
    while "{table_number}" in text and computed_table_numbers:
        text = text.replace("{table_number}", computed_table_numbers[0], 1)

    # Replace `{figure_number}` placeholders
    while "{figure_number}" in text and computed_figure_numbers:
        text = text.replace("{figure_number}", computed_figure_numbers[0], 1)

    # Add the final processed text to the document
    if text:
        doc.add_paragraph(replace_placeholders(text, placeholders))


def process_special_sections(section_title, section_data, placeholders, doc):
    """Handle special sections like Scope of Work and Regulatory Standards."""
    if section_title.lower() == "scope of work" and placeholders.get("report_parameters"):
        parameter_list = [p.strip() for p in placeholders["report_parameters"].split(",")]
        formatted_parameters = [format_parameter_section(param) for param in parameter_list]
        for param in formatted_parameters:
            doc.add_paragraph(param, style="List Bullet")

    if section_title.lower() == "regulatory standards" and placeholders.get("report_parameters"):
        parameter_list = [p.strip().lower() for p in placeholders["report_parameters"].split(",")]
        filtered_subsections = {
            key: value for key, value in section_data.get("subsections", {}).items()
            if key.lower() in parameter_list
        }
        if not filtered_subsections:
            print(f"⚠ Warning: No matching regulatory standard found for parameters {parameter_list}.")
        return filtered_subsections

    if section_title.lower() == "conclusion" and placeholders.get("report_parameters"):
        parameter_list = [p.strip().lower() for p in placeholders["report_parameters"].split(",")]

        # Load conclusions and verdict from constants.json
        conclusion_texts = CONSTANTS.get("conclusions", {})
        verdict_text = conclusion_texts.get("verdict",
                                            "This analysis revealed that the observed monitoring parameter(s) consistently adhered to the national standards across all monitored locations at the project site.")

        # Append the relevant conclusions based on selected parameters
        conclusion_paragraphs = []
        for param in parameter_list:
            if param in conclusion_texts:
                conclusion_paragraphs.append(conclusion_texts[param])

        # Add conclusions and verdict to the document
        if conclusion_paragraphs:
            doc.add_paragraph("\n\n".join(conclusion_paragraphs))  # Ensures spacing between paragraphs
            doc.add_paragraph(verdict_text)
        else:
            print(f"⚠ Warning: No matching conclusions found for parameters {parameter_list}.")

    return section_data.get("subsections", {})


def add_bullet_list(doc, section_data, placeholders):
    """Add bullet lists to the document if available."""
    bullet_points = section_data.get("bullet_list", [])
    for point in bullet_points:
        formatted_point = replace_placeholders(point, placeholders)
        doc.add_paragraph(formatted_point, style="List Bullet")



def insert_tables(doc, section_data, placeholders, computed_table_numbers):
    """Insert tables using precomputed table numbers and dynamically inject monitoring data when applicable."""

    # 🔹 Check if section contains a single table or multiple tables
    if "table" in section_data:
        tables = [section_data["table"]]
    elif "tables" in section_data:
        tables = section_data["tables"]
    else:
        return  # No table data present

    # ✅ Known headers for data injection
    air_quality_headers = ["Monitoring Locations", "Date and Time", "CO", "O3", "NO2", "SO2", "PM2.5", "PM10"]
    noise_quality_headers = ["Monitoring Location", "Date and Time", "EQ", "MAX", "AE", "10", "50", "90"]

    # 🔹 Check if this section needs dynamic data injection
    if "title" in section_data:
        section_title = section_data["title"].lower()

        # ✅ Monitoring Locations (Scope of Work)
        if "monitoring locations" in section_title and "monitoring_locations" in placeholders:
            tables[0]["data"] = placeholders["monitoring_locations"]

        # ✅ Inject Air & Noise Quality Data
        if "table" in section_data and "data" in section_data["table"] and section_data["table"]["data"]:
            header_row = section_data["table"]["data"][0]

            if header_row == air_quality_headers and "air_monitoring_data" in placeholders:
                section_data["table"]["data"] = placeholders["air_monitoring_data"]

            elif header_row == noise_quality_headers and "noise_monitoring_data" in placeholders:
                section_data["table"]["data"] = placeholders["noise_monitoring_data"]

    # 🔹 Ensure Correct Number of Table Numbers Are Available
    if len(computed_table_numbers) < len(tables):
        print(f"⚠ Warning: Mismatch between precomputed table numbers and actual tables in section.")
        return  # Avoid index errors

    # 🔹 Insert All Tables
    for index, table_data in enumerate(tables):
        if not isinstance(table_data, dict) or "data" not in table_data or not table_data["data"]:
            print(f"⚠ Warning: Unexpected table format in section. Skipping.")
            continue

        table_number = computed_table_numbers[index]  # ✅ Use correct precomputed number
        table_title = replace_placeholders(table_data.get("title", "Table"), placeholders).replace("{table_number}", table_number)

        doc.add_heading(table_title, level=4)

        # ✅ Insert Updated Table into Document
        table = doc.add_table(rows=len(table_data["data"]), cols=len(table_data["data"][0]))
        table.style = 'Table Grid'

        for row_idx, row_data in enumerate(table_data["data"]):
            for col_idx, cell_data in enumerate(row_data):
                table.cell(row_idx, col_idx).text = replace_placeholders(str(cell_data), placeholders)

        doc.add_paragraph("")


def insert_images_and_graphs(doc, section_data, computed_figure_numbers, placeholders):
    """Insert multiple images and graphs with descriptions, ensuring they are centered and appear below."""

    if section_data.get("title") == "Scope of Work":
        monitoring_locations_section = section_data.get("subsections", {}).get("monitoring_locations", {})

        if "images" not in monitoring_locations_section:
            monitoring_locations_section["images"] = []  # Ensure the key exists

        # ✅ Insert Monitoring Location Map at the beginning (if provided)
        if placeholders.get("monitoring_location_map") and not any(
                img.get("description") == "Monitoring Location Map showing marked locations." for img in
                monitoring_locations_section["images"]):
            monitoring_locations_section["images"].insert(0, {
                "path": placeholders["monitoring_location_map"],
                "description": "Environmental Monitoring Location Map"
            })

        # ✅ Insert Site Images for each monitoring location (if provided)
        if placeholders.get("monitoring_location_images"):
            for location, image_path in placeholders["monitoring_location_images"].items():
                monitoring_locations_section["images"].append({
                    "path": image_path,
                    "description": f"Location {location}"
                })

    # 🔹 Handle Multiple Images
    if "images" in section_data:
        for image_data in section_data["images"]:
            if not isinstance(image_data, dict) or "path" not in image_data:
                print("⚠ Warning: Image data format incorrect. Skipping.")
                continue

            if not computed_figure_numbers:
                print("⚠ Warning: Not enough figure numbers for images.")
                continue

            figure_number = computed_figure_numbers.pop(0)
            image_path = image_data["path"]
            image_description = image_data.get("description", f"Figure {figure_number} - Image Description")

            if os.path.exists(image_path):
                try:
                    with Image.open(image_path) as img:
                        dpi = img.info.get('dpi')  # Extract DPI metadata
                        if dpi is None or dpi[0] == 0 or dpi[1] == 0:
                            print(f"⚠ Warning: DPI metadata missing for {image_path}. Setting default DPI to 96.")
                            dpi = (96, 96)
                        img.save(image_path, dpi=dpi)  # Save with corrected DPI

                    # 🔹 Determine Image Size
                    image_width = Inches(6) if "Location Map" in image_description else Inches(3)  # Larger for Location Map

                    # 🔹 Insert Image and Center Align
                    image_paragraph = doc.add_paragraph()
                    image_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                    run = image_paragraph.add_run()
                    run.add_picture(image_path, width=image_width)  # ✅ Adjust width dynamically

                    # 🔹 Add Image Description Below
                    desc_paragraph = doc.add_heading(f"Figure {figure_number} - {image_description}", level=5)
                    desc_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

                except Exception as e:
                    print(f"⚠ Warning: Failed to insert image {image_path}. Error: {e}")

    # 🔹 Handle Single Image (for backward compatibility)
    elif "image" in section_data:
        figure_number = computed_figure_numbers.pop(0)
        image_path = section_data["image"]
        image_description = section_data.get("image_description", f"Figure {figure_number} - Image Description")

        if os.path.exists(image_path):
            try:
                with Image.open(image_path) as img:
                    dpi = img.info.get('dpi')
                    if dpi is None or dpi[0] == 0 or dpi[1] == 0:
                        print(f"⚠ Warning: DPI metadata missing for {image_path}. Setting default DPI to 96.")
                        dpi = (96, 96)
                    img.save(image_path, dpi=dpi)

                # 🔹 Determine Image Size
                image_width = Inches(3) if "Location Map" in image_description else Inches(1.5)  # Larger for Location Map

                # 🔹 Insert Image and Center Align
                image_paragraph = doc.add_paragraph()
                image_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER
                run = image_paragraph.add_run()
                run.add_picture(image_path, width=image_width)

                # 🔹 Add Image Description Below
                desc_paragraph = doc.add_paragraph(f"Figure {figure_number} - {image_description}")
                desc_paragraph.alignment = WD_PARAGRAPH_ALIGNMENT.CENTER

            except Exception as e:
                print(f"⚠ Warning: Failed to insert image {image_path}. Error: {e}")



def format_parameter_section(parameter):
    """Formats user input parameters into proper section titles."""
    formatted_parameters = {
        "air": "Ambient Air Quality Monitoring",
        "noise": "Noise Monitoring",
        "soil": "Soil Quality Monitoring",
        "water": "Water Quality Monitoring"
    }
    return formatted_parameters.get(parameter.lower(), parameter.capitalize() + " Monitoring")

def generate_report():
    """Generates a monitoring report dynamically based on input data."""

    structure_file = CONSTANTS["structure_file"]


    placeholders = {'consultancy_name': 'Green Fields Environmental Consulting',
                    'contractor_name': 'sdfsafasf',
                    'project_name': 'sfsaf',
                    'project_number': '45343',
                    'reference_number': '23414',
                    'report_frequency': 'Weekly',
                    'report_date': '324234',
                    'report_number': '4324231',
                    'report_parameters': 'Air, Noise',
                    'monitoring_frequency': '30 mins',
                    'monitoring_locations': [['Monitoring Location', 'Description', 'Latitude', 'Longitude'],
                                             ['ml01', 'fdsafda', '421432', '432421'], ['ml02', 'dfsfddsaf', '23424', '432423']],
                    "monitoring_location_map": "monitoring/test_data/map.png",
                    "monitoring_location_images": {'ml01': 'monitoring/test_data/ml01.png',
                                                   'ml02': 'monitoring/test_data/ml02.png'},
                    'air_monitoring_data': [['Monitoring Location', 'Time', 'CO', 'O3', 'NO2', 'SO2', 'PM2.5', 'PM10'],
                                            ['ml01', '423141', '23', '32', '43', '32', '21', '432'],
                                            ['ml02', 'r3242', '432', '23', '32', '43', '23', '34']],
                    'noise_monitoring_data': [['Monitoring Location', 'Time', 'EQ', 'Max', 'AE', '10', '50', '90'],
                                              ['ml01', '3421', '32', '32', '32', '32', '32', '32'],
                                              ['ml02', '342234', '54', '54', '45', '45', '54', '45']]}

    # Load structured JSON
    with open(structure_file, 'r') as file:
        section_data = json.load(file)

    # Convert JSON keys to lowercase for **case-insensitive** lookup
    section_data = {key.lower(): value for key, value in section_data.items()}

    # Define **explicit** section order
    sections = [
        "Introduction",
        "Scope of Work",
        "Regulatory Standards"
    ]

    # 📌 Insert User-defined Sections Dynamically
    if placeholders["report_parameters"]:
        parameter_list = [p.strip() for p in placeholders["report_parameters"].split(",")]
        formatted_parameters = [format_parameter_section(param) for param in parameter_list]
        sections.extend(formatted_parameters)

    # 📌 Add Final Sections
    sections.extend(["Conclusion", "Appendices"])

    # Initialize table/figure/graph numbering tracker
    numbering_tracker = {
        "table": {},
        "figure": {},
        "graph": {}
    }

    # Create Word document
    doc = Document()
    # template_path = "monitoring/config/template.docx"
    # doc = Document(template_path)


    add_page_number(doc)

    # 📌 Title Page
    # add_title_page(doc, placeholders["report_frequency"])

    # 📌 Table of Contents
    add_table_of_contents(doc)

    # 📌 Generate Sections
    for i, section_key in enumerate(sections, start=1):
        # Convert key to lowercase for safe lookup
        section_key_lower = section_key.lower().replace(" ", "_")

        if section_key_lower in section_data:
            section_title = section_key  # ✅ Fetch correct title
            add_section(doc, section_title, section_data[section_key_lower], str(i), placeholders, numbering_tracker)
        else:
            print(f"⚠ Warning: Section '{section_key}' not found in JSON.")


    # 📌 Save Document
    output_dir = CONSTANTS["output_dir"]
    os.makedirs(output_dir, exist_ok=True)
    report_path = f"{output_dir}/{placeholders['report_frequency'].capitalize()}_Monitoring_Report.docx"
    doc.save(report_path)

    print(f"✅ {placeholders['report_frequency'].capitalize()} Monitoring Report generated: {report_path}")
    return report_path
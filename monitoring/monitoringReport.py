import os
import json
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.enum.text import WD_PARAGRAPH_ALIGNMENT
from docx.shared import Inches

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


def add_table_of_contents(doc):
    """Adds a TOC field that updates when `F9` is pressed in Word."""
    doc.add_paragraph("Table of Contents", "Title")

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

    # Add page break only for main sections (level 1 heading)
    if heading_level == 1:
        doc.add_page_break()

    # Format section title properly (capitalize words)
    formatted_title = " ".join(word.capitalize() for word in section_title.split("_"))

    # Add section heading
    doc.add_heading(f"{section_number} {formatted_title}", level=heading_level)

    # Add text if available
    text = replace_placeholders(section_data.get("text", ""), placeholders)
    if text:
        doc.add_paragraph(text)

    # **🔹 Add Bullet List if Available**
    bullet_points = section_data.get("bullet_list", [])
    for point in bullet_points:
        formatted_point = replace_placeholders(point, placeholders)
        doc.add_paragraph(formatted_point, style="List Bullet")  # ✅ Adds bullet points correctly

    # **Global Table & Figure Numbering Logic**
    main_section_number = section_number.split(".")[0]  # Extract only main section (e.g., `4` from `4.1.2`)

    # Add table if available
    if "table" in section_data:
        numbering_tracker["table"][main_section_number] = numbering_tracker["table"].get(main_section_number, 0) + 1
        table_number = f"{main_section_number}.{numbering_tracker['table'][main_section_number]}"

        table_data = section_data["table"]
        table_title = replace_placeholders(table_data.get("title", "Table"), placeholders).replace("{table_number}",
                                                                                                   table_number)
        doc.add_paragraph(table_title, style="Heading 3")

        table = doc.add_table(rows=len(table_data["data"]), cols=len(table_data["data"][0]))
        table.style = 'Table Grid'

        for row_idx, row_data in enumerate(table_data["data"]):
            for col_idx, cell_data in enumerate(row_data):
                table.cell(row_idx, col_idx).text = replace_placeholders(str(cell_data), placeholders)

    # Add image if available
    if "image" in section_data:
        numbering_tracker["figure"][main_section_number] = numbering_tracker["figure"].get(main_section_number, 0) + 1
        figure_number = f"{main_section_number}.{numbering_tracker['figure'][main_section_number]}"

        doc.add_paragraph(f"Figure {figure_number} - Image Description", style="Heading 3")
        image_path = section_data["image"]
        if os.path.exists(image_path):
            doc.add_picture(image_path, width=Inches(5))

    # **Recursive Call for Subsections**
    subsections = section_data.get("subsections", {})
    for idx, (sub_key, sub_data) in enumerate(subsections.items(), start=1):
        sub_section_number = f"{section_number}.{idx}"
        add_section(doc, sub_key, sub_data, sub_section_number, placeholders, numbering_tracker)


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

    placeholders = {
        "contractor_name": "Amala",
        "project_name": "Concrete structure work of the HW1 Shura Island.",
        "report_frequency": "Monthly",
        "report_date": "06th January 2025",
        "report_number": "Twenty-third",
        "report_parameters": "Air, Noise"
    }

    constant_text_json_file = CONSTANTS["constant_text_json_file"]

    # Load structured JSON
    with open(constant_text_json_file, 'r') as file:
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
    add_page_number(doc)

    # 📌 Title Page
    add_title_page(doc, placeholders["report_frequency"])

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

    # 📌 Collect User Inputs and Store in Placeholders Dictionary
    # placeholders = {
    #     "contractor_name": input("Enter the contractor's name: ").strip(),
    #     "project_name": input("Enter the project name: ").strip(),
    #     "report_frequency": input("Enter report frequency (Weekly, Monthly): ").strip().lower(),
    #     "report_date": input("Enter the report date (e.g., 06th January 2025): ").strip(),
    #     "report_number": input("Enter the report number (e.g., Twenty-third): ").strip(),
    #     "report_parameters": input("Enter report parameters (comma-separated, e.g., Air, Noise, Soil): ").strip()
    # }
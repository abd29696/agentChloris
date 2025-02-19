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

    # Add a page break only for main sections (level 1 heading)
    if heading_level == 1:
        doc.add_page_break()

    # Fetch the correct section title from JSON (fallback to formatted key name if not found)
    json_title = section_data.get("title", section_title.replace("_", " ").title())

    # Add section heading
    doc.add_heading(f"{section_number} {json_title}", level=heading_level)

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
    insert_images_and_graphs(doc, section_data, computed_figure_numbers)

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
    if "table" in section_data:  # Single table case
        numbering_tracker["table"][main_section_number] = numbering_tracker["table"].get(main_section_number, 0) + 1
        computed_table_numbers.append(f"{main_section_number}.{numbering_tracker['table'][main_section_number]}")

    elif "tables" in section_data:  # Multiple tables case
        num_tables = len(section_data["tables"])
        for _ in range(num_tables):
            numbering_tracker["table"][main_section_number] = numbering_tracker["table"].get(main_section_number, 0) + 1
            computed_table_numbers.append(f"{main_section_number}.{numbering_tracker['table'][main_section_number]}")

    # ✅ Precompute Figure Numbers (For Both Images & Graphs)
    num_figures = sum(1 for key in ["image", "graph"] if key in section_data)
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
    """Insert tables using precomputed table numbers."""

    # ✅ Handle Single Table
    if "table" in section_data:
        tables = [section_data["table"]]  # Convert single table into a list
    elif "tables" in section_data:
        tables = section_data["tables"]  # Use list of tables
    else:
        return  # No table data present

    # ✅ Ensure Correct Number of Table Numbers Are Available
    if len(computed_table_numbers) < len(tables):
        print(f"⚠ Warning: Mismatch between precomputed table numbers and actual tables in section.")
        return  # Avoid index errors

    # ✅ Insert All Tables
    for index, table_data in enumerate(tables):
        if not isinstance(table_data, dict):
            print(f"⚠ Warning: Unexpected table format in section. Skipping.")
            continue

        table_number = computed_table_numbers[index]  # ✅ Use correct precomputed number
        table_title = replace_placeholders(table_data.get("title", "Table"), placeholders).replace("{table_number}", table_number)

        doc.add_paragraph(table_title, style="Heading 3")

        # ✅ Insert Table
        table = doc.add_table(rows=len(table_data["data"]), cols=len(table_data["data"][0]))
        table.style = 'Table Grid'

        for row_idx, row_data in enumerate(table_data["data"]):
            for col_idx, cell_data in enumerate(row_data):
                table.cell(row_idx, col_idx).text = replace_placeholders(str(cell_data), placeholders)



def insert_images_and_graphs(doc, section_data, computed_figure_numbers):
    """Insert images and graphs using precomputed figure numbers."""
    if "image" in section_data:
        figure_number = computed_figure_numbers.pop(0)
        doc.add_paragraph(f"Figure {figure_number} - Image Description", style="Heading 3")

        image_path = section_data["image"]
        if os.path.exists(image_path):
            doc.add_picture(image_path, width=Inches(5))

    if "graph" in section_data:
        figure_number = computed_figure_numbers.pop(0)
        doc.add_paragraph(f"Figure {figure_number} - Graph Description", style="Heading 3")

        graph_path = section_data["graph"]
        if os.path.exists(graph_path):
            doc.add_picture(graph_path, width=Inches(5))


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

    placeholders = {
        "consultancy_name": CONSTANTS["consultancy_name"],
        "contractor_name": "Amala",
        "project_name": "Concrete structure work of the HW1 Shura Island.",
        "report_frequency": "Monthly",
        "report_date": "06th January 2025",
        "report_number": "Twenty-third",
        "report_parameters": "Air, Noise"
    }


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
import gradio as gr
from monitoring.monitoringReport import generate_report

# Global storage for monitoring data
monitoring_data = []
air_data = []
noise_data = []


def add_monitoring_location(location, description, latitude, longitude):
    """Adds a monitoring location entry to the table and resets input fields."""
    if location and description and latitude and longitude:
        monitoring_data.append([location, description, latitude, longitude])

    return monitoring_data, "", "", "", ""  # Resets input fields


def add_air_data(location, datetime, co, o3, no2, so2, pm25, pm10):
    """Adds air quality monitoring data to the table and resets input fields."""
    if location and datetime and co and o3 and no2 and so2 and pm25 and pm10:
        air_data.append([location, datetime, co, o3, no2, so2, pm25, pm10])

    return air_data, "", "", "", "", "", "", "", ""  # Resets input fields


def add_noise_data(location, datetime, eq, max_val, ae, val10, val50, val90):
    """Adds noise monitoring data to the table and resets input fields."""
    if location and datetime and eq and max_val and ae and val10 and val50 and val90:
        noise_data.append([location, datetime, eq, max_val, ae, val10, val50, val90])

    return noise_data, "", "", "", "", "", "", "", ""  # Resets input fields


def toggle_air_section(selected_parameters):
    """Toggles the Air Monitoring input fields visibility based on checkbox selection."""
    return gr.update(visible="Air" in selected_parameters)


def toggle_noise_section(selected_parameters):
    """Toggles the Noise Monitoring input fields visibility based on checkbox selection."""
    return gr.update(visible="Noise" in selected_parameters)


def generate_and_download_report(contractor_name, project_name, project_number, reference_number, report_frequency,
                                 report_date, report_number, monitoring_frequency, report_parameters):
    """Handles report generation and provides a download link."""

    # Ensure report_parameters is always a string
    parameters_text = ", ".join(report_parameters) if report_parameters else "None"

    placeholders = {
        "consultancy_name": "Green Fields Environmental Consulting",
        "contractor_name": contractor_name,
        "project_name": project_name,
        "project_number": project_number,
        "reference_number": reference_number,
        "report_frequency": report_frequency,
        "report_date": report_date,
        "report_number": report_number,
        "report_parameters": parameters_text,
        "monitoring_frequency": monitoring_frequency,
        "monitoring_locations": [["Monitoring Location", "Description", "Latitude", "Longitude"]] + monitoring_data,
        "air_monitoring_data": [["Monitoring Location", "Time", "CO", "O3", "NO2", "SO2", "PM2.5", "PM10"]] + air_data,
        "noise_monitoring_data": [["Monitoring Location", "Time", "EQ", "Max", "AE", "10", "50", "90"]] + noise_data,
    }

    # ✅ Generate report
    report_path = generate_report(placeholders)

    return report_path


# ✅ Create UI
with gr.Blocks() as demo:
    gr.Markdown("## Agent Chloris")

    report_type = gr.Dropdown(["Monitoring", "CESMP"], label="Select Report Type")

    contractor_name = gr.Textbox(label="Contractor Name")
    project_name = gr.Textbox(label="Project Name")
    project_number = gr.Textbox(label="Project Number")
    reference_number = gr.Textbox(label="Reference Number")
    report_frequency = gr.Dropdown(["Weekly", "Monthly"], label="Report Frequency")
    report_date = gr.Textbox(label="Report Date (e.g., 06Jan2025)")
    report_number = gr.Textbox(label="Report Number")

    gr.Markdown("### Add Monitoring Location Data")
    monitoring_location = gr.Textbox(label="Monitoring Location")
    monitoring_description = gr.Textbox(label="Description")
    monitoring_latitude = gr.Textbox(label="Latitude")
    monitoring_longitude = gr.Textbox(label="Longitude")

    add_data_button = gr.Button("Add Data")
    monitoring_table = gr.Dataframe(headers=["Monitoring Location", "Description", "Latitude", "Longitude"],
                                    datatype=["str", "str", "str", "str"],
                                    label="Monitoring Locations Table")

    add_data_button.click(fn=add_monitoring_location,
                          inputs=[monitoring_location, monitoring_description, monitoring_latitude,
                                  monitoring_longitude],
                          outputs=[monitoring_table, monitoring_location, monitoring_description, monitoring_latitude,
                                   monitoring_longitude])

    monitoring_frequency = gr.Dropdown(["15 mins", "30 mins", "1 hr", "24 hr"], label="Monitoring Frequency")
    report_parameters = gr.CheckboxGroup(
        ["Air", "Noise", "Soil", "Ground Water", "Sea Water", "Emission", "Vibration"],
        label="Monitoring Parameters"
    )

    # ✅ Air Monitoring Section (Hidden by default)
    with gr.Column(visible=False) as air_section:
        gr.Markdown("### Add Air Monitoring Data")
        air_location = gr.Textbox(label="Monitoring Location")
        air_datetime = gr.Textbox(label="Date and Time")
        air_co = gr.Textbox(label="CO")
        air_o3 = gr.Textbox(label="O3")
        air_no2 = gr.Textbox(label="NO2")
        air_so2 = gr.Textbox(label="SO2")
        air_pm25 = gr.Textbox(label="PM2.5")
        air_pm10 = gr.Textbox(label="PM10")

        add_air_button = gr.Button("Add Air Data")
        air_table = gr.Dataframe(
            headers=["Monitoring Location", "Date and Time", "CO", "O3", "NO2", "SO2", "PM2.5", "PM10"],
            datatype=["str", "str", "str", "str", "str", "str", "str", "str"],
            label="Air Monitoring Table"
        )

        add_air_button.click(fn=add_air_data,
                             inputs=[air_location, air_datetime, air_co, air_o3, air_no2, air_so2, air_pm25, air_pm10],
                             outputs=[air_table, air_location, air_datetime, air_co, air_o3, air_no2, air_so2, air_pm25,
                                      air_pm10])

    # ✅ Noise Monitoring Section (Hidden by default)
    with gr.Column(visible=False) as noise_section:
        gr.Markdown("### Add Noise Monitoring Data")
        noise_location = gr.Textbox(label="Monitoring Location")
        noise_datetime = gr.Textbox(label="DateTime")
        noise_eq = gr.Textbox(label="EQ")
        noise_max = gr.Textbox(label="Max")
        noise_ae = gr.Textbox(label="AE")
        noise_val10 = gr.Textbox(label="10")
        noise_val50 = gr.Textbox(label="50")
        noise_val90 = gr.Textbox(label="90")

        add_noise_button = gr.Button("Add Noise Data")
        noise_table = gr.Dataframe(
            headers=["Monitoring Location", "DateTime", "EQ", "Max", "AE", "10", "50", "90"],
            datatype=["str", "str", "str", "str", "str", "str", "str", "str"],
            label="Noise Monitoring Table"
        )

        add_noise_button.click(fn=add_noise_data,
                               inputs=[noise_location, noise_datetime, noise_eq, noise_max, noise_ae, noise_val10,
                                       noise_val50, noise_val90],
                               outputs=[noise_table, noise_location, noise_datetime, noise_eq, noise_max, noise_ae,
                                        noise_val10, noise_val50, noise_val90])

    # ✅ Show Air & Noise Sections Dynamically
    report_parameters.change(fn=toggle_air_section, inputs=[report_parameters], outputs=[air_section])
    report_parameters.change(fn=toggle_noise_section, inputs=[report_parameters], outputs=[noise_section])

    generate_button = gr.Button("Generate Report as Word")
    download_output = gr.File(label="Download Report")

    generate_button.click(fn=generate_and_download_report,
                          inputs=[contractor_name, project_name, project_number, reference_number, report_frequency,
                                  report_date, report_number, monitoring_frequency, report_parameters],
                          outputs=[download_output])

# ✅ Launch UI
demo.launch()

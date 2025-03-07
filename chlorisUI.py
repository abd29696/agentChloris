import gradio as gr
from monitoring.monitoringReport import generate_report

# Global storage for monitoring data
monitoring_data = []
air_data = []
noise_data = []
location_images = {}
monitoring_location_map = None

class OceanDefaultTheme(gr.themes.Default):
    def __init__(self, **kwargs):
        super().__init__(**kwargs)
        self.primary_hue = "#047857"  # Set primary color to teal
        self.button_radius = "lg"  # Larger rounded buttons
        self.button_shadow = "md"  # Soft button shadow
        self.button_primary_background_fill = "#047857"  # Custom button fill color
        self.button_primary_background_fill_hover = "teal"  # Darker hover effect
        self.button_primary_text_color = "white"  # White text for contrast
        self.button_primary_border_color = "#047857"  # ✅ Match border color to button
        self.button_primary_border_color_hover = "#047857"  # ✅ Keep it subtle
        self.button_primary_focus_ring_color = "#047857"  # ✅ Soft focus glow instead of orange




custom_theme = OceanDefaultTheme()

def upload_monitoring_map(file):
    """Stores the uploaded Monitoring Location Map image path."""
    global monitoring_location_map
    if file:
        monitoring_location_map = file
        return "✅ Map uploaded successfully!"
    return "⚠ Please upload a valid image file."

def add_monitoring_location(location, description, latitude, longitude, image):
    """Adds a monitoring location entry to the table and resets input fields."""
    if location and description and latitude and longitude:
        monitoring_data.append([location, description, latitude, longitude])

    if image:
        location_images[location] = image

    return monitoring_data, "", "", "", "", None  # Resets input fields


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

def show_image():
    return "chloris.png"


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
        "monitoring_location_map": monitoring_location_map,
        "monitoring_location_images": location_images,
        "air_monitoring_data": [["Monitoring Location", "Time", "CO", "O3", "NO2", "SO2", "PM2.5", "PM10"]] + air_data,
        "noise_monitoring_data": [["Monitoring Location", "Time", "EQ", "Max", "AE", "10", "50", "90"]] + noise_data,
    }

    # ✅ Generate report
    report_path = generate_report(placeholders)

    return report_path, gr.update(visible=True)


# ✅ Create UI
with gr.Blocks(theme=custom_theme) as demo:

    with gr.Column():
        with gr.Row():
            image = gr.Image(value="chloris.png", label="Agent Chloris", interactive=False)

            with gr.Column():
                contractor_name = gr.Textbox(label="Contractor Name")
                with gr.Column():
                    reference_number = gr.Textbox(label="Reference Number")
                with gr.Column():
                    project_name = gr.Textbox(label="Project Name")
                    project_number = gr.Textbox(label="Project Number")

            with gr.Column():
                gr.Markdown("Add Report Details")
                report_type = gr.Dropdown(["Monitoring", "CESMP"], label="Select Report Type")
                report_date = gr.Textbox(label="Report Date (e.g., 06Jan2025)")
                report_frequency = gr.Dropdown(["Weekly", "Monthly"], label="Report Frequency")
                report_number = gr.Textbox(label="Report Number")



        with gr.Column():
            with gr.Row():
                with gr.Column():
                    report_parameters = gr.CheckboxGroup(
                        ["Air", "Noise", "Soil Quality", "Ground Water", "Sea Water", "Emission", "Vibration"],
                        label="Monitoring Parameters"
                    )
                    with gr.Column():
                        monitoring_frequency = gr.Dropdown(["15 mins", "30 mins", "1 hr", "24 hr"],
                                                       label="Monitoring Frequency")
                with gr.Column():
                    monitoring_map_upload = gr.File(label="Upload Monitoring Location Map")
                    monitoring_map_upload.change(fn=upload_monitoring_map, inputs=[monitoring_map_upload])


    with gr.Column():
        gr.Markdown("### Add Monitoring Location Data")

        with gr.Row():
            location_image = gr.File(label="Upload Location Image")

        with gr.Row():
            monitoring_location = gr.Textbox(label="Monitoring Location")
            monitoring_description = gr.Textbox(label="Description")
            monitoring_latitude = gr.Textbox(label="Latitude")
            monitoring_longitude = gr.Textbox(label="Longitude")
            add_data_button = gr.Button("Add Data", variant='primary')



    with gr.Row():
        monitoring_table = gr.Dataframe(headers=["Monitoring Location", "Description", "Latitude", "Longitude"],
                                    datatype=["str", "str", "str", "str"],
                                    label="Monitoring Locations Table")
        add_data_button.click(fn=add_monitoring_location,
                          inputs=[monitoring_location, monitoring_description, monitoring_latitude,
                                  monitoring_longitude, location_image],
                          outputs=[monitoring_table, monitoring_location, monitoring_description, monitoring_latitude,
                                   monitoring_longitude, location_image])


    # ✅ Air Monitoring Section (Hidden by default)
    with gr.Column(visible=False) as air_section:
        gr.Markdown("### Add Air Monitoring Data")
        with gr.Row():
            air_location = gr.Textbox(label="Monitoring Location")
            air_datetime = gr.Textbox(label="Date and Time")
        with gr.Row():
            air_co = gr.Textbox(label="CO")
            air_o3 = gr.Textbox(label="O3")
            air_no2 = gr.Textbox(label="NO2")
            air_so2 = gr.Textbox(label="SO2")
            air_pm25 = gr.Textbox(label="PM2.5")
            air_pm10 = gr.Textbox(label="PM10")
            add_air_button = gr.Button("Add Air Data", variant='primary')


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
        with gr.Row():
            noise_location = gr.Textbox(label="Monitoring Location")
            noise_datetime = gr.Textbox(label="DateTime")
        with gr.Row():
            noise_eq = gr.Textbox(label="EQ")
            noise_max = gr.Textbox(label="Max")
            noise_ae = gr.Textbox(label="AE")
            noise_val10 = gr.Textbox(label="10")
            noise_val50 = gr.Textbox(label="50")
            noise_val90 = gr.Textbox(label="90")
            add_noise_button = gr.Button("Add Noise Data", variant='primary')


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

    with gr.Row():
        generate_button = gr.Button("Generate Report as Word", variant="primary")
        generate_button2 = gr.Button("Generate Report as PDF")


    download_output = gr.File(label="Download Report", visible=False)



    generate_button.click(fn=generate_and_download_report,
                          inputs=[contractor_name, project_name, project_number, reference_number, report_frequency,
                                  report_date, report_number, monitoring_frequency, report_parameters],
                          outputs=[download_output, download_output])

# ✅ Launch UI
demo.launch(share=True)

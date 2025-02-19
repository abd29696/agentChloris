from monitoring.monitoringReport import generate_report

def select_report_type():
    # report_type = input("Select report type (monitoring, CESMP): ").strip().lower()
    report_type = "monitoring"
    if report_type not in ["monitoring", "cesmp"]:
        print("Invalid report type. Please choose monitoring or CESMP.")
        return None
    return report_type

def main():
    report_type = select_report_type()
    if report_type == "monitoring":
        generate_report()
    elif report_type == "cesmp":
        print("WIP")

if __name__ == "__main__":
    main()

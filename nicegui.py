import polars as pl
from nicegui import ui

# Import your existing pure-Polars engines directly:
# from hrdash.attendance import engine as attendance_engine
# from hrdash.payroll import engine as payroll_engine
# from hrdash.reporting import engine as reporting_engine


def load_attendance_tab():
    ui.label("Attendance Tracking").classes("text-xl font-bold mb-2")

    # Reactive controls update only the container, not the entire script
    with ui.row().classes("items-center mb-4 gap-4"):
        date_input = ui.input("Date", value="2026-09-18")
        ui.button("Apply Filter", on_click=lambda: update_attendance_view())

    table_container = ui.element("div").classes("w-full")

    def update_attendance_view():
        # Example Polars interaction: Replace with call to attendance_engine
        df = pl.DataFrame(
            {
                "Employee ID": ["EMP001", "EMP002", "EMP003"],
                "Status": ["Present", "Late", "Present"],
                "Check-In": ["08:55", "09:15", "08:50"],
            }
        )

        table_container.clear()
        with table_container:
            ui.table(
                columns=[
                    {"name": c, "label": c, "field": c, "align": "left"}
                    for c in df.columns
                ],
                rows=df.to_dicts(),
                row_key="Employee ID",
            ).classes("w-full")

    update_attendance_view()


def load_payroll_tab():
    ui.label("Payroll Engine").classes("text-xl font-bold mb-2")
    with ui.row().classes("items-center gap-4"):
        ui.number(
            label="Overtime Multiplier",
            value=1.5,
            format="%.2f",
            on_change=lambda e: ui.notify(f"Updated rate: {e.value}"),
        )
        ui.button(
            "Calculate Payroll",
            on_click=lambda: ui.notify("Payroll processing completed."),
        )


def load_reporting_tab():
    ui.label("Reporting & Analytics").classes("text-xl font-bold mb-2")
    ui.button(
        "Export Summary PDF",
        icon="download",
        on_click=lambda: ui.notify("Report generated successfully."),
    )


# --- Application Shell ---
with ui.header().classes("bg-slate-800 text-white justify-between items-center"):
    ui.label("HRDash Engine").classes("text-lg font-semibold px-2")

with ui.tabs().classes("w-full border-b") as tabs:
    tab_att = ui.tab("Attendance", icon="badge")
    tab_pay = ui.tab("Payroll", icon="payments")
    tab_rep = ui.tab("Reporting", icon="analytics")

with ui.tab_panels(tabs, value=tab_att).classes("w-full p-6"):
    with ui.tab_panel(tab_att):
        load_attendance_tab()
    with ui.tab_panel(tab_pay):
        load_payroll_tab()
    with ui.tab_panel(tab_rep):
        load_reporting_tab()

# For PyInstaller desktop execution set native=True (requires pywebview)
ui.run(title="HRDash", native=False, port=8080, reload=False)

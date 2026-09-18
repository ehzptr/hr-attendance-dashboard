import polars as pl
from taipy.gui import Gui, notify

# Import pure-Polars engines directly:
# from hrdash.attendance import engine as attendance_engine
# from hrdash.payroll import engine as payroll_engine
# from hrdash.reporting import engine as reporting_engine

# --- State Variables ---
date_input = "2026-09-18"
overtime_mult = 1.5

# Polars data converted to Pandas for Taipy's native table component
df_attendance = pl.DataFrame(
    {
        "Employee ID": ["EMP001", "EMP002", "EMP003"],
        "Status": ["Present", "Late", "Present"],
        "Check-In": ["08:55", "09:15", "08:50"],
    }
).to_pandas()


# --- Callbacks ---
def apply_attendance_filter(state):
    notify(state, "info", f"Filtering attendance for date: {state.date_input}")
    # Example: state.df_attendance = attendance_engine.run(state.date_input).to_pandas()


def calculate_payroll(state):
    notify(
        state, "success", f"Payroll calculated with {state.overtime_mult}x multiplier"
    )


def export_report(state):
    notify(state, "success", "Report PDF exported successfully!")


# --- Page Definitions (Taipy Markdown Syntax) ---
root_page = """
<|navbar|>
"""

attendance_page = """
## Attendance Tracking

<|layout|columns=1 3|gap=1rem|
<|{date_input}|input|label=Date|>
<|Apply Filter|button|on_action=apply_attendance_filter|>
|>

<|{df_attendance}|table|show_all=True|rebuild|>
"""

payroll_page = """
## Payroll Engine

<|layout|columns=1 2|gap=1rem|
<|{overtime_mult}|number|label=Overtime Multiplier|step=0.1|>
<|Calculate Payroll|button|on_action=calculate_payroll|>
|>
"""

reporting_page = """
## Reporting & Analytics

<|Export Summary PDF|button|on_action=export_report|>
"""

# Register multi-page navigation layout
pages = {
    "/": root_page,
    "Attendance": attendance_page,
    "Payroll": payroll_page,
    "Reporting": reporting_page,
}

if __name__ == "__main__":
    Gui(pages=pages).run(
        title="HRDash Engine", port=8080, use_reloader=False, dark_mode=False
    )

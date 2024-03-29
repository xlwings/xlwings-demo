import xlwings as xw
from xlwings.pro import reports
import pandas as pd
from pathlib import Path


def table(rng: xw.Range, df: pd.DataFrame):
    """This is the formatter function"""
    # Header
    rng[0, :].color = "#A9D08E"
    rng[0, :].font.bold = True

    # Rows
    for ix, row in enumerate(rng.rows[1:]):
        if ix % 2 == 0:
            row.color = "#D0CECE"  # Even rows

    # Columns
    for ix, col in enumerate(df.columns):
        if "Weight" in col:
            rng[1:, ix].number_format = "0.0%"
    
    rng.autofit()

reports.register_formatter(table)

def main():
    template_sheet = xw.Book.caller().sheets.active
    report_sheet = template_sheet.copy()
    csv_path = Path(__file__).resolve().parent / 'holdings.csv'
    with template_sheet.book.app.properties(screen_updating=False):
        report_sheet.render_template(holdings=pd.read_csv(csv_path))

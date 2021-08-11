import xlwings as xw
import pandas as pd
from pathlib import Path


def main():
    template_sheet = xw.Book.caller().sheets.active
    report_sheet = template_sheet.copy()
    csv_path = Path(__file__).resolve().parent / 'holdings.csv'
    with template_sheet.book.app.properties(screen_updating=False):
        report_sheet.render_template(holdings=pd.read_csv(csv_path))


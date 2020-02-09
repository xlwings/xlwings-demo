import os
import pickle
import datetime as dt
import xlwings as xw
from xlwings.reports import create_report  # part of xlwings PRO


def main():
    # Files
    template = xw.Book.caller()
    template_path = template.fullname
    report_path = os.path.join(os.path.dirname(template.fullname), 'fund_report.xlsx')

    # Get your data via SQL, APIs, text files, Excel files etc.
    with open(os.path.join(os.path.dirname(__file__), 'data.pickle'), 'rb') as f:
        historical_perf, asset_allocation, top_ten_holdings, calendar_year_tot_ret, tot_ret = pickle.load(f)

    # Data wrangling (with pandas, obviously)
    historical_perf = historical_perf.resample('M').last()

    # Configuration (optional)
    date_format = template.sheets['Config']['date_format'].value
    if date_format == 'UK':
        fmt = '%e %b %Y'
    elif date_format == 'US':
        fmt = '%b %e, %Y'
    else:
        fmt = '%e %b %Y'

    # Collect all data
    data = dict(perf_start_date=dt.datetime(2009, 1, 1).strftime(fmt),
                perf_end_date=dt.date.today().strftime(fmt),
                reference_date=dt.date.today().strftime(fmt),
                total_net_assets=123,
                historical_perf=historical_perf,
                asset_allocation=asset_allocation,
                top_ten_holdings=top_ten_holdings,
                tot_ret=tot_ret,
                calendar_year_tot_ret=calendar_year_tot_ret,
                fund_name='xlwings Fund'
                )

    app = template.app
    app.screen_updating = False

    # Create the Excel report
    wb = create_report(template_path, report_path, app=app, **data)
    wb.sheets.active['A1'].select()


if __name__ == '__main__':
    # This part is to run the script directly from Python, not via Excel
    xw.Book(os.path.join(os.path.dirname(__file__), 'fund_template.xlsx')).set_mock_caller()
    main()

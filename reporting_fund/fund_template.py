import os
import pickle
import datetime as dt
from xlwings_reports import create_report  # not part of the open-source xlwings package


def main():
    # Get your data via SQL, APIs, text files, Excel files etc.
    with open(os.path.join(os.path.dirname(__file__), 'data.pickle'), 'rb') as f:
        historical_perf, asset_allocation, top_ten_holdings, calendar_year_tot_ret, tot_ret = pickle.load(f)

    # Manipulate your data (with pandas, obviously)
    historical_perf = historical_perf.resample('M').last()

    # Collect all data
    fmt = '%e %b %Y'
    data = dict(perf_start_date=dt.datetime(2009, 1, 1).strftime(fmt),
                perf_end_date=dt.date.today().strftime(fmt),
                reference_date=dt.date.today().strftime(fmt),
                total_net_assets=123,
                historical_perf=historical_perf,
                asset_allocation=asset_allocation,
                top_ten_holdings=top_ten_holdings,
                tot_ret=tot_ret,
                calendar_year_tot_ret=calendar_year_tot_ret,
                )

    # Create the Excel report
    template = os.path.join(os.path.dirname(__file__), 'fund_template.xlsx')
    report = os.path.join(os.path.dirname(__file__), 'fund_report.xlsx')
    wb = create_report(template, report, **data)


if __name__ == '__main__':
    main()

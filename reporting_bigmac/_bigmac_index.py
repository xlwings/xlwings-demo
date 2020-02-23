import os
import pandas as pd
import xlwings as xw
from xlwings.reports import create_report  # part of xlwings PRO


def main():
    template = xw.Book.caller()

    # Config
    date = template.sheets['Config']['date'].value.date()
    currency = template.sheets['Config']['currency'].value
    nb_weak = int(template.sheets['Config']['nb_weak'].value)
    nb_strong = int(template.sheets['Config']['nb_strong'].value)
    nb_cheap = int(template.sheets['Config']['nb_cheap'].value)
    nb_expensive = int(template.sheets['Config']['nb_expensive'].value)

    # Get Big Mac index data from GitHub
    # url = 'https://raw.githubusercontent.com/TheEconomist/big-mac-data/master/output-data/big-mac-raw-index.csv'
    url = os.path.join(os.path.dirname(template.fullname), 'big-mac-raw-index.csv')
    raw = pd.read_csv(url, index_col=0, parse_dates=True)

    # Data wrangling
    summary = raw.loc[date, ['name', 'currency_code', 'dollar_price', 'USD']]
    summary = summary.set_index('name')
    summary.index.name = 'Country'
    summary = summary.rename(columns={"currency_code": "Currency", "dollar_price": "Big Mac Price (USD)",
                                      "USD": "Over/undervalued"})

    valuation = summary.drop(columns=['Big Mac Price (USD)'], index=['United States'])
    valuation = valuation.sort_values(by=['Over/undervalued'], ascending=False)
    strong_valuation = valuation.head(nb_strong)
    weak_valuation = valuation.tail(nb_weak)

    price = summary.drop(columns=['Over/undervalued'])
    price = price.sort_values(by=['Big Mac Price (USD)'], ascending=False)
    expensive = price.head(nb_expensive)
    cheap = price.tail(nb_cheap)

    valuation_history = raw.reset_index().set_index('currency_code')
    valuation_history = valuation_history.loc[currency, ['date', 'USD']]
    valuation_history = valuation_history.set_index('date').sort_index()

    # Create Report
    template_path = template.fullname
    report_path = os.path.join(os.path.dirname(template.fullname), f'report_{date}.xlsx')
    app = template.app
    app.screen_updating = False

    data = dict(date=date.strftime('%b %e, %Y'), chart_currency=currency,
                strong_valuation=strong_valuation, weak_valuation=weak_valuation,
                nb_strong=nb_strong, nb_weak=nb_weak, valuation_history=valuation_history,
                expensive=expensive, nb_expensive=nb_expensive,
                cheap=cheap, nb_cheap=nb_cheap)
    wb = create_report(template_path, report_path, app=app, **data)

    app.screen_updating = True
    wb.sheets.active['A1'].select()


if __name__ == '__main__':
    xw.Book('bigmac_index.xlsx').set_mock_caller()
    main()

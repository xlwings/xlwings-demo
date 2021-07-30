import os
from pathlib import Path
from textwrap import dedent

import pandas as pd
from matplotlib.figure import Figure
import xlwings as xw
# Requires a license key: https://www.xlwings.org/trial
from xlwings.pro import Markdown, MarkdownStyle
from xlwings.pro.reports import create_report, Image


def main():
    template = xw.Book.caller()
    template_path = template.fullname
    report_path = os.path.join(os.path.dirname(template.fullname), 'report.xlsx')

    # Matplotlib
    fig = Figure(figsize=(4, 3))
    ax = fig.add_subplot(111)
    ax.plot([1, 2, 3, 4, 5])

    # Pandas DataFrame
    perf_data = pd.DataFrame(index=['r1', 'r1'],
                             columns=['c0', 'c1'],
                             data=[[1., 2.], [3., 4.]])

    # Picture
    logo = Image(os.path.join(os.path.dirname(template.fullname), 'xlwings.jpg'))

    # Float
    perf = 0.12

    # Markdown

    mytext = dedent("""\
                # Q1 2021 Results

                The perfomance was {{ perf }}.
                This was due to the following points:

                * More sales
                * Cost cuts

                # Sales were strong

                *Automation was the most important driver*.
                More info on request.
                """)

    style = MarkdownStyle()
    style.h1.font.color = (21, 164, 58)
    style.h1.font.size = 14

    app = template.app
    app.screen_updating = False

    wb = create_report(template_path, report_path,
        perf_data=perf_data,
        logo=logo,
        perf=perf,
        fig=fig,
        summary=Markdown(mytext, style)
    )

    wb.sheets.active['A1'].select()
    app.screen_updating = True

if __name__ == '__main__':
    # This part is to run the script directly from Python, not via Excel
    xw.Book('report_template.xlsx').set_mock_caller()
    main()

import os
import pandas as pd
from PIL import Image
from matplotlib.figure import Figure

import xlwings as xw
from xlwings_reports import create_report  # not part of the open-source xlwings package


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

    wb = create_report(template_path,
                       report_path,
                       perf=0.12 * 100,
                       perf_data=perf_data,
                       logo=Image.open(os.path.join(os.path.dirname(template.fullname), 'xlwings.jpg')),
                       fig=fig)


if __name__ == '__main__':
    # This part is to run the script directly from Python, not via Excel
    xw.Book(os.path.join(os.path.dirname(__file__), 'report_template.xlsx')).set_mock_caller()
    main()


import os
import shutil
import xlwings as xw
from pytest import approx, fixture


@fixture(scope="module")
def book():
    # Setup
    filepath = 'cash_flow_statement.xlsx'
    test_filepath = 'test_' + filepath
    shutil.copyfile(filepath, test_filepath)
    app = xw.App(visible=False)
    wb = app.books.open(test_filepath)
    yield wb

    # Teardown
    wb.save()
    app.quit()
    try:
        os.remove(test_filepath)
    except:
        pass


def test_cash_flow_formula_integrity(book):
    sheet = book.sheets[0]
    sheet['B2'].value = 100
    sheet['B3:M3'].value = 10
    sheet['B4:M4'].value = -5
    assert sheet['M5'].value == approx(160)


from pytest import fixture
import shutil
import xlwings as xw
import os

FILE_PATH = 'cash_flow_statement.xlsx'


@fixture(scope="module")
def book():
    # Setup
    test_filepath = 'test_' + FILE_PATH
    shutil.copyfile(FILE_PATH, test_filepath)
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
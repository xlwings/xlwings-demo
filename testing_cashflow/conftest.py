from pytest import fixture
import shutil
import xlwings as xw
import os


def pytest_addoption(parser):
    # This allows to pass in the file path via "pytest --book file.xlsx"
    parser.addoption("--book", action="store", help="Path of workbook")


@fixture(scope="module")
def book(pytestconfig):
    # Setup
    test_filepath = 'test_' + pytestconfig.getoption("book")
    shutil.copyfile(pytestconfig.getoption("book"), test_filepath)
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
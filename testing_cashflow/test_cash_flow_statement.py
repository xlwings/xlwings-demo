from pytest import approx
from utils import book


def test_cash_flow_formula_integrity(book):
    sheet = book.sheets[0]
    sheet['B2'].value = 100
    sheet['B3:M3'].value = 10
    sheet['B4:M4'].value = -5
    assert sheet['M5'].value == approx(160)

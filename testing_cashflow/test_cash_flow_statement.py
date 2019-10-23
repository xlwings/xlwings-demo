import os
import shutil
import unittest
import xlwings as xw


class TestMyBook(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        cls.test_copy = 'test_cash_flow_statement.xlsx'
        shutil.copyfile('cash_flow_statement.xlsx', cls.test_copy)
        cls.app = xw.App(visible=False)
        cls.wb = cls.app.books.open(cls.test_copy)

    @classmethod
    def tearDownClass(cls):
        cls.wb.save()
        cls.app.quit()
        try:
            os.remove(cls.test_copy)
        except:
            pass

    def test_cash_flow_formula_integrity(self):
        sheet = self.wb.sheets[0]
        sheet['B2'].value = 100
        sheet['B3:M3'].value = 10
        sheet['B4:M4'].value = -5
        self.assertAlmostEqual(sheet['M5'].value, 160)


if __name__ == '__main__':
    unittest.main()


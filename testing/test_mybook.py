import os
import unittest
import xlwings as xw


class TestMyBook(unittest.TestCase):

    @classmethod
    def setUpClass(cls):
        cls.app = xw.App(visible=False)
        cls.wb = cls.app.books.open('mybook.xlsm')

        # Map VBA functions
        cls.mysum = cls.wb.macro('Module1.MySum')
        cls.export_sheet_to_pdf = cls.wb.macro('Module2.ExportSheetToPDF')

    @classmethod
    def tearDownClass(cls):
        cls.wb.save()
        cls.app.quit()

    def test_vba_unittest(self):
        result = self.mysum(1, 2)
        self.assertAlmostEqual(3, result)

    def test_formula_unittest(self):
        for value in [-2, 0, 5]:
            self.wb.sheets[1].range('A1').value = value
            result = self.wb.sheets[1].range('B1').value
            self.assertAlmostEqual(value * 2, result)

    def test_xlwings_udf_unittest(self):
        sheet = self.wb.sheets[2]
        sheet.range('B1').value = 'xlwings'
        self.assertEqual(sheet.range('A1').value, 'hello xlwings')

    def test_vba_integrationtest(self):
        self.export_sheet_to_pdf()
        self.assertTrue(os.path.isfile('mybook.pdf'))
        os.remove('mybook.pdf')

    def test_cell_logic(self):
        values = self.wb.sheets[0].range('A1').expand().value
        for col in range(1, len(values[0])):
            sum_inputs = values[1][col] + values[2][col]
            total_row = values[3][col]
            self.assertAlmostEqual(total_row, sum_inputs)

    def test_vba_alternative_implementation(self):
        arg1, arg2 = 1, 2
        self.assertAlmostEqual(self.mysum(arg1, arg2), sum((arg1, arg2)))


if __name__ == '__main__':
    unittest.main()


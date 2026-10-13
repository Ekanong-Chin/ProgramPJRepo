"""Approved ROUND_HALF_UP policy, applied before monthly aggregation."""
from datetime import date
from decimal import Decimal
from pathlib import Path
import tempfile
import unittest

import excel_sales as sales
from test_excel_sales import fixture


class DailyRounding(unittest.TestCase):
    def test_positive_negative_ties_and_precision(self):
        for source, expected in [('1.005','1.01'),('-1.005','-1.01'),
                                 ('2.125','2.13'),('-2.125','-2.13'),
                                 ('0.004','0.00'),('0.005','0.01'),
                                 ('0.30000000000000004','0.30'),
                                 ('1E-1000','0.00'),('0','0.00')]:
            with self.subTest(source=source):
                self.assertEqual(sales.round_daily(Decimal(source)),Decimal(expected))

    def test_monthly_sum_is_sum_of_rounded_days(self):
        result = sales.analyze(fixture([(None,'HATO','Total',None,Decimal('.005'),Decimal('.005'))],days=('2026-10-01','2026-10-02')))
        self.assertFalse(result.critical,result.issues)
        self.assertEqual(result.months[2026,10,3],Decimal('.02'))
        self.assertEqual(result.records[0].values[date(2026,10,1)],Decimal('.01'))
        self.assertIn('WARN-DAILY-ROUNDED',[i.code for i in result.issues])
        cells = sales.workbook_cells(result)
        with tempfile.TemporaryDirectory(dir=Path.cwd()) as folder:
            output = Path(folder)/'rounded.xlsx'
            sales.save_workbook(output,cells)
            sales.validate_output(output,cells)

    def test_rounding_does_not_relax_output_limits(self):
        result = sales.analyze(fixture([(None,'HATO','Total',None,1234567890123456)]))
        self.assertIn('ERR-PRECISION',[i.code for i in result.issues])
        for value in ['Infinity','NaN','1E309']:
            with self.assertRaises(ValueError):
                sales.round_daily(Decimal(value))


if __name__ == '__main__':
    unittest.main()

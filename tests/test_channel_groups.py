"""Channel columns follow stable hierarchy groups across upsert runs."""
from datetime import date
from decimal import Decimal
from pathlib import Path
import tempfile
import unittest

import excel_sales as sales


def source_file(path, rows):
    cells = {(1,5):('2026-10-01',0),(2,5):('Sales',0)}
    for row,values in enumerate(rows,3):
        for col,value in enumerate(values,1):
            if value is not None:
                cells[row,col] = (value,0)
    cells[len(rows)+3,1] = ('Total',0)
    sales.save_workbook(path,cells)


class ChannelGroups(unittest.TestCase):
    def test_grouped_daily_monthly_and_persistent_upsert(self):
        with tempfile.TemporaryDirectory(dir=Path.cwd()) as folder:
            root = Path(folder)
            src,target = root/'sales.xlsx',root/'target'
            source_file(src,[(None,'OMS','OnDemand','Gofresh',1),
                             (None,'Shop','Online','Checkout',2),
                             (None,'OMS','Scheduled','Scheduled CC',3),
                             (None,'OMS','OnDemand','Hyper',4)])
            result,first = sales.transform(src,target)
            self.assertIsNotNone(first,result.issues)
            expected = ['Gofresh','Hyper','Scheduled CC','Checkout']
            self.assertEqual(sales.read_master(first)[0],expected)
            sheet = sales.read_source(first)
            self.assertEqual([sheet.cell(2,col).value for col in range(6,10)],expected)
            self.assertEqual([sheet.cell(6,col).value for col in range(6,10)],expected)
            original = first.read_bytes()
            source_file(src,[(None,'OMS','OnDemand','New Channel',5)])
            result,second = sales.transform(src,target)
            self.assertIsNotNone(second,result.issues)
            labels,days,values = sales.read_master(second)
            self.assertEqual(labels,['Gofresh','Hyper','New Channel','Scheduled CC','Checkout'])
            self.assertEqual(values['checkout',date(2026,10,1)],Decimal(2))
            self.assertEqual(first.read_bytes(),original)
            groups = sales.read_channel_groups(second)
            self.assertEqual(groups['hyper'],['','OMS','OnDemand'])
            self.assertEqual(groups['new channel'],['','OMS','OnDemand'])
            combined = sales.consolidate(sales.analyze(sales.read_source(src)),second)
            sales.validate_output(second,sales.workbook_cells(combined),combined.channel_groups)

    def test_stable_exact_groups_and_unknown_history(self):
        labels = ['Old unknown','A1','B1','A2','Case group']
        groups = {'a1':['','OMS','OnDemand'],'a2':['','OMS','OnDemand'],
                  'b1':['','Shop','Online'],'case group':['','oms','OnDemand']}
        self.assertEqual(sales.grouped_channel_order(labels,groups),
                         ['Old unknown','A1','A2','B1','Case group'])

    def test_metadata_is_validated(self):
        with tempfile.TemporaryDirectory(dir=Path.cwd()) as folder:
            path = Path(folder)/'master.xlsx'
            cells = {(1,1):('Example',0)}
            sales.save_workbook(path,cells,{'a':['','OMS']})
            with self.assertRaises(ValueError):
                sales.validate_output(path,cells,{'a':['','Shop']})


if __name__ == '__main__':
    unittest.main()

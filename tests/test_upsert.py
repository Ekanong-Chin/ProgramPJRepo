"""Synthetic (channel, calendar date) insert/replace acceptance tests."""
from datetime import date
from decimal import Decimal
from pathlib import Path
import concurrent.futures
import tempfile
import unittest

import excel_sales as sales


def transaction(path, days, channels):
    cells = {(1,c):(day,0) for c,day in enumerate(days,5)}
    for row,(name,amounts) in enumerate(channels,3):
        cells[row,2] = (name,0)
        cells[row,3] = ('Total',0)
        for col,value in enumerate(amounts,5):
            if value is not None:
                cells[row,col] = (Decimal(value),0)
    cells[len(channels)+3,1] = ('Total',0)
    sales.save_workbook(path,cells)


class Upsert(unittest.TestCase):
    def test_replace_insert_preserve_and_idempotency(self):
        with tempfile.TemporaryDirectory(dir=Path.cwd()) as folder:
            root = Path(folder)
            source, target = root/'source.xlsx', root/'target'
            transaction(source,['2026-10-01','2026-10-02'],[('HATO',['1','2']),('Old',['3','4'])])
            first,old = sales.transform(source,target)
            original = old.read_bytes()
            transaction(source,['2026-10-02','2026-11-01'],[('hato',['9.005','5']),('New',['7','8'])])
            source_bytes = source.read_bytes()
            second,new = sales.transform(source,target)
            self.assertIsNotNone(new,second.issues)
            self.assertEqual(second.upsert['updated_keys'],1)
            self.assertEqual(second.upsert['inserted_keys'],3)
            labels,days,values = sales.read_master(new)
            self.assertEqual(labels,['HATO','Old','New'])
            self.assertEqual(len(days),3)
            self.assertEqual(values['hato',date(2026,10,1)],Decimal('1'))
            self.assertEqual(values['hato',date(2026,10,2)],Decimal('9.01'))
            self.assertEqual(values['old',date(2026,10,2)],Decimal('4'))
            self.assertEqual(values['new',date(2026,11,1)],Decimal('8'))
            self.assertEqual(values['new',date(2026,10,1)],Decimal(0))
            self.assertEqual(old.read_bytes(),original)
            self.assertEqual(source.read_bytes(),source_bytes)
            third,repeat = sales.transform(source,target)
            self.assertEqual(third.upsert['updated_keys'],4)
            self.assertEqual(third.upsert['inserted_keys'],0)
            self.assertEqual(sales.read_master(repeat),sales.read_master(new))
            combined = sales.consolidate(sales.analyze(sales.read_source(source)),new)
            self.assertEqual(combined.months[2026,10,3],Decimal('10.01'))
            sales.validate_output(repeat,sales.workbook_cells(combined))

    def test_explicit_master_zero_replace_and_invalid_master(self):
        with tempfile.TemporaryDirectory(dir=Path.cwd()) as folder:
            root = Path(folder)
            source = root/'source.xlsx'
            transaction(source,['2026-10-01'],[('A',['10'])])
            _,master = sales.transform(source,root/'first')
            original = master.read_bytes()
            transaction(source,['2026-10-01'],[('A',[None])])
            result,output = sales.transform(source,root/'second',master)
            self.assertIsNotNone(output,result.issues)
            self.assertEqual(sales.read_master(output)[2]['a',date(2026,10,1)],Decimal(0))
            self.assertEqual(master.read_bytes(),original)
            prior = {p:p.read_bytes() for p in (root/'second').glob('*.xlsx')}
            result,output = sales.transform(source,root/'second',root/'missing.xlsx')
            self.assertIsNone(output)
            self.assertIn('ERR-MASTER-UPSERT',[i.code for i in result.issues])
            self.assertEqual(prior,{p:p.read_bytes() for p in (root/'second').glob('*.xlsx')})
            self.assertFalse((root/'second'/'.sales-upsert.lock').exists())

    def test_concurrent_updates_preserve_both_sources(self):
        with tempfile.TemporaryDirectory(dir=Path.cwd()) as folder:
            root = Path(folder)
            (root/'one').mkdir()
            (root/'two').mkdir()
            paths = [root/'one'/'shared.xlsx',root/'two'/'shared.xlsx']
            transaction(paths[0],['2026-10-01'],[('One',['1'])])
            transaction(paths[1],['2026-10-02'],[('Two',['2'])])
            with concurrent.futures.ThreadPoolExecutor(max_workers=2) as pool:
                outputs = list(pool.map(lambda p:sales.transform(p,root/'target'),paths))
            self.assertTrue(all(p for r,p in outputs),[r.issues for r,p in outputs])
            latest = sales.latest_master(root/'target')
            labels,days,values = sales.read_master(latest)
            self.assertEqual(set(labels),{'One','Two'})
            self.assertEqual(len(days),2)
            self.assertEqual(values['one',date(2026,10,1)],Decimal('1'))
            self.assertEqual(values['two',date(2026,10,2)],Decimal('2'))


if __name__ == '__main__':
    unittest.main()

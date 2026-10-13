"""Synthetic acceptance tests; never read transaction files."""
import concurrent.futures
from datetime import date
from decimal import Decimal
from pathlib import Path
import tempfile
import unittest
import zipfile
import xml.etree.ElementTree as ET

import excel_sales as sales


def fixture(rows, merges=(), days=('2026-10-01',)):
    cells = {(1, c): sales.Cell('text', d) for c, d in enumerate(days, 5)}
    cells.update({(2, c): sales.Cell('text', 'Sales') for c in range(5, 5 + len(days))})
    for row, data in enumerate(rows, 3):
        for col, value in enumerate(data, 1):
            if isinstance(value, sales.Cell):
                cells[row, col] = value
            elif value is not None:
                cells[row, col] = sales.Cell('text' if col <= 4 or isinstance(value, str) else 'n', str(value))
    cells[len(rows) + 3, 1] = sales.Cell('text', 'Total')
    return sales.Source(cells, list(merges))


def selected(result):
    return [r.label for r in result.records if r.status == 'final_selected']


def codes(result):
    return [(i.row, i.code) for i in result.issues if i.category == 'hierarchy']


class Hierarchy(unittest.TestCase):
    def test_basic_and_subtotals_before_after(self):
        for rows in [
            [(None,'OMS','Total',None,100),(None,None,'Scheduled','Total',20)],
            [(None,'OMS','Scheduled','Total',20),(None,None,'Total',None,100)],
        ]:
            result = sales.analyze(fixture(rows))
            self.assertFalse(result.critical, result.issues)
            self.assertEqual(selected(result), ['Scheduled'])
            self.assertEqual(sum(r.status == 'excluded_subtotal' for r in result.records), 1)
        result = sales.analyze(fixture([(None,'HATO','Total',None,10)]))
        self.assertEqual(selected(result), ['HATO'])

    def test_latest_matrix(self):
        matrix = [
            ('21', [(None,'OMS','OnDemand','Gofresh',10),(None,None,'Scheduled','Total',20)], ['Gofresh','Scheduled'], []),
            ('22', [(None,'OMS','OnDemand','Gofresh',10),(None,'Lazada','Total',None,20)], ['Gofresh','Lazada'], []),
            ('23', [(None,'OMS','Total',None,100),(None,'Lazada','Total',None,20),(None,'OMS','OnDemand','Super',30)], ['OMS','Lazada','Super'], []),
            ('24', [('Group','Total','Total',None,10)], [], [(3,'ERR-HIER-MALFORMED')]),
            ('25', [('Group','Total','Total',None,10),(None,'Lazada','Total',None,20)], ['Lazada'], [(3,'ERR-HIER-MALFORMED')]),
            ('26', [('Group','Total','Total',None,10),(None,None,'Total',None,20)], [], [(3,'ERR-HIER-MALFORMED'),(4,'ERR-HIER-CONTEXT-BLOCKED')]),
            ('27/35', [(None,'OMS','OnDemand','Gofresh',10),(None,None,None,None,None),(None,None,'Scheduled','Total',20)], ['Gofresh'], [(5,'ERR-HIER-MISSING-PARENT')]),
            ('28', [(None,'OMS',None,None,0),(None,None,'OnDemand','Gofresh',10)], ['Gofresh'], [(3,'ERR-HIER-UNASSIGNED-SALES')]),
            ('31', [(None,'OMS','Total',None,100),(None,None,'Scheduled','Total',20)], ['Scheduled'], []),
            ('32', [(None,'OMS','OnDemand','Gofresh',10),(None,None,None,None,None),(None,'OMS','OnDemand','Hyper',20)], ['Gofresh','Hyper'], []),
            ('34', [(None,'OMS','Total',None,100),(None,None,'Total','Total',20)], [], [(4,'ERR-HIER-MALFORMED')]),
            ('36/46', [('Group','Total','Channel','Total',20),(None,None,'Scheduled','Total',30),(None,'OMS','Total',None,40)], ['OMS'], [(3,'ERR-HIER-MISSING-PARENT'),(4,'ERR-HIER-MISSING-PARENT')]),
            ('37', [(None,None,'Scheduled','Total',20),(None,None,'OnDemand','Hyper',10),(None,'OMS','OnDemand','Gofresh',30)], ['Gofresh'], [(3,'ERR-HIER-MISSING-PARENT'),(4,'ERR-HIER-MISSING-PARENT')]),
            ('38', [(None,'OMS','OnDemand','Gofresh',10),('Group','Total','Total',None,20),(None,'OMS','OnDemand','Hyper',30)], ['Gofresh','Hyper'], [(4,'ERR-HIER-MALFORMED')]),
            ('39/45', [(None,'OMS','Total',None,100),(None,None,'Scheduled','Total',20),(None,None,'Total','Total',30)], ['Scheduled'], [(5,'ERR-HIER-MALFORMED')]),
            ('40', [(None,'OMS','Total',None,100),(None,None,'Scheduled','Total',20),(None,'Lazada','Total',None,30),(None,None,'Total','Total',40)], ['Scheduled'], [(6,'ERR-HIER-MALFORMED')]),
            ('41', [(None,'OMS','Total',None,100),(None,'Lazada','Total','Total',20)], ['OMS'], [(4,'ERR-HIER-MALFORMED')]),
        ]
        for name, rows, names, errors in matrix:
            with self.subTest(fixture=name):
                result = sales.analyze(fixture(rows))
                self.assertEqual(selected(result), names)
                self.assertEqual(codes(result), errors)

    def test_instances_and_reset(self):
        result = sales.analyze(fixture([(None,'OMS','OnDemand','Gofresh',10),(None,None,'Scheduled','Total',20)]))
        a, b = result.records
        self.assertEqual(a.instances[1], b.instances[1])
        self.assertNotEqual(a.instances[2], b.instances[2])
        result = sales.analyze(fixture([('Group','Total','Channel','Total',20),(None,'OMS','Total',None,40)]))
        self.assertEqual(result.traces[3]['after'][1], ('VERIFIED','ABSENT','ABSENT'))

    def test_merges(self):
        cases = [
            ('43', [(None,'OMS',None,'Gofresh',10),(None,None,'OnDemand','Hyper',20),(None,'Lazada','Total',None,30)], [(3,2,3,3)], ['Lazada'], [(3,'ERR-HIER-HORIZONTAL-MERGE'),(4,'ERR-HIER-CONTEXT-BLOCKED')]),
            ('44', [(None,'OMS','OnDemand','Gofresh',10),(None,'Lazada',None,'Hyper',20)], [(3,3,4,3)], ['Gofresh'], [(4,'ERR-HIER-MERGE-CONFLICT')]),
            ('48', [(None,'OMS','OnDemand','Gofresh',10),(None,'Lazada',None,'Hyper',20),(None,None,'Scheduled','Total',30)], [(3,3,4,3)], ['Gofresh'], [(4,'ERR-HIER-MERGE-CONFLICT'),(5,'ERR-HIER-CONTEXT-BLOCKED')]),
            ('49', [(None,'OMS','OnDemand','Gofresh',10),(None,'Lazada',None,'Hyper',20),(None,'Shopee','Total',None,30)], [(3,3,4,3)], ['Gofresh','Shopee'], [(4,'ERR-HIER-MERGE-CONFLICT')]),
            ('47', [(None,'OMS','OnDemand','Gofresh',10),(None,None,None,None,None),(None,'Lazada','Total',None,30)], [(3,2,4,2),(3,3,4,3)], ['Gofresh','Lazada'], []),
            ('42', [(None,'OMS','OnDemand','Gofresh',10),(None,None,'OnDemand','Hyper',20)], [(3,2,4,2)], ['Gofresh','Hyper'], []),
        ]
        for name, rows, merges, names, errors in cases:
            with self.subTest(fixture=name):
                result = sales.analyze(fixture(rows, merges))
                self.assertEqual(selected(result), names)
                self.assertEqual(codes(result), errors)

    def test_duplicate_and_invalid_child(self):
        result = sales.analyze(fixture([(None,'OMS','Gofresh','Total',10),(None,'Lazada','Total',None,20),(None,'OMS','gofresh','Total',30)]))
        self.assertEqual([i.row for i in result.issues if i.code == 'ERR-CHANNEL-DUPLICATE'], [3,5])
        result = sales.analyze(fixture([(None,'OMS','Total',None,100),(None,None,'Scheduled','Total','INVALID')]))
        self.assertEqual(result.records[0].status, 'excluded_subtotal')
        self.assertIn('ERR-SALES-INVALID-NUMERIC', [i.code for i in result.issues])


class Parsing(unittest.TestCase):
    def test_dates_summaries_and_duplicates(self):
        for label in ['Total Sales','Grand Total','Total','ยอดขายรวม',' TOTAL SALES ']:
            src = fixture([(None,'HATO','Total',None,1)])
            src.cells[1,6] = sales.Cell('text',label)
            src.cells[3,6] = sales.Cell('text','IGNORED')
            self.assertFalse(sales.analyze(src).critical)
        for label in ['Total  Sales','Total Sales Adjusted','Sales Amount']:
            src = fixture([(None,'HATO','Total',None,1)])
            src.cells[1,6] = sales.Cell('text',label)
            self.assertIn('ERR-DATE-INVALID',[i.code for i in sales.analyze(src).issues])
        self.assertEqual(sales.parse_date(sales.Cell('text','13/10/2026')),date(2026,10,13))
        src = fixture([(None,'HATO','Total',None,1,2)], days=('13/10/2026','2026-10-13'))
        self.assertIn('ERR-DATE-DUPLICATE',[i.code for i in sales.analyze(src).issues])
        self.assertEqual(sales.parse_date(sales.Cell('date_serial','1.5'),True),date(1904,1,2))
        with self.assertRaises(ValueError):
            sales.parse_date(sales.Cell('date_serial','60'))

    def test_formulas_and_precision(self):
        for cell, okay in [(sales.Cell('n','0.125',True),True),(sales.Cell('n','',True),False),(sales.Cell('text','123',True),False),(sales.Cell('e','#VALUE!',True),False),(sales.Cell('b','1'),False)]:
            result = sales.analyze(fixture([(None,'HATO','Total',None,cell)]))
            self.assertEqual(not result.critical,okay)
        self.assertEqual(sales.exact_sum([Decimal('.1'),Decimal('.2')]),Decimal('.3'))
        self.assertEqual(sales.exact_sum([Decimal('1E308'),Decimal('-1E308'),Decimal('.004')]),Decimal('.004'))
        for value, okay in [('123456789012345',True),('1234567890123456',False),('1.2300',True),('1E-309',False),('Infinity',False)]:
            self.assertEqual(sales.precision_ok(Decimal(value)),okay)

    def test_boundaries_and_footer(self):
        src = fixture([(None,'HATO','Total',None,1)])
        src.cells[10,1] = sales.Cell('text','Applied filters:')
        src.cells[10,100] = sales.Cell('text','footer')
        self.assertFalse(sales.analyze(src).critical)
        del src.cells[4,1]
        self.assertTrue(sales.analyze(src).critical)


class Workbooks(unittest.TestCase):
    def test_end_to_end_collision_and_failure(self):
        with tempfile.TemporaryDirectory(dir=Path.cwd()) as folder:
            root = Path(folder)
            source = root/'TransactionDetail.xlsx'
            sales.save_workbook(source, {(1,5):('2026-10-01',0),(1,6):('2026-10-05',0),(1,7):('2026-11-01',0),(2,5):('Sales',0),(3,2):('HATO',0),(3,3):('Total',0),(3,5):(Decimal('.004'),0),(3,6):(Decimal('.004'),0),(3,7):(Decimal('-1.234'),0),(4,1):('Total',0)})
            original = source.read_bytes()
            target = root/'outputs'
            with concurrent.futures.ThreadPoolExecutor(max_workers=2) as pool:
                runs = list(pool.map(lambda _: sales.transform(source,target),range(2)))
            self.assertTrue(all(output for result, output in runs), [r.issues for r,o in runs])
            self.assertNotEqual(runs[0][1],runs[1][1])
            self.assertEqual(source.read_bytes(),original)
            for result, output in runs:
                self.assertEqual(result.months[2026,10,3],Decimal('0.00'))
                sales.validate_output(output,sales.workbook_cells(result))
            prior = {p:p.read_bytes() for p in target.iterdir()}
            bad = root/'bad.xlsx'
            sales.save_workbook(bad,{(1,5):('2026-10-01',0),(3,2):('HATO',0),(3,3):('Total',0),(3,5):('INVALID',0),(4,1):('Total',0)})
            bad_original = bad.read_bytes()
            result, output = sales.transform(bad,target)
            self.assertIsNone(output)
            self.assertTrue(result.critical)
            self.assertEqual(bad.read_bytes(),bad_original)
            self.assertEqual(prior,{p:p.read_bytes() for p in target.iterdir()})

    def test_serialization_tamper(self):
        with tempfile.TemporaryDirectory(dir=Path.cwd()) as folder:
            path = Path(folder)/'test.xlsx'
            expected = {(1,1):(Decimal('.3'),2)}
            sales.save_workbook(path,{(1,1):(Decimal('.300000000000001'),2)})
            with self.assertRaises(ValueError):
                sales.validate_output(path,expected)


if __name__ == '__main__':
    unittest.main()

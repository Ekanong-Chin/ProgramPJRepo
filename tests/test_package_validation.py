"""Exercise real XLSX cache tokens, failure safety, and numeric boundaries."""
from datetime import date
from decimal import Decimal
from pathlib import Path
import tempfile
import unittest
import zipfile
import xml.etree.ElementTree as ET

import excel_sales as sales
from test_excel_sales import fixture


class PackageValidation(unittest.TestCase):
    def test_header_merge_scope(self):
        for merge in [(1,1,1,4),(1,1,2,4)]:
            with self.subTest(allowed=merge):
                result = sales.analyze(fixture([(None,'HATO','Total',None,1)], [merge]))
                self.assertFalse(result.critical, result.issues)
        for merge in [(1,1,3,1),(1,4,1,5),(1,5,1,6),(2,1,3,1)]:
            with self.subTest(rejected=merge):
                result = sales.analyze(fixture([(None,'HATO','Total',None,1)], [merge]))
                self.assertIn('ERR-HIER-MERGE-CONFLICT',[i.code for i in result.issues])

    def test_independent_source_reconciliation(self):
        result = sales.analyze(fixture([(None,'HATO','Total',None,Decimal('.1'))]))
        result.records[0].values[date(2026,10,1)] = Decimal('.2')
        with self.assertRaises(ValueError):
            sales.workbook_cells(result)

    def test_hier_50_defensive_invariant(self):
        self.assertEqual(sales.candidate_error(['one','two']),'ERR-HIER-MULTI-CANDIDATE')
        self.assertIsNone(sales.candidate_error(['one']))

    def test_formula_caches(self):
        with tempfile.TemporaryDirectory(dir=Path.cwd()) as folder:
            root = Path(folder)
            base = root/'base.xlsx'
            sales.save_workbook(base,{(1,5):('2026-10-01',0),(3,2):('HATO',0),(3,3):('Total',0),(3,5):(Decimal('.125'),0),(4,1):('Total',0)})
            with zipfile.ZipFile(base) as z:
                package = {name:z.read(name) for name in z.namelist()}
            for cache,kind,ftype,okay in [('.125','n','normal',True),('0','n','shared',True),('1.25','n','array',True),('','n','normal',False),('123','str','normal',False),('#VALUE!','e','normal',False),('1','b','normal',False)]:
                with self.subTest(kind=kind,formula=ftype,empty=not cache):
                    sheet = ET.fromstring(package['xl/worksheets/sheet1.xml'])
                    node = next(c for c in sheet.iter(sales.NS+'c') if c.get('r') == 'E3')
                    node.set('t',kind)
                    ET.SubElement(node,sales.NS+'f',{'t':ftype}).text = 'SYNTHETIC_FORMULA'
                    node.find(sales.NS+'v').text = cache
                    path = root/'formula.xlsx'
                    with zipfile.ZipFile(path,'w') as z:
                        for name,data in package.items():
                            z.writestr(name,ET.tostring(sheet) if name == 'xl/worksheets/sheet1.xml' else data)
                    before = path.read_bytes()
                    result,output = sales.transform(path,root/'out')
                    self.assertEqual(output is not None,okay,result.issues)
                    self.assertEqual(path.read_bytes(),before)

    def test_multiple_errors_and_month_precision(self):
        src = fixture([(None,'Lazada','Total',None,'INVALID',1),(None,'lazada','Total',None,1,1)],days=('2026-10-01','2026-10-01'))
        result = sales.analyze(src)
        self.assertTrue({'ERR-DATE-DUPLICATE','ERR-CHANNEL-DUPLICATE','ERR-SALES-INVALID-NUMERIC'} <= {i.code for i in result.issues})
        src = fixture([(None,'HATO','Total',None,999999999999999,2)],days=('2026-10-01','2026-10-02'))
        self.assertIn('ERR-PRECISION-MONTH',[i.code for i in sales.analyze(src).issues])
        src = fixture([(None,'HATO','Total',None,None,-1)],days=('2026-10-01','2027-10-01'))
        result = sales.analyze(src)
        self.assertEqual(result.months[2026,10,3],Decimal(0))
        self.assertEqual(result.months[2027,10,3],Decimal(-1))

    def test_missing_parent_control_and_context(self):
        src = fixture([(None,'HATO',None,None,None),(None,None,' total ',None,1)])
        self.assertEqual([r.label for r in sales.analyze(src).records if r.status == 'final_selected'],['HATO'])
        src = fixture([(None,'OMS','HI','Child',1),(None,None,None,None,None)])
        result = sales.analyze(src)
        self.assertEqual(result.traces[4]['after'][1],('ABSENT','ABSENT','ABSENT'))
        result = sales.analyze(fixture([('Total','OMS','HI','Child',1)]))
        self.assertIn('ERR-HIER-MALFORMED',[i.code for i in result.issues])
        result = sales.analyze(fixture([(None,'OMS','HI','Child',1),(None,None,None,None,2)],[(3,4,4,4)]))
        self.assertIn('ERR-HIER-MERGED-D',[i.code for i in result.issues])

    def test_datetime_and_blank_header(self):
        src = fixture([(None,'HATO','Total',None,1,2)])
        src.cells[1,5] = sales.Cell('d','2026-10-01T09:00:00')
        src.cells[1,6] = sales.Cell('d','2026-10-01T15:00:00')
        self.assertIn('ERR-DATE-DUPLICATE',[i.code for i in sales.analyze(src).issues])
        src.cells[1,6] = sales.Cell('text',' \t\n')
        self.assertIn('ERR-DATE-BLANK-HEADER',[i.code for i in sales.analyze(src).issues])

    def test_absent_invalid_unwritable_and_styles(self):
        with tempfile.TemporaryDirectory(dir=Path.cwd()) as folder:
            root = Path(folder)
            for path in [root/'absent.xlsx',root/'invalid.xlsx']:
                if path.name == 'invalid.xlsx':
                    path.write_bytes(b'not a zip')
                result,output = sales.transform(path,root/'out')
                self.assertIsNone(output)
                self.assertTrue(result.critical)
            src = root/'source.xlsx'
            sales.save_workbook(src,{(1,5):('2026-10-01',0),(3,2):('HATO',0),(3,3):('Total',0),(3,5):(1,0),(4,1):('Total',0)})
            blocked = root/'blocked'
            blocked.write_bytes(b'keep')
            result,output = sales.transform(src,blocked)
            self.assertIsNone(output)
            self.assertEqual(blocked.read_bytes(),b'keep')
            result,output = sales.transform(src,root/'out')
            with zipfile.ZipFile(output) as z:
                styles = z.read('xl/styles.xml').decode()
                for token in ['FF17365D','FFFFFFFF','#,##0.00','DD/MM/YYYY']:
                    self.assertIn(token,styles)
                sheet = ET.fromstring(z.read('xl/worksheets/sheet1.xml'))
                refs = {c.get('r') for c in sheet.iter(sales.NS+'c')}
                self.assertNotIn('D7',refs)
                self.assertNotIn('E7',refs)


if __name__ == '__main__':
    unittest.main()

"""Per-source workbook families and directory batch execution."""
from datetime import date, datetime
from decimal import Decimal
from pathlib import Path
import tempfile
import subprocess
import sys
import json
import unittest
from unittest.mock import patch

import excel_sales as sales
from test_upsert import transaction


class Batch(unittest.TestCase):
    def test_cli_default_directory(self):
        with tempfile.TemporaryDirectory(dir=Path.cwd()) as folder:
            root = Path(folder)
            (root/'source').mkdir()
            transaction(root/'source'/'aaa.xlsx',['2026-10-01'],[('One',['1'])])
            proc = subprocess.run([sys.executable,str(Path(sales.__file__).resolve())],cwd=root,capture_output=True,text=True)
            self.assertEqual(proc.returncode,0,proc.stderr)
            report = json.loads(proc.stdout)
            self.assertEqual(report['files'],1)
            self.assertEqual(report['succeeded'],1)
            self.assertTrue(Path(report['results'][0]['output']).name.startswith('aaa_'))

    def test_multiple_sources_isolated_replacements(self):
        with tempfile.TemporaryDirectory(dir=Path.cwd()) as folder:
            root = Path(folder)
            src,target = root/'source',root/'target'
            src.mkdir()
            transaction(src/'aaa.xlsx',['2026-10-01'],[('Shared',['1'])])
            transaction(src/'bbb.XLSX',['2026-10-01'],[('Shared',['2'])])
            (src/'~$locked.xlsx').write_bytes(b'ignore')
            (src/'readme.txt').write_text('ignore')
            (src/'nested').mkdir()
            transaction(src/'nested'/'ignored.xlsx',['2026-10-01'],[('Ignored',['3'])])
            first = sales.transform_directory(src,target)
            self.assertEqual(len(first),2)
            self.assertTrue(all(r['status']=='success' for r in first))
            old = {Path(r['output']):Path(r['output']).read_bytes() for r in first}
            self.assertTrue(all(Path(r['output']).name.startswith(Path(r['source']).stem+'_') for r in first))
            transaction(src/'aaa.xlsx',['2026-10-01','2026-10-02'],[('Shared',['9','10'])])
            second = sales.transform_directory(src,target)
            byname = {Path(r['source']).stem:r for r in second}
            self.assertEqual(byname['aaa']['upsert']['updated_keys'],1)
            self.assertEqual(byname['aaa']['upsert']['inserted_keys'],1)
            self.assertEqual(byname['bbb']['upsert']['updated_keys'],1)
            a = sales.read_master(Path(byname['aaa']['output']))
            b = sales.read_master(Path(byname['bbb']['output']))
            self.assertEqual(a[2]['shared',date(2026,10,1)],Decimal('9'))
            self.assertEqual(b[2]['shared',date(2026,10,1)],Decimal('2'))
            self.assertEqual(len(a[1]),2)
            self.assertEqual(len(b[1]),1)
            self.assertTrue(all(p.read_bytes()==value for p,value in old.items()))

    def test_failure_does_not_stop_other_sources(self):
        with tempfile.TemporaryDirectory(dir=Path.cwd()) as folder:
            root = Path(folder)
            src = root/'source'
            src.mkdir()
            (src/'aaa.xlsx').write_bytes(b'invalid')
            transaction(src/'bbb.xlsx',['2026-10-01'],[('One',['1'])])
            reports = sales.transform_directory(src,root/'target')
            self.assertEqual([r['status'] for r in reports],['failed','success'])
            self.assertIsNone(reports[0]['output'])
            self.assertIsNotNone(reports[1]['output'])

    def test_filename_matching_and_numeric_collision_suffix(self):
        with tempfile.TemporaryDirectory(dir=Path.cwd()) as folder:
            root = Path(folder)
            names = ['aaa_20261014_100000_999.xlsx','aaa_20261014_100000_1000.xlsx',
                     'aaa_more_20991231_235959.xlsx','MasterDetail_20991231_235959.xlsx',
                     'aaa_20261013_100000.xlsx']
            for name in names:
                (root/name).write_bytes(b'only lookup')
            self.assertEqual(sales.latest_master(root,'aaa').name,names[1])
            self.assertIsNone(sales.latest_master(root,'a.a'))
            source = root/'a.a.xlsx'
            transaction(source,['2026-10-01'],[('One',['1'])])
            with patch('excel_sales.datetime',wraps=datetime) as clock:
                clock.now.return_value = datetime(2026,10,14,12,0,0)
                _,first = sales.transform(source,root/'out')
                _,second = sales.transform(source,root/'out')
            self.assertEqual(first.name,'a.a_20261014_120000.xlsx')
            self.assertEqual(second.name,'a.a_20261014_120000_001.xlsx')

    def test_empty_missing_and_same_directory(self):
        with tempfile.TemporaryDirectory(dir=Path.cwd()) as folder:
            root = Path(folder)
            for source,target in [(root,root/'out'),(root/'missing',root/'out'),(root,root)]:
                with self.subTest(source=source,target=target):
                    with self.assertRaises(ValueError):
                        sales.transform_directory(source,target)


if __name__ == '__main__':
    unittest.main()

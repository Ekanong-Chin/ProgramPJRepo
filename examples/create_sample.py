"""Generate entirely synthetic XLSX input/output without reading user data."""
from decimal import Decimal
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))
from excel_sales import save_workbook, transform


def main():
    folder = ROOT / 'target' / 'synthetic'
    folder.mkdir(parents=True, exist_ok=True)
    source = folder / 'TransactionDetail.synthetic.xlsx'
    if source.exists():
        raise SystemExit('Synthetic source already exists; it will not be overwritten.')
    cells = {
        (1, 5): ('2026-10-01', 0), (1, 6): ('2026-10-02', 0),
        (1, 7): ('2026-11-01', 0), (1, 8): ('Total Sales', 0),
        (2, 5): ('Sales', 0), (2, 6): ('Sales', 0), (2, 7): ('Sales', 0),
        (3, 2): ('HATO', 0), (3, 3): ('Total', 0),
        (4, 2): ('OMS', 0), (4, 3): ('Total', 0),
        (5, 3): ('Scheduled', 0), (5, 4): ('Total', 0),
        (6, 3): ('OnDemand', 0), (6, 4): ('Gofresh', 0),
        (7, 1): ('Total', 0), (8, 1): ('Applied filters: synthetic example', 0),
    }
    for row, values in [(3, ['.1', '.2', '-.004']), (4, ['999', '999', '999']),
                        (5, ['.004', '.004', '2.125']), (6, ['1', None, '-1'])]:
        for col, value in enumerate(values, 5):
            if value is not None:
                cells[row, col] = (Decimal(value), 0)
    save_workbook(source, cells)
    result, output = transform(source, folder)
    if output is None:
        raise SystemExit('Synthetic transformation failed: ' + ', '.join(i.code for i in result.issues))
    print('Synthetic source:', source)
    print('Validated output:', output)
    print('3 channels; 3 actual dates; 2 months. Rounded October totals: 0.30, 0.00, 1.00.')


if __name__ == '__main__':
    main()

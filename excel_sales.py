"""Local, standard-library-only XLSX sales transformation."""
from __future__ import annotations

import argparse
import calendar
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta
from decimal import Decimal, InvalidOperation, ROUND_HALF_UP, localcontext
import json
import os
from pathlib import Path
import posixpath
import re
import tempfile
import time
import zipfile
import xml.etree.ElementTree as ET
from xml.sax.saxutils import escape

NS = '{http://schemas.openxmlformats.org/spreadsheetml/2006/main}'
REL = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
SUMMARIES = {'total sales', 'grand total', 'total', 'ยอดขายรวม'}
MINIMUM = Decimal('2.2250738585072E-308')
MAXIMUM = Decimal('1.79769313486231E+308')
GROUP_NS = 'urn:programpj:channel-groups:v1'
NUMERIC = re.compile(r'[+-]?(?:[0-9]+(?:\.[0-9]*)?|\.[0-9]+)(?:[eE][+-]?[0-9]+)?\Z')


def column(n):
    result = ''
    while n:
        n, remainder = divmod(n - 1, 26)
        result = chr(65 + remainder) + result
    return result


def coordinates(ref):
    match = re.fullmatch(r'([A-Z]+)([1-9][0-9]*)', ref)
    if not match:
        raise ValueError('Invalid cell reference')
    col = 0
    for char in match[1]:
        col = col * 26 + ord(char) - 64
    return int(match[2]), col


@dataclass
class Issue:
    code: str
    category: str
    explanation: str
    row: int | None = None
    col: int | None = None
    severity: str = 'critical'
    action: str = 'Correct the indicated source structure or value and retry.'


@dataclass
class Cell:
    kind: str = 'blank'
    value: str = ''
    formula: bool = False
    unsupported: bool = False

    def populated(self):
        return self.formula or self.value != ''


@dataclass
class Source:
    cells: dict
    merges: list
    sheet: str = 'Sheet1'
    epoch1904: bool = False

    def cell(self, row, col):
        return self.cells.get((row, col), Cell())


@dataclass
class Record:
    row: int
    level: int
    label: str
    path: tuple
    instances: tuple
    status: str = 'provisional'
    values: dict = field(default_factory=dict)


@dataclass
class Analysis:
    source: Source
    issues: list = field(default_factory=list)
    dates: dict = field(default_factory=dict)
    records: list = field(default_factory=list)
    traces: dict = field(default_factory=dict)
    months: dict = field(default_factory=dict)
    complete: bool = True
    upsert: dict = field(default_factory=dict)
    channel_groups: dict = field(default_factory=dict)

    def error(self, code, category, explanation, row=None, col=None):
        self.issues.append(Issue(code, category, explanation, row, col))

    @property
    def critical(self):
        return any(i.severity == 'critical' for i in self.issues)


def read_source(path):
    if path.suffix.lower() != '.xlsx' or path.stat().st_size > 50 * 1024**2:
        raise ValueError('Source must be XLSX and at most 50 MiB')
    with zipfile.ZipFile(path) as archive:
        members = archive.infolist()
        if len({m.filename for m in members}) != len(members):
            raise ValueError('Duplicate XLSX package members')
        if sum(m.file_size for m in members) > 250 * 1024**2:
            raise ValueError('Expanded XLSX exceeds 250 MiB')
        def xml(name):
            raw = archive.read(name)
            if b'<!DOCTYPE' in raw or b'<!ENTITY' in raw:
                raise ValueError('Unsupported XML declarations')
            return ET.fromstring(raw)
        workbook = xml('xl/workbook.xml')
        sheets = workbook.find(NS + 'sheets')
        if sheets is None or not len(sheets):
            raise ValueError('No first worksheet')
        first = sheets[0]
        relations = xml('xl/_rels/workbook.xml.rels')
        relation = next((x for x in relations if x.get('Id') == first.get(REL + 'id')), None)
        if relation is None or relation.get('TargetMode') == 'External' or not relation.get('Type', '').endswith('/worksheet'):
            raise ValueError('First sheet is not a local worksheet')
        target = relation.get('Target', '')
        target = posixpath.normpath(target.lstrip('/') if target.startswith('/') else 'xl/' + target)
        if not target.startswith('xl/'):
            raise ValueError('Invalid worksheet relationship')
        strings = []
        if 'xl/sharedStrings.xml' in archive.namelist():
            strings = [''.join(t.text or '' for t in si.iter(NS + 't')) for si in xml('xl/sharedStrings.xml')]
        date_styles = set()
        if 'xl/styles.xml' in archive.namelist():
            styles = xml('xl/styles.xml')
            custom = {int(x.get('numFmtId')): x.get('formatCode', '') for x in styles.findall(NS + 'numFmts/' + NS + 'numFmt')}
            xfs = styles.find(NS + 'cellXfs')
            for index, xf in enumerate(xfs if xfs is not None else []):
                num = int(xf.get('numFmtId', '0'))
                fmt = re.sub(r'"[^"]*"|\\.|\[[^\]]*\]', '', custom.get(num, '')).lower()
                if num in set(range(14, 23)) | set(range(27, 37)) | set(range(50, 59)) or ('d' in fmt and ('y' in fmt or 'm' in fmt)):
                    date_styles.add(index)
        root = xml(target)
        cells = {}
        for node in root.findall('.//' + NS + 'sheetData/' + NS + 'row/' + NS + 'c'):
            row, col = coordinates(node.get('r', ''))
            if row > 100000 or col > 16384 or (row, col) in cells:
                raise ValueError('Invalid or excessive source dimensions')
            kind = node.get('t', 'n')
            value = node.find(NS + 'v')
            token = value.text or '' if value is not None else ''
            formula = node.find(NS + 'f')
            unsupported = node.get('cm') is not None or node.get('vm') is not None or (formula is not None and formula.get('t', 'normal') not in {'normal', 'shared', 'array'})
            if kind == 's':
                token = strings[int(token)]
                kind = 'text'
            elif kind == 'inlineStr':
                token = ''.join(t.text or '' for t in node.iter(NS + 't'))
                kind = 'text'
            elif kind == 'str':
                kind = 'text'
            elif kind == 'n' and int(node.get('s', '0')) in date_styles:
                kind = 'date_serial'
            cells[row, col] = Cell(kind, token, formula is not None, unsupported)
        merges = []
        for merge in root.findall(NS + 'mergeCells/' + NS + 'mergeCell'):
            start, end = merge.get('ref').split(':')
            r1, c1 = coordinates(start)
            r2, c2 = coordinates(end)
            if r2 < r1 or c2 < c1 or r2 > 100000:
                raise ValueError('Invalid merge range')
            merges.append((r1, c1, r2, c2))
        if len(merges) > 10000:
            raise ValueError('Excessive merge ranges')
        props = workbook.find(NS + 'workbookPr')
        return Source(cells, merges, first.get('name', 'Sheet1'), props is not None and props.get('date1904') in {'1', 'true'})


def blank(value):
    return not value.strip()


def total(value):
    return value.strip().casefold() == 'total'


def parse_date(cell, epoch1904=False):
    if cell.formula or cell.unsupported:
        raise ValueError('Unsupported date header formula/metadata')
    if cell.kind == 'date_serial':
        number = Decimal(cell.value)
        if not number.is_finite() or number < 0:
            raise ValueError('Invalid Excel date serial')
        days = int(number)
        if not epoch1904 and days == 60:
            raise ValueError('Invalid Excel leap-day serial')
        base = date(1904, 1, 1) if epoch1904 else date(1899, 12, 31)
        return base + timedelta(days=days if epoch1904 or days < 60 else days - 1)
    if cell.kind == 'd':
        return datetime.fromisoformat(cell.value).date()
    if cell.kind == 'text':
        value = cell.value.strip()
        fmt = '%d/%m/%Y' if re.fullmatch(r'\d{2}/\d{2}/\d{4}', value) else '%Y-%m-%d' if re.fullmatch(r'\d{4}-\d{2}-\d{2}', value) else None
        if fmt:
            return datetime.strptime(value, fmt).date()
    raise ValueError('Unsupported date header')


def exact_sum(values):
    values = list(values)
    if not values:
        return Decimal(0)
    exponent = min(v.as_tuple().exponent for v in values)
    coefficient = 0
    for value in values:
        parts = value.as_tuple()
        digits = int(''.join(map(str, parts.digits)))
        coefficient += (-1 if parts.sign else 1) * digits * 10 ** (parts.exponent - exponent)
    return Decimal((int(coefficient < 0), tuple(map(int, str(abs(coefficient)))), exponent))


def precision_ok(value):
    if not value.is_finite():
        return False
    if not value:
        return True
    digits = ''.join(map(str, value.as_tuple().digits)).rstrip('0')
    return len(digits) <= 15 and MINIMUM <= value.copy_abs() <= MAXIMUM


def round_daily(value):
    """User-approved daily rounding; monthly totals sum these rounded values."""
    if not value.is_finite() or value.copy_abs() > MAXIMUM:
        raise ValueError('Source amount outside supported numeric range')
    with localcontext() as context:
        context.prec = max(32, len(value.as_tuple().digits), value.adjusted() + 4)
        return value.quantize(Decimal('0.01'), rounding=ROUND_HALF_UP)


def candidate_error(valid_candidates):
    return 'ERR-HIER-MULTI-CANDIDATE' if len(valid_candidates) > 1 else None


def analyze(source):
    result = Analysis(source)
    sentinel = []
    for row in sorted({r for r, c in source.cells if r >= 3 and c <= 4}):
        raw = [source.cell(row, c) for c in range(1, 5)]
        if not any(x.formula for x in raw) and total(raw[0].value) and all(blank(x.value) for x in raw[1:]):
            sentinel.append(row)
    if len(sentinel) != 1:
        result.error('ERR-SOURCE-SENTINEL', 'structure', 'Exactly one raw grand-total sentinel is required.')
        result.complete = False
        end = max((r for r, c in source.cells), default=3) + 1
    else:
        end = sentinel[0]
    boundary_valid = len(sentinel) == 1
    last = max((c for (r, c), cell in source.cells.items() if (r < end if boundary_valid else r == 1) and cell.populated()), default=4)
    populated_columns = {c for (r, c), cell in source.cells.items() if 2 <= r < end and cell.populated()}
    seen = {}
    for col in range(5, last + 1):
        header = source.cell(1, col)
        if header.kind == 'text' and not header.formula and header.value.strip().casefold() in SUMMARIES:
            continue
        if not header.populated() or blank(header.value):
            if col in populated_columns:
                result.error('ERR-DATE-BLANK-HEADER', 'date', 'Blank header over populated cells.', 1, col)
            continue
        try:
            day = parse_date(header, source.epoch1904)
        except (ValueError, InvalidOperation, OverflowError):
            result.error('ERR-DATE-INVALID', 'date', 'Unsupported or invalid date header.', 1, col)
            result.complete = False
            continue
        result.dates[col] = day
        if day in seen:
            result.error('ERR-DATE-DUPLICATE', 'date', 'Duplicate normalized source date.', 1, col)
        seen[day] = col
    if not result.dates:
        result.error('ERR-DATE-NONE', 'date', 'No valid source dates.')
        result.complete = False
    if len(result.dates) > 366:
        result.error('ERR-LIMIT-DATES', 'limit', 'More than 366 date columns.')
        return result
    if not boundary_valid:
        return result
    labels, states, ids = [''] * 3, ['ABSENT'] * 3, [None] * 3
    serial = 0
    evidence, uncertain = [], set()
    merge_map = {}
    for merge in source.merges:
        r1, c1, r2, c2 = merge
        # Rows 1-2 in A-D are outside channel classification and date headers.
        if r2 <= 2 and c2 <= 4:
            continue
        if r1 < 3 or r2 >= end or c2 >= 5:
            if r1 <= end and r2 >= 1:
                result.error('ERR-HIER-MERGE-CONFLICT', 'merge', 'Merge crosses a header, sentinel, or date/sales region.', r1, c1)
            continue
        for row in range(r1, r2 + 1):
            merge_map.setdefault(row, []).append(merge)
    damaged = set()
    def clear(start, state):
        for level in range(start, 3):
            labels[level], states[level], ids[level] = '', state, None
    for row in range(3, end):
        raw_cells = [source.cell(row, c) for c in range(1, 5)]
        raw = [x.value for x in raw_cells]
        sales = any(source.cell(row, c).populated() for c in result.dates)
        before = (tuple(labels), tuple(states), tuple(ids))
        merges = sorted(merge_map.get(row, []), key=lambda m: (m[0], m[1]))
        merged = {}
        error, error_col, invalidate = None, None, None
        for merge in merges:
            r1, c1, r2, c2 = merge
            code = 'ERR-HIER-HORIZONTAL-MERGE' if c1 != c2 else 'ERR-HIER-MERGED-D' if c1 == 4 else None
            if merge in damaged and c1 == c2 == 4 and row > r1:
                error, error_col = 'ERR-HIER-CONTEXT-BLOCKED', 4
                break
            if code and merge not in damaged:
                error, error_col, invalidate = code, c1, 0 if c1 != c2 else None
                damaged.add(merge)
                break
            if merge not in damaged and c1 <= 3:
                merged[c1 - 1] = (source.cell(r1, c1).value, merge)
        if not error and all(blank(x) for x in raw) and not sales and not merged:
            clear(0, 'ABSENT')
            result.traces[row] = {'before': before, 'after': (tuple(labels), tuple(states), tuple(ids)), 'classification': 'separator'}
            continue
        if not error and any(x.formula or x.unsupported or (x.populated() and x.kind != 'text') for x in raw_cells):
            error, error_col, invalidate = 'ERR-HIER-MALFORMED', 1, 0
        if not error and total(raw[0]):
            error, error_col, invalidate = 'ERR-HIER-MALFORMED', 1, 0
        if not error:
            for level in range(3):
                value = raw[level]
                if blank(value) and level in merged:
                    value = merged[level][0]
                if blank(value) or total(value):
                    continue
                if value != labels[level] or states[level] != 'VERIFIED':
                    conflict = next((m for deep, (_, m) in merged.items() if deep > level and m[0] < row), None)
                    if conflict:
                        error, error_col, invalidate = 'ERR-HIER-MERGE-CONFLICT', level + 1, level
                        damaged.add(conflict)
                        clear(level, 'UNTRUSTED')
                        break
                    clear(level, 'ABSENT')
                    serial += 1
                    labels[level], states[level], ids[level] = value, 'VERIFIED', serial
        association = list(ids)
        for level in range(3):
            if total(raw[level]):
                association[level:] = [None] * (3 - level)
                break
        resolved = [raw[n] if not blank(raw[n]) else labels[n] if states[n] == 'VERIFIED' else '' for n in range(3)] + [raw[3]]
        candidates = []
        missing = set()
        if not error:
            for marker in range(1, 4):
                if total(raw[marker]):
                    predecessor = resolved[marker - 1]
                    if total(predecessor):
                        error, error_col, invalidate = 'ERR-HIER-MALFORMED', marker + 1, 0
                        break
                    if blank(predecessor):
                        missing.add(marker - 1)
                    else:
                        candidates.append((marker - 1, predecessor))
            if raw[3] and not blank(raw[3]) and not total(raw[3]):
                candidates.append((3, raw[3]))
            for level, name in candidates:
                for parent in ([1] if level == 2 else [1, 2] if level == 3 else []):
                    if states[parent] != 'VERIFIED' or total(resolved[parent]):
                        missing.add(parent)
            if not error and missing:
                blocked = any(states[n] == 'UNTRUSTED' and blank(raw[n]) for n in missing)
                error = 'ERR-HIER-CONTEXT-BLOCKED' if blocked else 'ERR-HIER-MISSING-PARENT'
                error_col = min(missing) + 1
            elif not error and candidate_error(candidates):
                error, error_col, invalidate = 'ERR-HIER-MULTI-CANDIDATE', 1, 0
            elif not error and not candidates and sales:
                error, error_col = 'ERR-HIER-UNASSIGNED-SALES', 4
        if error:
            result.error(error, 'hierarchy', 'Hierarchy row cannot be assigned safely.', row, error_col)
            if error != 'ERR-HIER-UNASSIGNED-SALES':
                uncertain.update(i for i in association if i is not None)
            if invalidate is not None:
                clear(invalidate, 'UNTRUSTED')
            elif error == 'ERR-HIER-MISSING-PARENT':
                clear(min(missing), 'ABSENT')
            result.complete = result.complete and error != 'ERR-HIER-CONTEXT-BLOCKED'
        else:
            deepest = 3 if not blank(raw[3]) and not total(raw[3]) else max((n for n in range(3) if states[n] == 'VERIFIED' and not total(resolved[n])), default=-1)
            evidence.append((tuple(ids), deepest))
            if candidates:
                level, name = candidates[0]
                result.records.append(Record(row, level, name, tuple(resolved[:level + 1]), tuple(ids)))
        if error not in {'ERR-HIER-MALFORMED', 'ERR-HIER-MULTI-CANDIDATE', 'ERR-HIER-HORIZONTAL-MERGE', 'ERR-HIER-MERGE-CONFLICT', 'ERR-HIER-MERGED-D', 'ERR-HIER-CONTEXT-BLOCKED'}:
            for n in range(1, 4):
                if total(raw[n]):
                    clear(n, 'ABSENT')
        result.traces[row] = {'before': before, 'after': (tuple(labels), tuple(states), tuple(ids)), 'error': error, 'candidates': candidates}
    for rec in result.records:
        if rec.level < 3 and any(eids[rec.level] == rec.instances[rec.level] and depth > rec.level for eids, depth in evidence):
            rec.status = 'excluded_subtotal'
        elif rec.level < 3 and rec.instances[rec.level] in uncertain:
            rec.status = 'blocked'
        else:
            rec.status = 'final_selected'
    selected = [r for r in result.records if r.status == 'final_selected']
    groups = {}
    for rec in selected:
        groups.setdefault(rec.label.casefold(), []).append(rec.row)
    for rows in groups.values():
        if len(rows) > 1:
            for row in rows:
                result.error('ERR-CHANNEL-DUPLICATE', 'channel', 'Duplicate selected name; affected rows: ' + ', '.join(map(str, rows)), row)
    if not selected:
        result.error('ERR-CHANNEL-NONE', 'channel', 'No selected channels.')
    if len(selected) > 1000 or len(selected) * len(result.dates) > 1000000:
        result.error('ERR-LIMIT-CHANNELS', 'limit', 'Channel or intersection limit exceeded.')
        return result
    for rec in selected:
        rounded_cells = 0
        for col, day in result.dates.items():
            cell = source.cell(rec.row, col)
            try:
                if cell.unsupported or cell.kind not in {'n', 'blank'} or (cell.formula and not cell.value) or (cell.value and not NUMERIC.fullmatch(cell.value)):
                    raise ValueError()
                value = Decimal(cell.value) if cell.value else Decimal(0)
                if not value.is_finite():
                    raise ValueError()
            except (ValueError, InvalidOperation):
                result.error('ERR-SALES-INVALID-NUMERIC', 'sales', 'Expected numeric literal or usable numeric formula cache.', rec.row, col)
                continue
            try:
                original = value
                value = round_daily(value)
            except (ValueError, InvalidOperation):
                result.error('ERR-PRECISION', 'precision', 'Source amount cannot be rounded within supported numeric range.', rec.row, col)
                continue
            rounded_cells += value != original
            if not precision_ok(value):
                result.error('ERR-PRECISION', 'precision', 'Value exceeds approved numeric precision/range.', rec.row, col)
                continue
            rec.values[day] = value
        if rounded_cells:
            result.issues.append(Issue('WARN-DAILY-ROUNDED', 'rounding', f'{rounded_cells} daily cells rounded to two decimal places using ROUND_HALF_UP.', rec.row, severity='warning', action='Review the approved daily rounding policy.'))
        if len(rec.values) != len(set(result.dates.values())):
            result.complete = False
        for month in sorted({(d.year, d.month) for d in rec.values}):
            amount = exact_sum(v for d, v in rec.values.items() if (d.year, d.month) == month)
            result.months[month[0], month[1], rec.row] = amount
            if not precision_ok(amount):
                result.error('ERR-PRECISION-MONTH', 'precision', 'Monthly total exceeds approved numeric precision/range.', rec.row)
    return result


def workbook_cells(result):
    selected = [r for r in result.records if r.status == 'final_selected']
    days = sorted(set(result.dates.values()))
    for rec in selected:
        if set(rec.values) != set(days):
            raise ValueError('Incomplete daily matrix')
        for col, day in result.dates.items():
            cell = result.source.cell(rec.row, col)
            expected = round_daily(Decimal(cell.value)) if cell.value else Decimal('0.00')
            if rec.values[day] != expected:
                raise ValueError('Source-to-daily reconciliation failed')
    months = sorted({(d.year, d.month) for d in days})
    cells = {}
    def put(row, col, value, style=0):
        cells[row, col] = (value, style)
    put(1, 1, 'MONTHLY TOTAL', 1)
    put(2, 1, 'Month', 1)
    put(2, 2, 'Days', 1)
    for col, rec in enumerate(selected, 6):
        put(2, col, rec.label, 1)
    for row, month in enumerate(months, 3):
        put(row, 1, date(*month, 1), 4)
        put(row, 2, sum((d.year, d.month) == month for d in days), 5)
        for col, rec in enumerate(selected, 6):
            value = exact_sum(rec.values[d] for d in days if (d.year, d.month) == month)
            if value != result.months[month[0], month[1], rec.row]:
                raise ValueError('Monthly reconciliation failed')
            put(row, col, value, 2)
    title = len(months) + 4
    put(title, 1, 'DAILY SALES', 1)
    for col, label in enumerate(['Month / Date', 'Day', 'Weekday', 'Tab WK', 'Tab Promotion'] + [r.label for r in selected], 1):
        put(title + 1, col, label, 1)
    for row, day in enumerate(days, title + 2):
        put(row, 1, day, 3)
        put(row, 2, day.day, 5)
        put(row, 3, calendar.day_name[day.weekday()], 5)
        for col, rec in enumerate(selected, 6):
            put(row, col, rec.values[day], 2)
    return cells


def numeric_token(value):
    if isinstance(value, date):
        return str((value - date(1899, 12, 30)).days - (1 if value < date(1900, 3, 1) else 0))
    return str(value)


def save_workbook(path, cells, channel_groups=None):
    rows = {}
    for (row, col), (value, style) in sorted(cells.items()):
        ref = column(col) + str(row)
        body = '<is><t xml:space="preserve">' + escape(value) + '</t></is>' if isinstance(value, str) else '<v>' + numeric_token(value) + '</v>'
        rows.setdefault(row, []).append(f'<c r="{ref}" s="{style}"' + (' t="inlineStr"' if isinstance(value, str) else '') + '>' + body + '</c>')
    main = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
    sheet = f'<worksheet xmlns="{main}"><cols><col min="1" max="1" width="18" customWidth="1"/><col min="2" max="5" width="16" customWidth="1"/><col min="6" max="16384" width="26" customWidth="1"/></cols><sheetData>' + ''.join(f'<row r="{row}">' + ''.join(body) + '</row>' for row, body in rows.items()) + '</sheetData></worksheet>'
    styles = f'<styleSheet xmlns="{main}"><numFmts count="3"><numFmt numFmtId="164" formatCode="#,##0.00"/><numFmt numFmtId="165" formatCode="DD/MM/YYYY"/><numFmt numFmtId="166" formatCode="YYYY-MM"/></numFmts><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF17365D"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="2"><border/><border><left style="thin"/><right style="thin"/><top style="thin"/><bottom style="thin"/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="6">' + ''.join(f'<xf numFmtId="{num}" fontId="{font}" fillId="{fill}" borderId="{border}" xfId="0" applyNumberFormat="1" applyAlignment="1"><alignment horizontal="{align}" vertical="center" wrapText="1"/></xf>' for num, font, fill, border, align in [(0,0,0,0,'left'),(0,1,2,1,'center'),(164,0,0,1,'right'),(165,0,0,1,'center'),(166,0,0,1,'center'),(0,0,0,1,'center')]) + '</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>'
    with zipfile.ZipFile(path, 'w', zipfile.ZIP_DEFLATED) as z:
        z.writestr('[Content_Types].xml', '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>')
        z.writestr('_rels/.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>')
        z.writestr('xl/workbook.xml', f'<workbook xmlns="{main}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets><sheet name="MasterDetail" sheetId="1" r:id="rId1"/></sheets></workbook>')
        z.writestr('xl/_rels/workbook.xml.rels', '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>')
        z.writestr('xl/styles.xml', styles)
        if channel_groups is not None:
            metadata = '<extLst><ext uri="{C026266F-7B44-4F61-AF0C-22D99C55CC52}"><groups xmlns="' + GROUP_NS + '">' + escape(json.dumps(channel_groups,ensure_ascii=False)) + '</groups></ext></extLst>'
            sheet = sheet.replace('</worksheet>',metadata+'</worksheet>')
        z.writestr('xl/worksheets/sheet1.xml', sheet)


def validate_output(path, expected, channel_groups=None):
    with zipfile.ZipFile(path) as z:
        workbook = ET.fromstring(z.read('xl/workbook.xml'))
        sheets = workbook.find(NS + 'sheets')
        if len(sheets) != 1 or sheets[0].get('name') != 'MasterDetail':
            raise ValueError('Invalid output worksheet structure')
        root = ET.fromstring(z.read('xl/worksheets/sheet1.xml'))
        if channel_groups is not None:
            metadata = root.find('.//{' + GROUP_NS + '}groups')
            if metadata is None or json.loads(metadata.text or '') != channel_groups:
                raise ValueError('Output channel grouping metadata mismatch')
        actual = {}
        for cell in root.findall('.//' + NS + 'c'):
            key = coordinates(cell.get('r'))
            if key in actual or key not in expected:
                raise ValueError('Unexpected/duplicate output cell')
            value, style = expected[key]
            if cell.get('s') != str(style):
                raise ValueError('Output style mismatch')
            if isinstance(value, str):
                if cell.get('t') != 'inlineStr' or ''.join(x.text or '' for x in cell.iter(NS + 't')) != value:
                    raise ValueError('Output label mismatch')
            else:
                if cell.get('t') not in {None, 'n'} or cell.find(NS + 'f') is not None or Decimal(cell.find(NS + 'v').text) != Decimal(numeric_token(value)):
                    raise ValueError('Strict stored numeric equality failed')
            actual[key] = True
        if actual.keys() != expected.keys():
            raise ValueError('Missing output cells')
        styles = ET.fromstring(z.read('xl/styles.xml'))
        formats = {x.get('numFmtId'): x.get('formatCode') for x in styles.findall(NS + 'numFmts/' + NS + 'numFmt')}
        if formats != {'164': '#,##0.00', '165': 'DD/MM/YYYY', '166': 'YYYY-MM'}:
            raise ValueError('Output number formats mismatch')
        if styles.find(NS + 'fonts')[1].find(NS + 'color').get('rgb') != 'FFFFFFFF' or styles.find(NS + 'fills')[2].find(NS + 'patternFill/' + NS + 'fgColor').get('rgb') != 'FF17365D':
            raise ValueError('Output header colors mismatch')


def read_master(path):
    """Read the Daily Sales section only; monthly totals are always rebuilt."""
    source = read_source(path)
    with zipfile.ZipFile(path) as archive:
        sheets = ET.fromstring(archive.read('xl/workbook.xml')).find(NS + 'sheets')
        if len(sheets) != 1 or source.sheet != 'MasterDetail':
            raise ValueError('Master must contain one MasterDetail worksheet')
    titles = [r for (r,c), cell in source.cells.items() if c == 1 and cell.value == 'DAILY SALES' and not cell.formula]
    if len(titles) != 1 or source.merges:
        raise ValueError('Invalid master Daily Sales structure')
    header = titles[0] + 1
    expected = ['Month / Date','Day','Weekday','Tab WK','Tab Promotion']
    if [source.cell(header,c).value for c in range(1,6)] != expected:
        raise ValueError('Invalid master headers')
    last_col = max(c for (r,c),cell in source.cells.items() if r == header and cell.populated())
    labels = [source.cell(header,c).value for c in range(6,last_col+1)]
    if not labels or len(labels)>1000 or any(blank(x) or total(x) for x in labels) or len({x.casefold() for x in labels}) != len(labels):
        raise ValueError('Invalid/duplicate master channels')
    if any(source.cell(header,c).formula or source.cell(header,c).kind != 'text' for c in range(1,last_col+1)):
        raise ValueError('Master headers must be literal text')
    if any(r>header and c>last_col and cell.populated() for (r,c),cell in source.cells.items()):
        raise ValueError('Master contains daily data without a channel header')
    values, days = {}, set()
    last_row = max(r for (r,c),cell in source.cells.items() if cell.populated())
    for row in range(header+1,last_row+1):
        if not any(source.cell(row,c).populated() for c in range(1,last_col+1)):
            continue
        day = parse_date(source.cell(row,1),source.epoch1904)
        if day in days:
            raise ValueError('Duplicate master dates')
        days.add(day)
        if len(days)>366 or len(days)*len(labels)>1000000:
            raise ValueError('Master dimensions exceed approved limits')
        for col,label in enumerate(labels,6):
            cell = source.cell(row,col)
            if cell.formula or cell.unsupported or cell.kind not in {'n','blank'} or (cell.value and not NUMERIC.fullmatch(cell.value)):
                raise ValueError('Invalid master sales value')
            value = round_daily(Decimal(cell.value) if cell.value else Decimal(0))
            if not precision_ok(value):
                raise ValueError('Unsupported master precision')
            values[label.casefold(),day] = value
    if not days:
        raise ValueError('Master contains no dates')
    return labels, days, values


def latest_master(target_dir, source_stem=None):
    matches = []
    prefix = re.escape(source_stem) if source_stem is not None else r'.+'
    for path in target_dir.iterdir():
        match = re.fullmatch(prefix + r'_(\d{8}_\d{6})(?:_(\d+))?\.xlsx',path.name,re.IGNORECASE)
        if path.is_file() and match:
            matches.append(((match[1],int(match[2] or 0)),path))
    return max(matches,key=lambda item:item[0])[1] if matches else None


def read_channel_groups(path):
    with zipfile.ZipFile(path) as archive:
        workbook = ET.fromstring(archive.read('xl/workbook.xml'))
        first = workbook.find(NS+'sheets')[0]
        relations = ET.fromstring(archive.read('xl/_rels/workbook.xml.rels'))
        relation = next(x for x in relations if x.get('Id')==first.get(REL+'id'))
        target = relation.get('Target','')
        target = posixpath.normpath(target.lstrip('/') if target.startswith('/') else 'xl/'+target)
        sheet = ET.fromstring(archive.read(target))
        metadata = sheet.find('.//{'+GROUP_NS+'}groups')
        if metadata is None:
            return {}
        groups = json.loads(metadata.text or '')
        if not isinstance(groups,dict) or any(not isinstance(k,str) or not isinstance(v,list) or len(v)>3 or any(not isinstance(x,str) for x in v) for k,v in groups.items()):
            raise ValueError('Invalid channel grouping metadata')
        return groups


def grouped_channel_order(labels, groups):
    """Stable hierarchy traversal: siblings retain their first-seen order."""
    tree = {}
    for label in labels:
        node = tree
        for level,parent in enumerate(groups.get(label.casefold(),[])):
            if not blank(parent) and not total(parent):
                node = node.setdefault(('group',level,parent),{})
        node[('channel',label.casefold())] = label
    ordered = []
    def visit(node):
        for key,value in node.items():
            if key[0]=='channel':
                ordered.append(value)
            else:
                visit(value)
    visit(tree)
    return ordered


def consolidate(result, master_path=None):
    # Validate incoming source mapping before constructing the combined matrix.
    workbook_cells(result)
    labels, days, values = read_master(master_path) if master_path else ([],set(),{})
    groups = read_channel_groups(master_path) if master_path else {}
    originals = {label.casefold():label for label in labels}
    inserted = updated = 0
    incoming = [r for r in result.records if r.status == 'final_selected']
    for rec in incoming:
        key = rec.label.casefold()
        groups[key] = list(rec.path[:-1])
        if key not in originals:
            originals[key] = rec.label
            labels.append(rec.label)
        for day,value in rec.values.items():
            if (key,day) in values:
                updated += 1
            else:
                inserted += 1
            values[key,day] = value
            days.add(day)
    if len(labels)>1000 or len(days)>366 or len(labels)*len(days)>1000000:
        raise ValueError('Combined workbook exceeds approved limits')
    labels = grouped_channel_order(labels,groups)
    merged = Analysis(Source({},[],'MasterDetail'))
    merged.channel_groups = {label.casefold():groups.get(label.casefold(),[]) for label in labels}
    merged.dates = {col:day for col,day in enumerate(sorted(days),5)}
    for row,label in enumerate(labels,3):
        rec = Record(row,3,label,(),(),status='final_selected')
        for col,day in merged.dates.items():
            value = values.get((label.casefold(),day),Decimal('0.00'))
            rec.values[day] = value
            merged.source.cells[row,col] = Cell('n',str(value))
        merged.records.append(rec)
        for month in {(d.year,d.month) for d in days}:
            amount = exact_sum(v for d,v in rec.values.items() if (d.year,d.month)==month)
            if not precision_ok(amount):
                raise ValueError('Combined monthly precision exceeded')
            merged.months[month[0],month[1],row] = amount
    result.upsert = {'master': str(master_path.resolve()) if master_path else None,
                     'updated_keys':updated,'inserted_keys':inserted,
                     'channels':len(labels),'dates':len(days),
                     'months':len({(d.year,d.month) for d in days})}
    return merged


def transform(source_path, target_dir, master_path=None):
    source_path, target_dir = Path(source_path), Path(target_dir)
    try:
        result = analyze(read_source(source_path))
    except (OSError, ValueError, KeyError, zipfile.BadZipFile, ET.ParseError, IndexError, InvalidOperation, OverflowError):
        result = Analysis(Source({}, []), complete=False)
        result.error('ERR-SOURCE-READ', 'source', 'Source is absent, unreadable, malformed, or exceeds limits.')
    output = None
    if not result.critical:
        temp = None
        lock = None
        try:
            target_dir.mkdir(parents=True, exist_ok=True)
            lock_path = target_dir / '.sales-upsert.lock'
            deadline = time.monotonic() + 30
            while True:
                try:
                    descriptor = os.open(lock_path, os.O_CREAT | os.O_EXCL | os.O_WRONLY)
                    os.close(descriptor)
                    lock = lock_path
                    break
                except FileExistsError:
                    if time.monotonic() >= deadline:
                        raise ValueError('Target is locked by another transformation')
                    time.sleep(0.05)
            base = Path(master_path) if master_path is not None else latest_master(target_dir,source_path.stem)
            try:
                combined = consolidate(result,base)
            except (OSError, ValueError, KeyError, InvalidOperation, ET.ParseError, zipfile.BadZipFile, IndexError):
                result.error('ERR-MASTER-UPSERT','master','Existing master or combined data failed validation; no output published.')
                return result, None
            with tempfile.NamedTemporaryFile(prefix='.sales-', suffix='.xlsx', dir=target_dir, delete=False) as handle:
                temp = Path(handle.name)
            cells = workbook_cells(combined)
            save_workbook(temp, cells, combined.channel_groups)
            validate_output(temp, cells, combined.channel_groups)
            stem = source_path.stem + '_' + datetime.now().strftime('%Y%m%d_%H%M%S')
            suffix = 0
            while True:
                candidate = target_dir / (stem + (f'_{suffix:03d}' if suffix else '') + '.xlsx')
                try:
                    os.link(temp, candidate)
                    output = candidate.resolve()
                    break
                except FileExistsError:
                    suffix += 1
        except (OSError, ValueError, KeyError, InvalidOperation, ET.ParseError):
            result.error('ERR-OUTPUT', 'output', 'Output save, validation, or race-safe publication failed.')
        finally:
            if temp is not None:
                try:
                    temp.unlink(missing_ok=True)
                except OSError:
                    result.issues.append(Issue('WARN-TEMP-CLEANUP', 'output', 'Current-run temporary file could not be removed.', severity='warning'))
            if lock is not None:
                try:
                    lock.unlink(missing_ok=True)
                except OSError:
                    result.issues.append(Issue('WARN-LOCK-CLEANUP','output','Current-run lock could not be removed.',severity='warning',action='Remove the stale lock only after confirming no run is active.'))
    return result, output


def status_report(result, output):
    selected = [r for r in result.records if r.status == 'final_selected']
    return {
        'status': 'success' if output else 'failed',
        'output': str(output) if output else None,
        'channels': result.upsert.get('channels',len(selected)),
        'dates': result.upsert.get('dates',len(set(result.dates.values()))),
        'months': result.upsert.get('months',len({(d.year,d.month) for d in result.dates.values()})),
        'upsert': result.upsert,
        'validation': 'complete' if result.complete else 'blocked',
        'formula_cache_policy': 'Last saved numeric cache only; freshness not verified.',
        'blocked_parent_count': sum(r.status == 'blocked' for r in result.records),
        'critical_errors': sum(i.severity == 'critical' for i in result.issues),
        'warnings': sum(i.severity == 'warning' for i in result.issues),
        'issues': [dict(vars(i), worksheet=result.source.sheet,
                       cell=column(i.col)+str(i.row) if i.col and i.row else None)
                   for i in result.issues],
    }


def transform_directory(source_dir, target_dir):
    source_dir, target_dir = Path(source_dir), Path(target_dir)
    if not source_dir.is_dir():
        raise ValueError('Source directory does not exist')
    if source_dir.resolve() == target_dir.resolve():
        raise ValueError('Source and target directories must be different')
    # Snapshot once; do not recurse or import files created during this run.
    paths = sorted((p for p in source_dir.iterdir()
                    if p.is_file() and p.suffix.lower() == '.xlsx'
                    and not p.name.startswith(('~$', '.sales-'))),
                   key=lambda p:(p.name.casefold(),p.name))
    if not paths:
        raise ValueError('No source XLSX files found')
    if len({p.stem.casefold() for p in paths}) != len(paths):
        raise ValueError('Ambiguous source names differing only by case')
    reports = []
    for path in paths:
        result, output = transform(path,target_dir)
        reports.append(dict(status_report(result,output),source=str(path.resolve())))
    return reports


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--source', default='source', help='Source XLSX file or directory (default: source)')
    parser.add_argument('--target', default='target')
    parser.add_argument('--master', help='Explicit baseline for a single source file only')
    args = parser.parse_args()
    source = Path(args.source)
    if source.is_dir() or source.suffix.lower() != '.xlsx':
        if args.master:
            parser.error('--master requires one XLSX source file, not a directory')
        try:
            reports = transform_directory(source,args.target)
        except (OSError,ValueError):
            print(json.dumps({'status':'failed','code':'ERR-SOURCE-DIRECTORY',
                              'explanation':'Source directory must exist, contain unambiguous XLSX inputs, and differ from target.'}))
            return 1
        successful = sum(r['status']=='success' for r in reports)
        report = {'status':'success' if successful==len(reports) else 'failed',
                  'files':len(reports),'succeeded':successful,
                  'failed':len(reports)-successful,'results':reports}
        print(json.dumps(report,ensure_ascii=False,indent=2))
        return 0 if successful==len(reports) else 1
    result, output = transform(source,args.target,args.master)
    print(json.dumps(dict(status_report(result,output),source=str(source.resolve())),ensure_ascii=False,indent=2))
    return 0 if output else 1


if __name__ == '__main__':
    raise SystemExit(main())

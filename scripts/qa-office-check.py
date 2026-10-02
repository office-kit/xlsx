"""Check a canonical semantic projection using Python stdlib only."""
import hashlib
import json
import posixpath
import pathlib
import sys
import xml.etree.ElementTree as ET
import zipfile

source, output = map(pathlib.Path, sys.argv[1:3])
application = sys.argv[3] if len(sys.argv) > 3 else 'libreoffice'
if application not in ('libreoffice', 'excel'): raise ValueError('Unknown application profile')
ns = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
results = []
failures = []
def check(actual, expected, context):
    if actual != expected: failures.append((context, actual, expected))
def features(package):
    # Follow relationships from the workbook; orphan parts cannot satisfy a claim.
    r = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    n = {'s': ns['s'], 'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart',
         'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
         'x': 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'}
    def xml(path): return ET.fromstring(package.read(path))
    def target(path, identifier, kind):
        relpath = posixpath.join(posixpath.dirname(path), '_rels', posixpath.basename(path) + '.rels')
        rel = next(node for node in xml(relpath) if node.get('Id') == identifier)
        if rel.get('Type') != r + '/' + kind or rel.get('TargetMode') == 'External':
            raise ValueError('Wrong feature relationship type/mode')
        to = rel.get('Target')
        return posixpath.normpath(to[1:] if to.startswith('/') else posixpath.join(posixpath.dirname(path), to))
    workbook = xml('xl/workbook.xml')
    sheet_ref = workbook.find('s:sheets/s:sheet', n)
    path = target('xl/workbook.xml', sheet_ref.get('{' + r + '}id'), 'worksheet')
    sheet = xml(path)
    tables, charts, images, anchors = [], [], [], []
    for link in sheet.findall('s:tableParts/s:tablePart', n):
        table = xml(target(path, link.get('{' + r + '}id'), 'table'))
        tables.append({'name': table.get('name'), 'ref': table.get('ref'),
                       'columns': [col.get('name') for col in table.findall('s:tableColumns/s:tableColumn', n)],
                       'autoFilter': table.find('s:autoFilter', n).get('ref') if table.find('s:autoFilter', n) is not None else None})
    drawing_link = sheet.find('s:drawing', n)
    if drawing_link is not None:
        drawing_path = target(path, drawing_link.get('{' + r + '}id'), 'drawing')
        drawing = xml(drawing_path)
        for anchor in drawing:
            chart_link = anchor.find('.//c:chart', n)
            image_link = anchor.find('.//a:blip', n)
            if chart_link is None and image_link is None: continue
            def marker(tag, coordinate):
                node = anchor.find('x:' + tag + '/x:' + coordinate, n)
                return int(node.text) if node is not None else None
            anchors.append(['chart' if chart_link is not None else 'image', marker('from', 'col'), marker('from', 'row'), marker('to', 'col'), marker('to', 'row')])
            if chart_link is not None:
                chart = xml(target(drawing_path, chart_link.get('{' + r + '}id'), 'chart'))
                plot = chart.find('c:chart/c:plotArea', n)
                kinds = [node for node in plot if node.tag.endswith('Chart')]
                for kind in kinds:
                    for ser in kind.findall('c:ser', n):
                        def text(query):
                            node = ser.find(query, n)
                            return node.text if node is not None else None
                        charts.append({'type': kind.tag.split('}')[1], 'valueRef': text('c:val/c:numRef/c:f'),
                                       'categoryRef': text('c:cat/c:strRef/c:f'),
                                       'values': [float(node.text) for node in ser.findall('c:val/c:numRef/c:numCache/c:pt/c:v', n)],
                                       'categories': [node.text for node in ser.findall('c:cat/c:strRef/c:strCache/c:pt/c:v', n)]})
            if image_link is not None:
                image = package.read(target(drawing_path, image_link.get('{' + r + '}embed'), 'image'))
                images.append(hashlib.sha256(image).hexdigest())
    def attributes(tag, keys, convert=str):
        node = sheet.find('s:' + tag, n)
        return {key: convert(node.get(key)) for key in keys if node is not None and node.get(key) is not None}
    page_setup = attributes('pageSetup', ('paperSize', 'orientation', 'fitToWidth', 'fitToHeight'))
    # ECMA-376 CT_PageSetup defaults: Excel may omit an explicit value of 1.
    if sheet.find('s:pageSetup', n) is not None:
        for key in ('fitToWidth', 'fitToHeight'): page_setup.setdefault(key, '1')
    return {'tables': tables, 'charts': charts, 'images': images, 'anchors': anchors,
            'pageSetup': page_setup,
            'pageMargins': attributes('pageMargins', ('left', 'right', 'top', 'bottom', 'header', 'footer'), float),
            'printOptions': attributes('printOptions', ('horizontalCentered',))}
for case in json.loads((source / 'manifest.json').read_text()):
    with zipfile.ZipFile(output / (case['id'] + '.output.xlsx')) as package:
        sheet = ET.fromstring(package.read('xl/worksheets/sheet1.xml'))
        cells = {cell.attrib['r']: cell for cell in sheet.findall('.//s:c', ns)}
        shared = []
        if 'xl/sharedStrings.xml' in package.namelist():
            shared = [''.join(t.text or '' for t in si.findall('.//s:t', ns)) for si in ET.fromstring(package.read('xl/sharedStrings.xml'))]
        def value(cell):
            kind = cell.get('t')
            v = cell.find('s:v', ns)
            if kind == 's':
                return shared[int(v.text)]
            if kind == 'inlineStr':
                return ''.join(t.text or '' for t in cell.findall('.//s:t', ns))
            if kind == 'b':
                return v.text in ('1', 'true')
            if kind in (None, 'n') and v is not None and v.text is not None: return float(v.text)
            return v.text if v is not None else None
        if 'expectedFeatures' in case: check(features(package), case['expectedFeatures'], case['id'] + ':features')
        actual = value(cells['A1'])
        expected = case.get('officeValue', case.get('expectedValue', 'audit'))
        check(actual, expected, case['id'])
        for i, wanted in enumerate(case.get('officeValues', []), 1): check(value(cells['A' + str(i)]), wanted, case['id'] + ':A' + str(i))
        check(value(cells['B2']), 'edited', case['id'] + ':edited')
        workbook = ET.fromstring(package.read('xl/workbook.xml'))
        if 'expectedDate1904' in case:
            properties = workbook.find('s:workbookPr', ns)
            observed = properties is not None and properties.get('date1904', '0') in ('1', 'true')
            check(observed, case['expectedDate1904'], case['id'] + ':date1904')
        if 'expectedNames' in case:
            sheet_names = [n.get('name') for n in workbook.findall('s:sheets/s:sheet', ns)]
            def name_value(text, scope):
                if scope is not None:
                    name = sheet_names[int(scope)]
                    # Normalize only this exact sheet's prefix; localSheetId is not a reference.
                    for prefix in (name + '!', "'" + name.replace("'", "''") + "'!"):
                        if text.startswith(prefix): return text[len(prefix):]
                return text
            observed = sorted((n.get('name'), name_value(n.text, n.get('localSheetId')), n.get('localSheetId')) for n in workbook.findall('s:definedNames/s:definedName', ns))
            wanted = sorted((n['name'], name_value(n['value'], str(n['scope']) if 'scope' in n else None), str(n['scope']) if 'scope' in n else None) for n in case['expectedNames'])
            check(observed, wanted, case['id'] + ':names')
        if isinstance(case.get('expectedValue'), dict) and case['expectedValue'].get('kind') == 'rich-text':
            cell = cells['A1']
            container = ET.fromstring(package.read('xl/sharedStrings.xml'))[int(cell.find('s:v', ns).text)] if cell.get('t') == 's' else cell.find('s:is', ns)
            runs = []
            for run in container.findall('s:r', ns):
                properties = run.find('s:rPr', ns)
                font = {}
                for key in ('b', 'i'):
                    prop = properties.find('s:' + key, ns) if properties is not None else None
                    if prop is not None and prop.get('val', 'true') in ('1', 'true'): font[key] = True
                runs.append({'text': ''.join(t.text or '' for t in run.findall('s:t', ns)), 'font': font})
            check(runs, case['expectedValue']['runs'], case['id'] + ':runs')
        styles = ET.fromstring(package.read('xl/styles.xml'))
        xf = styles.find('s:cellXfs', ns)[int(cells['A1'].get('s', '0'))]
        if 'numFmtId' in case:
            if case['numFmtId'] != 14: raise ValueError('Unprofiled office number format')
            format_id = int(xf.get('numFmtId', '0'))
            custom = {int(n.get('numFmtId')): n.get('formatCode') for n in styles.findall('s:numFmts/s:numFmt', ns)}
            # Built-in 14 is locale-dependent; Calc expands it to a custom date format.
            observed = 'date' if format_id == 14 or custom.get(format_id, '').lower() in ('mm/dd/yyyy', 'm/d/yyyy', 'mm/dd/yy', 'm/d/yy', 'mm-dd-yy') else custom.get(format_id, str(format_id))
            if xf.get('applyNumberFormat') in ('0', 'false'): observed = 'not-applied'
            check(observed, 'date', case['id'] + ':numberFormat')
        font = styles.find('s:fonts', ns)[int(xf.get('fontId', '0'))]
        tags = {'bold': 'b', 'italic': 'i', 'strike': 'strike', 'outline': 'outline', 'shadow': 'shadow', 'condense': 'condense', 'extend': 'extend'}
        for key, wanted in case.get('expectedFont', {}).items():
            element = font.find('s:' + tags.get(key, 'u'), ns)
            # Application rendering treats absent false and explicit false alike.
            if key == 'underline':
                observed = element.get('val', 'single') if element is not None else 'none'
            else:
                observed = element is not None and element.get('val', 'true') in ('1', 'true')
            check(observed, wanted, case['id'] + ':' + key)
        for key in case.get('expectedMissingFont', []):
            element = font.find('s:' + tags[key], ns)
            observed = element is not None and element.get('val', 'true') in ('1', 'true')
            check(observed, False, case['id'] + ':' + key)
        if 'expectedBorder' in case:
            border = styles.find('s:borders', ns)[int(xf.get('borderId', '0'))]
            side = border.find('s:left', ns)
            observed = side.get('style', 'none') if side is not None else 'none'
            check(observed, case['expectedBorder'], case['id'] + ':border')
        results.append({'id': case['id'], 'status': 'pass'})
profile = json.loads(pathlib.Path('tests/conformance/corpus/libreoffice-profile.json').read_text()) if application == 'libreoffice' else {}
known = []
unexpected = []
for context, actual, expected in failures:
    baseline = profile.get(context)
    if baseline and actual == baseline['actual'] and expected == baseline['expected']:
        known.append({'context': context, **baseline})
    else:
        unexpected.append((context, actual, expected))
missing = set(profile) - {entry['context'] for entry in known}
for entry in results:
    if any(context == entry['id'] or context.startswith(entry['id'] + ':') for context, _, _ in unexpected):
        entry['status'] = 'fail'
    elif any(k['context'].startswith(entry['id'] + ':') for k in known):
        entry['status'] = 'known-compatibility-difference'
report = {'cases': results, 'knownDifferences': known, 'unexpectedDifferences': unexpected,
          'changedBaseline': sorted(missing)}
(output / 'results.json').write_text(json.dumps(report, indent=2))
print(f"{application}: {sum(entry['status'] == 'pass' for entry in results)} pass, {len(known)} known compatibility differences, {len(unexpected)} unexpected differences")
if unexpected or missing:
    print(json.dumps(report, indent=2))
    sys.exit(1)

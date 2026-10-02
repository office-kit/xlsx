"""Check a canonical semantic projection using Python stdlib only."""
import json
import pathlib
import sys
import xml.etree.ElementTree as ET
import zipfile

source, output = map(pathlib.Path, sys.argv[1:])
ns = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
results = []
failures = []
def check(actual, expected, context):
    if actual != expected: failures.append((context, actual, expected))
for case in json.loads((source / 'manifest.json').read_text()):
    with zipfile.ZipFile(output / (case['id'] + '.output.xlsx')) as package:
        sheet = ET.fromstring(package.read('xl/worksheets/sheet1.xml'))
        cells = {cell.attrib['r']: cell for cell in sheet.findall('.//s:c', ns)}
        shared = []
        if 'xl/sharedStrings.xml' in package.namelist():
            shared = [''.join(si.itertext()) for si in ET.fromstring(package.read('xl/sharedStrings.xml'))]
        def value(cell):
            kind = cell.get('t')
            v = cell.find('s:v', ns)
            if kind == 's':
                return shared[int(v.text)]
            if kind == 'inlineStr':
                return ''.join(cell.find('s:is', ns).itertext())
            if kind == 'b':
                return v.text in ('1', 'true')
            return v.text if v is not None else None
        actual = value(cells['A1'])
        expected = case.get('expectedValue', 'audit')
        check(actual, expected, case['id'])
        check(value(cells['B2']), 'edited', case['id'] + ':edited')
        styles = ET.fromstring(package.read('xl/styles.xml'))
        xf = styles.find('s:cellXfs', ns)[int(cells['A1'].get('s', '0'))]
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
profile = json.loads(pathlib.Path('tests/conformance/corpus/libreoffice-profile.json').read_text())
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
    if any(k['context'].startswith(entry['id'] + ':') for k in known):
        entry['status'] = 'known-compatibility-difference'
report = {'cases': results, 'knownDifferences': known, 'unexpectedDifferences': unexpected,
          'changedBaseline': sorted(missing)}
(output / 'results.json').write_text(json.dumps(report, indent=2))
print(f"LibreOffice: {len(results) - len(known)} pass, {len(known)} known compatibility differences, {len(unexpected)} unexpected differences")
if unexpected or missing:
    print(json.dumps(report, indent=2))
    sys.exit(1)

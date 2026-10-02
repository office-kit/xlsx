"""Same-environment differential rendering; no golden image is produced by the writer."""
import hashlib
import json
import os
from pathlib import Path
import platform
import re
import subprocess
import sys

root = Path(sys.argv[1]).resolve()
soffice = os.environ.get('SOFFICE', 'soffice')
env = {**os.environ, 'LC_ALL': 'C.UTF-8', 'TZ': 'UTC'}

def run(args, timeout=45):
    result = subprocess.run(args, capture_output=True, text=True, env=env, timeout=timeout, check=True)
    return result.stdout + result.stderr

def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()

report = {'status': 'running', 'platform': platform.platform(), 'dpi': 96, 'locale': env['LC_ALL'], 'timezone': env['TZ'], 'cases': {}}

def save():
    (root / 'results.json').write_text(json.dumps(report, indent=2) + '\n')

def render(case):
    folder = root / case
    profile = root / f'profile-{case}'
    log = run([soffice, f'-env:UserInstallation={profile.as_uri()}', '--headless', '--convert-to', 'pdf:calc_pdf_Export', '--outdir', str(folder), str(folder / 'sample.xlsx')])
    (folder / 'convert.log').write_text(log)
    pdf = folder / 'sample.pdf'
    if not pdf.is_file():
        raise AssertionError(f'Missing PDF: {case}')
    info = run(['pdfinfo', str(pdf)])
    count = re.search(r'^Pages:\s+(\d+)$', info, re.MULTILINE)
    if not count or int(count[1]) < 1:
        raise AssertionError('Missing/nonpositive page count')
    pages = int(count[1])
    fonts = run(['pdffonts', str(pdf)])
    if 'LiberationSans' not in fonts:
        raise AssertionError('Requested cell font is missing from rendered PDF')
    (folder / 'pdfinfo.txt').write_text(info)
    (folder / 'pdffonts.txt').write_text(fonts)
    run(['pdftotext', '-layout', str(pdf), str(folder / 'text.txt')])
    # PPM is a simple uncompressed raster; byte equality includes dimensions and every pixel.
    run(['pdftoppm', '-r', '96', str(pdf), str(folder / 'page')])
    run(['pdftoppm', '-r', '96', '-png', str(pdf), str(folder / 'preview')])
    rasters = sorted(folder.glob('page-*.ppm'))
    previews = sorted(folder.glob('preview-*.png'))
    if len(rasters) != pages or len(previews) != pages:
        raise AssertionError('Incomplete rendered page set')
    return {'pages': pages, 'text': (folder / 'text.txt').read_text(), 'rasterSha256': [sha(path) for path in rasters], 'xlsxSha256': sha(folder / 'sample.xlsx'), 'pdfSha256': sha(pdf)}

def equivalent(left, right):
    return all(left[key] == right[key] for key in ['pages', 'text', 'rasterSha256'])

save()
try:
    report['libreoffice'] = run([soffice, '--version']).strip()
    report['poppler'] = run(['pdftoppm', '-v']).splitlines()[0]
    bundled = Path(soffice).parent.parent / 'Resources/fonts/truetype/LiberationSans-Regular.ttf'
    font = Path(os.environ['QA_RENDER_FONT']) if 'QA_RENDER_FONT' in os.environ else bundled if bundled.is_file() else Path(run(['fc-match', 'Liberation Sans', '-f', '%{file}']).strip())
    if not font.is_file() or 'LiberationSans' not in font.name:
        raise AssertionError('Install Liberation Sans or set QA_RENDER_FONT to its exact file')
    report['font'] = {'path': str(font), 'sha256': sha(font)}
    for case in ['input', 'output', 'changed-orientation', 'lost-drawing', 'changed-chart-color']:
        report['cases'][case] = render(case)
        save()
    # Ensure the independently colored pie is actually painted, rather than an empty chart frame.
    raster = next((root / 'input').glob('page-*.ppm')).read_bytes().split(b'\n', 3)[3]
    for color in [bytes.fromhex('4472c4'), bytes.fromhex('ed7d31')]:
        if sum(raster[i:i + 3] == color for i in range(0, len(raster), 3)) < 500:
            raise AssertionError('Expected visible pie sector is missing')
    before = report['cases']['input']
    if before['pages'] != 1 or not all(text in before['text'] for text in ['audit', 'First', 'Second']):
        raise AssertionError('Independent input must render one page with all expected cell labels')
    if not equivalent(before, report['cases']['output']):
        raise AssertionError('Library save changed real page count, extracted text or raster pixels')
    for case in ['changed-orientation', 'lost-drawing', 'changed-chart-color']:
        if equivalent(before, report['cases'][case]):
            raise AssertionError(f'Renderer calibration failed to detect {case}')
    if before['text'] != report['cases']['changed-chart-color']['text']:
        raise AssertionError('Color calibration must isolate raster changes from extracted text')
    report['status'] = 'passed'
    save()
    print('PDF rendering: one-page input/output pixels and text match; orientation, detached drawings and chart colors are detected')
except Exception as error:
    report['status'] = 'failed'
    report['error'] = str(error)
    save()
    raise

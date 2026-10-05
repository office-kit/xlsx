// Excel's header/footer mini-language: `&L` / `&C` / `&R` split a header into
// sections, `&P` `&N` `&D` `&T` `&F` `&Z` `&A` are fields filled at print time,
// and the rest (`&B`, `&I`, `&"Font,Style"`, `&12`, `&KRRGGBB`, …) only format
// text, so they are dropped when the text is shown here.

import type { HeaderFooter } from '@office-kit/xlsx/worksheet';

export interface HeaderParts {
  left: string;
  center: string;
  right: string;
}

/** Split a header string into sections; text before any marker is centred. */
export function parseHeader(text: string | undefined): HeaderParts {
  const out: HeaderParts = { left: '', center: '', right: '' };
  if (!text) return out;
  let section: keyof HeaderParts = 'center';
  let i = 0;
  while (i < text.length) {
    const marker = text.slice(i, i + 2);
    if (marker === '&L' || marker === '&C' || marker === '&R') {
      section = marker === '&L' ? 'left' : marker === '&C' ? 'center' : 'right';
      i += 2;
      continue;
    }
    // "&&" is a literal ampersand and must not be read as a marker.
    if (marker === '&&') {
      out[section] += marker;
      i += 2;
      continue;
    }
    out[section] += text.charAt(i);
    i++;
  }
  return out;
}

export function buildHeader(parts: HeaderParts): string | undefined {
  const text = (parts.left ? `&L${parts.left}` : '') + (parts.center ? `&C${parts.center}` : '') + (parts.right ? `&R${parts.right}` : '');
  return text || undefined;
}

/** Values of the static fields; page numbers come from the printer (see {@link hfSegments}). */
export interface HfFields {
  readonly date: string;
  readonly time: string;
  readonly file: string;
  readonly path: string;
  readonly sheet: string;
}

/** A section as text runs and page counters, which only the print engine can number. */
export type HfSegment = string | { readonly counter: 'page' | 'pages'; readonly offset: number };

// &P+3 / &P-1 shift the printed page number.
const PAGE_OFFSET_RE = /^[+-]\d+/;
// &"Font,Style", &KRRGGBB or &KTT+NNN (theme colour), and font sizes like &12.
const FORMAT_RE = /^(?:"[^"]*"|K(?:[0-9A-Fa-f]{6}|\d\d[+-]\d{3})|\d+)/;

export function hfSegments(section: string, fields: HfFields): HfSegment[] {
  const out: HfSegment[] = [];
  let text = '';
  const flush = () => {
    if (text) out.push(text);
    text = '';
  };
  let i = 0;
  while (i < section.length) {
    const ch = section.charAt(i);
    if (ch !== '&' || i + 1 >= section.length) {
      text += ch;
      i++;
      continue;
    }
    const code = section.charAt(i + 1);
    i += 2;
    switch (code.toUpperCase()) {
      case '&':
        text += '&';
        break;
      case 'P':
      case 'N': {
        const m = code.toUpperCase() === 'P' ? PAGE_OFFSET_RE.exec(section.slice(i)) : null;
        if (m) i += m[0].length;
        flush();
        out.push({ counter: code.toUpperCase() === 'P' ? 'page' : 'pages', offset: m ? Number(m[0]) : 0 });
        break;
      }
      case 'D':
        text += fields.date;
        break;
      case 'T':
        text += fields.time;
        break;
      case 'F':
        text += fields.file;
        break;
      case 'Z':
        text += fields.path;
        break;
      case 'A':
        text += fields.sheet;
        break;
      default: {
        // Formatting code: skip it and any argument it takes.
        const m = FORMAT_RE.exec(section.slice(i - 1));
        if (m) i += m[0].length - 1;
      }
    }
  }
  flush();
  return out;
}

/** A section as plain text for page `page` of `pages`. */
export function renderHf(section: string, fields: HfFields, page: number, pages: number | string): string {
  return hfSegments(section, fields)
    .map((s) => (typeof s === 'string' ? s : s.counter === 'page' ? String(page + s.offset) : String(pages)))
    .join('');
}

/** Words the presets are built from, in the UI language ("Page &P", "Confidential"). */
export interface PresetWords {
  readonly page: string;
  readonly pageOf: string;
  readonly confidential: string;
}

/** Excel's Header / Footer drop-down entries, as header strings ('' is "(none)"). */
export function hfPresets(w: PresetWords): string[] {
  return [
    '',
    `&C${w.page}`,
    `&C${w.pageOf}`,
    '&C&A',
    `&L${w.confidential}&C&D&R${w.page}`,
    '&C&F',
    `&C&A&R${w.page}`,
    '&C&A&R&F',
    `&C&F&R${w.page}`,
    '&C&F&R&A',
    `&C${w.page}&R&A`,
    `&C${w.page}&R&F`,
  ];
}

/** How a preset reads in the list: its sections filled in for page 1, joined by commas. */
export function presetLabel(codes: string, fields: HfFields): string {
  const p = parseHeader(codes);
  return [p.left, p.center, p.right]
    .filter((s) => s !== '')
    .map((s) => renderHf(s, fields, 1, '?'))
    .join(', ');
}

/** The header or footer string that applies to 1-based `page` (first / even pages may differ). */
export function headerFooterFor(hf: HeaderFooter | undefined, page: number, kind: 'header' | 'footer'): string | undefined {
  if (!hf) return undefined;
  if (page === 1 && hf.differentFirst) return kind === 'header' ? hf.firstHeader : hf.firstFooter;
  if (page % 2 === 0 && hf.differentOddEven) return kind === 'header' ? hf.evenHeader : hf.evenFooter;
  return kind === 'header' ? hf.oddHeader : hf.oddFooter;
}

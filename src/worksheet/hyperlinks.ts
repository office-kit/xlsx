// Worksheet hyperlinks.
//
// External URLs land in `xl/worksheets/_rels/sheetN.xml.rels` (one rel per
// Hyperlink entry, type "...relationships/hyperlink", TargetMode="External").
// Internal jumps (`#'Sheet 2'!A1`) live entirely in the `<hyperlink
// location="..."/>` attribute and don't need a rel.

import { OpenXmlSchemaError } from '../utils/exceptions';

export interface Hyperlink {
  /** Cell or range the hyperlink covers — "A1" or "A1:B5". */
  ref: string;
  /** External URL or relative target path. Mutually exclusive with `location`-only links. */
  target?: string;
  /** Anchor inside the workbook (e.g. `'Sheet 2'!A1`). */
  location?: string;
  /** Tooltip shown on hover. */
  tooltip?: string;
  /** Visible link text — typically falls back to the referenced cell value. */
  display?: string;
  /** Worksheet-rels rId. Populated on read; assigned by the writer when missing. */
  rId?: string;
}

/**
 * Anything the OPC `Target` attribute (`xsd:anyURI`) cannot carry: whitespace,
 * C0/C1 controls, and DEL. Excel does not ignore the one bad link — it refuses
 * the whole package and drops content on repair.
 */
const ILLEGAL_TARGET_CHAR = /[\u0000-\u0020\u007f-\u009f]/;

/** The authority of a `scheme://authority/…` target, if the target has one. */
const AUTHORITY = /^[A-Za-z][A-Za-z0-9+.-]*:\/\/([^/?#]*)/;

const describeChar = (ch: string): string =>
  ch === ' ' ? 'a space' : `U+${ch.codePointAt(0)?.toString(16).toUpperCase().padStart(4, '0')}`;

/**
 * Reject a target the writer would emit verbatim into the rels and Excel would
 * then refuse. Applied when the hyperlink is authored, not when one is read
 * back off disk — a file that already carries a broken target still loads, so
 * it can be inspected and repaired.
 */
function assertValidTarget(target: string): void {
  if (target.length === 0) {
    throw new OpenXmlSchemaError(
      'Hyperlink: target is empty — pass a URI, or use `location` for an in-workbook jump',
    );
  }
  const illegal = ILLEGAL_TARGET_CHAR.exec(target);
  if (illegal?.[0] !== undefined) {
    throw new OpenXmlSchemaError(
      `Hyperlink: target "${target}" contains ${describeChar(illegal[0])} at index ${illegal.index}. ` +
        'The rels Target attribute is xsd:anyURI and Excel refuses the file; percent-encode it or drop it.',
    );
  }
  const authority = AUTHORITY.exec(target)?.[1];
  if (authority !== undefined && /[^\u0021-\u007e]/.test(authority)) {
    throw new OpenXmlSchemaError(
      `Hyperlink: target "${target}" has a non-ASCII host. Percent-encoding does not rescue it — ` +
        'convert the host to punycode (xn--…).',
    );
  }
}

/**
 * Build a validated {@link Hyperlink}. `target` is checked against what the
 * OPC `Target` attribute accepts, so a bad URL fails here rather than as an
 * Excel repair prompt on a package that has already been written.
 */
export function makeHyperlink(opts: Partial<Hyperlink> & { ref: string }): Hyperlink {
  if (opts.ref === undefined || opts.ref.length === 0) {
    throw new OpenXmlSchemaError('Hyperlink: ref is required');
  }
  if (opts.target !== undefined) assertValidTarget(opts.target);
  return {
    ref: opts.ref,
    ...(opts.target !== undefined ? { target: opts.target } : {}),
    ...(opts.location !== undefined ? { location: opts.location } : {}),
    ...(opts.tooltip !== undefined ? { tooltip: opts.tooltip } : {}),
    ...(opts.display !== undefined ? { display: opts.display } : {}),
    ...(opts.rId !== undefined ? { rId: opts.rId } : {}),
  };
}

// ---- Worksheet ergonomic helpers ----------------------------------------

import type { Worksheet } from './worksheet';

const replaceHyperlink = (ws: Worksheet, hl: Hyperlink): Hyperlink => {
  const idx = ws.hyperlinks.findIndex((h) => h.ref === hl.ref);
  if (idx >= 0) ws.hyperlinks[idx] = hl;
  else ws.hyperlinks.push(hl);
  return hl;
};

/**
 * Add an external URL hyperlink to a cell or range. The URL goes into the
 * worksheet rels as a hyperlink relationship; the writer generates an rId on
 * save.
 */
export const addUrlHyperlink = (
  ws: Worksheet,
  ref: string,
  url: string,
  opts: { tooltip?: string; display?: string } = {},
): Hyperlink => {
  return replaceHyperlink(
    ws,
    makeHyperlink({
      ref,
      target: url,
      ...(opts.tooltip !== undefined ? { tooltip: opts.tooltip } : {}),
      ...(opts.display !== undefined ? { display: opts.display } : {}),
    }),
  );
};

/**
 * Add an in-workbook jump hyperlink (e.g. to `'Sheet2'!A1` or a defined-name).
 * No rels entry is written — the location is inline in the `<hyperlink
 * location="…"/>` attribute.
 */
export const addInternalHyperlink = (
  ws: Worksheet,
  ref: string,
  location: string,
  opts: { tooltip?: string; display?: string } = {},
): Hyperlink => {
  return replaceHyperlink(
    ws,
    makeHyperlink({
      ref,
      location,
      ...(opts.tooltip !== undefined ? { tooltip: opts.tooltip } : {}),
      ...(opts.display !== undefined ? { display: opts.display } : {}),
    }),
  );
};

/** `mailto:` shortcut. */
export const addMailtoHyperlink = (
  ws: Worksheet,
  ref: string,
  email: string,
  opts: { subject?: string; tooltip?: string; display?: string } = {},
): Hyperlink => {
  const url = opts.subject
    ? `mailto:${email}?subject=${encodeURIComponent(opts.subject)}`
    : `mailto:${email}`;
  return addUrlHyperlink(ws, ref, url, {
    ...(opts.tooltip !== undefined ? { tooltip: opts.tooltip } : {}),
    ...(opts.display !== undefined ? { display: opts.display } : {}),
  });
};

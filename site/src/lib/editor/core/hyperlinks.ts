// Following hyperlinks from the grid.

import type { Hyperlink, Worksheet } from '@office-kit/xlsx/worksheet';
import { parseRangeAddress } from './address.ts';
import type { EditorController } from './controller.svelte.ts';
import { goToReference } from './names.ts';

/** Schemes safe to open from a link in an untrusted workbook. */
const SAFE_SCHEMES = new Set(['http:', 'https:', 'mailto:', 'ftp:']);

export function hyperlinkAt(ws: Worksheet, row: number, col: number): Hyperlink | undefined {
  return ws.hyperlinks.find((h) => {
    const r = parseRangeAddress(h.ref)?.range;
    return r !== undefined && row >= r.r1 && row <= r.r2 && col >= r.c1 && col <= r.c2;
  });
}

/** Open an external link in a new tab or jump to an in-workbook location. Returns whether it went anywhere. */
export function followHyperlink(ctl: EditorController, link: Hyperlink): boolean {
  if (link.location) return goToReference(ctl, link.location);
  if (!link.target) return false;
  let url: URL;
  try {
    url = new URL(link.target);
  } catch {
    // Relative file targets point next to the workbook on disk, which a browser cannot reach.
    return false;
  }
  if (!SAFE_SCHEMES.has(url.protocol)) return false;
  window.open(url.href, '_blank', 'noopener,noreferrer');
  return true;
}

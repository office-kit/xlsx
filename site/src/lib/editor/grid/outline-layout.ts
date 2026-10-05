// Where the outline gutters' level buttons, +/− buttons and brackets go for
// the current view. Shared by the painter and the pointer hit-test so the
// two can never disagree.

import type { Worksheet } from '@office-kit/xlsx/worksheet';
import { isCollapsed, type AxisOutline, type OutlineAxis, type OutlineRun } from '../core/outline.ts';
import { OUTLINE_STEP, type GridGeometry } from './geometry.ts';

export interface OutlineButton {
  readonly axis: OutlineAxis;
  readonly x: number;
  readonly y: number;
  readonly size: number;
  readonly action: { readonly kind: 'level'; readonly level: number } | { readonly kind: 'toggle'; readonly run: OutlineRun; readonly collapsed: boolean };
}

export interface OutlineBracket {
  readonly axis: OutlineAxis;
  /** Across-axis position of the bracket line. */
  readonly at: number;
  /** Along-axis span in px, ending where the summary button sits. */
  readonly from: number;
  readonly to: number;
}

export interface OutlineLayout {
  readonly buttons: readonly OutlineButton[];
  readonly brackets: readonly OutlineBracket[];
}

export function outlineLayout(geo: GridGeometry, ws: Worksheet, rows: AxisOutline, cols: AxisOutline): OutlineLayout {
  const buttons: OutlineButton[] = [];
  const brackets: OutlineBracket[] = [];
  const scale = Math.max(0.6, Math.min(geo.zoom, 2));
  const step = OUTLINE_STEP * scale;
  const size = Math.round(11 * scale);

  if (rows.maxLevel > 0) {
    for (let n = 1; n <= rows.maxLevel + 1; n++) {
      buttons.push({ axis: 'row', x: (n - 1) * step + (step - size) / 2, y: geo.headerH - size - 4, size, action: { kind: 'level', level: n } });
    }
    const [first, last] = visibleSpan(geo, 'row');
    for (const run of rows.runs) {
      if (run.end < first - 1 && run.summary < first) continue;
      if (run.start > last + 1 && run.summary > last) continue;
      const collapsed = isCollapsed(ws, 'row', run);
      const at = (run.level - 1) * step + step / 2;
      if (!collapsed) {
        const y1 = geo.rowY(run.start);
        const y2 = geo.rowY(run.end) + geo.rowH(run.end);
        brackets.push({ axis: 'row', at, from: y1, to: y2 });
      }
      if (run.summary >= 1) {
        const y = geo.rowY(run.summary) + (geo.rowH(run.summary) - size) / 2;
        buttons.push({ axis: 'row', x: at - size / 2, y, size, action: { kind: 'toggle', run, collapsed } });
      }
    }
  }

  if (cols.maxLevel > 0) {
    for (let n = 1; n <= cols.maxLevel + 1; n++) {
      buttons.push({ axis: 'col', x: geo.headerW - size - 4, y: (n - 1) * step + (step - size) / 2, size, action: { kind: 'level', level: n } });
    }
    const [first, last] = visibleSpan(geo, 'col');
    for (const run of cols.runs) {
      if (run.end < first - 1 && run.summary < first) continue;
      if (run.start > last + 1 && run.summary > last) continue;
      const collapsed = isCollapsed(ws, 'col', run);
      const at = (run.level - 1) * step + step / 2;
      if (!collapsed) {
        const x1 = geo.colX(run.start);
        const x2 = geo.colX(run.end) + geo.colW(run.end);
        brackets.push({ axis: 'col', at, from: x1, to: x2 });
      }
      if (run.summary >= 1) {
        const x = geo.colX(run.summary) + (geo.colW(run.summary) - size) / 2;
        buttons.push({ axis: 'col', x, y: at - size / 2, size, action: { kind: 'toggle', run, collapsed } });
      }
    }
  }
  return { buttons, brackets };
}

function visibleSpan(geo: GridGeometry, axis: OutlineAxis): [number, number] {
  const [first, last] = axis === 'row' ? geo.mainRows() : geo.mainCols();
  const frozen = axis === 'row' ? geo.frozenRows : geo.frozenCols;
  return [frozen > 0 ? 1 : first, last];
}

export function outlineButtonAt(layout: OutlineLayout, x: number, y: number): OutlineButton | undefined {
  // A little slack around the small buttons, as Excel's hit areas are generous.
  return layout.buttons.find((b) => x >= b.x - 2 && x <= b.x + b.size + 2 && y >= b.y - 2 && y <= b.y + b.size + 2);
}

// Defined names: resolving what the Name Box / Go To accept, and creating
// names for the selection.

import type { DefinedName } from '@office-kit/xlsx/workbook';
import { parseRangeAddress, quoteSheetName, rangeAddress, type Range } from './address.ts';
import type { EditorController } from './controller.svelte.ts';
import { selectRange } from './selection.ts';

/** Sheet + range a name refers to, when it is a plain reference. */
export function resolveName(ctl: EditorController, name: string): { sheetIndex: number; range: Range } | undefined {
  const lower = name.toLowerCase();
  const active = ctl.doc.activeSheetIndex;
  const candidates = ctl.doc.wb.definedNames.filter((d) => d.name.toLowerCase() === lower);
  // A sheet-scoped name on the active sheet wins over a workbook-scoped one.
  const dn = candidates.find((d) => d.scope === active) ?? candidates.find((d) => d.scope === undefined);
  if (!dn) return undefined;
  return resolveRefText(ctl, dn.value.replace(/^=/, ''));
}

function resolveRefText(ctl: EditorController, text: string): { sheetIndex: number; range: Range } | undefined {
  const parsed = parseRangeAddress(text);
  if (!parsed) return undefined;
  const sheetIndex = parsed.sheet === undefined ? ctl.doc.activeSheetIndex : ctl.doc.wb.sheets.findIndex((s) => s.sheet.title.toLowerCase() === parsed.sheet?.toLowerCase());
  if (sheetIndex < 0 || ctl.doc.wb.sheets[sheetIndex]?.kind !== 'worksheet') return undefined;
  return { sheetIndex, range: parsed.range };
}

/** Select what `text` refers to (address, sheet-qualified address or defined name). */
export function goToReference(ctl: EditorController, text: string): boolean {
  const target = resolveRefText(ctl, text) ?? resolveName(ctl, text);
  if (!target) return false;
  if (ctl.edit && !ctl.commitEdit()) return true;
  ctl.doc.activateSheet(target.sheetIndex);
  // As with the mouse, a range that cuts through a merged cell grows to take all of it.
  ctl.doc.setSelection(selectRange(ctl.doc.merges.expand(target.range)));
  ctl.reveal(target.range.r1, target.range.c1);
  return true;
}

const NAME_RE = /^[A-Za-z_\\À-￿][A-Za-z0-9_.\\À-￿]*$/;

export function validateName(name: string): 'invalidName' | undefined {
  if (!NAME_RE.test(name) || /^[A-Za-z]{1,3}\d+$/.test(name) || /^[RrCc]$/.test(name) || name.length > 255) return 'invalidName';
  return undefined;
}

export function addName(ctl: EditorController, dn: DefinedName): string | undefined {
  const err = validateName(dn.name);
  if (err) return err;
  const wb = ctl.doc.wb;
  if (wb.definedNames.some((d) => d.name.toLowerCase() === dn.name.toLowerCase() && d.scope === dn.scope)) return 'duplicateName';
  ctl.doc.transact('Define Name', (tx) => {
    tx.structural = true;
    tx.workbook('definedNames');
    wb.definedNames = [...wb.definedNames, dn];
  });
  return undefined;
}

export function selectionRefText(ctl: EditorController): string {
  const r = ctl.doc.selection.ranges[ctl.doc.selection.activeRange] ?? ctl.doc.selection.ranges[0];
  return `${quoteSheetName(ctl.doc.ws.title)}!${r ? rangeAddress(r, true) : ''}`;
}

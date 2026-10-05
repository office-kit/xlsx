// Keyboard shortcuts available while navigating (not editing a cell). Both the
// Windows (Ctrl) and Mac (Cmd) spellings are accepted; where Mac Excel binds a
// different key (Cmd+Shift+X strikethrough, Cmd+Shift+T AutoSum, Ctrl+G Go
// To while Cmd+G is Find Next) both forms are listed.

import * as A from './actions.ts';
import type { EditorController } from './controller.svelte.ts';
import { newComment } from './comments.ts';
import { findNext } from './find.ts';

type Handler = (ctl: EditorController) => void;

interface Binding {
  readonly key: string;
  readonly mod?: boolean;
  readonly shift?: boolean;
  readonly alt?: boolean;
  /** Require the Ctrl key specifically (Mac Excel distinguishes Ctrl from Cmd for some). */
  readonly ctrl?: boolean;
  readonly cmd?: boolean;
  /** Only on macOS, where Excel binds Ctrl and Cmd to different commands. */
  readonly macOnly?: boolean;
  readonly run: Handler;
}

export const IS_MAC = typeof navigator !== 'undefined' && /Mac|iPhone|iPad/.test(navigator.platform);

const open = (kind: Parameters<EditorController['openDialog']>[0]): Handler => (ctl) => ctl.openDialog(kind);

const BINDINGS: readonly Binding[] = [
  // Mac Excel: Ctrl+U edits the cell (F2); Cmd+U stays Underline.
  { key: 'u', ctrl: true, macOnly: true, run: (ctl) => ctl.startEdit() },
  { key: 'z', mod: true, run: A.undo },
  { key: 'z', mod: true, shift: true, run: A.redo },
  { key: 'y', mod: true, run: A.redo },
  { key: 'F4', run: (ctl) => ctl.repeatable?.() },
  { key: 'b', mod: true, run: A.toggleBold },
  { key: '2', mod: true, run: A.toggleBold },
  { key: 'i', mod: true, run: A.toggleItalic },
  { key: '3', mod: true, run: A.toggleItalic },
  { key: 'u', mod: true, run: (ctl) => A.toggleUnderline(ctl) },
  { key: '4', mod: true, run: (ctl) => A.toggleUnderline(ctl) },
  { key: '5', mod: true, run: A.toggleStrike },
  { key: 'x', mod: true, shift: true, run: A.toggleStrike },
  { key: '1', mod: true, run: open('formatCells') },
  { key: 'f', mod: true, run: open('find') },
  { key: 'h', mod: true, run: open('replace') },
  { key: 'g', ctrl: true, run: open('goto') },
  { key: 'g', cmd: true, run: (ctl) => findNext(ctl, 1) },
  { key: 'g', cmd: true, shift: true, run: (ctl) => findNext(ctl, -1) },
  { key: 'F5', run: open('goto') },
  { key: 'k', mod: true, run: open('hyperlink') },
  { key: 'd', mod: true, run: (ctl) => A.fillFrom(ctl, 'down') },
  { key: 'e', mod: true, run: A.flashFillActive },
  { key: '8', mod: true, run: (ctl) => A.setViewFlag(ctl, 'showOutlineSymbols', !ctl.showOutlineSymbols) },
  { key: 'r', mod: true, run: (ctl) => A.fillFrom(ctl, 'right') },
  { key: 'a', mod: true, run: (ctl) => ctl.selectAll() },
  { key: ' ', ctrl: true, run: (ctl) => ctl.selectEntire('cols') },
  { key: ' ', mod: true, shift: true, run: (ctl) => ctl.selectAll() },
  { key: '-', mod: true, run: A.deleteCellsSmart },
  { key: '=', mod: true, shift: true, run: A.insertCellsSmart },
  { key: '+', mod: true, run: A.insertCellsSmart },
  { key: '+', mod: true, shift: true, run: A.insertCellsSmart },
  { key: '9', mod: true, run: (ctl) => A.hideLines(ctl, 'row', true) },
  { key: '0', mod: true, run: (ctl) => A.hideLines(ctl, 'col', true) },
  { key: '(', mod: true, shift: true, run: (ctl) => A.hideLines(ctl, 'row', false) },
  { key: ')', mod: true, shift: true, run: (ctl) => A.hideLines(ctl, 'col', false) },
  { key: 'l', mod: true, shift: true, run: (ctl) => ctl.toggleFilter() },
  { key: 'f', mod: true, shift: true, run: (ctl) => ctl.toggleFilter() },
  { key: 't', mod: true, run: open('createTable') },
  { key: 'l', ctrl: true, run: open('createTable') },
  { key: '~', mod: true, shift: true, run: (ctl) => A.setNumberFormat(ctl, 'General') },
  { key: '!', mod: true, shift: true, run: (ctl) => A.setNumberFormat(ctl, '#,##0.00') },
  { key: '$', mod: true, shift: true, run: (ctl) => A.setNumberFormat(ctl, '"$"#,##0.00_);("$"#,##0.00)') },
  { key: '%', mod: true, shift: true, run: (ctl) => A.setNumberFormat(ctl, '0%') },
  { key: '^', mod: true, shift: true, run: (ctl) => A.setNumberFormat(ctl, '0.00E+00') },
  { key: '#', mod: true, shift: true, run: (ctl) => A.setNumberFormat(ctl, 'd-mmm-yy') },
  { key: '@', mod: true, shift: true, run: (ctl) => A.setNumberFormat(ctl, 'h:mm AM/PM') },
  { key: '&', mod: true, shift: true, run: (ctl) => A.applyBorder(ctl, 'outside') },
  { key: '_', mod: true, shift: true, run: (ctl) => A.applyBorder(ctl, 'none') },
  { key: '=', alt: true, run: (ctl) => A.autoSum(ctl) },
  { key: 't', cmd: true, shift: true, run: (ctl) => A.autoSum(ctl) },
  { key: ';', mod: true, run: (ctl) => ctl.startEdit(ctl.nowText('date')) },
  { key: ':', mod: true, shift: true, run: (ctl) => ctl.startEdit(ctl.nowText('time')) },
  { key: ';', mod: true, shift: true, run: (ctl) => ctl.startEdit(ctl.nowText('time')) },
  { key: 'F2', shift: true, run: open('note') },
  { key: 'F2', mod: true, shift: true, run: newComment },
  { key: 'F3', shift: true, run: open('insertFunction') },
  { key: 'F11', shift: true, run: (ctl) => A.insertSheet(ctl) },
  { key: 'F9', run: (ctl) => ctl.doc.recalculate() },
  { key: '*', mod: true, shift: true, run: (ctl) => ctl.selectAll() },
  { key: 's', mod: true, run: (ctl) => ctl.host?.save() },
  { key: 's', mod: true, shift: true, run: (ctl) => ctl.host?.saveAs() },
  { key: 'o', mod: true, run: (ctl) => ctl.host?.open() },
  { key: 'p', mod: true, run: (ctl) => ctl.host?.print() },
  { key: 'v', ctrl: true, cmd: true, run: open('pasteSpecial') },
  { key: 'v', mod: true, alt: true, run: open('pasteSpecial') },
  { key: 'r', mod: true, alt: true, run: (ctl) => (ctl.ribbonCollapsed = !ctl.ribbonCollapsed) },
  { key: 'ArrowDown', alt: true, run: (ctl) => ctl.openPickList() },
  { key: '>', mod: true, shift: true, run: (ctl) => A.stepFontSize(ctl, 1) },
  { key: '<', mod: true, shift: true, run: (ctl) => A.stepFontSize(ctl, -1) },
  { key: 'k', cmd: true, shift: true, run: (ctl) => A.groupSelection(ctl, true) },
  { key: 'j', cmd: true, shift: true, run: (ctl) => A.groupSelection(ctl, false) },
  { key: 'ArrowRight', alt: true, shift: true, run: (ctl) => A.groupSelection(ctl, true) },
  { key: 'ArrowLeft', alt: true, shift: true, run: (ctl) => A.groupSelection(ctl, false) },
  { key: '`', ctrl: true, run: (ctl) => A.setViewFlag(ctl, 'showFormulas', !ctl.showFormulas) },
  { key: '=', cmd: true, run: (ctl) => ctl.doc.recalculate() },
  { key: 'F9', shift: true, run: (ctl) => ctl.doc.recalculate() },
  { key: 'u', mod: true, shift: true, run: (ctl) => (ctl.formulaBarExpanded = !ctl.formulaBarExpanded) },
  { key: '0', cmd: true, alt: true, run: (ctl) => A.applyBorder(ctl, 'outside') },
  { key: 'ArrowUp', cmd: true, alt: true, run: (ctl) => A.applyBorder(ctl, 'top') },
  { key: 'ArrowDown', cmd: true, alt: true, run: (ctl) => A.applyBorder(ctl, 'bottom') },
  { key: 'ArrowLeft', cmd: true, alt: true, run: (ctl) => A.applyBorder(ctl, 'left') },
  { key: 'ArrowRight', cmd: true, alt: true, run: (ctl) => A.applyBorder(ctl, 'right') },
  { key: '-', cmd: true, alt: true, run: (ctl) => A.applyBorder(ctl, 'none') },
  { key: "'", mod: true, run: (ctl) => A.copyFromAbove(ctl, 'formula') },
  { key: '"', mod: true, shift: true, run: (ctl) => A.copyFromAbove(ctl, 'value') },
  { key: '.', ctrl: true, run: (ctl) => ctl.nextCorner() },
  { key: 'Backspace', shift: true, run: (ctl) => ctl.selectCell(ctl.doc.selection.active) },
  { key: 'Backspace', mod: true, run: (ctl) => ctl.reveal(ctl.doc.selection.active.row, ctl.doc.selection.active.col) },
];

function matches(b: Binding, ev: KeyboardEvent): boolean {
  if (b.macOnly && !IS_MAC) return false;
  const key = ev.key.length === 1 ? ev.key.toLowerCase() : ev.key;
  if (key !== b.key) return false;
  const mod = ev.ctrlKey || ev.metaKey;
  if (b.ctrl && b.cmd) {
    if (!(ev.ctrlKey && ev.metaKey)) return false;
  } else if (b.ctrl) {
    if (!ev.ctrlKey || ev.metaKey) return false;
  } else if (b.cmd) {
    if (!ev.metaKey || ev.ctrlKey) return false;
  } else if (Boolean(b.mod) !== mod) return false;
  if (Boolean(b.alt) !== ev.altKey) return false;
  // Shifted symbols (`!`, `$`, …) already encode Shift in `key`; letters must match explicitly.
  if (b.key.length === 1 && /[a-z0-9 =;\-+]/.test(b.key)) return Boolean(b.shift) === ev.shiftKey;
  return b.shift === undefined || b.shift === ev.shiftKey;
}

/** Run the shortcut bound to `ev`, if any. Returns true when handled. */
export function runShortcut(ctl: EditorController, ev: KeyboardEvent): boolean {
  // Ctrl+C / X / V go through the native clipboard events, never here.
  const k = ev.key.toLowerCase();
  if ((ev.ctrlKey || ev.metaKey) && !ev.altKey && !ev.shiftKey && (k === 'c' || k === 'x' || k === 'v')) return false;
  for (const b of BINDINGS) {
    if (matches(b, ev)) {
      b.run(ctl);
      return true;
    }
  }
  return false;
}

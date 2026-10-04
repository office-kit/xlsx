// Keyboard handling for the grid's input element, in Excel's three entry
// modes. Ctrl and Cmd are treated alike so Windows and Mac shortcuts both
// work (Mac Excel accepts Ctrl for most of them as well).

import { insertArgumentNames } from '../core/formula-assist.ts';
import { formatPainter } from '../core/format-painter.svelte.ts';
import type { EditorController } from '../core/controller.svelte.ts';
import { runShortcut } from '../core/shortcuts.ts';

type Dir = -1 | 0 | 1;

const ARROWS: Readonly<Record<string, [Dir, Dir]>> = {
  ArrowUp: [-1, 0],
  ArrowDown: [1, 0],
  ArrowLeft: [0, -1],
  ArrowRight: [0, 1],
};

export function handleGridKey(ctl: EditorController, ev: KeyboardEvent, input: HTMLTextAreaElement): void {
  const mod = ev.ctrlKey || ev.metaKey;
  const e = ctl.edit;

  if (e) {
    handleEditingKey(ctl, ev, input, mod);
    return;
  }

  // Shortcuts first (formatting, clipboard, dialogs, …).
  if (runShortcut(ctl, ev)) {
    ev.preventDefault();
    return;
  }

  const arrow = ARROWS[ev.key];
  if (arrow) {
    ev.preventDefault();
    ctl.move(arrow[0], arrow[1], { extend: ev.shiftKey, jump: mod });
    return;
  }
  switch (ev.key) {
    case 'Enter':
      ev.preventDefault();
      ctl.advance(ev.shiftKey ? 'up' : 'down');
      return;
    case 'Tab':
      ev.preventDefault();
      ctl.advance(ev.shiftKey ? 'left' : 'right');
      return;
    case 'Home':
      ev.preventDefault();
      ctl.home(mod, ev.shiftKey);
      return;
    case 'End':
      if (mod) {
        ev.preventDefault();
        ctl.end(ev.shiftKey);
      }
      return;
    case 'PageDown':
    case 'PageUp':
      ev.preventDefault();
      if (mod) ctl.switchSheet(ev.key === 'PageDown' ? 1 : -1);
      else ctl.pageMove(ev.key === 'PageDown' ? 1 : -1, ev.altKey, ev.shiftKey);
      return;
    case 'F2':
      ev.preventDefault();
      ctl.startEdit();
      return;
    case 'Delete':
      ev.preventDefault();
      ctl.clearSelection('contents');
      return;
    case 'Backspace':
      ev.preventDefault();
      // Mac's delete key and Windows' Backspace clear the active cell and start editing it.
      ctl.startEdit('');
      return;
    case 'Escape':
      if (formatPainter.active) formatPainter.cancel();
      if (ctl.clipboard) {
        ev.preventDefault();
        ctl.clipboard = null;
      }
      return;
    case 'ContextMenu':
      ev.preventDefault();
      ctl.openContextMenuAtActive();
      return;
    default:
  }
  if (ev.key === 'F10' && ev.shiftKey) {
    ev.preventDefault();
    ctl.openContextMenuAtActive();
    return;
  }
  if (ev.key === ' ' && ev.shiftKey && !mod) {
    ev.preventDefault();
    ctl.selectEntire('rows');
  }
  // Printable characters fall through to the textarea, whose input event starts Enter mode.
}

function handleEditingKey(ctl: EditorController, ev: KeyboardEvent, input: HTMLTextAreaElement, mod: boolean): void {
  const e = ctl.edit;
  if (!e) return;
  if (ctl.handleCompletionKey(ev)) return;
  const arrow = ARROWS[ev.key];

  if (ev.key === 'Escape') {
    ev.preventDefault();
    ctl.cancelEdit();
    return;
  }
  if (ev.key === 'Enter') {
    // Alt/Option+Enter (and Ctrl+Option+Enter on Mac) insert a line break.
    if (ev.altKey) {
      ev.preventDefault();
      const s = input.selectionStart;
      const t = input.selectionEnd;
      ctl.setEditText(e.text.slice(0, s) + '\n' + e.text.slice(t), s + 1);
      return;
    }
    ev.preventDefault();
    if (!ctl.commitEdit({ fillSelection: mod })) return;
    if (!mod) ctl.advance(ev.shiftKey ? 'up' : 'down');
    return;
  }
  if (ev.key === 'Tab') {
    ev.preventDefault();
    if (!ctl.commitEdit()) return;
    ctl.advance(ev.shiftKey ? 'left' : 'right');
    return;
  }
  if (ev.key === 'F2') {
    ev.preventDefault();
    e.mode = e.mode === 'edit' ? 'enter' : 'edit';
    e.point = null;
    return;
  }
  // F4 (and Cmd+T, Mac Excel's binding) cycles $ anchoring while editing a formula.
  if (ev.key === 'F4' || (ev.metaKey && ev.key.toLowerCase() === 't')) {
    ev.preventDefault();
    ctl.toggleAbsolute();
    return;
  }
  if (arrow) {
    if (ctl.canPoint()) {
      ev.preventDefault();
      if (mod) {
        // Ctrl+arrow in point mode jumps the reference like it jumps the selection.
        const from = e.point?.cursor ?? { row: e.row, col: e.col };
        ctl.selectCell(from);
        ctl.move(arrow[0], arrow[1], { jump: true });
        ctl.pointTo(ctl.doc.selection.active, ev.shiftKey);
      } else ctl.pointMove(arrow[0], arrow[1], ev.shiftKey);
      return;
    }
    if (e.mode === 'enter') {
      ev.preventDefault();
      if (!ctl.commitEdit()) return;
      ctl.move(arrow[0], arrow[1], { extend: ev.shiftKey, jump: mod });
      return;
    }
    // Edit mode: the caret moves inside the text (native behaviour).
    return;
  }
  if (ev.ctrlKey && ev.shiftKey && ev.key.toLowerCase() === 'a') {
    ev.preventDefault();
    const next = insertArgumentNames(e.text, input.selectionStart);
    if (next) ctl.setEditText(next.text, next.caret);
    return;
  }
  if (mod && (ev.key === ';' || ev.key === ':')) {
    ev.preventDefault();
    const stamp = ev.shiftKey || ev.key === ':' ? ctl.nowText('time') : ctl.nowText('date');
    const s = input.selectionStart;
    ctl.setEditText(e.text.slice(0, s) + stamp + e.text.slice(input.selectionEnd), s + stamp.length);
  }
}

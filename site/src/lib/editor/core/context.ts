// Svelte context access to the editor controller, so deeply nested ribbon
// and dialog components don't need it threaded through props.

import { getContext, setContext } from 'svelte';
import type { EditorController } from './controller.svelte.ts';

const KEY = Symbol('xlsx-editor');

export function setEditor(ctl: EditorController): void {
  setContext(KEY, ctl);
}

export function getEditor(): EditorController {
  const ctl = getContext<EditorController | undefined>(KEY);
  if (!ctl) throw new Error('getEditor() called outside <EditorApp>');
  return ctl;
}

// The REPL's CodeMirror editor. The REPL page imports this module dynamically,
// so CodeMirror stays out of the route's initial JavaScript.

import { javascript } from '@codemirror/lang-javascript';
import { syntaxTree } from '@codemirror/language';
import { EditorState } from '@codemirror/state';
import { oneDark } from '@codemirror/theme-one-dark';
import { EditorView, basicSetup } from 'codemirror';

export type ReplEditor = {
  setText(text: string): void;
  /**
   * The 1-based line of the first thing the parser could not read. Engines
   * report no position for a SyntaxError raised by `new Function`, so the
   * editor's own parse tree is where the line comes from.
   */
  firstSyntaxErrorLine(): number | undefined;
  destroy(): void;
};

export function createEditor(
  parent: HTMLElement,
  text: string,
  onChange: (text: string) => void,
): ReplEditor {
  const view = new EditorView({
    parent,
    state: EditorState.create({
      doc: text,
      extensions: [
        basicSetup,
        javascript(),
        // The pane is half a screen, or a phone: wrapping beats scrolling sideways.
        EditorView.lineWrapping,
        oneDark,
        // One Dark supplies the syntax colours; the surfaces are ours so the
        // editor matches every other code panel on the site.
        EditorView.theme(
          {
            '&': { height: '100%', fontSize: '13px', backgroundColor: 'var(--night)' },
            '.cm-gutters': { backgroundColor: 'var(--night)', border: 'none' },
            '.cm-activeLine, .cm-activeLineGutter': { backgroundColor: 'var(--night-2)' },
            '.cm-scroller': { fontFamily: 'var(--mono)', lineHeight: '1.6', overflow: 'auto' },
          },
          { dark: true },
        ),
        EditorView.updateListener.of((update) => {
          if (update.docChanged) onChange(update.state.doc.toString());
        }),
      ],
    }),
  });

  return {
    setText(next) {
      view.dispatch({ changes: { from: 0, to: view.state.doc.length, insert: next } });
    },
    firstSyntaxErrorLine() {
      let position: number | undefined;
      syntaxTree(view.state).iterate({
        enter(node) {
          if (position !== undefined) return false;
          if (node.type.isError) position = node.from;
          return undefined;
        },
      });
      return position === undefined ? undefined : view.state.doc.lineAt(position).number;
    },
    destroy() {
      view.destroy();
    },
  };
}

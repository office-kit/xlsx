<script lang="ts">
  // In-place text editing of a shape or text box. The text commits as one undo
  // step when the editor loses focus; Escape commits too and leaves the shape
  // selected, as in Excel.
  import type { ShapeReference } from '@office-kit/xlsx/drawing';
  import { onMount, untrack } from 'svelte';
  import { getEditor } from '../core/context.ts';
  import { setShapeText, shapeText } from '../core/shapes.ts';
  import { layoutText, shapePaint } from './shape-render.ts';

  let { index, shape, zoom }: { index: number; shape: ShapeReference; zoom: number } = $props();

  const ctl = getEditor();
  const palette = ctl.doc.styles.palette;
  // The text the editor opened with; later model changes don't reset what is being typed.
  const initial = untrack(() => shapeText(shape.txBody));
  let text = $state(initial);
  let area: HTMLTextAreaElement;

  const layout = $derived(shape.txBody ? layoutText(shape.txBody, shapePaint(shape, palette), palette, zoom) : undefined);
  const firstRun = $derived(layout?.paragraphs[0]?.runs[0]?.style ?? `font-size:${(11 * 96 * zoom) / 72}px;color:${shapePaint(shape, palette).fontColor}`);

  onMount(() => {
    area.focus();
    area.setSelectionRange(text.length, text.length);
  });

  let done = false;
  function commit(): void {
    if (done) return;
    done = true;
    if (ctl.shapeTextEdit === index) ctl.shapeTextEdit = null;
    if (text !== initial) setShapeText(ctl.doc, index, text);
  }

  function onKey(ev: KeyboardEvent): void {
    // Keys typed here belong to the text, not to the grid or the drawing layer.
    ev.stopPropagation();
    if (ev.key === 'Escape') {
      ev.preventDefault();
      commit();
    }
  }
</script>

<textarea
  bind:this={area}
  bind:value={text}
  class="editor"
  style="{firstRun};text-align:{layout?.paragraphs[0]?.align ?? 'left'};padding:{(layout?.insets[0] ?? 5) * zoom}px {(layout?.insets[1] ?? 10) * zoom}px"
  spellcheck="false"
  onkeydown={onKey}
  onblur={commit}
  onpointerdown={(ev) => ev.stopPropagation()}
  ondblclick={(ev) => ev.stopPropagation()}
></textarea>

<style>
  .editor {
    position: absolute;
    inset: 0;
    box-sizing: border-box;
    width: 100%;
    height: 100%;
    border: 0;
    margin: 0;
    background: transparent;
    resize: none;
    outline: none;
    line-height: 1.2;
    font-family: var(--xl-font, system-ui);
    cursor: text;
    overflow: hidden;
  }
</style>

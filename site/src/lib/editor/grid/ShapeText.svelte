<script lang="ts">
  // The text of a shape or text box, laid out as HTML over the shape's SVG so
  // the browser handles wrapping, fonts and CJK text.
  import type { ShapeReference } from '@office-kit/xlsx/drawing';
  import { getEditor } from '../core/context.ts';
  import { layoutText, shapePaint } from './shape-render.ts';

  let { shape, zoom }: { shape: ShapeReference; zoom: number } = $props();

  const palette = getEditor().doc.styles.palette;
  const layout = $derived(shape.txBody ? layoutText(shape.txBody, shapePaint(shape, palette), palette, zoom) : undefined);
</script>

{#if layout}
  <div
    class="text"
    class:nowrap={!layout.wrap}
    style:justify-content={layout.justify}
    style:padding="{layout.insets[0] * zoom}px {layout.insets[1] * zoom}px {layout.insets[2] * zoom}px {layout.insets[3] * zoom}px"
  >
    {#each layout.paragraphs as para, i (i)}
      <p style:text-align={para.align} style:min-height="{((para.emptySizePt * 96) / 72) * zoom * 1.2}px">
        {#each para.runs as run, j (j)}<span style={run.style}>{run.text}</span>{/each}
      </p>
    {/each}
  </div>
{/if}

<style>
  .text {
    position: absolute;
    inset: 0;
    display: flex;
    flex-direction: column;
    overflow: hidden;
    box-sizing: border-box;
    pointer-events: none;
    white-space: pre-wrap;
    overflow-wrap: anywhere;
    line-height: 1.2;
    font-family: var(--xl-font, system-ui);
  }
  .nowrap {
    white-space: pre;
  }
  p {
    margin: 0;
  }
</style>

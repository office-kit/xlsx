<script lang="ts">
  // Draw mode after Insert ▸ Shapes / Text Box: the cursor turns into a
  // crosshair, a drag on the sheet draws the shape and a click drops the
  // default size, as in Excel. Escape (or a press outside the sheet) disarms.
  import { getEditor } from '../core/context.ts';
  import { insertShape } from '../core/shapes.ts';

  const ctl = getEditor();
  const doc = ctl.doc;

  let drag = $state<{ x0: number; y0: number; x1: number; y1: number } | null>(null);
  let overlay: HTMLDivElement | undefined = $state();

  /** Screen offset inside the cell area → sheet content px at 100% zoom. */
  function contentPoint(ev: PointerEvent): { x: number; y: number } {
    const r = overlay?.getBoundingClientRect();
    const sx = ev.clientX - (r?.left ?? 0);
    const sy = ev.clientY - (r?.top ?? 0);
    const geo = ctl.geometry;
    // The frozen panes don't scroll.
    const x = sx < geo.frozenW ? sx : sx + doc.scrollX;
    const y = sy < geo.frozenH ? sy : sy + doc.scrollY;
    return { x: Math.max(0, x / doc.zoom), y: Math.max(0, y / doc.zoom) };
  }

  function down(ev: PointerEvent): void {
    if (ev.button !== 0) return;
    ev.preventDefault();
    ev.stopPropagation();
    overlay?.setPointerCapture(ev.pointerId);
    const p = contentPoint(ev);
    drag = { x0: p.x, y0: p.y, x1: p.x, y1: p.y };
  }

  function move(ev: PointerEvent): void {
    if (!drag) return;
    const p = contentPoint(ev);
    let { x, y } = p;
    // Shift keeps a shape square (or a line at 45° steps), as in Excel.
    if (ev.shiftKey) {
      const d = Math.max(Math.abs(x - drag.x0), Math.abs(y - drag.y0));
      x = drag.x0 + Math.sign(x - drag.x0 || 1) * d;
      y = drag.y0 + Math.sign(y - drag.y0 || 1) * d;
    }
    drag = { ...drag, x1: x, y1: y };
  }

  function up(): void {
    const d = drag;
    const tool = ctl.shapeTool;
    drag = null;
    if (!d || !tool) return;
    ctl.shapeTool = null;
    const index = insertShape(doc, ctl.geometry, tool, { x: d.x0, y: d.y0 }, { x: d.x1, y: d.y1 });
    if (index === undefined) return;
    ctl.selectedDrawing = index;
    ctl.ribbonTab = 'shapeFormat';
    ctl.ribbonCollapsed = false;
    if (tool.kind === 'textBox') ctl.shapeTextEdit = index;
  }

  function onKey(ev: KeyboardEvent): void {
    if (ev.key === 'Escape' && ctl.shapeTool) {
      ev.preventDefault();
      drag = null;
      ctl.shapeTool = null;
    }
  }

  const preview = $derived.by(() => {
    if (!drag) return null;
    const z = doc.zoom;
    const geo = ctl.geometry;
    const toScreen = (x: number, y: number) => ({
      x: x * z - (x * z < geo.frozenW ? 0 : doc.scrollX),
      y: y * z - (y * z < geo.frozenH ? 0 : doc.scrollY),
    });
    const a = toScreen(drag.x0, drag.y0);
    const b = toScreen(drag.x1, drag.y1);
    return { left: Math.min(a.x, b.x), top: Math.min(a.y, b.y), width: Math.abs(b.x - a.x), height: Math.abs(b.y - a.y) };
  });
</script>

<svelte:window onkeydown={onKey} />

{#if ctl.shapeTool}
  <div
    class="draw"
    bind:this={overlay}
    role="presentation"
    style:left="{ctl.geometry.headerW}px"
    style:top="{ctl.geometry.headerH}px"
    onpointerdown={down}
    onpointermove={move}
    onpointerup={up}
    onpointercancel={() => (drag = null)}
  >
    {#if preview}
      <div class="preview" style:left="{preview.left}px" style:top="{preview.top}px" style:width="{preview.width}px" style:height="{preview.height}px"></div>
    {/if}
  </div>
{/if}

<style>
  .draw {
    position: absolute;
    right: 0;
    bottom: 0;
    cursor: crosshair;
    touch-action: none;
    z-index: 4;
  }
  .preview {
    position: absolute;
    border: 1px solid #605e5c;
    background: rgba(68, 114, 196, 0.15);
    pointer-events: none;
  }
</style>

<script lang="ts" module>
  import type { DrawingItem } from '@office-kit/xlsx/drawing';

  const MIME: Readonly<Record<string, string>> = {
    png: 'image/png',
    jpeg: 'image/jpeg',
    gif: 'image/gif',
    bmp: 'image/bmp',
    webp: 'image/webp',
    tiff: 'image/tiff',
    svg: 'image/svg+xml',
  };

  function isVisible(item: DrawingItem): boolean {
    return item.content.kind !== 'picture' || item.content.picture.hidden !== true;
  }
</script>

<script lang="ts">
  // Pictures and charts floating over the grid canvas. Items are positioned
  // from their anchors through the same axis indices the painter uses, and
  // only those intersecting the viewport are mounted. A click selects an
  // item; dragging moves it, the corner handles resize it, Delete removes it.
  // Each gesture is one undo step that rewrites the item's anchor.
  import { onDestroy } from 'svelte';
  import { getEditor } from '../core/context.ts';
  import { chartData } from '../core/chart-data.ts';
  import { t } from '../i18n/i18n.svelte.ts';
  import { anchorFor, anchorRect, type Rect } from '../core/drawing-anchor.ts';
  import type { DrawingAnchor } from '@office-kit/xlsx/drawing';
  import { chartElements, editChart, setChartTitle } from '../core/chart-edit.ts';
  import { renderChartSvg, titleBandHeight } from './chart-render.ts';
  import { renderShapeSvg } from './shape-render.ts';
  import ShapeText from './ShapeText.svelte';
  import ShapeTextEditor from './ShapeTextEditor.svelte';
  import type { Worksheet } from '@office-kit/xlsx/worksheet';

  const ctl = getEditor();
  const doc = ctl.doc;

  let gesture = $state<{ index: number; kind: 'move' | 'resize'; handle: string; startX: number; startY: number; start: Rect; current: Rect } | null>(null);
  const urls = new Map<Uint8Array, string>();

  const items = $derived.by(() => {
    void doc.version;
    return doc.ws.drawing?.items ?? [];
  });

  // A selection belongs to one sheet's drawing; switching sheets or undoing past it drops it.
  let lastSheet: Worksheet | undefined;
  $effect(() => {
    const ws = doc.ws;
    const selected = ctl.selectedDrawing;
    if ((lastSheet && lastSheet !== ws) || (selected !== null && selected >= items.length)) ctl.selectedDrawing = null;
    lastSheet = ws;
  });

  const worksheets = $derived.by(() => {
    void doc.version;
    const map = new Map<string, Worksheet>();
    for (const s of doc.wb.sheets) if (s.kind === 'worksheet') map.set(s.sheet.title, s.sheet);
    return map;
  });

  function screenRect(r: Rect, a: DrawingAnchor): Rect {
    const geo = ctl.geometry;
    const z = doc.zoom;
    // Items anchored inside the frozen panes stay put while the rest scrolls.
    const fromCol = a.kind === 'absolute' ? Infinity : a.from.col + 1;
    const fromRow = a.kind === 'absolute' ? Infinity : a.from.row + 1;
    const sx = fromCol <= geo.frozenCols ? 0 : doc.scrollX;
    const sy = fromRow <= geo.frozenRows ? 0 : doc.scrollY;
    return { x: geo.headerW + r.x * z - sx, y: geo.headerH + r.y * z - sy, w: r.w * z, h: r.h * z };
  }

  const placed = $derived.by(() => {
    void doc.layoutVersion;
    const geo = ctl.geometry;
    const out: Array<{ index: number; item: DrawingItem; rect: Rect; screen: Rect }> = [];
    items.forEach((item, index) => {
      if (!isVisible(item)) return;
      const live = gesture?.index === index ? gesture.current : anchorRect(item.anchor, ctl.geometry.cols, ctl.geometry.rows);
      const screen = screenRect(live, item.anchor);
      if (screen.x + screen.w < geo.headerW || screen.y + screen.h < geo.headerH || screen.x > geo.width || screen.y > geo.height) return;
      out.push({ index, item, rect: live, screen });
    });
    return out;
  });

  function imageUrl(item: DrawingItem): string | undefined {
    if (item.content.kind !== 'picture') return undefined;
    const img = item.content.picture.image;
    if (!img) return undefined;
    const mime = MIME[img.format];
    if (!mime) return undefined;
    let url = urls.get(img.bytes);
    if (!url) {
      url = URL.createObjectURL(new Blob([img.bytes.slice()], { type: mime }));
      urls.set(img.bytes, url);
    }
    return url;
  }

  onDestroy(() => {
    for (const url of urls.values()) URL.revokeObjectURL(url);
  });

  // Scrolling re-places items every frame; the SVG only changes with the data or the size.
  const svgCache = new WeakMap<DrawingItem, { version: number; w: number; h: number; svg: string | undefined }>();

  function chartSvg(item: DrawingItem, width: number, height: number): string | undefined {
    if (item.content.kind !== 'chart') return undefined;
    const version = doc.version;
    const w = Math.max(40, Math.round(width));
    const h = Math.max(30, Math.round(height));
    const hit = svgCache.get(item);
    if (hit && hit.version === version && hit.w === w && hit.h === h) return hit.svg;
    const data = chartData(item.content.chart, doc.calc, worksheets, doc.styles.palette);
    const svg = data ? renderChartSvg(data, w, h, doc.styles.palette) : undefined;
    svgCache.set(item, { version, w, h, svg });
    return svg;
  }

  function placeholder(item: DrawingItem): string {
    switch (item.content.kind) {
      case 'chart':
        return item.content.chart.space?.plotArea.chart.kind ?? 'chartex';
      case 'picture':
        return item.content.picture.image?.format ?? 'picture';
      case 'shape':
        return 'shape';
      case 'unsupported':
        return item.content.rawTag;
    }
  }

  // ---- gestures ------------------------------------------------------------------

  function begin(ev: PointerEvent, index: number, rect: Rect, kind: 'move' | 'resize', handle = ''): void {
    if (ev.button !== 0) return;
    ev.preventDefault();
    ev.stopPropagation();
    if (ctl.edit && !ctl.commitEdit()) return;
    ctl.selectedDrawing = index;
    (ev.currentTarget as HTMLElement).setPointerCapture(ev.pointerId);
    gesture = { index, kind, handle, startX: ev.clientX, startY: ev.clientY, start: rect, current: rect };
    focusSelected();
  }

  function move(ev: PointerEvent): void {
    const g = gesture;
    if (!g) return;
    const dx = (ev.clientX - g.startX) / doc.zoom;
    const dy = (ev.clientY - g.startY) / doc.zoom;
    const s = g.start;
    let next: Rect;
    if (g.kind === 'move') next = { ...s, x: Math.max(0, s.x + dx), y: Math.max(0, s.y + dy) };
    else {
      let { x, y, w, h } = s;
      if (g.handle.includes('e')) w = s.w + dx;
      if (g.handle.includes('s')) h = s.h + dy;
      if (g.handle.includes('w')) {
        x = s.x + dx;
        w = s.w - dx;
      }
      if (g.handle.includes('n')) {
        y = s.y + dy;
        h = s.h - dy;
      }
      // Shift keeps a corner drag proportional, as in Excel.
      if (ev.shiftKey && g.handle.length === 2 && s.w > 0 && s.h > 0) {
        const k = Math.max(w / s.w, h / s.h);
        if (g.handle.includes('w')) x = s.x + s.w - s.w * k;
        if (g.handle.includes('n')) y = s.y + s.h - s.h * k;
        w = s.w * k;
        h = s.h * k;
      }
      next = { x: Math.max(0, x), y: Math.max(0, y), w: Math.max(8, w), h: Math.max(8, h) };
    }
    gesture = { ...g, current: next };
  }

  function end(): void {
    const g = gesture;
    gesture = null;
    if (!g) return;
    const c = g.current;
    if (c.x === g.start.x && c.y === g.start.y && c.w === g.start.w && c.h === g.start.h) return;
    const ws = doc.ws;
    doc.transact(g.kind === 'move' ? 'Move Object' : 'Resize Object', (tx) => {
      tx.sheet(ws, 'drawing');
      const item = ws.drawing?.items[g.index];
      if (item) item.anchor = anchorFor(item.anchor, c, ctl.geometry.cols, ctl.geometry.rows);
    });
  }

  function remove(): void {
    const index = ctl.selectedDrawing;
    const ws = doc.ws;
    if (index === null || !ws.drawing?.items[index]) return;
    ctl.selectedDrawing = null;
    doc.transact('Delete Object', (tx) => {
      tx.sheet(ws, 'drawing');
      ws.drawing?.items.splice(index, 1);
    });
  }

  function onKeyDown(ev: KeyboardEvent): void {
    if (titleEdit !== null || ctl.shapeTextEdit !== null) return;
    if (ev.key === 'Delete' || ev.key === 'Backspace') {
      ev.preventDefault();
      remove();
    } else if (ev.key === 'Escape') {
      ev.preventDefault();
      ctl.selectedDrawing = null;
    }
  }

  // Items sit outside the grid's scroll container, so wheel over a chart scrolls the sheet by hand.
  function onWheel(ev: WheelEvent): void {
    doc.setScroll(doc.scrollX + ev.deltaX, doc.scrollY + ev.deltaY);
  }

  let layer: HTMLDivElement;

  function focusSelected(): void {
    queueMicrotask(() => layer?.querySelector<HTMLElement>('.item.selected')?.focus({ preventScroll: true }));
  }

  // A press on the grid outside every drawing item drops the selection, like
  // clicking a cell in Excel. Presses on the ribbon or a dialog keep it, since
  // those are what edit the selected chart.
  function onWindowPointerDown(ev: PointerEvent): void {
    if (ctl.selectedDrawing === null || !(ev.target instanceof Node)) return;
    if (layer?.contains(ev.target) || !layer?.parentElement?.contains(ev.target)) return;
    ctl.selectedDrawing = null;
  }

  // ---- inline chart title editing ------------------------------------------------------

  /** Index of the chart whose title is being edited inline, and the text typed so far. */
  let titleEdit = $state<number | null>(null);
  let titleDraft = $state('');

  function onDblClick(ev: MouseEvent, index: number, item: DrawingItem, screen: Rect): void {
    if (item.content.kind === 'shape') {
      if (item.content.shape.connector) return;
      ev.stopPropagation();
      ctl.shapeTextEdit = index;
      return;
    }
    if (item.content.kind !== 'chart') return;
    const target = ev.currentTarget;
    if (!(target instanceof HTMLElement)) return;
    const y = ev.clientY - target.getBoundingClientRect().top;
    const el = chartElements(item.content.chart);
    if (el.title === 'none' || y > titleBandHeight(screen.w, screen.h)) return;
    ev.stopPropagation();
    titleEdit = index;
    titleDraft = el.titleText;
    queueMicrotask(() => {
      const input = layer?.querySelector<HTMLInputElement>('.title-edit');
      input?.focus();
      input?.select();
    });
  }

  function commitTitle(): void {
    const index = titleEdit;
    const text = titleDraft;
    titleEdit = null;
    if (index === null) return;
    const item = doc.ws.drawing?.items[index];
    if (item?.content.kind !== 'chart') return;
    const el = chartElements(item.content.chart);
    if (el.titleText === text) return;
    editChart(doc, index, 'Edit Chart Title', (chart) => setChartTitle(chart, el.title === 'none' ? 'above' : el.title, text));
    focusSelected();
  }

  function onTitleKey(ev: KeyboardEvent): void {
    ev.stopPropagation();
    if (ev.key === 'Enter') {
      ev.preventDefault();
      commitTitle();
    } else if (ev.key === 'Escape') {
      ev.preventDefault();
      titleEdit = null;
      focusSelected();
    }
  }

  const HANDLES = ['nw', 'n', 'ne', 'e', 'se', 's', 'sw', 'w'] as const;
  const handlePos = (h: string, r: Rect) => ({
    left: h.includes('w') ? 0 : h.includes('e') ? r.w : r.w / 2,
    top: h.includes('n') ? 0 : h.includes('s') ? r.h : r.h / 2,
  });
</script>

<svelte:window onpointerdown={onWindowPointerDown} />

<div class="layer" bind:this={layer} style:left="{ctl.geometry.headerW}px" style:top="{ctl.geometry.headerH}px">
  {#each placed as p (p.index)}
    {@const s = p.screen}
    {@const isSel = ctl.selectedDrawing === p.index}
    <div
      class="item"
      class:selected={isSel}
      role="button"
      tabindex={isSel ? 0 : -1}
      aria-label={placeholder(p.item)}
      style:left="{s.x - ctl.geometry.headerW}px"
      style:top="{s.y - ctl.geometry.headerH}px"
      style:width="{s.w}px"
      style:height="{s.h}px"
      onpointerdown={(ev) => begin(ev, p.index, p.rect, 'move')}
      onpointermove={move}
      onpointerup={end}
      onpointercancel={end}
      onkeydown={onKeyDown}
      onwheel={onWheel}
      ondblclick={(ev) => onDblClick(ev, p.index, p.item, s)}
    >
      {#if p.item.content.kind === 'picture' && imageUrl(p.item)}
        <img src={imageUrl(p.item)} alt={p.item.content.picture.descr ?? p.item.content.picture.name ?? ''} draggable="false" />
      {:else if p.item.content.kind === 'chart'}
        {@const svg = chartSvg(p.item, s.w, s.h)}
        {#if svg}
          <div class="chart">{@html svg}</div>
          {#if titleEdit === p.index}
            <input
              class="title-edit"
              style:height="{titleBandHeight(s.w, s.h)}px"
              bind:value={titleDraft}
              onkeydown={onTitleKey}
              onblur={commitTitle}
              onpointerdown={(ev) => ev.stopPropagation()}
              aria-label={t('chElChartTitle')}
            />
          {/if}
        {:else}
          <div class="placeholder">{placeholder(p.item)}</div>
        {/if}
      {:else if p.item.content.kind === 'shape'}
        {@const shape = p.item.content.shape}
        <div class="shape">{@html renderShapeSvg(shape, s.w, s.h, doc.styles.palette, `shp${p.index}`)}</div>
        {#if shape.connector}<span class="line-hit"></span>{/if}
        {#if ctl.shapeTextEdit === p.index}
          <ShapeTextEditor index={p.index} {shape} zoom={doc.zoom} />
        {:else if shape.txBody && !shape.connector}
          <ShapeText {shape} zoom={doc.zoom} />
        {/if}
      {:else}
        <div class="placeholder">{placeholder(p.item)}</div>
      {/if}
      {#if isSel}
        {#each HANDLES as h (h)}
          {@const pos = handlePos(h, s)}
          <span
            class="handle h-{h}"
            role="button"
            tabindex="-1"
            aria-label={h}
            style:left="{pos.left}px"
            style:top="{pos.top}px"
            onpointerdown={(ev) => begin(ev, p.index, p.rect, 'resize', h)}
            onpointermove={move}
            onpointerup={end}
            onpointercancel={end}
          ></span>
        {/each}
      {/if}
    </div>
  {/each}
</div>

<style>
  .layer {
    position: absolute;
    right: 0;
    bottom: 0;
    overflow: hidden;
    pointer-events: none;
  }
  .item {
    position: absolute;
    pointer-events: auto;
    cursor: move;
    outline: none;
    touch-action: none;
  }
  .item img {
    display: block;
    width: 100%;
    height: 100%;
    user-select: none;
  }
  .chart {
    width: 100%;
    height: 100%;
    overflow: hidden;
  }
  .title-edit {
    position: absolute;
    left: 10%;
    top: 0;
    width: 80%;
    box-sizing: border-box;
    border: 1px solid var(--xl-accent);
    background: #fff;
    font: inherit;
    font-size: 16px;
    text-align: center;
    color: #595959;
    outline: none;
    cursor: text;
  }
  .chart :global(svg) {
    display: block;
  }
  .shape {
    position: absolute;
    inset: 0;
    pointer-events: none;
  }
  /* A horizontal or vertical line has no area to click; widen its hit box. */
  .line-hit {
    position: absolute;
    inset: -4px;
  }
  .shape :global(svg) {
    display: block;
    overflow: visible;
  }
  .placeholder {
    display: flex;
    align-items: center;
    justify-content: center;
    width: 100%;
    height: 100%;
    background: #fafafa;
    border: 1px dashed #a6a6a6;
    color: #545b69;
    font-size: 12px;
  }
  .selected {
    outline: 1px solid #8a909c;
  }
  .handle {
    position: absolute;
    width: 9px;
    height: 9px;
    margin: -5px 0 0 -5px;
    background: #fff;
    border: 1px solid #545b69;
    border-radius: 50%;
  }
  .h-nw,
  .h-se {
    cursor: nwse-resize;
  }
  .h-ne,
  .h-sw {
    cursor: nesw-resize;
  }
  .h-n,
  .h-s {
    cursor: ns-resize;
  }
  .h-e,
  .h-w {
    cursor: ew-resize;
  }
</style>

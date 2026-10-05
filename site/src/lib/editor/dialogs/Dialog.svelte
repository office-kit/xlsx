<script lang="ts">
  // The frame every editor dialog shares: a compact Excel-for-Mac sheet with a
  // draggable title bar, the body, and an OK/Cancel footer. Enter runs OK
  // (except on buttons, which handle Enter themselves, and in multi-line
  // fields, where Cmd/Ctrl+Enter does); Escape cancels. `notice` shows Excel's
  // message box over the dialog — validation errors and results — and keeps
  // the dialog open behind it.
  import { onMount, tick, type Snippet } from 'svelte';
  import { getEditor } from '../core/context.ts';
  import { t } from '../i18n/i18n.svelte.ts';

  let {
    title,
    onok,
    oncancel,
    okLabel,
    cancelLabel,
    okDisabled = false,
    showCancel = true,
    modal = true,
    width = 420,
    notice = $bindable(null),
    children,
    footer,
    footerStart,
  }: {
    title: string;
    /** Return false to keep the dialog open (e.g. after setting `notice`). */
    onok?: () => boolean | void;
    /** Return false to stay open (a dialog stepping back from a sub-page). */
    oncancel?: () => boolean | void;
    okLabel?: string;
    cancelLabel?: string;
    okDisabled?: boolean;
    showCancel?: boolean;
    /** Non-modal dialogs (Find and Replace) leave the sheet clickable behind them. */
    modal?: boolean;
    width?: number;
    notice?: string | null;
    children: Snippet;
    /** Replaces the OK/Cancel buttons entirely. */
    footer?: Snippet;
    /** Extra buttons on the left of the footer (Clear All, Options…). */
    footerStart?: Snippet;
  } = $props();

  const ctl = getEditor();
  // The dialog this frame belongs to; closing must not clobber a dialog an OK handler opened.
  const own = ctl.dialog;
  let box: HTMLDivElement;
  let noticeButton = $state<HTMLButtonElement>();
  let pos = $state<{ left: number; top: number } | null>(null);
  const titleId = `xl-dlg-${Math.random().toString(36).slice(2)}`;

  function close(): void {
    if (ctl.dialog === own) ctl.closeDialog();
  }

  export function ok(): void {
    if (okDisabled || notice) return;
    const result = onok?.();
    if (result !== false) close();
  }

  export function cancel(): void {
    if (oncancel?.() !== false) close();
  }

  onMount(() => {
    const r = box.getBoundingClientRect();
    pos = { left: Math.max(8, (window.innerWidth - r.width) / 2), top: Math.max(8, Math.min(window.innerHeight * 0.16, window.innerHeight - r.height - 8)) };
    const first =
      box.querySelector<HTMLElement>('.body [data-autofocus]') ??
      box.querySelector<HTMLElement>('.body input:not([disabled]):not([type="hidden"]), .body select:not([disabled]), .body textarea:not([disabled]), .body button:not([disabled])');
    // A message-only dialog focuses its default button so Enter and Escape still work.
    // A disabled button cannot take focus; fall back to the dialog itself so Escape still reaches it.
    (first ?? box.querySelector<HTMLElement>('.footer button:not([disabled])') ?? box).focus();
    if (first instanceof HTMLInputElement && first.type === 'text') first.select();
  });

  // Where the keyboard was when the notice opened; it goes back there when the notice closes.
  let beforeNotice: HTMLElement | null = null;
  $effect(() => {
    if (!notice) return;
    const active = document.activeElement;
    beforeNotice = active instanceof HTMLElement && box.contains(active) && active !== noticeButton ? active : null;
    void tick().then(() => noticeButton?.focus());
  });

  function dismissNotice(): void {
    notice = null;
    void tick().then(() => (beforeNotice?.isConnected ? beforeNotice : box).focus());
  }

  function onkeydown(e: KeyboardEvent): void {
    // Keys typed in a dialog never reach the grid's shortcuts.
    e.stopPropagation();
    if (e.isComposing) return;
    if (e.key === 'Escape') {
      e.preventDefault();
      if (notice) dismissNotice();
      else cancel();
      return;
    }
    if (e.key !== 'Enter' || e.shiftKey || e.altKey) return;
    const target = e.target as HTMLElement;
    // Multi-line fields take Enter as a newline; Cmd/Ctrl+Enter still confirms.
    if (target instanceof HTMLTextAreaElement && !(e.metaKey || e.ctrlKey)) return;
    if (target instanceof HTMLButtonElement || (target instanceof HTMLSelectElement && target.multiple)) return;
    e.preventDefault();
    if (notice) dismissNotice();
    else ok();
  }

  function startDrag(e: PointerEvent): void {
    if (e.button !== 0 || !pos) return;
    const start = { x: e.clientX, y: e.clientY, left: pos.left, top: pos.top };
    const bar = e.currentTarget as HTMLElement;
    bar.setPointerCapture(e.pointerId);
    const move = (ev: PointerEvent) => {
      pos = {
        left: Math.min(window.innerWidth - 60, Math.max(60 - box.offsetWidth, start.left + ev.clientX - start.x)),
        top: Math.min(window.innerHeight - 30, Math.max(0, start.top + ev.clientY - start.y)),
      };
    };
    const up = () => {
      bar.removeEventListener('pointermove', move);
      bar.removeEventListener('pointerup', up);
    };
    bar.addEventListener('pointermove', move);
    bar.addEventListener('pointerup', up);
  }
</script>

{#if modal}
  <div class="backdrop" aria-hidden="true"></div>
{/if}
<div
  class="dlg"
  bind:this={box}
  role="dialog"
  aria-modal={modal}
  aria-labelledby={titleId}
  tabindex="-1"
  style:width="{width}px"
  style:left={pos ? `${pos.left}px` : '50%'}
  style:top={pos ? `${pos.top}px` : '16%'}
  style:transform={pos ? 'none' : 'translateX(-50%)'}
  {onkeydown}
>
  <div class="titlebar" onpointerdown={startDrag} role="presentation">
    <span class="lights" aria-hidden="true"><button class="light close" tabindex="-1" onclick={cancel} onpointerdown={(e) => e.stopPropagation()} aria-label={t('dlgClose')}></button><i></i><i></i></span>
    <span class="title" id={titleId}>{title}</span>
  </div>
  <div class="body">
    {@render children()}
  </div>
  <div class="footer">
    {#if footer}
      {@render footer()}
    {:else}
      <div class="start">{@render footerStart?.()}</div>
      {#if showCancel}<button class="xl-btn outlined" onclick={cancel}>{cancelLabel ?? t('dlgCancel')}</button>{/if}
      <button class="xl-btn primary" disabled={okDisabled} onclick={ok}>{okLabel ?? t('dlgOk')}</button>
    {/if}
  </div>
  {#if notice}
    <div class="notice-layer">
      <div class="notice" role="alertdialog" aria-label={notice}>
        <div class="notice-icon" aria-hidden="true">!</div>
        <p>{notice}</p>
        <button class="xl-btn primary" bind:this={noticeButton} onclick={dismissNotice}>{t('dlgOk')}</button>
      </div>
    </div>
  {/if}
</div>

<style>
  .backdrop {
    position: fixed;
    inset: 0;
    z-index: 900;
    background: rgba(0, 0, 0, 0.08);
  }
  .dlg {
    position: fixed;
    z-index: 901;
    max-width: calc(100vw - 16px);
    max-height: calc(100vh - 16px);
    display: flex;
    flex-direction: column;
    background: #ececec;
    border: 1px solid rgba(0, 0, 0, 0.25);
    border-radius: 10px;
    box-shadow: 0 18px 48px rgba(0, 0, 0, 0.28), 0 2px 6px rgba(0, 0, 0, 0.12);
    font-size: 12.5px;
    color: var(--xl-text);
    outline: none;
  }
  .titlebar {
    position: relative;
    display: flex;
    align-items: center;
    justify-content: center;
    height: 28px;
    flex: none;
    cursor: default;
    user-select: none;
    touch-action: none;
    border-bottom: 1px solid rgba(0, 0, 0, 0.08);
  }
  .title {
    font-weight: 600;
    font-size: 12.5px;
    color: var(--xl-text);
  }
  .lights {
    position: absolute;
    left: 10px;
    top: 8px;
    display: flex;
    gap: 7px;
  }
  .lights i,
  .light {
    display: block;
    width: 12px;
    height: 12px;
    border-radius: 50%;
    background: #d3d3d3;
    border: 1px solid rgba(0, 0, 0, 0.12);
    padding: 0;
  }
  .light.close {
    background: #ff5f57;
    cursor: pointer;
  }
  .body {
    padding: 12px 16px 4px;
    overflow: auto;
    min-height: 0;
  }
  .footer {
    display: flex;
    align-items: center;
    justify-content: flex-end;
    gap: 8px;
    padding: 10px 16px 14px;
    flex: none;
  }
  .footer .start {
    margin-right: auto;
    display: flex;
    gap: 8px;
  }
  .footer :global(.xl-btn) {
    min-width: 72px;
  }
  .notice-layer {
    position: absolute;
    inset: 0;
    display: flex;
    align-items: center;
    justify-content: center;
    background: rgba(236, 236, 236, 0.6);
    border-radius: 10px;
  }
  .notice {
    width: min(320px, 90%);
    background: #f6f6f6;
    border: 1px solid rgba(0, 0, 0, 0.2);
    border-radius: 10px;
    box-shadow: var(--xl-shadow-lg);
    padding: 16px;
    display: flex;
    flex-direction: column;
    align-items: center;
    gap: 10px;
    text-align: center;
  }
  .notice p {
    margin: 0;
    white-space: pre-line;
    line-height: 1.45;
  }
  .notice-icon {
    width: 36px;
    height: 36px;
    border-radius: 8px;
    background: var(--xl-accent);
    color: #fff;
    font-weight: 700;
    font-size: 20px;
    display: flex;
    align-items: center;
    justify-content: center;
  }
  .notice .xl-btn {
    min-width: 96px;
  }

  /* Shared form layout for dialog bodies. */
  /* Tab strip as Excel for Mac draws it: a gray track with the selected segment raised in white. */
  .body :global(.seg) {
    display: flex;
    justify-content: center;
    margin: 0 auto 10px;
    width: fit-content;
    padding: 2px;
    background: rgba(0, 0, 0, 0.08);
    border-radius: 6px;
  }
  .body :global(.seg > button) {
    border: 0;
    background: transparent;
    border-radius: 5px;
    padding: 3px 12px;
    font: inherit;
    font-size: 12px;
    color: var(--xl-text);
    cursor: default;
  }
  .body :global(.seg > button[aria-selected='true']) {
    background: #fff;
    box-shadow: 0 0.5px 2px rgba(0, 0, 0, 0.25);
  }
  .body :global(.row) {
    display: flex;
    align-items: center;
    gap: 8px;
    margin: 6px 0;
  }
  .body :global(.row > label:first-child),
  .body :global(.lbl) {
    min-width: 96px;
    color: var(--xl-text);
  }
  .body :global(fieldset) {
    border: 1px solid rgba(0, 0, 0, 0.14);
    border-radius: 6px;
    padding: 6px 10px 8px;
    margin: 8px 0;
    background: rgba(255, 255, 255, 0.45);
  }
  .body :global(legend) {
    padding: 0 4px;
    font-weight: 600;
    font-size: 12px;
  }
  .body :global(.check) {
    display: flex;
    align-items: center;
    gap: 6px;
    margin: 4px 0;
  }
  .body :global(.grow) {
    flex: 1;
    min-width: 0;
  }
  .body :global(.list) {
    border: 1px solid var(--xl-border-strong);
    background: #fff;
    border-radius: 3px;
    overflow: auto;
  }
  .body :global(.list button) {
    display: block;
    width: 100%;
    text-align: left;
    border: 0;
    background: transparent;
    padding: 1px 6px;
    line-height: 16px;
    font: inherit;
    color: inherit;
    cursor: default;
  }
  .body :global(.list button[aria-selected='true']) {
    background: var(--xl-pressed);
  }
  .body :global(.hint) {
    color: var(--xl-text-2);
    font-size: 11.5px;
    line-height: 1.4;
  }
  .body :global(input[type='number']) {
    width: 72px;
  }
</style>

<script lang="ts" module>
  // Only one reference box listens to the sheet at a time: the one focused last.
  let listening = $state<symbol | null>(null);
</script>

<script lang="ts">
  // Excel's RefEdit: a reference box that, once focused, takes the range the
  // user selects on the sheet. The dialog using it must be non-modal so the
  // grid stays clickable.
  import { untrack } from 'svelte';
  import { quoteSheetName, rangeAddress } from '../core/address.ts';
  import { getEditor } from '../core/context.ts';
  import { currentRange } from '../core/selection.ts';

  let {
    value = $bindable(''),
    id,
    label,
    withSheet = false,
    autofocus = false,
  }: { value?: string; id?: string; label?: string; withSheet?: boolean; autofocus?: boolean } = $props();

  const ctl = getEditor();
  const me = Symbol('ref');
  /** Sheet the dialog opened on; references to other sheets carry their sheet name. */
  const homeSheet = untrack(() => ctl.doc.activeSheetIndex);
  /** Selection when this box started listening; only later changes are taken. */
  let startSelection: unknown = null;

  function listen(): void {
    listening = me;
    startSelection = ctl.doc.selection;
  }

  $effect(() => {
    const sel = ctl.doc.selection;
    void ctl.doc.activeSheetIndex;
    untrack(() => {
      if (listening !== me || sel === startSelection) return;
      const range = currentRange(sel);
      const prefix = withSheet || ctl.doc.activeSheetIndex !== homeSheet ? `${quoteSheetName(ctl.doc.ws.title)}!` : '';
      value = prefix + rangeAddress(range, true);
    });
  });

  $effect(() => () => {
    if (listening === me) listening = null;
  });
</script>

<input
  {id}
  class="xl-input ref"
  class:listening={listening === me}
  bind:value
  spellcheck="false"
  autocomplete="off"
  aria-label={label}
  data-autofocus={autofocus ? '' : undefined}
  onfocus={listen}
/>

<style>
  .ref {
    font-family: var(--xl-mono);
    flex: 1;
    width: 100%;
    min-width: 0;
    box-sizing: border-box;
  }
  .listening {
    outline: 1px dashed var(--xl-accent);
    outline-offset: 1px;
  }
</style>

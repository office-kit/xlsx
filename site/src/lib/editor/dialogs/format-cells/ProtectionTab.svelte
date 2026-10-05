<script lang="ts">
  import { t } from '../../i18n/i18n.svelte.ts';
  import type { ProtectionState } from '../format-cells.ts';

  let { model = $bindable() }: { model: ProtectionState } = $props();
</script>

<label class="check"><input type="checkbox" bind:checked={model.locked} />{t('dlgLocked')}</label>
<label class="check"><input type="checkbox" bind:checked={model.hidden} />{t('dlgHidden')}</label>
<p class="hint">{t('dlgProtectionHint')}</p>

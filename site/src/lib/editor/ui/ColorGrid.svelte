<script lang="ts">
  // Excel's colour gallery: theme colours with their tint/shade rows, the ten
  // standard colours, "Automatic"/"No Fill", and a custom colour well.
  import { applyTint, type ThemePalette } from '../core/theme.ts';
  import { t } from '../i18n/i18n.svelte.ts';

  let { palette, onpick, noneLabel }: { palette: ThemePalette; onpick: (rgb: string | null) => void; noneLabel: string } = $props();

  // Columns in Excel's order: lt1, dk1, lt2, dk2, accent1-6.
  const ORDER = [0, 1, 2, 3, 4, 5, 6, 7, 8, 9];
  const TINTS = [0, 0.8, 0.6, 0.4, -0.25, -0.5];
  const DARK_TINTS = [0, 0.5, 0.35, 0.25, 0.15, 0.05];
  const STANDARD = ['C00000', 'FF0000', 'FFC000', 'FFFF00', '92D050', '00B050', '00B0F0', '0070C0', '002060', '7030A0'];

  function shade(base: string, row: number, col: number): string {
    // The two dark slots lighten in every row; the rest tint then shade.
    const isDark = col === 1 || col === 3;
    const tint = isDark ? DARK_TINTS[row] ?? 0 : col === 0 ? [0, -0.05, -0.15, -0.25, -0.35, -0.5][row] ?? 0 : TINTS[row] ?? 0;
    return applyTint(base, tint);
  }
</script>

<div class="cg">
  <button class="xl-menu-item none" onclick={() => onpick(null)}>{noneLabel}</button>
  <div class="section">{t('themeColors')}</div>
  <div class="grid">
    {#each [0, 1, 2, 3, 4, 5] as row (row)}
      {#each ORDER as col (col)}
        {@const hex = shade(palette[col] ?? '000000', row, col)}
        <button class="sw" class:gap={row === 0} style:background="#{hex}" title="#{hex}" aria-label="#{hex}" onclick={() => onpick(hex)}></button>
      {/each}
    {/each}
  </div>
  <div class="section">{t('standardColors')}</div>
  <div class="grid">
    {#each STANDARD as hex (hex)}
      <button class="sw" style:background="#{hex}" title="#{hex}" aria-label="#{hex}" onclick={() => onpick(hex)}></button>
    {/each}
  </div>
  <label class="xl-menu-item more">
    {t('moreColors')}
    <input type="color" onchange={(e) => onpick((e.currentTarget as HTMLInputElement).value.slice(1).toUpperCase())} />
  </label>
</div>

<style>
  .cg {
    padding: 2px 8px 6px;
    width: 236px;
  }
  .section {
    font-weight: 600;
    font-size: 11px;
    color: var(--xl-text-2);
    margin: 6px 0 4px;
  }
  .grid {
    display: grid;
    grid-template-columns: repeat(10, 18px);
    gap: 2px 4px;
  }
  .sw {
    width: 18px;
    height: 16px;
    border: 1px solid rgba(0, 0, 0, 0.15);
    padding: 0;
    cursor: pointer;
  }
  .sw:hover {
    outline: 2px solid #f5a623;
    outline-offset: 0;
  }
  .sw.gap {
    margin-bottom: 4px;
  }
  .none,
  .more {
    padding-left: 2px;
  }
  .more input {
    width: 28px;
    height: 18px;
    border: 0;
    padding: 0;
    margin-left: auto;
  }
</style>

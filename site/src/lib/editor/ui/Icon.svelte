<script lang="ts">
  // Stroke icons (24×24, currentColor) for the ribbon, menus and bars. Paths
  // are simple primitives so they stay legible at 16 px.
  let { name, size = 18 }: { name: string; size?: number } = $props();
</script>

<svg width={size} height={size} viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="1.6" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true">
  {#if name === 'paste'}
    <rect x="5" y="4" width="14" height="17" rx="2" /><rect x="9" y="2.5" width="6" height="3.5" rx="1" /><line x1="9" y1="11" x2="15" y2="11" /><line x1="9" y1="15" x2="15" y2="15" />
  {:else if name === 'cut'}
    <circle cx="6.5" cy="17.5" r="2.5" /><circle cx="17.5" cy="17.5" r="2.5" /><line x1="8.3" y1="15.7" x2="17" y2="4" /><line x1="15.7" y1="15.7" x2="7" y2="4" />
  {:else if name === 'copy'}
    <rect x="9" y="9" width="11" height="11" rx="2" /><path d="M5 15V5a2 2 0 0 1 2-2h8" />
  {:else if name === 'brush'}
    <path d="M14 4l6 6-7 7-6-6z" /><path d="M7 11l-3 3a2.5 2.5 0 0 0 0 3.5l1.5 1.5a2.5 2.5 0 0 0 3.5 0l3-3" />
  {:else if name === 'undo'}
    <path d="M9 14L4 9l5-5" /><path d="M4 9h10a6 6 0 0 1 0 12h-3" />
  {:else if name === 'redo'}
    <path d="M15 14l5-5-5-5" /><path d="M20 9H10a6 6 0 0 0 0 12h3" />
  {:else if name === 'save'}
    <path d="M5 3h11l3 3v13a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2z" /><rect x="7" y="3" width="8" height="5" /><rect x="7" y="13" width="10" height="8" />
  {:else if name === 'open'}
    <path d="M3 7a2 2 0 0 1 2-2h4l2 2h8a2 2 0 0 1 2 2v8a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2z" />
  {:else if name === 'new'}
    <path d="M6 3h8l4 4v14H6z" /><path d="M14 3v4h4" /><line x1="12" y1="11" x2="12" y2="17" /><line x1="9" y1="14" x2="15" y2="14" />
  {:else if name === 'print'}
    <rect x="6" y="3" width="12" height="6" /><rect x="3" y="9" width="18" height="8" rx="2" /><rect x="7" y="14" width="10" height="7" />
  {:else if name === 'bold'}
    <path d="M7 4h6a4 4 0 0 1 0 8H7z" stroke-width="2.4" /><path d="M7 12h7a4 4 0 0 1 0 8H7z" stroke-width="2.4" />
  {:else if name === 'italic'}
    <line x1="10" y1="4" x2="18" y2="4" /><line x1="6" y1="20" x2="14" y2="20" /><line x1="14" y1="4" x2="10" y2="20" />
  {:else if name === 'underline'}
    <path d="M7 4v7a5 5 0 0 0 10 0V4" /><line x1="5" y1="21" x2="19" y2="21" />
  {:else if name === 'strike'}
    <line x1="4" y1="12" x2="20" y2="12" /><path d="M16.5 7.5C16 5.5 14 4.5 12 4.5c-2.5 0-4.5 1.3-4.5 3.4 0 1.5.9 2.4 2.5 3" /><path d="M8 16.5c.6 2 2.4 3 4.4 3 2.6 0 4.6-1.4 4.6-3.6 0-1-.4-1.8-1.1-2.4" />
  {:else if name === 'font-grow'}
    <path d="M3 20l5-14h2l5 14" /><line x1="5.5" y1="14" x2="12.5" y2="14" /><polyline points="17,9 19.5,6 22,9" />
  {:else if name === 'font-shrink'}
    <path d="M3 20l5-14h2l5 14" /><line x1="5.5" y1="14" x2="12.5" y2="14" /><polyline points="17,6 19.5,9 22,6" />
  {:else if name === 'font-color'}
    <path d="M6 16L11 4h2l5 12" /><line x1="8" y1="11.5" x2="16" y2="11.5" />
  {:else if name === 'fill'}
    <path d="M5 11l6-6 7 7-6 6a2 2 0 0 1-3 0l-4-4a2 2 0 0 1 0-3z" /><path d="M19 14c1 1.5 2 2.5 2 3.5a2 2 0 1 1-4 0c0-1 1-2 2-3.5z" />
  {:else if name === 'borders'}
    <rect x="4" y="4" width="16" height="16" stroke-dasharray="2 2" /><line x1="4" y1="20" x2="20" y2="20" stroke-width="2.4" />
  {:else if name === 'align-top'}
    <line x1="4" y1="4" x2="20" y2="4" /><line x1="12" y1="8" x2="12" y2="20" /><polyline points="8,12 12,8 16,12" />
  {:else if name === 'align-middle'}
    <line x1="4" y1="12" x2="20" y2="12" /><polyline points="9,5 12,8 15,5" /><polyline points="9,19 12,16 15,19" /><line x1="12" y1="2" x2="12" y2="8" /><line x1="12" y1="16" x2="12" y2="22" />
  {:else if name === 'align-bottom'}
    <line x1="4" y1="20" x2="20" y2="20" /><line x1="12" y1="4" x2="12" y2="16" /><polyline points="8,12 12,16 16,12" />
  {:else if name === 'align-left'}
    <line x1="4" y1="6" x2="20" y2="6" /><line x1="4" y1="10" x2="14" y2="10" /><line x1="4" y1="14" x2="20" y2="14" /><line x1="4" y1="18" x2="14" y2="18" />
  {:else if name === 'align-center'}
    <line x1="4" y1="6" x2="20" y2="6" /><line x1="7" y1="10" x2="17" y2="10" /><line x1="4" y1="14" x2="20" y2="14" /><line x1="7" y1="18" x2="17" y2="18" />
  {:else if name === 'align-right'}
    <line x1="4" y1="6" x2="20" y2="6" /><line x1="10" y1="10" x2="20" y2="10" /><line x1="4" y1="14" x2="20" y2="14" /><line x1="10" y1="18" x2="20" y2="18" />
  {:else if name === 'wrap'}
    <line x1="4" y1="6" x2="20" y2="6" /><path d="M4 12h13a3 3 0 0 1 0 6h-4" /><polyline points="15,16 13,18 15,20" /><line x1="4" y1="18" x2="9" y2="18" />
  {:else if name === 'merge'}
    <rect x="3" y="6" width="18" height="12" /><polyline points="7,10 10,12 7,14" /><polyline points="17,10 14,12 17,14" />
  {:else if name === 'indent-in'}
    <line x1="4" y1="6" x2="20" y2="6" /><line x1="10" y1="12" x2="20" y2="12" /><line x1="4" y1="18" x2="20" y2="18" /><polyline points="4,9.5 7,12 4,14.5" />
  {:else if name === 'indent-out'}
    <line x1="4" y1="6" x2="20" y2="6" /><line x1="10" y1="12" x2="20" y2="12" /><line x1="4" y1="18" x2="20" y2="18" /><polyline points="7,9.5 4,12 7,14.5" />
  {:else if name === 'orientation'}
    <path d="M5 19L15 5" /><path d="M9 19l2-5" /><polyline points="14,9 15,5 11,6" />
  {:else if name === 'currency'}
    <path d="M16 7.5c-.6-1.6-2.2-2.5-4-2.5-2.3 0-4 1.2-4 3s1.6 2.6 4 3 4 1.3 4 3-1.7 3-4 3c-1.9 0-3.5-.9-4-2.5" /><line x1="12" y1="3" x2="12" y2="21" />
  {:else if name === 'percent'}
    <line x1="19" y1="5" x2="5" y2="19" /><circle cx="7" cy="7" r="2.3" /><circle cx="17" cy="17" r="2.3" />
  {:else if name === 'comma'}
    <text x="4" y="17" font-size="13" fill="currentColor" stroke="none" font-family="sans-serif">,000</text>
  {:else if name === 'dec-inc'}
    <text x="2" y="11" font-size="8" fill="currentColor" stroke="none" font-family="sans-serif">.0</text><text x="9" y="20" font-size="8" fill="currentColor" stroke="none" font-family="sans-serif">.00</text><polyline points="3,16 6,19 3,22" stroke-width="1.3" />
  {:else if name === 'dec-dec'}
    <text x="2" y="11" font-size="8" fill="currentColor" stroke="none" font-family="sans-serif">.00</text><text x="11" y="20" font-size="8" fill="currentColor" stroke="none" font-family="sans-serif">.0</text><polyline points="7,16 4,19 7,22" stroke-width="1.3" />
  {:else if name === 'cond-format'}
    <rect x="3" y="4" width="18" height="16" rx="1" /><line x1="3" y1="9.5" x2="21" y2="9.5" /><line x1="3" y1="15" x2="21" y2="15" /><rect x="5" y="11" width="8" height="2.5" fill="currentColor" stroke="none" /><rect x="5" y="16.5" width="12" height="2.5" fill="currentColor" stroke="none" />
  {:else if name === 'table'}
    <rect x="3" y="4" width="18" height="16" rx="1" /><line x1="3" y1="9" x2="21" y2="9" /><line x1="3" y1="14.5" x2="21" y2="14.5" /><line x1="9" y1="4" x2="9" y2="20" /><line x1="15" y1="4" x2="15" y2="20" />
  {:else if name === 'cell-styles'}
    <rect x="3" y="5" width="8" height="6" /><rect x="13" y="5" width="8" height="6" fill="currentColor" opacity="0.3" /><rect x="3" y="13" width="8" height="6" fill="currentColor" opacity="0.6" /><rect x="13" y="13" width="8" height="6" />
  {:else if name === 'insert'}
    <rect x="3" y="4" width="18" height="16" rx="1" /><line x1="12" y1="8" x2="12" y2="16" /><line x1="8" y1="12" x2="16" y2="12" />
  {:else if name === 'delete'}
    <rect x="3" y="4" width="18" height="16" rx="1" /><line x1="8.5" y1="8.5" x2="15.5" y2="15.5" /><line x1="15.5" y1="8.5" x2="8.5" y2="15.5" />
  {:else if name === 'format'}
    <rect x="3" y="4" width="18" height="16" rx="1" /><line x1="3" y1="10" x2="21" y2="10" /><line x1="10" y1="4" x2="10" y2="20" /><path d="M14 14l3 3m0-3l-3 3" />
  {:else if name === 'sum'}
    <path d="M18 5H6l6 7-6 7h12" />
  {:else if name === 'fill-down'}
    <line x1="12" y1="4" x2="12" y2="18" /><polyline points="7,13 12,18 17,13" /><line x1="5" y1="21" x2="19" y2="21" />
  {:else if name === 'eraser'}
    <path d="M16 4l5 5-10 10H6l-3-3z" /><line x1="10" y1="10" x2="15" y2="15" /><line x1="11" y1="21" x2="21" y2="21" />
  {:else if name === 'sort'}
    <path d="M7 4v16" /><polyline points="4,17 7,20 10,17" /><text x="12" y="10" font-size="8" fill="currentColor" stroke="none" font-family="sans-serif">A</text><text x="12" y="20" font-size="8" fill="currentColor" stroke="none" font-family="sans-serif">Z</text>
  {:else if name === 'sort-asc'}
    <path d="M7 4v16" /><polyline points="4,17 7,20 10,17" /><text x="12" y="10" font-size="8" fill="currentColor" stroke="none" font-family="sans-serif">A</text><text x="12" y="20" font-size="8" fill="currentColor" stroke="none" font-family="sans-serif">Z</text>
  {:else if name === 'sort-desc'}
    <path d="M7 4v16" /><polyline points="4,17 7,20 10,17" /><text x="12" y="10" font-size="8" fill="currentColor" stroke="none" font-family="sans-serif">Z</text><text x="12" y="20" font-size="8" fill="currentColor" stroke="none" font-family="sans-serif">A</text>
  {:else if name === 'filter'}
    <path d="M3 5h18l-7 8v6l-4 2v-8z" />
  {:else if name === 'search'}
    <circle cx="10.5" cy="10.5" r="6.5" /><line x1="15.5" y1="15.5" x2="21" y2="21" />
  {:else if name === 'chart'}
    <line x1="4" y1="20" x2="20" y2="20" /><rect x="6" y="11" width="3" height="9" /><rect x="11" y="6" width="3" height="14" /><rect x="16" y="14" width="3" height="6" />
  {:else if name === 'chart-line'}
    <line x1="4" y1="20" x2="20" y2="20" /><polyline points="4,16 9,10 13,13 20,5" />
  {:else if name === 'chart-pie'}
    <path d="M12 3a9 9 0 1 0 9 9h-9z" /><path d="M14 3.2A9 9 0 0 1 20.8 10H14z" />
  {:else if name === 'image'}
    <rect x="3" y="5" width="18" height="14" rx="2" /><circle cx="8.5" cy="10" r="1.6" /><polyline points="21,16 15,11 6,19" />
  {:else if name === 'link'}
    <path d="M10 14a4 4 0 0 0 5.7 0l3-3a4 4 0 0 0-5.7-5.7l-1 1" /><path d="M14 10a4 4 0 0 0-5.7 0l-3 3a4 4 0 0 0 5.7 5.7l1-1" />
  {:else if name === 'note'}
    <path d="M4 4h16v12l-4 4H4z" /><path d="M16 20v-4h4" />
  {:else if name === 'function'}
    <path d="M14 4c-2 0-3 1-3.5 3L8 18c-.5 2-1.5 2.5-3 2.5" /><line x1="7" y1="10" x2="15" y2="10" /><path d="M15 13l5 6m0-6l-5 6" />
  {:else if name === 'name'}
    <rect x="3" y="7" width="18" height="10" rx="2" /><line x1="7" y1="12" x2="17" y2="12" />
  {:else if name === 'freeze'}
    <rect x="3" y="4" width="18" height="16" rx="1" /><line x1="3" y1="9" x2="21" y2="9" stroke-width="2.4" /><line x1="8" y1="4" x2="8" y2="20" stroke-width="2.4" />
  {:else if name === 'zoom-in'}
    <circle cx="10.5" cy="10.5" r="6.5" /><line x1="15.5" y1="15.5" x2="21" y2="21" /><line x1="10.5" y1="7.5" x2="10.5" y2="13.5" /><line x1="7.5" y1="10.5" x2="13.5" y2="10.5" />
  {:else if name === 'navigation'}
    <rect x="3" y="4" width="18" height="16" rx="1.5" /><line x1="9" y1="4" x2="9" y2="20" /><line x1="4.5" y1="8" x2="7.5" y2="8" /><line x1="4.5" y1="11" x2="7.5" y2="11" /><line x1="4.5" y1="14" x2="7.5" y2="14" />
  {:else if name === 'watch'}
    <rect x="3" y="5" width="18" height="14" rx="1.5" /><circle cx="9" cy="12" r="2.2" /><circle cx="15" cy="12" r="2.2" /><line x1="11.2" y1="12" x2="12.8" y2="12" />
  {:else if name === 'crosshair'}
    <rect x="3" y="3" width="18" height="18" rx="1" /><rect x="3" y="9.5" width="18" height="5" fill="currentColor" opacity=".25" stroke="none" /><rect x="9.5" y="3" width="5" height="18" fill="currentColor" opacity=".25" stroke="none" /><rect x="9.5" y="9.5" width="5" height="5" />
  {:else if name === 'grid'}
    <rect x="3" y="3" width="18" height="18" rx="1" /><line x1="3" y1="9" x2="21" y2="9" /><line x1="3" y1="15" x2="21" y2="15" /><line x1="9" y1="3" x2="9" y2="21" /><line x1="15" y1="3" x2="15" y2="21" />
  {:else if name === 'validation'}
    <rect x="3" y="4" width="18" height="16" rx="1" /><polyline points="7,12 10,15 17,8" />
  {:else if name === 'duplicates'}
    <rect x="3" y="5" width="12" height="5" /><rect x="3" y="12" width="12" height="5" /><line x1="18" y1="8" x2="22" y2="12" /><line x1="22" y1="8" x2="18" y2="12" />
  {:else if name === 'text-columns'}
    <rect x="3" y="5" width="7" height="14" /><rect x="14" y="5" width="7" height="14" /><polyline points="10,12 14,12" />
  {:else if name === 'lock'}
    <rect x="5" y="11" width="14" height="10" rx="2" /><path d="M8 11V8a4 4 0 0 1 8 0v3" />
  {:else if name === 'page-layout'}
    <rect x="2" y="4" width="9" height="16" rx="1" /><rect x="13" y="4" width="9" height="16" rx="1" /><line x1="4" y1="8" x2="9" y2="8" /><line x1="15" y1="8" x2="20" y2="8" />
  {:else if name === 'page'}
    <path d="M6 3h8l4 4v14H6z" /><path d="M14 3v4h4" />
  {:else if name === 'margins'}
    <rect x="5" y="3" width="14" height="18" /><rect x="8" y="6" width="8" height="12" stroke-dasharray="2 2" />
  {:else if name === 'landscape'}
    <rect x="3" y="6" width="18" height="12" rx="1" />
  {:else if name === 'portrait'}
    <rect x="6" y="3" width="12" height="18" rx="1" />
  {:else if name === 'chevron-down'}
    <polyline points="6,9 12,15 18,9" />
  {:else if name === 'chevron-up'}
    <polyline points="6,15 12,9 18,15" />
  {:else if name === 'chevron-left'}
    <polyline points="15,6 9,12 15,18" />
  {:else if name === 'chevron-right'}
    <polyline points="9,6 15,12 9,18" />
  {:else if name === 'plus'}
    <line x1="12" y1="5" x2="12" y2="19" /><line x1="5" y1="12" x2="19" y2="12" />
  {:else if name === 'minus'}
    <line x1="5" y1="12" x2="19" y2="12" />
  {:else if name === 'close'}
    <line x1="6" y1="6" x2="18" y2="18" /><line x1="18" y1="6" x2="6" y2="18" />
  {:else if name === 'check'}
    <polyline points="5,12 10,17 19,7" />
  {:else if name === 'more'}
    <circle cx="6" cy="12" r="1" /><circle cx="12" cy="12" r="1" /><circle cx="18" cy="12" r="1" />
  {:else if name === 'group'}
    <path d="M7 4H4v16h3" /><line x1="10" y1="8" x2="20" y2="8" /><line x1="10" y1="12" x2="20" y2="12" /><line x1="10" y1="16" x2="20" y2="16" />
  {:else if name === 'ungroup'}
    <line x1="4" y1="8" x2="20" y2="8" /><line x1="4" y1="12" x2="20" y2="12" /><line x1="4" y1="16" x2="20" y2="16" />
  {:else if name === 'calc'}
    <rect x="5" y="3" width="14" height="18" rx="2" /><rect x="8" y="6" width="8" height="3" /><line x1="8.5" y1="13" x2="9.5" y2="13" /><line x1="14.5" y1="13" x2="15.5" y2="13" /><line x1="8.5" y1="17" x2="9.5" y2="17" /><line x1="14.5" y1="17" x2="15.5" y2="17" />
  {:else if name === 'trace'}
    <circle cx="6" cy="6" r="2" /><circle cx="18" cy="18" r="2" /><path d="M8 7l8 9" /><polyline points="12.5,15.5 16,16 15.5,12.5" />
  {:else if name === 'spell'}
    <path d="M4 15l3.5-9h1L12 15" /><line x1="5.3" y1="12" x2="10.7" y2="12" /><polyline points="12,17 15,20 21,12" />
  {:else if name === 'sheet'}
    <rect x="4" y="3" width="16" height="18" rx="1" /><line x1="4" y1="9" x2="20" y2="9" /><line x1="10" y1="3" x2="10" y2="21" />
  {:else if name === 'shapes'}
    <circle cx="8" cy="8" r="4" /><rect x="12" y="12" width="8" height="8" />
  {:else if name === 'textbox'}
    <rect x="3" y="6" width="18" height="12" rx="1" stroke-dasharray="3 2" /><line x1="12" y1="9" x2="12" y2="15" /><line x1="9.5" y1="9" x2="14.5" y2="9" />
  {:else if name === 'symbol'}
    <path d="M7 19h4v-2a6 6 0 1 1 2 0v2h4" />
  {:else if name === 'gridlines'}
    <rect x="3" y="3" width="18" height="18" stroke-dasharray="2 2" /><line x1="3" y1="12" x2="21" y2="12" stroke-dasharray="2 2" /><line x1="12" y1="3" x2="12" y2="21" stroke-dasharray="2 2" />
  {:else if name === 'headings'}
    <rect x="3" y="3" width="18" height="18" /><rect x="3" y="3" width="18" height="4" fill="currentColor" opacity="0.3" /><rect x="3" y="3" width="4" height="18" fill="currentColor" opacity="0.3" />
  {:else if name === 'formula-bar'}
    <rect x="3" y="7" width="18" height="10" rx="1" /><text x="6" y="15" font-size="7" fill="currentColor" stroke="none" font-family="serif" font-style="italic">fx</text>
  {:else if name === 'show-formulas'}
    <text x="3" y="16" font-size="10" fill="currentColor" stroke="none" font-family="serif" font-style="italic">fx</text><line x1="14" y1="9" x2="21" y2="9" /><line x1="14" y1="14" x2="21" y2="14" />
  {:else if name === 'recent'}
    <rect x="4" y="3" width="12" height="16" /><circle cx="16.5" cy="16.5" r="4.5" /><polyline points="16.5,14.5 16.5,16.5 18,17.5" />
  {:else if name === 'chart-column'}
    <rect x="4" y="11" width="4" height="9" fill="currentColor" fill-opacity="0.25" /><rect x="10" y="5" width="4" height="15" fill="var(--xl-accent)" fill-opacity="0.5" /><rect x="16" y="9" width="4" height="11" fill="currentColor" fill-opacity="0.25" /><line x1="2" y1="20" x2="22" y2="20" />
  {:else if name === 'chart-scatter'}
    <line x1="3" y1="3" x2="3" y2="21" /><line x1="3" y1="21" x2="21" y2="21" /><circle cx="8" cy="15" r="1.4" fill="currentColor" /><circle cx="12" cy="10" r="1.4" fill="var(--xl-accent)" /><circle cx="16" cy="13" r="1.4" fill="currentColor" /><circle cx="18" cy="6" r="1.4" fill="var(--xl-accent)" />
  {:else if name === 'chart-stock'}
    <line x1="6" y1="4" x2="6" y2="18" /><rect x="4.5" y="8" width="3" height="6" fill="var(--xl-accent)" fill-opacity="0.5" /><line x1="12" y1="6" x2="12" y2="20" /><rect x="10.5" y="10" width="3" height="7" fill="currentColor" fill-opacity="0.25" /><line x1="18" y1="3" x2="18" y2="15" /><rect x="16.5" y="5" width="3" height="6" fill="var(--xl-accent)" fill-opacity="0.5" />
  {:else if name === 'filter-clear'}
    <path d="M3 4h14l-5.5 7v6l-3 2v-8z" /><line x1="16" y1="15" x2="21" y2="20" /><line x1="21" y1="15" x2="16" y2="20" />
  {:else if name === 'filter-reapply'}
    <path d="M3 4h14l-5.5 7v6l-3 2v-8z" /><path d="M21 15a3.5 3.5 0 1 1-1-2.5" /><polyline points="20.5,10.5 20.5,13 18,13" />
  {:else if name === 'filter-advanced'}
    <path d="M3 4h14l-5.5 7v6l-3 2v-8z" /><circle cx="18.5" cy="17.5" r="2.5" /><line x1="18.5" y1="13" x2="18.5" y2="15" /><line x1="18.5" y1="20" x2="18.5" y2="22" />
  {:else if name === 'show-detail'}
    <line x1="10" y1="6" x2="21" y2="6" /><line x1="10" y1="12" x2="21" y2="12" /><line x1="10" y1="18" x2="21" y2="18" /><line x1="3" y1="12" x2="7" y2="12" /><line x1="5" y1="10" x2="5" y2="14" />
  {:else if name === 'hide-detail'}
    <line x1="10" y1="6" x2="21" y2="6" /><line x1="10" y1="12" x2="21" y2="12" /><line x1="10" y1="18" x2="21" y2="18" /><line x1="3" y1="12" x2="7" y2="12" />
  {:else if name === 'flash-fill'}
    <rect x="3" y="4" width="9" height="16" /><line x1="3" y1="9" x2="12" y2="9" /><line x1="3" y1="14" x2="12" y2="14" /><path d="M18 3l-3 8h4l-3 8" />
  {:else if name === 'consolidate'}
    <rect x="3" y="3" width="7" height="7" /><rect x="3" y="14" width="7" height="7" /><rect x="15" y="8" width="6" height="8" /><path d="M10 6.5h2.5v11H10M12.5 12H15" />
  {:else if name === 'define-name'}
    <path d="M4 7h11l5 5-5 5H4z" /><line x1="7" y1="12" x2="12" y2="12" />
  {:else if name === 'use-in-formula'}
    <text x="2" y="17" font-size="11" fill="currentColor" stroke="none" font-family="serif" font-style="italic">fx</text><path d="M13 8h5l3 4-3 4h-5z" />
  {:else if name === 'create-from-selection'}
    <rect x="3" y="4" width="18" height="16" /><line x1="3" y1="9" x2="21" y2="9" /><line x1="9" y1="4" x2="9" y2="20" /><path d="M13 13h5l2 2.5-2 2.5h-5z" />
  {:else if name === 'trace-dependents'}
    <rect x="3" y="4" width="6" height="5" /><rect x="15" y="15" width="6" height="5" /><path d="M9 6.5c5 0 8 3 9 8.5" /><polyline points="15.8,12.8 18,15 20,12.6" />
  {:else if name === 'remove-arrows'}
    <path d="M4 7h10" /><polyline points="11,4 14,7 11,10" /><line x1="14" y1="14" x2="20" y2="20" /><line x1="20" y1="14" x2="14" y2="20" />
  {:else if name === 'error-check'}
    <path d="M12 3l9 17H3z" /><line x1="12" y1="9" x2="12" y2="14" /><line x1="12" y1="17" x2="12" y2="17.5" />
  {:else if name === 'evaluate'}
    <text x="2" y="16" font-size="10" fill="currentColor" stroke="none" font-family="serif" font-style="italic">fx</text><circle cx="17" cy="11" r="3.5" /><line x1="19.5" y1="13.5" x2="22" y2="16" />
  {:else if name === 'calc-now'}
    <rect x="4" y="3" width="13" height="18" rx="1.5" /><line x1="7" y1="7" x2="14" y2="7" /><path d="M21 13a3.5 3.5 0 1 1-1-2.5" /><polyline points="20.5,8.5 20.5,11 18,11" />
  {:else if name === 'calc-sheet'}
    <rect x="3" y="4" width="13" height="16" /><line x1="3" y1="9" x2="16" y2="9" /><line x1="8" y1="4" x2="8" y2="20" /><path d="M22 15a3.5 3.5 0 1 1-1-2.5" /><polyline points="21.5,10.5 21.5,13 19,13" />
  {:else if name === 'comment'}
    <path d="M4 4h16v12H10l-5 4v-4H4z" />
  {:else if name === 'comment-new'}
    <path d="M4 6h12v10H10l-4 3v-3H4z" /><line x1="19" y1="2" x2="19" y2="8" /><line x1="16" y1="5" x2="22" y2="5" />
  {:else if name === 'comment-delete'}
    <path d="M4 6h12v10H10l-4 3v-3H4z" /><line x1="16" y1="2" x2="22" y2="8" /><line x1="22" y1="2" x2="16" y2="8" />
  {:else if name === 'comment-prev'}
    <path d="M8 6h12v10h-6l-4 3v-3H8z" /><polyline points="5,8 2,11 5,14" />
  {:else if name === 'comment-next'}
    <path d="M4 6h12v10h-6l-4 3v-3H4z" /><polyline points="19,8 22,11 19,14" />
  {:else if name === 'spelling'}
    <text x="2" y="11" font-size="9" fill="currentColor" stroke="none">abc</text><polyline points="5,16 9,20 19,10" />
  {:else if name === 'thesaurus'}
    <path d="M3 5h7a2 2 0 0 1 2 2v13a2 2 0 0 0-2-2H3z" /><path d="M21 5h-7a2 2 0 0 0-2 2v13a2 2 0 0 1 2-2h7z" />
  {:else if name === 'accessibility'}
    <circle cx="12" cy="4.5" r="1.8" /><path d="M5 8l7 1.5L19 8" /><path d="M12 9.5V14l-3 6M12 14l3 6" />
  {:else if name === 'protect-sheet'}
    <rect x="3" y="3" width="12" height="16" /><line x1="3" y1="8" x2="15" y2="8" /><line x1="8" y1="3" x2="8" y2="19" /><rect x="14" y="14" width="8" height="7" rx="1" /><path d="M16 14v-2a2 2 0 0 1 4 0v2" />
  {:else if name === 'protect-workbook'}
    <rect x="3" y="3" width="12" height="16" /><line x1="3" y1="8" x2="15" y2="8" /><line x1="3" y1="13" x2="15" y2="13" /><rect x="14" y="14" width="8" height="7" rx="1" /><path d="M16 14v-2a2 2 0 0 1 4 0v2" />
  {:else if name === 'page-break'}
    <rect x="5" y="3" width="14" height="18" /><line x1="3" y1="12" x2="21" y2="12" stroke-dasharray="2 2" />
  {:else if name === 'background'}
    <rect x="3" y="4" width="18" height="16" /><path d="M3 17l5-6 4 4 3-3 6 6" /><circle cx="16" cy="8.5" r="1.5" />
  {:else if name === 'header-footer'}
    <rect x="5" y="3" width="14" height="18" /><line x1="5" y1="7" x2="19" y2="7" /><line x1="5" y1="17" x2="19" y2="17" />
  {:else if name === 'recommended-charts'}
    <rect x="3" y="12" width="3.5" height="8" /><rect x="9" y="8" width="3.5" height="12" /><rect x="15" y="13" width="3.5" height="7" /><path d="M17 3.5c2 0 3 1 3 2.5s-1.5 2-2.5 2.5v1" /><line x1="17.5" y1="10.5" x2="17.5" y2="11" />
  {:else if name === 'sparkline-line'}
    <polyline points="3,16 8,10 12,13 16,6 21,9" /><line x1="3" y1="20" x2="21" y2="20" />
  {:else if name === 'sparkline-column'}
    <rect x="4" y="11" width="3" height="9" /><rect x="10.5" y="6" width="3" height="14" /><rect x="17" y="9" width="3" height="11" />
  {:else if name === 'sparkline-winloss'}
    <rect x="4" y="6" width="3" height="6" /><rect x="10.5" y="12" width="3" height="6" /><rect x="17" y="6" width="3" height="6" /><line x1="3" y1="12" x2="21" y2="12" />
  {:else if name === 'new-window'}
    <rect x="3" y="5" width="14" height="14" /><line x1="3" y1="9" x2="17" y2="9" /><line x1="20" y1="2" x2="20" y2="8" /><line x1="17" y1="5" x2="23" y2="5" />
  {:else if name === 'unhide'}
    <rect x="3" y="5" width="18" height="14" /><line x1="3" y1="9" x2="21" y2="9" /><path d="M8 14c1.3-1.5 2.6-2.2 4-2.2s2.7.7 4 2.2c-1.3 1.5-2.6 2.2-4 2.2s-2.7-.7-4-2.2z" />
  {:else if name === 'custom-views'}
    <rect x="3" y="4" width="18" height="16" /><line x1="3" y1="9" x2="21" y2="9" /><line x1="8" y1="9" x2="8" y2="20" /><line x1="11" y1="13" x2="18" y2="13" /><line x1="11" y1="16" x2="18" y2="16" />
  {:else}
    <circle cx="12" cy="12" r="3" />
  {/if}
</svg>

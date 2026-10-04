<script lang="ts">
  // Mirror of the in-cell editor's text with formula references coloured to
  // match their outlines on the grid. The textarea above it keeps the caret
  // and IME; its own text is transparent.
  let { text, refs }: { text: string; refs: ReadonlyArray<{ start: number; end: number; color: string }> } = $props();

  const parts = $derived.by(() => {
    const out: Array<{ text: string; color?: string }> = [];
    let pos = 0;
    for (const ref of [...refs].sort((a, b) => a.start - b.start)) {
      if (ref.start < pos) continue;
      if (ref.start > pos) out.push({ text: text.slice(pos, ref.start) });
      out.push({ text: text.slice(ref.start, ref.end), color: ref.color });
      pos = ref.end;
    }
    if (pos < text.length) out.push({ text: text.slice(pos) });
    // A trailing newline needs a character after it to occupy a line.
    if (text.endsWith('\n')) out.push({ text: ' ' });
    return out;
  });
</script>

<div class="mirror" aria-hidden="true">{#each parts as part, i (i)}{#if part.color}<span style:color={part.color}>{part.text}</span>{:else}{part.text}{/if}{/each}</div>

<style>
  .mirror {
    white-space: pre-wrap;
    word-break: break-all;
    line-height: 1.25;
    min-height: 100%;
    padding-right: 6px;
  }
</style>

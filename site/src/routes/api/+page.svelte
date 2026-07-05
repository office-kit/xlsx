<script lang="ts">
  import { base } from '$app/paths';
  import type { PageProps } from './$types';

  const { data }: PageProps = $props();
</script>

<svelte:head>
  <title>API reference — @office-kit/xlsx</title>
</svelte:head>

<div class="content">
  <p class="eyebrow">Reference</p>
  <h1>API reference</h1>
  <p class="lede">
    Every public export of <code>@office-kit/xlsx</code>, <code>@office-kit/xlsx/streaming</code>, and
    <code>@office-kit/xlsx/node</code>, organized by section. Generated from the source via typedoc
    and rendered inline so types stay in lockstep with the package.
  </p>

  <div class="grid">
    {#each data.sections as section, i (section.id)}
      <a href="{base}/api/{section.id}" class="card">
        <span class="card-num">{String(i + 1).padStart(2, '0')}</span>
        <div class="card-head">
          <h3>{section.title}</h3>
          <span class="count">{section.itemCount}</span>
        </div>
        <p>{section.description}</p>
      </a>
    {/each}
  </div>
</div>

<style>
  .content {
    max-width: var(--max-wide);
    padding: 2.25rem 1.5rem 5rem;
  }

  h1 {
    margin: 0 0 0.6rem;
  }

  .lede {
    color: var(--fg-soft);
    font-size: 1.05rem;
    margin: 0 0 2.25rem;
    max-width: 720px;
  }

  .grid {
    display: grid;
    grid-template-columns: repeat(auto-fill, minmax(260px, 1fr));
    border: 1px solid var(--border);
    border-radius: var(--radius);
    overflow: hidden;
  }

  .card {
    position: relative;
    display: block;
    padding: 1.15rem 1.25rem 1.25rem;
    background: var(--bg-elev);
    border-bottom: 1px solid var(--border);
    color: var(--fg);
    transition: background 140ms ease;
  }

  .card:hover {
    text-decoration: none;
    background: var(--bg-soft);
  }

  .card:hover h3 {
    color: var(--accent);
  }

  .card-num {
    display: block;
    font-family: var(--mono);
    font-size: 11px;
    font-weight: 500;
    letter-spacing: 0.1em;
    color: var(--accent);
    margin-bottom: 0.5rem;
  }

  .card-head {
    display: flex;
    align-items: baseline;
    justify-content: space-between;
    gap: 0.5rem;
    margin-bottom: 0.4rem;
  }

  .card h3 {
    margin: 0;
    font-family: var(--display);
    font-size: 1.12rem;
    font-weight: 540;
    line-height: 1.2;
    color: var(--fg);
    transition: color 140ms ease;
    font-variation-settings: 'opsz' 32, 'SOFT' 25;
  }

  .card p {
    margin: 0;
    color: var(--fg-soft);
    font-size: 0.92rem;
    line-height: 1.5;
  }

  .count {
    font-family: var(--mono);
    font-size: 11px;
    color: var(--fg-muted);
    padding: 0.05em 0.4em;
    background: var(--bg-paper);
    border: 1px solid var(--border);
    border-radius: 3px;
    font-variant-numeric: tabular-nums;
    letter-spacing: 0.04em;
  }
</style>

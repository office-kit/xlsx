<script lang="ts">
  import ApiItem from '$lib/components/ApiItem.svelte';
  import type { PageProps } from './$types';

  const { data }: PageProps = $props();
</script>

<svelte:head>
  <title>{data.section.title} — API — @office-kit/xlsx</title>
</svelte:head>

<div class="content">
  <header class="section-head">
    <p class="eyebrow">API reference</p>
    <h1>{data.section.title}</h1>
    <p class="lede">{data.section.description}</p>
    <p class="meta">
      {data.section.itemCount} export{data.section.itemCount === 1 ? '' : 's'} ·
      {data.subgroups.length} source file{data.subgroups.length === 1 ? '' : 's'}
    </p>
  </header>

  <nav class="toc" aria-label="On this page" data-pagefind-ignore>
    <h4>On this page</h4>
    {#each data.subgroups as group (group.id)}
      <div class="toc-group">
        <a class="toc-group-link" href="#{group.id}">{group.label}</a>
        <ul>
          {#each group.items as item (item.id)}
            <li>
              <a href="#{group.id}-{item.name}">
                <span class="kind-dot kind-{item.kind}"></span>
                <span class="t-name">{item.name}</span>
              </a>
            </li>
          {/each}
        </ul>
      </div>
    {/each}
  </nav>

  <div class="items">
    {#each data.subgroups as group (group.id)}
      <section class="subgroup" id={group.id}>
        <header class="subgroup-head">
          <h2>{group.label}</h2>
          <a class="subfile" href="https://github.com/office-kit/xlsx/blob/main/{group.sourceFile}" target="_blank" rel="noopener">
            {group.sourceFile}
          </a>
        </header>
        {#each group.items as item (item.id)}
          <ApiItem {item} anchorId="{group.id}-{item.name}" />
        {/each}
      </section>
    {/each}
  </div>
</div>

<style>
  .content {
    display: grid;
    grid-template-columns: minmax(0, 1fr) 240px;
    grid-template-areas:
      'head head'
      'items toc';
    column-gap: 2rem;
    max-width: 1280px;
    padding: 2rem 1.5rem 5rem;
  }

  .section-head {
    grid-area: head;
    margin-bottom: 1.5rem;
  }

  .items {
    grid-area: items;
    min-width: 0;
  }

  .toc {
    grid-area: toc;
    position: sticky;
    top: calc(var(--header-h) + 1.5rem);
    align-self: start;
    max-height: calc(100vh - var(--header-h) - 3rem);
    overflow-y: auto;
    font-size: 13px;
    padding-left: 1rem;
  }

  h1 {
    margin: 0 0 0.4rem;
  }

  .lede {
    color: var(--fg-soft);
    margin: 0;
    max-width: 640px;
  }

  .meta {
    color: var(--fg-muted);
    font-size: 0.85rem;
    margin: 0.6rem 0 0;
  }

  .toc h4 {
    font-size: 13px;
    text-transform: uppercase;
    letter-spacing: 0.08em;
    color: var(--fg-muted);
    margin: 0 0 0.6rem;
    border: none;
  }

  .toc-group {
    margin-bottom: 1rem;
  }

  .toc-group-link {
    display: block;
    font-weight: 600;
    color: var(--fg);
    padding: 0.2rem 0.4rem;
    border-radius: 4px;
    margin-bottom: 0.15rem;
  }

  .toc-group-link:hover {
    background: var(--bg-soft);
    text-decoration: none;
  }

  .toc ul {
    list-style: none;
    padding: 0 0 0 0.7rem;
    margin: 0;
  }

  .toc li {
    margin-bottom: 0.05rem;
  }

  .toc a {
    display: flex;
    align-items: center;
    gap: 0.45rem;
    padding: 0.18rem 0.4rem;
    border-radius: 4px;
    color: var(--fg-soft);
    line-height: 1.3;
  }

  .toc a:hover {
    color: var(--fg);
    background: var(--bg-soft);
    text-decoration: none;
  }

  .t-name {
    font-family: var(--mono);
    font-size: 13px;
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
  }

  .kind-dot {
    display: inline-block;
    width: 6px;
    height: 6px;
    border-radius: 50%;
    flex: none;
  }
  .kind-function {
    background: #a5d6a7;
  }
  .kind-interface {
    background: #90caf9;
  }
  .kind-type {
    background: #ce93d8;
  }
  .kind-class {
    background: #ffab91;
  }
  .kind-variable {
    background: #ffd54f;
  }

  .subgroup {
    margin-bottom: 3rem;
    scroll-margin-top: calc(var(--header-h) + 1rem);
  }

  .subgroup-head {
    display: flex;
    align-items: baseline;
    justify-content: space-between;
    gap: 1rem;
    flex-wrap: wrap;
    border-bottom: 2px solid var(--accent);
    padding-bottom: 0.5rem;
    margin-bottom: 0.5rem;
  }

  .subgroup-head h2 {
    margin: 0;
    border: none;
    padding: 0;
    font-size: 1.65rem;
  }

  .subfile {
    font-family: var(--mono);
    font-size: 13px;
    color: var(--fg-muted);
  }

  .subfile:hover {
    color: var(--accent);
  }

  @media (max-width: 1000px) {
    .content {
      grid-template-columns: 1fr;
      grid-template-areas:
        'head'
        'toc'
        'items';
    }
    .toc {
      position: static;
      max-height: none;
      padding-left: 0;
      border-top: 1px solid var(--border);
      border-bottom: 1px solid var(--border);
      padding: 0.8rem 0;
    }
  }
</style>

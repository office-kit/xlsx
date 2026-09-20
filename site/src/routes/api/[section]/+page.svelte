<script lang="ts">
  import ApiItem from '$lib/components/ApiItem.svelte';
  import type { PageProps } from './$types';

  const { data }: PageProps = $props();

  const REPO_BLOB = 'https://github.com/office-kit/xlsx/blob/main';
  const plural = (n: number, word: string): string => `${n} ${word}${n === 1 ? '' : 's'}`;
</script>

<svelte:head>
  <title>{data.section.title} · API · @office-kit/xlsx</title>
</svelte:head>

<div class="content">
  <header>
    <h1>{data.section.title}</h1>
    <p class="lede">{data.section.description}</p>
    <p class="meta">
      {plural(data.section.itemCount, 'export')} from {plural(data.subgroups.length, 'source file')}
    </p>
    <!-- The full index only fits beside the content; narrower screens get the files. -->
    <nav class="jump" aria-label="Source files on this page" data-pagefind-ignore>
      {#each data.subgroups as group (group.id)}
        <a href="#{group.id}">{group.label}</a>
      {/each}
    </nav>
  </header>

  <div class="items">
    {#each data.subgroups as group (group.id)}
      <section class="subgroup" id={group.id}>
        <header class="subgroup-head">
          <h2>{group.label}</h2>
          <a href="{REPO_BLOB}/{group.sourceFile}" target="_blank" rel="noopener">
            {group.sourceFile}
          </a>
        </header>
        {#each group.items as item (item.id)}
          <ApiItem {item} anchorId="{group.id}-{item.name}" />
        {/each}
      </section>
    {/each}
  </div>

  <nav class="toc" aria-label="On this page" data-pagefind-ignore>
    <h2>On this page</h2>
    {#each data.subgroups as group (group.id)}
      <a class="toc-group" href="#{group.id}">{group.label}</a>
      <ul>
        {#each group.items as item (item.id)}
          <li><a href="#{group.id}-{item.name}">{item.name}</a></li>
        {/each}
      </ul>
    {/each}
  </nav>
</div>

<style>
  .content {
    display: grid;
    grid-template-columns: minmax(0, 1fr) 13rem;
    grid-template-areas:
      'head head'
      'items toc';
    column-gap: 2.5rem;
  }

  .content > header {
    grid-area: head;
  }

  .lede {
    max-width: 68ch;
    margin: 0;
    color: var(--ink-2);
    font-size: 1.05rem;
  }

  .meta {
    margin: 0.6rem 0 0;
    color: var(--ink-3);
    font-size: 0.88rem;
  }

  .jump {
    display: none;
    flex-wrap: wrap;
    gap: 0.4rem;
    margin: 1.5rem 0 0;
  }

  .jump a {
    padding: 0.35rem 0.75rem;
    border: 1px solid var(--line);
    border-radius: 999px;
    color: var(--ink-2);
    font-size: 0.9rem;
    font-weight: 500;
  }

  .jump a:hover {
    color: var(--ink);
    border-color: var(--ink-3);
    text-decoration: none;
  }

  .items {
    grid-area: items;
    min-width: 0;
  }

  .subgroup {
    margin-top: 3rem;
  }

  .subgroup-head {
    display: flex;
    flex-wrap: wrap;
    align-items: baseline;
    justify-content: space-between;
    gap: 0.25rem 1rem;
    padding-bottom: 0.6rem;
    border-bottom: 1px solid var(--line-strong);
  }

  .subgroup-head h2 {
    margin: 0;
    font-size: 1.4rem;
  }

  .subgroup-head a {
    color: var(--ink-3);
    font-family: var(--mono);
    font-size: 0.8rem;
    overflow-wrap: anywhere;
  }

  .subgroup-head a:hover {
    color: var(--accent-ink);
  }

  .toc {
    grid-area: toc;
    position: sticky;
    top: calc(var(--header-h) + 1.5rem);
    align-self: start;
    max-height: calc(100vh - var(--header-h) - 3rem);
    margin-top: 3rem;
    overflow-y: auto;
  }

  .toc h2 {
    margin: 0 0 0.5rem;
    font-family: var(--sans);
    font-size: 0.85rem;
    font-weight: 600;
    letter-spacing: 0;
    color: var(--ink-3);
  }

  .toc-group {
    display: block;
    margin-top: 0.9rem;
    color: var(--ink);
    font-size: 0.88rem;
    font-weight: 600;
  }

  .toc ul {
    list-style: none;
    margin: 0.25rem 0 0;
    padding: 0;
    border-left: 1px solid var(--line);
  }

  .toc li {
    margin: 0;
  }

  .toc li a {
    display: block;
    padding: 0.2rem 0 0.2rem 0.8rem;
    color: var(--ink-2);
    font-family: var(--mono);
    font-size: 0.78rem;
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
  }

  .toc a:hover {
    color: var(--accent-ink);
    text-decoration: none;
  }

  @media (max-width: 1100px) {
    .content {
      display: block;
    }

    .toc {
      display: none;
    }

    .jump {
      display: flex;
    }
  }
</style>

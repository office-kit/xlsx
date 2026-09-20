<script lang="ts">
  import { base } from '$app/paths';
  import CodeBlock from '$lib/components/CodeBlock.svelte';
  import type { PageProps } from './$types';

  const { data }: PageProps = $props();

  const groupId = (title: string): string => 'group-' + title.toLowerCase().replace(/\s+/g, '-');
</script>

<svelte:head>
  <title>Recipes · @office-kit/xlsx</title>
</svelte:head>

<h1>Recipes</h1>

<p class="lede">
  Working code for the tasks people ask about most. Every snippet is a real file under
  <code>site/src/lib/examples/</code> that is type-checked against the library on every build, so
  an API rename breaks this page before it ships. For a one-line lookup use the
  <a href="{base}/docs/cheatsheet">cheatsheet</a>; for a specific function, the
  <a href="{base}/api">API reference</a>.
</p>

<nav class="jump" aria-label="Recipe groups">
  {#each data.groups as group (group.title)}
    <a href="#{groupId(group.title)}">{group.title}</a>
  {/each}
</nav>

{#each data.groups as group (group.title)}
  <section class="group" id={groupId(group.title)}>
    <h2>{group.title}</h2>

    {#each group.recipes as r (r.slug)}
      <section class="recipe" id={r.slug}>
        <h3><a href="#{r.slug}">{r.title}</a></h3>
        <p>{r.teaser}</p>
        <CodeBlock html={r.html} title={r.path} />
        {#if r.notes?.length}
          <ul class="notes">
            {#each r.notes as n (n)}
              <li>{n}</li>
            {/each}
          </ul>
        {/if}
        {#if r.relatedApi?.length}
          <p class="related">
            Related API:
            {#each r.relatedApi as name, k (name)}
              <code>{name}</code>{#if k < r.relatedApi.length - 1},{' '}{/if}
            {/each}
          </p>
        {/if}
      </section>
    {/each}
  </section>
{/each}

<p class="more">
  To drive a workbook end to end, walk through
  <a href="{base}/docs/getting-started">Getting started</a>. To experiment without installing
  anything, open <a href="{base}/repl">the REPL</a>; to inspect a real file, open
  <a href="{base}/playground">the playground</a>.
</p>

<style>
  .lede {
    color: var(--ink-2);
    font-size: 1.08rem;
  }

  .jump {
    display: flex;
    flex-wrap: wrap;
    gap: 0.4rem;
    margin: 1.75rem 0 0;
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

  .group > h2 {
    margin-top: 3.5rem;
    padding-bottom: 0.6rem;
    border-bottom: 1px solid var(--line);
  }

  .recipe {
    margin-top: 2.5rem;
  }

  .recipe h3 {
    margin: 0 0 0.4rem;
  }

  .recipe h3 a {
    color: inherit;
  }

  .recipe > p {
    margin: 0;
    color: var(--ink-2);
  }

  .notes {
    color: var(--ink-2);
    font-size: 0.95rem;
  }

  .related {
    font-size: 0.9rem;
    line-height: 1.9;
  }

  .more {
    margin-top: 3.5rem;
    padding-top: 1.5rem;
    border-top: 1px solid var(--line);
    color: var(--ink-2);
  }
</style>

<script lang="ts">
  import { base } from '$app/paths';
  import type { PageProps } from './$types';

  const { data }: PageProps = $props();

  const total = $derived(data.sections.reduce((sum, section) => sum + section.itemCount, 0));
</script>

<svelte:head>
  <title>API reference · @office-kit/xlsx</title>
</svelte:head>

<h1>API reference</h1>
<p class="lede">
  The {total} functions, classes, and constants that <code>@office-kit/xlsx</code> exports, in
  {data.sections.length} sections. It is generated from the source with typedoc on every build, so
  the signatures cannot drift from the package. Types are left to your editor.
</p>
<p class="lede">
  For the conceptual map see the <a href="{base}/docs/api">API overview</a>. For code you can paste,
  see the <a href="{base}/docs/cheatsheet">cheatsheet</a> and
  <a href="{base}/docs/recipes">recipes</a>.
</p>

<ul class="sections">
  {#each data.sections as section (section.id)}
    <li>
      <a href="{base}/api/{section.id}">
        <span class="title">{section.title}</span>
        <span class="count">{section.itemCount} export{section.itemCount === 1 ? '' : 's'}</span>
        <span class="description">{section.description}</span>
      </a>
    </li>
  {/each}
</ul>

<style>
  .lede {
    max-width: 68ch;
    color: var(--ink-2);
    font-size: 1.05rem;
  }

  .sections {
    list-style: none;
    margin: 2.5rem 0 0;
    padding: 0;
    border-top: 1px solid var(--line);
    columns: 2;
    column-gap: 3rem;
  }

  .sections li {
    margin: 0;
    border-bottom: 1px solid var(--line);
    break-inside: avoid;
  }

  .sections a {
    display: grid;
    grid-template-columns: minmax(0, 1fr) auto;
    gap: 0.2rem 1rem;
    align-items: baseline;
    padding: 1rem 0 1.1rem;
    color: var(--ink);
  }

  .sections a:hover {
    text-decoration: none;
  }

  .title {
    font-family: var(--display);
    font-size: 1.12rem;
    font-weight: 600;
    letter-spacing: -0.01em;
  }

  .sections a:hover .title {
    color: var(--accent-ink);
    text-decoration: underline;
    text-decoration-thickness: 1px;
    text-underline-offset: 0.2em;
  }

  .count {
    color: var(--ink-3);
    font-family: var(--mono);
    font-size: 0.78rem;
    font-variant-numeric: tabular-nums;
  }

  .description {
    grid-column: 1 / -1;
    color: var(--ink-2);
    font-size: 0.93rem;
    line-height: 1.5;
  }

  @media (max-width: 1000px) {
    .sections {
      columns: 1;
    }
  }
</style>

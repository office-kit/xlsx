<script lang="ts">
  import { base } from '$app/paths';
  import CodeBlock from '$lib/components/CodeBlock.svelte';
  import type { PageProps } from './$types';

  const { data }: PageProps = $props();

  // Flatten groups so we can stamp a continuous "01, 02, …" coordinate
  // across the whole page (matches word-kit's recipes layout).
  const flat = $derived(
    data.groups.flatMap((g) => g.recipes.map((r) => ({ ...r, group: g.title }))),
  );

  function indexOf(slug: string): number {
    return flat.findIndex((r) => r.slug === slug);
  }
</script>

<svelte:head>
  <title>Recipes · @office-kit/xlsx</title>
</svelte:head>

<article class="prose">
  <p class="eyebrow">§ 02 · Recipes</p>
  <h1>Common scenarios, type-checked snippets.</h1>

  <p class="lede">
    Every snippet below lives under <code>site/src/lib/examples/</code> and is type-checked by
    <code>svelte-check</code> against the live <code>@office-kit/xlsx</code> surface on every build —
    an API rename breaks this page before anything ships. Need a one-line lookup? See the
    <a href="{base}/docs/cheatsheet">Cheatsheet</a>. Looking up a specific function? Jump to
    the <a href="{base}/api">API reference</a>.
  </p>

  {#each data.groups as group (group.title)}
    <section
      class="group"
      id={'group-' + group.title.toLowerCase().replace(/\s+/g, '-')}
    >
      <h2>{group.title}</h2>

      {#each group.recipes as r (r.slug)}
        {@const i = indexOf(r.slug)}
        <section class="recipe" id={r.slug}>
          <header class="r-head">
            <span class="r-num">{String(i + 1).padStart(2, '0')}</span>
            <div class="r-text">
              <h3>
                <a href="#{r.slug}" class="hash" aria-label="Permalink">#</a>
                {r.title}
              </h3>
              <p>{r.teaser}</p>
              <p class="r-pointer">
                <span class="label">where:</span>
                <code>{r.path}</code>
              </p>
            </div>
          </header>
          <CodeBlock
            html={r.html}
            title={r.path}
            coord={String.fromCharCode(64 + ((i % 26) + 1)) + '1'}
          />
          {#if r.notes?.length}
            <ul class="notes">
              {#each r.notes as n (n)}
                <li>{n}</li>
              {/each}
            </ul>
          {/if}
          {#if r.relatedApi?.length}
            <p class="related">
              <span class="related-label">Related API:</span>
              {#each r.relatedApi as name, k (name)}
                <code>{name}</code>{#if k < r.relatedApi.length - 1}, {/if}
              {/each}
            </p>
          {/if}
        </section>
      {/each}
    </section>
  {/each}

  <p class="more">
    Want to drive a workbook end-to-end? Walk the
    <a href="{base}/docs/getting-started">Getting started</a> page, then open
    <a href="{base}/playground">the playground</a> to inspect a real file.
  </p>
</article>

<style>
  .prose {
    max-width: var(--max-content);
    margin: 0 auto;
    padding: 3rem 1.5rem 5rem;
  }

  .eyebrow {
    font-family: var(--mono);
    font-size: 11.5px;
    color: var(--fg-muted);
    text-transform: uppercase;
    letter-spacing: 0.14em;
    margin: 0 0 0.85rem;
  }

  h1 {
    font-family: var(--display);
    font-weight: 460;
    font-size: clamp(2rem, 4.6vw, 2.95rem);
    line-height: 1.05;
    letter-spacing: -0.026em;
    margin: 0 0 1rem;
    font-variation-settings: 'opsz' 144, 'SOFT' 30;
    max-width: 22ch;
    border: none;
    padding: 0;
  }

  .lede {
    color: var(--fg-soft);
    font-size: 1.06rem;
    line-height: 1.55;
    max-width: 64ch;
    margin: 0 0 2rem;
  }

  .group {
    margin-top: 3rem;
  }

  .group > h2 {
    font-family: var(--display);
    font-weight: 500;
    font-size: 1.55rem;
    letter-spacing: -0.015em;
    margin: 0 0 0.5rem;
    padding: 0 0 0.5rem;
    border-bottom: 1px solid var(--border);
    font-variation-settings: 'opsz' 96, 'SOFT' 25;
  }

  .recipe {
    margin: 2.25rem 0 2.75rem;
    padding-top: 1.75rem;
    border-top: 1px solid var(--border);
    scroll-margin-top: calc(var(--header-h) + 1rem);
  }

  .group > h2 + .recipe {
    border-top: none;
    padding-top: 1.25rem;
  }

  .r-head {
    display: grid;
    grid-template-columns: 4ch 1fr;
    gap: 1.25rem;
    margin: 0 0 0.75rem;
  }

  .r-num {
    font-family: var(--mono);
    font-size: 11.5px;
    color: var(--accent);
    font-weight: 500;
    letter-spacing: 0.06em;
    margin-top: 0.55rem;
  }

  .r-text h3 {
    margin: 0 0 0.5rem;
    border: none;
    padding: 0;
    font-family: var(--display);
    font-size: 1.4rem;
    font-weight: 540;
    line-height: 1.2;
    font-variation-settings: 'opsz' 64, 'SOFT' 25;
  }

  .hash {
    color: var(--fg-muted);
    font-weight: 400;
    margin-right: 0.25rem;
    text-decoration: none;
    visibility: hidden;
  }

  .r-text h3:hover .hash {
    visibility: visible;
  }

  .r-text > p {
    margin: 0;
    color: var(--fg-soft);
    font-size: 1rem;
    line-height: 1.55;
  }

  .r-pointer {
    margin-top: 0.55rem !important;
    font-family: var(--mono);
    font-size: 11.5px;
    color: var(--fg-muted) !important;
  }

  .r-pointer .label {
    text-transform: uppercase;
    letter-spacing: 0.12em;
    margin-right: 0.4rem;
    color: var(--fg-faint);
  }

  .notes {
    margin: 1rem 0 0.4rem;
    padding-left: 1.2rem;
    color: var(--fg-soft);
    font-size: 0.92rem;
  }

  .notes li {
    margin-bottom: 0.35rem;
  }

  .related {
    margin: 0.9rem 0 0;
    color: var(--fg-muted);
    font-size: 0.85rem;
  }

  .related-label {
    color: var(--fg-muted);
    margin-right: 0.4rem;
  }

  .related code {
    margin-right: 0.15rem;
  }

  .more {
    margin-top: 3rem;
    border-top: 1px solid var(--border);
    padding-top: 1.25rem;
    color: var(--fg-soft);
    font-size: 0.95rem;
  }
</style>

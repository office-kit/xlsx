<script lang="ts">
  import type { ApiKind } from '$lib/api/types';
  import type { RenderedItem } from '../../routes/api/[section]/+page.server';

  type Props = {
    item: RenderedItem;
    /** Anchor id for the heading; usually `${subgroup}-${item.name}` so
     *  that same-name exports from different files don't collide. */
    anchorId: string;
  };

  const { item, anchorId }: Props = $props();

  const kindLabel: Record<ApiKind, string> = {
    function: 'function',
    class: 'class',
    variable: 'const',
  };

  // Doc comments are hard-wrapped at the source's line length and mark code
  // with backticks. Reflow each paragraph to the page's measure and set the
  // code in <code>; a paragraph holding a list keeps its line breaks.
  const LIST_LINE = /^\s*(?:[-*]|\d+\.)\s/m;
  const blocks = $derived(
    item.description.split(/\n{2,}/).map((text) => ({
      verbatim: LIST_LINE.test(text),
      // Odd positions are the backticked spans.
      spans: text.split(/`([^`]+)`/),
    })),
  );
</script>

<section id={anchorId} class="item">
  <header>
    <h3>
      <a href="#{anchorId}">{item.name}</a>
      <span class="kind">{kindLabel[item.kind]}</span>
    </h3>
    <a href={item.sourceUrl} target="_blank" rel="noopener" class="source">
      {item.sourceFile}:{item.sourceLine}
    </a>
  </header>

  {#if item.description}
    {#each blocks as block, i (i)}
      <p class="description" class:verbatim={block.verbatim}>
        {#each block.spans as span, j (j)}{#if j % 2 === 1}<code>{span}</code>{:else}{span}{/if}{/each}
      </p>
    {/each}
  {/if}

  <div class="signature">{@html item.signatureHtml}</div>

  {#if item.parameters?.length}
    <h4>Parameters</h4>
    <div class="table-scroll">
      <table>
        <thead>
          <tr>
            <th scope="col">Name</th>
            <th scope="col">Type</th>
            <th scope="col">Description</th>
          </tr>
        </thead>
        <tbody>
          {#each item.parameters as p (p.name)}
            <tr>
              <td>
                <code>{p.name}{p.optional ? '?' : ''}</code>
                {#if p.defaultValue}
                  <span class="default">= {p.defaultValue}</span>
                {/if}
              </td>
              <td><code>{p.type}</code></td>
              <td>{p.description ?? ''}</td>
            </tr>
          {/each}
        </tbody>
      </table>
    </div>
  {/if}

  {#if item.returnType && item.kind === 'function'}
    <h4>Returns</h4>
    <p class="returns">
      <code>{item.returnType}</code>{item.returnDescription ? ` — ${item.returnDescription}` : ''}
    </p>
  {/if}
</section>

<style>
  .item {
    padding: 1.75rem 0 1.5rem;
    border-bottom: 1px solid var(--line);
  }

  header {
    display: flex;
    flex-wrap: wrap;
    align-items: baseline;
    justify-content: space-between;
    gap: 0.25rem 1rem;
    margin-bottom: 0.6rem;
  }

  h3 {
    display: flex;
    flex-wrap: wrap;
    align-items: baseline;
    gap: 0.25rem 0.6rem;
    min-width: 0;
    margin: 0;
    font-family: var(--mono);
    font-size: 1.08rem;
    font-weight: 600;
    letter-spacing: 0;
  }

  h3 a {
    color: var(--ink);
    overflow-wrap: anywhere;
  }

  .kind {
    color: var(--ink-3);
    font-size: 0.78rem;
    font-weight: 400;
  }

  .source {
    color: var(--ink-3);
    font-family: var(--mono);
    font-size: 0.78rem;
    overflow-wrap: anywhere;
  }

  .source:hover {
    color: var(--accent-ink);
  }

  .description {
    max-width: 72ch;
    margin: 0 0 0.8rem;
    color: var(--ink-2);
  }

  .description.verbatim {
    white-space: pre-wrap;
  }

  .signature :global(pre) {
    margin: 0.5rem 0 0;
  }

  h4 {
    margin: 1.4rem 0 0.25rem;
    font-family: var(--sans);
    font-size: 0.85rem;
    font-weight: 600;
    color: var(--ink-2);
  }

  table {
    min-width: 28rem;
    margin: 0;
    font-size: 0.9rem;
  }

  .default {
    margin-left: 0.4em;
    color: var(--ink-3);
    font-family: var(--mono);
    font-size: 0.8rem;
  }

  .returns {
    margin: 0.25rem 0 0;
    font-size: 0.95rem;
  }
</style>

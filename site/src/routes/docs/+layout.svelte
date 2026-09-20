<script lang="ts">
  import { base } from '$app/paths';
  import { page } from '$app/state';
  import Sidebar from '$lib/components/Sidebar.svelte';
  import SidebarShell from '$lib/components/SidebarShell.svelte';
  import { allDocLinks } from '$lib/docs-nav';

  type Props = {
    children?: import('svelte').Snippet;
  };

  const { children }: Props = $props();

  const mdHref = $derived(`${page.url.pathname.replace(/\/$/, '')}.md`);
  const currentRoute = $derived(
    page.url.pathname.replace(new RegExp(`^${base}`), '').replace(/\/$/, '') || '/',
  );
  const index = $derived(allDocLinks.findIndex((l) => l.href === currentRoute));
  const current = $derived(allDocLinks[index]);
  const prev = $derived(index > 0 ? allDocLinks[index - 1] : undefined);
  const next = $derived(index >= 0 ? allDocLinks[index + 1] : undefined);
</script>

<SidebarShell label={current?.title ?? 'Documentation'}>
  {#snippet nav()}
    <Sidebar />
  {/snippet}

  <article class="doc-content">
    {@render children?.()}
  </article>

  {#if current}
    <nav class="pager" aria-label="Previous and next page" data-pagefind-ignore>
      {#if prev}
        <a href="{base}{prev.href}" class="prev">
          <span>Previous</span>
          {prev.title}
        </a>
      {/if}
      {#if next}
        <a href="{base}{next.href}" class="next">
          <span>Next</span>
          {next.title}
        </a>
      {/if}
    </nav>
    <p class="md-link" data-pagefind-ignore>
      <a href={mdHref}>View this page as Markdown</a>. Models and tools can fetch that URL, or
      <a href="{base}/llms.txt">/llms.txt</a> for the full index.
    </p>
  {/if}
</SidebarShell>

<style>
  .doc-content,
  .pager,
  .md-link {
    max-width: var(--measure);
  }

  /* Wide content scrolls in its own box so a long table or signature never
   * pushes the page sideways on a phone. */
  .doc-content :global(table) {
    display: block;
    overflow-x: auto;
  }

  .doc-content :global(h2) {
    margin-top: 3rem;
  }

  .pager {
    display: grid;
    grid-template-columns: 1fr 1fr;
    gap: 0.75rem;
    margin-top: 3.5rem;
  }

  .pager a {
    display: flex;
    flex-direction: column;
    gap: 0.15rem;
    padding: 0.9rem 1.1rem;
    border: 1px solid var(--line);
    border-radius: var(--radius);
    color: var(--ink);
    font-weight: 600;
  }

  .pager a:hover {
    border-color: var(--accent);
    text-decoration: none;
  }

  .pager span {
    color: var(--ink-3);
    font-size: 0.82rem;
    font-weight: 450;
  }

  .pager .next {
    grid-column: 2;
    text-align: right;
  }

  .md-link {
    margin: 1.5rem 0 0;
    color: var(--ink-3);
    font-size: 0.88rem;
  }

  @media (max-width: 480px) {
    .pager {
      grid-template-columns: 1fr;
    }

    .pager .next {
      grid-column: 1;
      text-align: left;
    }
  }
</style>

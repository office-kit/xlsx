<script lang="ts">
  import { base } from '$app/paths';
  import { onMount } from 'svelte';

  type PagefindResult = {
    id: string;
    data: () => Promise<{
      url: string;
      excerpt: string;
      meta: { title?: string };
      sub_results?: Array<{ title: string; url: string; excerpt: string }>;
    }>;
  };

  type PagefindModule = {
    search(query: string): Promise<{ results: PagefindResult[] }>;
    options?(opts: { baseUrl?: string }): Promise<void>;
  };

  type Hit = {
    url: string;
    title: string;
    excerpt: string;
  };

  let open = $state(false);
  // eslint-disable-next-line prefer-const -- reassigned by `bind:value` in template
  let query = $state('');
  let hits = $state<Hit[]>([]);
  let pagefind = $state<PagefindModule | null>(null);
  let status = $state<'idle' | 'loading' | 'ready' | 'unavailable'>('idle');
  // eslint-disable-next-line prefer-const -- reassigned by `bind:this` in template
  let inputEl = $state<HTMLInputElement | null>(null);

  async function loadPagefind(): Promise<PagefindModule | null> {
    if (pagefind) return pagefind;
    if (status === 'unavailable') return null;
    status = 'loading';
    try {
      // SvelteKit's `paths.relative: true` means `base` is relative to the
      // current page (e.g. `..`/`../..`). A bare dynamic import would resolve
      // that against the JS chunk's URL, not the page's — so we anchor to
      // `document.baseURI` to land at the site-root pagefind directory.
      // Pagefind generates these files at build time; in `vite dev` they
      // don't exist, so the import fails and we surface "unavailable".
      const pagefindUrl = new URL(`${base}/pagefind/pagefind.js`, document.baseURI);
      const siteRoot = new URL(`${base}/`, document.baseURI);
      const mod = (await import(/* @vite-ignore */ pagefindUrl.href)) as PagefindModule;
      await mod.options?.({ baseUrl: siteRoot.pathname });
      pagefind = mod;
      status = 'ready';
      return mod;
    } catch (err) {
      console.error('[@office-kit/xlsx search] failed to load pagefind index', err);
      status = 'unavailable';
      return null;
    }
  }

  async function runSearch(q: string): Promise<void> {
    const trimmed = q.trim();
    if (!trimmed) {
      hits = [];
      return;
    }
    const pf = await loadPagefind();
    if (!pf) return;
    const search = await pf.search(trimmed);
    if (q !== query) return; // a newer keystroke landed; drop stale results
    const top = await Promise.all(search.results.slice(0, 10).map((r) => r.data()));
    hits = top.map((d) => ({
      url: normalizeRouteUrl(d.url),
      title: d.meta.title ?? d.url,
      excerpt: d.excerpt,
    }));
  }

  // Pagefind records the on-disk file path it indexed (e.g.
  // `/docs/streaming.html`), but SvelteKit's dev/preview servers only
  // resolve the canonical route (`/docs/streaming`). Strip the `.html`
  // suffix so result links work in dev and on the deployed static host
  // alike.
  function normalizeRouteUrl(url: string): string {
    return url.replace(/\/index\.html(?=$|\?|#)/, '/').replace(/\.html(?=$|\?|#)/, '');
  }

  $effect(() => {
    void runSearch(query);
  });

  function openModal(): void {
    open = true;
    void loadPagefind();
    queueMicrotask(() => inputEl?.focus());
  }

  function closeModal(): void {
    open = false;
  }

  function onKeyDown(e: KeyboardEvent): void {
    if ((e.metaKey || e.ctrlKey) && e.key === 'k') {
      e.preventDefault();
      if (open) closeModal();
      else openModal();
    } else if (e.key === 'Escape' && open) {
      closeModal();
    } else if (e.key === '/' && !open && !isTypingTarget(e.target)) {
      e.preventDefault();
      openModal();
    }
  }

  function isTypingTarget(t: EventTarget | null): boolean {
    if (!(t instanceof HTMLElement)) return false;
    const tag = t.tagName;
    return tag === 'INPUT' || tag === 'TEXTAREA' || t.isContentEditable;
  }

  onMount(() => {
    document.addEventListener('keydown', onKeyDown);
    return () => document.removeEventListener('keydown', onKeyDown);
  });
</script>

<button
  type="button"
  class="trigger"
  onclick={openModal}
  aria-label="Search docs (press / or Cmd-K)"
  data-pagefind-ignore
>
  <span class="trigger-label">Search</span>
  <kbd class="trigger-kbd">⌘K</kbd>
</button>

{#if open}
  <!-- svelte-ignore a11y_click_events_have_key_events -->
  <!-- svelte-ignore a11y_no_static_element_interactions -->
  <div class="overlay" role="presentation" onclick={closeModal} data-pagefind-ignore>
    <!-- svelte-ignore a11y_click_events_have_key_events -->
    <!-- svelte-ignore a11y_no_static_element_interactions -->
    <div
      class="modal"
      role="dialog"
      aria-modal="true"
      aria-label="Search"
      tabindex="-1"
      onclick={(e) => e.stopPropagation()}
    >
      <div class="input-row">
        <input
          bind:this={inputEl}
          bind:value={query}
          type="search"
          placeholder="Search docs and API…"
          autocomplete="off"
          spellcheck="false"
          class="input"
        />
        <kbd class="esc">Esc</kbd>
      </div>

      <div class="results" aria-live="polite">
        {#if status === 'unavailable'}
          <p class="status">
            Search index not available. Run <code>pnpm --filter @office-kit/xlsx-site build</code>
            to generate it (the dev server skips indexing).
          </p>
        {:else if status === 'loading' && hits.length === 0}
          <p class="status">Loading search index…</p>
        {:else if query.trim() && hits.length === 0 && status === 'ready'}
          <p class="status">No results for "{query}".</p>
        {:else if !query.trim()}
          <p class="status">
            Type to search across docs, recipes, and the full API reference.
          </p>
        {:else}
          <ul class="hits">
            {#each hits as hit (hit.url)}
              <li>
                <a href={hit.url} onclick={closeModal}>
                  <span class="hit-title">{hit.title}</span>
                  <span class="hit-excerpt">{@html hit.excerpt}</span>
                </a>
              </li>
            {/each}
          </ul>
        {/if}
      </div>
    </div>
  </div>
{/if}

<style>
  .trigger {
    display: inline-flex;
    align-items: center;
    gap: 0.5rem;
    padding: 0.32rem 0.55rem 0.32rem 0.7rem;
    background: var(--bg-soft);
    border: 1px solid var(--border);
    border-radius: var(--radius-sm);
    color: var(--fg-muted);
    font-family: var(--mono);
    font-size: 12px;
    font-weight: 500;
    letter-spacing: 0.04em;
    cursor: pointer;
    transition:
      color 120ms ease,
      border-color 120ms ease;
  }

  .trigger:hover {
    color: var(--fg);
    border-color: var(--accent-soft);
  }

  .trigger-kbd {
    font-family: var(--mono);
    font-size: 10.5px;
    font-weight: 500;
    color: var(--fg-muted);
    border: 1px solid var(--border);
    border-radius: 3px;
    padding: 0 5px;
    background: var(--bg);
  }

  .overlay {
    position: fixed;
    inset: 0;
    z-index: 100;
    background: color-mix(in oklab, black 55%, transparent);
    display: flex;
    align-items: flex-start;
    justify-content: center;
    padding: 6vh 1rem 1rem;
  }

  .modal {
    width: min(640px, 100%);
    max-height: calc(100vh - 8vh);
    background: var(--bg-elev);
    border: 1px solid var(--border);
    border-radius: var(--radius);
    box-shadow: 0 24px 60px -20px rgba(0, 0, 0, 0.6);
    display: flex;
    flex-direction: column;
    overflow: hidden;
  }

  .input-row {
    display: flex;
    align-items: center;
    gap: 0.5rem;
    padding: 0.85rem 1rem;
    border-bottom: 1px solid var(--border);
  }

  .input {
    flex: 1;
    appearance: none;
    background: transparent;
    border: none;
    outline: none;
    color: var(--fg);
    font-family: inherit;
    font-size: 0.98rem;
    padding: 0.2rem 0;
  }

  .input::placeholder {
    color: var(--fg-muted);
  }

  .esc {
    font-family: var(--mono);
    font-size: 10.5px;
    color: var(--fg-muted);
    border: 1px solid var(--border);
    border-radius: 3px;
    padding: 0 5px;
  }

  .results {
    overflow-y: auto;
    max-height: calc(100vh - 14vh);
  }

  .status {
    color: var(--fg-muted);
    font-size: 0.9rem;
    margin: 0;
    padding: 1.4rem 1rem;
    text-align: center;
  }

  .status code {
    font-family: var(--mono);
    background: var(--bg-soft);
    padding: 0.1em 0.35em;
    border-radius: 3px;
  }

  .hits {
    list-style: none;
    padding: 0.4rem 0.4rem 0.6rem;
    margin: 0;
  }

  .hits li a {
    display: block;
    padding: 0.55rem 0.65rem;
    border-radius: var(--radius-sm);
    color: var(--fg);
    text-decoration: none;
    line-height: 1.35;
  }

  .hits li a:hover,
  .hits li a:focus-visible {
    background: var(--bg-soft);
    outline: none;
  }

  .hit-title {
    display: block;
    font-weight: 600;
    font-size: 0.95rem;
    margin-bottom: 0.15rem;
  }

  .hit-excerpt {
    display: block;
    font-size: 0.85rem;
    color: var(--fg-soft);
  }

  .hit-excerpt :global(mark) {
    background: var(--accent-soft);
    color: var(--fg);
    padding: 0 0.1em;
    border-radius: 2px;
  }

  @media (max-width: 640px) {
    .trigger-label {
      display: none;
    }

    .trigger-kbd {
      display: none;
    }

    .trigger {
      padding: 0.32rem 0.5rem;
    }

    .trigger::after {
      content: '🔍';
      font-size: 13px;
    }
  }
</style>

<script lang="ts">
  import { base } from '$app/paths';
  import { page } from '$app/state';
  import Search from './Search.svelte';

  const links: Array<{ path: string; label: string; external?: boolean }> = [
    { path: '/docs/getting-started', label: 'Docs' },
    { path: '/docs/recipes', label: 'Recipes' },
    { path: '/playground', label: 'Playground' },
    { path: '/api', label: 'API' },
    { path: 'https://github.com/office-kit/xlsx', label: 'GitHub', external: true },
  ];

  function resolve(link: (typeof links)[number]): string {
    return link.external ? link.path : `${base}${link.path}`;
  }

  function isActive(link: (typeof links)[number]): boolean {
    if (link.external) return false;
    return page.url.pathname.startsWith(`${base}${link.path}`);
  }
</script>

<header class="site-header" data-pagefind-ignore>
  <div class="inner">
    <a href="{base}/" class="brand">
      <img src="{base}/logo.png" alt="" class="brand-mark" width="32" height="32" />
      <span class="brand-name">@office-kit/xlsx</span>
      <span class="brand-tag">spec</span>
    </a>
    <nav>
      {#each links as link, i (link.path)}
        <a
          href={resolve(link)}
          class="nav-link"
          class:active={isActive(link)}
          class:external={link.external}
        >
          <span class="nav-num">{String(i + 1).padStart(2, '0')}</span>
          <span class="nav-label">{link.label}</span>
          {#if link.external}<span class="nav-arrow">↗</span>{/if}
        </a>
      {/each}
      <Search />
    </nav>
  </div>
</header>

<style>
  .site-header {
    position: sticky;
    top: 0;
    z-index: 10;
    background: color-mix(in oklab, var(--bg) 85%, transparent);
    backdrop-filter: blur(10px) saturate(1.2);
    -webkit-backdrop-filter: blur(10px) saturate(1.2);
    border-bottom: 1px solid var(--border);
    height: var(--header-h);
    display: flex;
    align-items: center;
  }

  .inner {
    width: 100%;
    max-width: var(--max-wide);
    margin: 0 auto;
    padding: 0 1.5rem;
    display: flex;
    align-items: center;
    gap: 2rem;
  }

  .brand {
    display: inline-flex;
    align-items: center;
    gap: 0.6rem;
    color: var(--fg);
    font-family: var(--display);
    font-weight: 540;
    font-size: 1.18rem;
    letter-spacing: -0.02em;
    font-variation-settings: 'opsz' 96, 'SOFT' 30;
  }

  .brand:hover {
    text-decoration: none;
  }

  .brand-mark {
    display: block;
    width: 32px;
    height: 32px;
    object-fit: contain;
    flex: none;
  }

  .brand-tag {
    font-family: var(--mono);
    font-size: 10px;
    font-weight: 500;
    letter-spacing: 0.12em;
    text-transform: uppercase;
    color: var(--fg-muted);
    padding: 2px 6px;
    border: 1px solid var(--border);
    border-radius: 3px;
    line-height: 1;
    transform: translateY(1px);
  }

  nav {
    display: flex;
    gap: 0.25rem;
    align-items: center;
    margin-left: auto;
  }

  .nav-link {
    display: inline-flex;
    align-items: baseline;
    gap: 0.5rem;
    padding: 0.4rem 0.7rem;
    color: var(--fg-soft);
    font-family: var(--mono);
    font-size: 12px;
    font-weight: 500;
    letter-spacing: 0.06em;
    text-transform: uppercase;
    border-radius: var(--radius-sm);
    transition:
      color 120ms ease,
      background 120ms ease;
  }

  .nav-link:hover {
    color: var(--fg);
    background: var(--bg-soft);
    text-decoration: none;
  }

  .nav-link.active {
    color: var(--fg);
    background: var(--bg-soft);
  }

  .nav-link.active .nav-num {
    color: var(--accent);
  }

  .nav-num {
    color: var(--fg-muted);
    font-size: 10px;
    font-weight: 500;
    letter-spacing: 0;
  }

  .nav-arrow {
    color: var(--fg-muted);
    font-size: 11px;
    transform: translateY(-1px);
  }

  @media (max-width: 640px) {
    .nav-num {
      display: none;
    }

    .nav-link {
      padding: 0.4rem 0.5rem;
    }

    .brand-tag {
      display: none;
    }

    .inner {
      gap: 1rem;
      padding: 0 1rem;
    }
  }
</style>

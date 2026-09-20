<script lang="ts">
  import { afterNavigate } from '$app/navigation';
  import type { Snippet } from 'svelte';

  type Props = {
    /** Shown on the phone toggle that opens the nav; names the current page. */
    label: string;
    nav: Snippet;
    children: Snippet;
  };

  const { label, nav, children }: Props = $props();

  // The sidebar is always open on wide screens; this only governs the phone drawer.
  let navOpen = $state(false);
  afterNavigate(() => (navOpen = false));
</script>

<div class="shell frame">
  <aside class="sidebar" data-pagefind-ignore>
    <button
      type="button"
      class="sidebar-toggle"
      aria-expanded={navOpen}
      aria-controls="section-nav"
      onclick={() => (navOpen = !navOpen)}
    >
      <span>{label}</span>
      <svg width="12" height="12" viewBox="0 0 10 10" aria-hidden="true">
        <path d="M1.5 3.5 5 7l3.5-3.5" fill="none" stroke="currentColor" stroke-width="1.6" />
      </svg>
    </button>
    <div id="section-nav" class="sidebar-body" class:open={navOpen}>
      {@render nav()}
    </div>
  </aside>

  <div class="main">
    {@render children()}
  </div>
</div>

<style>
  .shell {
    display: flex;
    align-items: stretch;
  }

  .sidebar {
    flex: 0 0 var(--sidebar-w);
    width: var(--sidebar-w);
    border-right: 1px solid var(--line);
  }

  .sidebar-body {
    position: sticky;
    top: var(--header-h);
    max-height: calc(100vh - var(--header-h));
    overflow-y: auto;
    padding: 2.25rem 1rem 3rem var(--gutter);
  }

  .sidebar-toggle {
    display: none;
  }

  .main {
    flex: 1;
    min-width: 0;
    padding: 2.75rem clamp(1.25rem, 4vw, 3.5rem) 4rem;
  }

  @media (max-width: 860px) {
    .shell {
      flex-direction: column;
    }

    .sidebar {
      position: sticky;
      top: var(--header-h);
      z-index: 20;
      flex: none;
      width: 100%;
      border-right: none;
      border-bottom: 1px solid var(--line);
      background: var(--paper);
    }

    .sidebar-toggle {
      display: flex;
      align-items: center;
      justify-content: space-between;
      gap: 1rem;
      width: 100%;
      min-height: 48px;
      padding: 0 var(--gutter);
      border: none;
      background: transparent;
      color: var(--ink);
      font-family: var(--sans);
      font-size: 0.95rem;
      font-weight: 600;
      cursor: pointer;
    }

    .sidebar-toggle svg {
      flex: none;
    }

    .sidebar-toggle[aria-expanded='true'] svg {
      transform: rotate(180deg);
    }

    .sidebar-body {
      display: none;
      position: static;
      max-height: calc(100dvh - var(--header-h) - 48px);
      padding: 0.5rem var(--gutter) 1.5rem;
      border-top: 1px solid var(--line);
    }

    .sidebar-body.open {
      display: block;
    }

    .main {
      padding-top: 2rem;
    }
  }
</style>

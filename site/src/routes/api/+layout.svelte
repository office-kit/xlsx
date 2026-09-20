<script lang="ts">
  import { base } from '$app/paths';
  import { page } from '$app/state';
  import SidebarShell from '$lib/components/SidebarShell.svelte';
  import type { LayoutProps } from './$types';

  const { data, children }: LayoutProps = $props();

  const currentSection = $derived(
    data.sections.find((section) => section.id === page.params.section),
  );
</script>

<SidebarShell label={currentSection?.title ?? 'API reference'}>
  {#snippet nav()}
    <nav class="api-nav" aria-label="API sections">
      <h2>API reference</h2>
      <ul>
        <li>
          <a href="{base}/api" aria-current={currentSection ? undefined : 'page'}>Overview</a>
        </li>
        {#each data.sections as section (section.id)}
          <li>
            <a
              href="{base}/api/{section.id}"
              aria-current={currentSection?.id === section.id ? 'page' : undefined}
            >
              {section.title}
              <span class="count">{section.itemCount}</span>
            </a>
          </li>
        {/each}
      </ul>
    </nav>
  {/snippet}

  {@render children?.()}
</SidebarShell>

<style>
  h2 {
    margin: 0 0 0.5rem;
    font-family: var(--sans);
    font-size: 0.85rem;
    font-weight: 600;
    letter-spacing: 0;
    color: var(--ink-3);
  }

  ul {
    list-style: none;
    padding: 0;
    margin: 0;
    border-left: 1px solid var(--line);
  }

  li {
    margin: 0;
  }

  a {
    display: flex;
    align-items: baseline;
    justify-content: space-between;
    gap: 0.5rem;
    margin-left: -1px;
    padding: 0.4rem 0 0.4rem 0.95rem;
    border-left: 1px solid transparent;
    color: var(--ink-2);
    font-size: 0.95rem;
    line-height: 1.35;
  }

  a:hover {
    color: var(--ink);
    text-decoration: none;
  }

  a[aria-current='page'] {
    border-left-color: var(--accent);
    color: var(--accent-ink);
    font-weight: 600;
  }

  .count {
    color: var(--ink-3);
    font-family: var(--mono);
    font-size: 0.78rem;
    font-weight: 400;
    font-variant-numeric: tabular-nums;
  }
</style>

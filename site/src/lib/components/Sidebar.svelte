<script lang="ts">
  import { base } from '$app/paths';
  import { page } from '$app/state';
  import { docSections } from '$lib/docs-nav';
</script>

<nav class="sidebar-nav" aria-label="Documentation">
  {#each docSections as section (section.title)}
    <section>
      <h2>{section.title}</h2>
      <ul>
        {#each section.links as link (link.href)}
          {@const active = page.url.pathname.replace(/\/$/, '') === `${base}${link.href}`}
          <li>
            <a href="{base}{link.href}" aria-current={active ? 'page' : undefined}>{link.title}</a>
          </li>
        {/each}
      </ul>
    </section>
  {/each}
</nav>

<style>
  section + section {
    margin-top: 1.75rem;
  }

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
    display: block;
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
</style>

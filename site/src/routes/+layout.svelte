<script lang="ts">
  import '@office-kit/site-kit/site.css';
  import { KitFooter, KitHeader } from '@office-kit/site-kit';
  import CodeCopyEnhancer from '$lib/components/CodeCopyEnhancer.svelte';
  import Search from '$lib/components/Search.svelte';

  type Props = {
    children?: import('svelte').Snippet;
  };

  const { children }: Props = $props();

  const links = [
    { path: '/docs/getting-started', label: 'Docs', section: '/docs' },
    { path: '/docs/recipes', label: 'Recipes' },
    { path: '/api', label: 'API' },
    { path: '/playground', label: 'Playground' },
  ];
</script>

<a class="skip" href="#main">Skip to content</a>

<KitHeader product="xlsx" {links}>
  {#snippet search()}
    <Search />
  {/snippet}
</KitHeader>

<main id="main">
  {@render children?.()}
</main>

<CodeCopyEnhancer />

<KitFooter product="xlsx" {links} />

<style>
  main {
    min-height: calc(100vh - var(--header-h));
  }

  .skip {
    position: absolute;
    left: 1rem;
    top: -4rem;
    z-index: 40;
    padding: 0.6rem 1rem;
    border-radius: var(--radius);
    background: var(--ink);
    color: var(--paper);
  }

  .skip:focus {
    top: 0.75rem;
  }
</style>

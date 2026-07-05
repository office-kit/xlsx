<script lang="ts">
  import '../app.css';
  import { base } from '$app/paths';
  import CodeCopyEnhancer from '$lib/components/CodeCopyEnhancer.svelte';
  import SiteHeader from '$lib/components/SiteHeader.svelte';

  type Props = {
    children?: import('svelte').Snippet;
  };

  const { children }: Props = $props();
</script>

<SiteHeader />

<main>
  {@render children?.()}
</main>

<CodeCopyEnhancer />

<footer>
  <div class="footer-inner">
    <div class="footer-grid">
      <div class="cell">
        <span class="label">Project</span>
        <a href="https://github.com/office-kit/xlsx">@office-kit/xlsx</a>
      </div>
      <div class="cell">
        <span class="label">License</span>
        <span>MIT</span>
      </div>
      <div class="cell">
        <span class="label">Runtime</span>
        <span>Node 22+ · browsers</span>
      </div>
    </div>
    <div class="footer-llms">
      <span class="muted">For LLMs:</span>
      <a href="{base}/llms.txt">/llms.txt</a>
      <span class="dot">·</span>
      <a href="{base}/llms-full.txt">/llms-full.txt</a>
      <span class="muted">— or append <code>.md</code> to any docs URL.</span>
    </div>
  </div>
</footer>

<style>
  main {
    min-height: calc(100vh - var(--header-h));
  }

  footer {
    border-top: 1px solid var(--border);
    padding: 2.25rem 1.5rem 3rem;
    background: var(--bg-elev);
  }

  .footer-inner {
    max-width: var(--max-wide);
    margin: 0 auto;
  }

  .footer-grid {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
    gap: 0;
    border: 1px solid var(--border);
    border-radius: var(--radius);
    overflow: hidden;
    background: var(--bg);
  }

  .cell {
    display: flex;
    flex-direction: column;
    gap: 0.35rem;
    padding: 0.95rem 1.1rem;
    font-size: 0.92rem;
    color: var(--fg);
  }

  .label {
    font-family: var(--mono);
    font-size: 10px;
    font-weight: 500;
    letter-spacing: 0.14em;
    text-transform: uppercase;
    color: var(--fg-muted);
  }

  .cell a {
    font-family: var(--display);
    font-weight: 540;
    font-size: 1.05rem;
    color: var(--fg);
    font-variation-settings: 'opsz' 64, 'SOFT' 30;
  }

  .cell a:hover {
    color: var(--accent);
    text-decoration: none;
  }

  .footer-llms {
    margin-top: 1.25rem;
    font-family: var(--mono);
    font-size: 12px;
    color: var(--fg-soft);
    display: flex;
    align-items: center;
    flex-wrap: wrap;
    gap: 0.5rem;
  }

  .footer-llms .muted {
    color: var(--fg-muted);
  }

  .footer-llms .dot {
    color: var(--fg-faint);
  }

  .footer-llms a {
    color: var(--accent);
  }

  .footer-llms code {
    font-size: 11px;
    padding: 0.05em 0.3em;
  }

  @media (max-width: 600px) {
    .cell {
      border-bottom: 1px solid var(--border);
    }

    .cell:last-child {
      border-bottom: none;
    }
  }
</style>

<script lang="ts">
  type Props = {
    /** Pre-rendered Shiki HTML (the full <pre>...</pre>). */
    html: string;
    /** Optional file path / caption shown above the snippet. */
    title?: string;
  };

  const { html, title }: Props = $props();

  // eslint-disable-next-line prefer-const -- reassigned by `bind:this` in template
  let bodyEl = $state<HTMLDivElement | undefined>();
  let status = $state<'idle' | 'copied' | 'failed'>('idle');
  let timer: ReturnType<typeof setTimeout> | undefined;

  async function copy() {
    const text = bodyEl?.querySelector('pre')?.innerText ?? '';
    if (!text) return;
    try {
      await navigator.clipboard.writeText(text);
      status = 'copied';
    } catch {
      status = 'failed';
    }
    clearTimeout(timer);
    timer = setTimeout(() => {
      status = 'idle';
    }, 1500);
  }
</script>

<figure class="code-block">
  <figcaption>
    {#if title}
      <span class="path">{title}</span>
    {:else}
      <span class="path" aria-hidden="true"></span>
    {/if}
    <button
      type="button"
      class="copy"
      class:copied={status === 'copied'}
      class:failed={status === 'failed'}
      onclick={copy}
      aria-label="Copy code to clipboard"
    >
      {#if status === 'copied'}Copied{:else if status === 'failed'}Failed{:else}Copy{/if}
    </button>
  </figcaption>
  <div class="body" bind:this={bodyEl}>{@html html}</div>
</figure>

<style>
  .code-block {
    margin: 1.5rem 0;
    border: 1px solid var(--night-line);
    border-radius: var(--radius);
    background: var(--night);
    overflow: hidden;
  }

  figcaption {
    display: flex;
    align-items: center;
    gap: 0.6rem;
    padding: 0.4rem 0.45rem 0.4rem 1rem;
    border-bottom: 1px solid var(--night-line);
    color: var(--night-ink-2);
    font-family: var(--mono);
    font-size: 0.8rem;
  }

  .path {
    flex: 1;
    min-width: 0;
    overflow: hidden;
    text-overflow: ellipsis;
    white-space: nowrap;
  }

  .copy {
    flex: none;
    padding: 0.3rem 0.6rem;
    border: 1px solid var(--night-line);
    border-radius: var(--radius-sm);
    background: var(--night-2);
    color: var(--night-ink-2);
    font-family: var(--sans);
    font-size: 0.78rem;
    font-weight: 550;
    cursor: pointer;
  }

  .copy:hover {
    color: #fff;
    border-color: var(--night-ink-2);
  }

  .copy.copied {
    color: #fff;
    border-color: var(--accent);
    background: var(--accent);
  }

  .body :global(pre) {
    margin: 0;
    border: none;
    border-radius: 0;
  }
</style>

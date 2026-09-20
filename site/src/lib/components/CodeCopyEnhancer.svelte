<script lang="ts">
  // Adds a floating "Copy" button to every <pre> rendered outside a CodeBlock
  // figure — i.e. every code fence inside an .svx / .md doc. CodeBlock has its
  // own copy button in the figcaption, so we skip those.
  import { afterNavigate } from '$app/navigation';
  import { onMount, tick } from 'svelte';

  const ENHANCED = 'data-copy-enhanced';

  async function enhance() {
    await tick();
    const pres = document.querySelectorAll<HTMLPreElement>(`pre:not([${ENHANCED}])`);
    for (const pre of pres) {
      if (pre.closest('.code-block')) continue;
      attach(pre);
    }
  }

  function attach(pre: HTMLPreElement) {
    pre.setAttribute(ENHANCED, '');

    // Wrap the <pre> so the absolutely-positioned button anchors to its corner
    // without depending on a containing block we don't control.
    const wrap = document.createElement('div');
    wrap.className = 'code-copy-wrap';
    pre.parentNode?.insertBefore(wrap, pre);
    wrap.appendChild(pre);

    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'code-copy-btn';
    btn.setAttribute('aria-label', 'Copy code to clipboard');
    btn.textContent = 'Copy';

    let timer: ReturnType<typeof setTimeout> | undefined;
    btn.addEventListener('click', async () => {
      const text = pre.innerText;
      if (!text) return;
      try {
        await navigator.clipboard.writeText(text);
        btn.textContent = 'Copied';
        btn.classList.add('copied');
      } catch {
        btn.textContent = 'Failed';
        btn.classList.add('failed');
      }
      clearTimeout(timer);
      timer = setTimeout(() => {
        btn.textContent = 'Copy';
        btn.classList.remove('copied', 'failed');
      }, 1500);
    });

    wrap.appendChild(btn);
  }

  onMount(() => void enhance());
  afterNavigate(() => void enhance());
</script>

<style>
  :global(.code-copy-wrap) {
    position: relative;
  }

  :global(.code-copy-btn) {
    position: absolute;
    top: 0.55rem;
    right: 0.55rem;
    padding: 0.3rem 0.6rem;
    border: 1px solid var(--night-line);
    border-radius: var(--radius-sm);
    background: var(--night-2);
    color: var(--night-ink-2);
    font-family: var(--sans);
    font-size: 0.78rem;
    font-weight: 550;
    cursor: pointer;
    opacity: 0;
    transition: opacity 120ms ease;
  }

  :global(.code-copy-wrap:hover .code-copy-btn),
  :global(.code-copy-btn:focus-visible),
  :global(.code-copy-btn.copied),
  :global(.code-copy-btn.failed) {
    opacity: 1;
  }

  /* No hover on touch screens, so the button has to be there from the start. */
  @media (hover: none) {
    :global(.code-copy-btn) {
      opacity: 1;
    }
  }

  :global(.code-copy-btn:hover) {
    color: #fff;
    border-color: var(--night-ink-2);
  }

  :global(.code-copy-btn.copied) {
    color: #fff;
    border-color: var(--accent);
    background: var(--accent);
  }
</style>

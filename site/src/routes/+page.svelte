<script lang="ts">
  import { base } from '$app/paths';
  import CodeBlock from '$lib/components/CodeBlock.svelte';
  import type { PageProps } from './$types';

  const { data }: PageProps = $props();

  const features: Array<{ num: string; title: string; body: string }> = [
    {
      num: '01',
      title: 'Round-trips real workbooks',
      body: "Pivot tables, macro-enabled .xlsm, threaded comments, Power Query metadata, custom XML — anything we don't model is preserved byte-identical so Excel 365 still renders it.",
    },
    {
      num: '02',
      title: 'Streaming, both directions',
      body: 'createWriteOnlyWorkbook deflates rows as they arrive. loadWorkbookStream walks a file once and yields each row. Browser-safe via @office-kit/xlsx/streaming.',
    },
    {
      num: '03',
      title: 'Charts & drawings, modeled',
      body: '16 legacy c: chart kinds plus 8 cx: chartex kinds (Sunburst, Treemap, Waterfall, Histogram, Pareto, Funnel, BoxWhisker, RegionMap). Images auto-detect format and dimensions.',
    },
    {
      num: '04',
      title: 'Tiny & tree-shakeable',
      body: '@office-kit/xlsx ≤ 120 KB brotli (currently ~78 KB). @office-kit/xlsx/streaming ≤ 80 KB brotli (~47 KB). Every export is side-effect-free.',
    },
  ];

  const stats = [
    { label: 'rows', value: '10,000,000', sub: 'streamed under 100 MB heap' },
    { label: 'bundle', value: '~78 KB', sub: 'brotli, tree-shakeable' },
    { label: 'targets', value: 'Node + browser', sub: 'no Python, no Excel' },
  ];
</script>

<svelte:head>
  <title>@office-kit/xlsx — Read and write Excel .xlsx in Node and the browser</title>
</svelte:head>

<section class="hero">
  <div class="hero-bg" aria-hidden="true"></div>

  <div class="hero-inner">
    <h1 class="display">
      <span class="line">
        Read &amp; write
        <em>Excel</em>
        <span class="dot-sep" aria-hidden="true"></span>
        <code class="filetype">.xlsx</code>
      </span>
      <span class="line muted-line">
        from
        <span class="hl">Node</span>
        and
        <span class="hl">the browser</span>.
      </span>
    </h1>

    <p class="lede">
      Full workbook model — values, formulas, styles, charts, drawings, pivots, VBA — plus a
      streaming writer that pushes <strong>10M rows under a 100&nbsp;MB heap</strong>. No
      Python. No Excel. No runtime native modules.
    </p>

    <div class="cta">
      <a href="{base}/docs/getting-started" class="btn primary">
        <span>Get started</span>
        <span class="arrow">→</span>
      </a>
      <a href="{base}/docs/recipes" class="btn ghost">Recipes</a>
      <a href="{base}/api" class="btn ghost">API reference</a>
      <a href="https://github.com/office-kit/xlsx" class="btn ghost">
        GitHub <span class="ext">↗</span>
      </a>
    </div>

    <div class="install" role="group" aria-label="Install command">
      <span class="install-coord">$</span>
      <code class="install-cmd">pnpm add @office-kit/xlsx</code>
      <span class="install-alt">
        <span class="alt-sep">/</span>
        <code>npm i @office-kit/xlsx</code>
      </span>
    </div>

    <div class="stats" aria-label="Project stats">
      {#each stats as stat, i (stat.label)}
        <div class="stat" style="--i: {i}">
          <span class="stat-label">{stat.label}</span>
          <span class="stat-value">{stat.value}</span>
          <span class="stat-sub">{stat.sub}</span>
        </div>
      {/each}
    </div>
  </div>
</section>

<section class="features">
  <div class="features-inner">
    <header class="section-head">
      <h2>What it does well</h2>
      <span class="row-count">04 / 04</span>
    </header>

    <div class="features-grid">
      {#each features as f (f.num)}
        <article class="feature">
          <span class="feature-num">{f.num}</span>
          <h3>{f.title}</h3>
          <p>{f.body}</p>
        </article>
      {/each}
    </div>
  </div>
</section>

<section class="examples">
  <div class="examples-inner">
    <header class="section-head">
      <h2>Two snippets to get the shape</h2>
      <span class="row-count">live · type-checked</span>
    </header>

    <p class="lede examples-lede">
      Both files below live under <code>site/src/lib/examples/</code> and are type-checked by
      <code>svelte-check</code> against the real library on every build — if an API renames,
      the docs build fails.
    </p>

    {#each data.hero as ex, i (ex.key)}
      <div class="example">
        <header class="example-head">
          <span class="example-num">{String(i + 1).padStart(2, '0')}</span>
          <div class="example-text">
            <h3>{ex.title}</h3>
            <p>{ex.description}</p>
          </div>
        </header>
        <CodeBlock html={ex.html} title={ex.path} />
      </div>
    {/each}

    <p class="more">
      More in <a href="{base}/docs/getting-started">Getting started</a> &amp;
      <a href="{base}/docs/streaming">Streaming</a>.
    </p>
  </div>
</section>

<style>
  .hero {
    position: relative;
    padding: 3.5rem 1.5rem 4.5rem;
    border-bottom: 1px solid var(--border);
    overflow: hidden;
  }

  .hero-bg {
    position: absolute;
    inset: 0;
    pointer-events: none;
    background:
      radial-gradient(circle at 18% -10%, var(--accent-glow), transparent 45%),
      radial-gradient(circle at 100% 30%, var(--brass-soft), transparent 50%);
  }

  .hero-inner {
    position: relative;
    max-width: var(--max-wide);
    margin: 0 auto;
  }

  .display {
    font-family: var(--display);
    font-size: clamp(2.4rem, 6.5vw, 4.6rem);
    font-weight: 460;
    line-height: 1;
    letter-spacing: -0.035em;
    margin: 0 0 1.5rem;
    font-variation-settings: 'opsz' 144, 'SOFT' 30;
    color: var(--fg);
    max-width: 14ch;
  }

  .display .line {
    display: block;
  }

  .display em {
    font-style: italic;
    font-weight: 400;
    color: var(--accent);
    font-variation-settings: 'opsz' 144, 'SOFT' 70;
    margin: 0 0.05em;
  }

  .display .filetype {
    display: inline-block;
    font-family: var(--mono);
    font-weight: 500;
    font-size: 0.55em;
    color: var(--fg);
    background: var(--bg-paper);
    border: 1px solid var(--border);
    padding: 0.05em 0.4em;
    border-radius: var(--radius-sm);
    transform: translateY(-0.18em);
    letter-spacing: 0;
  }

  .display .dot-sep {
    display: inline-block;
    width: 0.18em;
    height: 0.18em;
    background: var(--brass);
    border-radius: 50%;
    transform: translateY(-0.4em);
    margin: 0 0.18em;
  }

  .display .muted-line {
    color: var(--fg-soft);
    font-size: 0.78em;
    font-weight: 380;
    margin-top: 0.15em;
  }

  .display .hl {
    position: relative;
    color: var(--fg);
    font-style: italic;
    font-variation-settings: 'opsz' 144, 'SOFT' 70;
    font-weight: 400;
    background-image: linear-gradient(0deg, var(--brass-soft) 0%, var(--brass-soft) 18%, transparent 18%);
    background-position: 0 88%;
    background-repeat: no-repeat;
    padding: 0 0.05em;
  }

  .lede {
    color: var(--fg-soft);
    font-size: 1.08rem;
    line-height: 1.55;
    max-width: 56ch;
    margin: 0 0 1.75rem;
  }

  .lede strong {
    color: var(--fg);
    font-weight: 650;
  }

  .cta {
    display: flex;
    gap: 0.5rem;
    flex-wrap: wrap;
    margin-bottom: 1.5rem;
  }

  .btn {
    display: inline-flex;
    align-items: center;
    gap: 0.5rem;
    padding: 0.7rem 1.05rem;
    border-radius: var(--radius-sm);
    border: 1px solid var(--border);
    background: var(--bg-elev);
    color: var(--fg);
    font-family: var(--sans);
    font-size: 0.94rem;
    font-weight: 540;
    letter-spacing: -0.005em;
    transition:
      transform 120ms ease,
      background 120ms ease,
      border-color 120ms ease,
      color 120ms ease;
  }

  .btn:hover {
    text-decoration: none;
    background: var(--bg-soft);
    border-color: var(--border-strong);
    transform: translateY(-1px);
  }

  .btn.primary {
    background: var(--accent);
    border-color: var(--accent);
    color: var(--bg);
    font-weight: 580;
    box-shadow: 0 8px 30px -12px var(--accent-glow);
  }

  .btn.primary:hover {
    background: var(--accent-hot);
    border-color: var(--accent-hot);
  }

  .btn .arrow {
    transition: transform 160ms ease;
  }

  .btn.primary:hover .arrow {
    transform: translateX(3px);
  }

  .btn .ext {
    color: var(--fg-muted);
    font-size: 0.85em;
  }

  .install {
    display: inline-flex;
    align-items: center;
    gap: 0.6rem;
    padding: 0.6rem 0.9rem;
    border: 1px solid var(--border);
    border-radius: var(--radius-sm);
    background: var(--code-bg);
    font-family: var(--mono);
    font-size: 0.88rem;
    color: var(--fg);
    margin-bottom: 2.5rem;
    flex-wrap: wrap;
  }

  .install-coord {
    color: var(--accent);
    font-weight: 600;
  }

  .install-cmd {
    background: transparent;
    border: none;
    padding: 0;
    color: var(--fg);
  }

  .install-alt {
    display: inline-flex;
    align-items: center;
    gap: 0.55rem;
    color: var(--fg-muted);
  }

  .install-alt .alt-sep {
    color: var(--fg-faint);
  }

  .install-alt code {
    background: transparent;
    border: none;
    padding: 0;
    color: var(--fg-soft);
  }

  /* Stats strip — like a spreadsheet status bar. */
  .stats {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(180px, 1fr));
    border: 1px solid var(--border);
    border-radius: var(--radius);
    overflow: hidden;
    background: var(--bg-paper);
  }

  .stat {
    display: flex;
    flex-direction: column;
    gap: 0.1rem;
    padding: 1rem 1.1rem;
    position: relative;
  }

  .stat::before {
    content: '';
    position: absolute;
    top: 0;
    left: 0;
    width: 28px;
    height: 1px;
    background: var(--accent);
    opacity: calc(1 - var(--i) * 0.18);
  }

  .stat-label {
    font-family: var(--mono);
    font-size: 10.5px;
    font-weight: 500;
    letter-spacing: 0.16em;
    text-transform: uppercase;
    color: var(--fg-muted);
  }

  .stat-value {
    font-family: var(--display);
    font-size: 1.55rem;
    font-weight: 480;
    color: var(--fg);
    line-height: 1.1;
    margin-top: 0.15rem;
    font-variation-settings: 'opsz' 96, 'SOFT' 30, 'wght' 480;
    letter-spacing: -0.02em;
  }

  .stat-sub {
    font-family: var(--mono);
    font-size: 11px;
    color: var(--fg-muted);
    margin-top: 0.25rem;
  }

  /* Section heads — used by features + examples + future sections. */
  .section-head {
    display: flex;
    align-items: baseline;
    gap: 0.85rem;
    margin: 0 0 1.5rem;
    border-bottom: 1px solid var(--border);
    padding-bottom: 0.85rem;
  }

  .section-head h2 {
    margin: 0;
    border: none;
    padding: 0;
    font-size: clamp(1.6rem, 3vw, 2.1rem);
    flex: 1;
  }

  .section-head .row-count {
    font-family: var(--mono);
    font-size: 11.5px;
    color: var(--fg-muted);
    text-transform: uppercase;
    letter-spacing: 0.12em;
  }

  /* Features. */
  .features {
    padding: 4.5rem 1.5rem;
    border-bottom: 1px solid var(--border);
  }

  .features-inner {
    max-width: var(--max-wide);
    margin: 0 auto;
  }

  .features-grid {
    display: grid;
    grid-template-columns: repeat(auto-fit, minmax(260px, 1fr));
    border: 1px solid var(--border);
    border-radius: var(--radius);
    overflow: hidden;
    background: var(--bg);
  }

  .feature {
    position: relative;
    padding: 1.35rem 1.4rem 1.5rem;
    border-bottom: 1px solid var(--border);
    background: var(--bg-elev);
    transition: background 160ms ease;
  }

  .feature:hover {
    background: var(--bg-soft);
  }

  .feature:nth-last-child(-n + 2) {
    border-bottom: none;
  }

  @media (max-width: 580px) {
    .feature:nth-last-child(-n + 2) {
      border-bottom: 1px solid var(--border);
    }

    .feature:last-child {
      border-bottom: none;
    }
  }

  .feature-num {
    display: inline-block;
    font-family: var(--mono);
    font-size: 11px;
    font-weight: 500;
    letter-spacing: 0.08em;
    color: var(--accent);
    margin-bottom: 0.6rem;
  }

  .feature h3 {
    font-family: var(--display);
    font-size: 1.18rem;
    font-weight: 540;
    margin: 0 0 0.45rem;
    line-height: 1.2;
    color: var(--fg);
    font-variation-settings: 'opsz' 32, 'SOFT' 30;
  }

  .feature p {
    color: var(--fg-soft);
    font-size: 0.92rem;
    line-height: 1.55;
    margin: 0;
  }

  /* Examples. */
  .examples {
    padding: 4.5rem 1.5rem 6rem;
  }

  .examples-inner {
    max-width: var(--max-content);
    margin: 0 auto;
  }

  .examples-lede {
    margin-bottom: 2rem;
  }

  .example {
    margin: 0 0 2.5rem;
  }

  .example-head {
    display: flex;
    align-items: flex-start;
    gap: 0.9rem;
    margin: 0 0 0.65rem;
  }

  .example-num {
    flex: none;
    margin-top: 0.55rem;
    font-family: var(--mono);
    font-size: 11.5px;
    font-weight: 500;
    letter-spacing: 0.08em;
    color: var(--accent);
  }

  .example-text {
    flex: 1;
  }

  .example-head h3 {
    margin: 0 0 0.25rem;
    font-size: 1.32rem;
    font-family: var(--display);
    font-weight: 540;
    font-variation-settings: 'opsz' 32, 'SOFT' 25;
  }

  .example-head p {
    color: var(--fg-soft);
    margin: 0;
    font-size: 0.95rem;
  }

  .more {
    margin-top: 2rem;
    color: var(--fg-soft);
    border-top: 1px solid var(--border);
    padding-top: 1.25rem;
    font-size: 0.95rem;
  }

  @media (max-width: 720px) {
    .display {
      max-width: none;
    }
  }
</style>

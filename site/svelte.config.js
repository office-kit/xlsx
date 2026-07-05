import adapter from '@sveltejs/adapter-static';
import { vitePreprocess } from '@sveltejs/vite-plugin-svelte';
import { mdsvex, escapeSvelte } from 'mdsvex';
import { createHighlighter } from 'shiki';

const theme = 'github-dark';
const langs = ['ts', 'tsx', 'js', 'json', 'sh', 'bash', 'xml', 'svelte', 'html'];

const highlighter = await createHighlighter({ themes: [theme], langs });

/** @type {import('mdsvex').MdsvexOptions} */
const mdsvexOptions = {
  extensions: ['.svx', '.md'],
  highlight: {
    highlighter: async (code, lang = 'text') => {
      const safeLang = langs.includes(lang) ? lang : 'text';
      const html = escapeSvelte(highlighter.codeToHtml(code, { lang: safeLang, theme }));
      return `{@html \`${html}\`}`;
    },
  },
};

// BASE_PATH lets the same build run locally (=''), on a GitHub user page
// or custom domain (=''), or on a project page (e.g. '/xlsx'). Set
// it in CI for GitHub Actions deploys.
const basePath = process.env.BASE_PATH ?? '';

/** @type {import('@sveltejs/kit').Config} */
const config = {
  extensions: ['.svelte', '.svx', '.md'],
  preprocess: [vitePreprocess(), mdsvex(mdsvexOptions)],
  kit: {
    adapter: adapter({ fallback: '404.html' }),
    prerender: { entries: ['*'] },
    paths: { base: basePath, relative: true },
    alias: {
      '@office-kit/xlsx/cell': '../src/cell/index.ts',
      '@office-kit/xlsx/chart': '../src/chart/index.ts',
      '@office-kit/xlsx/chartsheet': '../src/chartsheet/index.ts',
      '@office-kit/xlsx/drawing': '../src/drawing/index.ts',
      '@office-kit/xlsx/io': '../src/io/index.ts',
      '@office-kit/xlsx/node': '../src/node.ts',
      '@office-kit/xlsx/packaging': '../src/packaging/index.ts',
      '@office-kit/xlsx/schema': '../src/schema/index.ts',
      '@office-kit/xlsx/streaming': '../src/streaming/index.ts',
      '@office-kit/xlsx/styles': '../src/styles/index.ts',
      '@office-kit/xlsx/utils': '../src/utils/index.ts',
      '@office-kit/xlsx/workbook': '../src/workbook/index.ts',
      '@office-kit/xlsx/worksheet': '../src/worksheet/index.ts',
      '@office-kit/xlsx/xml': '../src/xml/index.ts',
      '@office-kit/xlsx/zip': '../src/zip/index.ts',
    },
  },
};

export default config;

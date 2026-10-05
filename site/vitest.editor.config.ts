import { fileURLToPath } from 'node:url';
import { svelte } from '@sveltejs/vite-plugin-svelte';
import { defineConfig } from 'vitest/config';

// Unit tests for the spreadsheet editor's framework-free logic
// (site/src/lib/editor/**). The editor consumes the library through its public
// subpaths, so the aliases mirror svelte.config.js and point at ../src/.
// The Svelte plugin compiles the runes in *.svelte.ts model files.

const SUBPATHS = [
  'cell',
  'chart',
  'chartsheet',
  'drawing',
  'io',
  'packaging',
  'schema',
  'streaming',
  'styles',
  'utils',
  'workbook',
  'worksheet',
  'xml',
  'zip',
];

const alias = Object.fromEntries(
  SUBPATHS.map((name) => [`@office-kit/xlsx/${name}`, fileURLToPath(new URL(`../src/${name}/index.ts`, import.meta.url))]),
);

export default defineConfig({
  plugins: [svelte({ compilerOptions: { runes: true } })],
  resolve: { alias },
  test: {
    environment: 'node',
    include: ['src/lib/editor/**/*.test.ts'],
  },
});

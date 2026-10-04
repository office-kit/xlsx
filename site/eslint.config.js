// Lints .svelte files with Svelte's own parser, which reads the template too.
// oxlint only sees the <script> block, so a variable written through `bind:`
// or an inline handler looks constant or unassigned to it; the two rules that
// depends on (prefer-const, no-unassigned-vars) are checked here instead.
// Every other rule on .svelte scripts stays with oxlint.
import svelte from 'eslint-plugin-svelte';
import ts from 'typescript-eslint';
import svelteConfig from './svelte.config.js';

export default [
  ...svelte.configs['flat/base'],
  {
    files: ['**/*.svelte'],
    languageOptions: {
      parserOptions: { parser: ts.parser, extraFileExtensions: ['.svelte'], svelteConfig },
    },
    rules: {
      'svelte/prefer-const': 'error',
      'no-unassigned-vars': 'error',
    },
  },
];

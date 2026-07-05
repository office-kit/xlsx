// Registry of every example file. The ?raw imports give us the on-disk
// source verbatim; the imports of the modules themselves (handled by ?raw
// not actually pulling them in at runtime, but listed here as the
// canonical spelling of each example's identity) keep this list in sync
// with the files on disk via the `as const` shape.

import basicReadWrite from './basic-read-write.ts?raw';
import browserFetch from './browser-fetch.ts?raw';
import nodeFs from './node-fs.ts?raw';
import streamingRead from './streaming-read.ts?raw';
import streamingWrite from './streaming-write.ts?raw';

export type Example = {
  /** Human title for the snippet (shown above the code block). */
  title: string;
  /** Repo-relative path, also used for the file-tab caption. */
  path: string;
  /** Verbatim source text (already type-checked by svelte-check / tsc). */
  source: string;
  /** Short description used in docs. */
  description: string;
};

export const examples = {
  basicReadWrite: {
    title: 'Read + edit + write',
    path: 'site/src/lib/examples/basic-read-write.ts',
    description:
      'Open an xlsx, mutate one cell, write it back — the canonical full-library round-trip.',
    source: basicReadWrite,
  },
  nodeFs: {
    title: 'Direct fs helpers',
    path: 'site/src/lib/examples/node-fs.ts',
    description:
      'fromFile / toFile from @office-kit/xlsx/node skip the manual readFile + writeFile glue.',
    source: nodeFs,
  },
  browserFetch: {
    title: 'Fetch in the browser',
    path: 'site/src/lib/examples/browser-fetch.ts',
    description:
      'fromResponse streams the workbook out of a fetch response — no full download buffer.',
    source: browserFetch,
  },
  streamingWrite: {
    title: 'Streaming write — 10M rows',
    path: 'site/src/lib/examples/streaming-write.ts',
    description:
      'createWriteOnlyWorkbook deflates rows as they arrive. Heap stays under 100 MB.',
    source: streamingWrite,
  },
  streamingRead: {
    title: 'Streaming read — iter rows',
    path: 'site/src/lib/examples/streaming-read.ts',
    description:
      'loadWorkbookStream + iterRows walks the file once and yields each row.',
    source: streamingRead,
  },
} as const satisfies Record<string, Example>;

export type ExampleKey = keyof typeof examples;

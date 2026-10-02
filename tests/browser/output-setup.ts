import { mkdirSync, rmSync, writeFileSync } from 'node:fs';
import { corpus } from '../conformance/corpus.js';
export const engines = ['chromium', 'firefox', 'webkit'] as const;
export const adapters = ['arrayBuffer', 'blob', 'response', 'stream'] as const;
export const outputDirectory = '.qa/browser-packages';
export default function setup(): void {
  rmSync(outputDirectory, { recursive: true, force: true });
  mkdirSync(outputDirectory, { recursive: true });
  // Expected outputs are declared before tests run; missing outputs cannot pass
  // SDK validation by shrinking a manifest to just the successful cases.
  const expected = corpus.flatMap(c => engines.flatMap(engine => adapters.map(adapter => ({ id: `${c.id}.${engine}.${adapter}`, sourceCase: c.id, engine, adapter }))));
  writeFileSync(`${outputDirectory}/manifest.json`, JSON.stringify(expected, null, 2));
}

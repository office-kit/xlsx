import { createServer } from 'node:http';
import { readFileSync } from 'node:fs';
import { build } from 'esbuild';
const manifest = JSON.parse(readFileSync('tests/conformance/corpus/manifest.json', 'utf8'));
const cases = new Set(manifest.cases.map(c => c.id));
const result = await build({ entryPoints: ['tests/browser/harness.mjs'], bundle: true, format: 'esm', platform: 'browser', external: ['node:*'], write: false });
const bundle = result.outputFiles[0].contents;
const server = createServer((req, res) => {
  if (req.url === '/') { res.setHeader('Content-Type', 'text/html'); res.end('<script type="module">import * as qa from "/harness.js"; window.qa = qa;</script>'); return; }
  if (req.url === '/harness.js') { res.setHeader('Content-Type', 'text/javascript'); res.end(bundle); return; }
  const id = /^\/cases\/([A-Za-z0-9-]+)\.xlsx$/.exec(req.url ?? '')?.[1];
  if (id && cases.has(id)) { res.end(readFileSync(`.qa/corpus/${id}.input.xlsx`)); return; }
  res.writeHead(404); res.end();
});
server.listen(41749, '127.0.0.1');
for (const signal of ['SIGINT', 'SIGTERM']) process.on(signal, () => server.close());

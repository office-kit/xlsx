import { sveltekit } from '@sveltejs/kit/vite';
import { defineConfig, type Connect, type Plugin } from 'vite';
import * as fs from 'node:fs';
import * as path from 'node:path';

// Pagefind writes its index into `build/pagefind/` AFTER `vite build` has
// already cataloged the static assets — so SvelteKit's dev/preview
// middleware doesn't know about those files and 404s requests for them.
// This middleware serves `/pagefind/**` straight from disk so the site's
// own `<Search>` component can fetch the index in `pnpm dev` / `pnpm
// preview` once a production build has been run at least once. On the
// real static host (GitHub Pages) the directory is served as plain
// files; this plugin is a local-only convenience.
function servePagefindFromBuild(): Plugin {
  const buildPagefind = path.resolve('build/pagefind');

  const handler: Connect.NextHandleFunction = (req, res, next) => {
    const url = (req.url ?? '').split('?')[0];
    const match = url.match(/\/pagefind\/(.+)$/);
    if (!match) return next();

    const rel = match[1];
    if (!rel) return next();

    const filePath = path.join(buildPagefind, rel);
    if (!filePath.startsWith(buildPagefind + path.sep)) return next();

    let stat: fs.Stats;
    try {
      stat = fs.statSync(filePath);
    } catch {
      return next();
    }
    if (!stat.isFile()) return next();

    res.setHeader('Content-Type', contentTypeFor(filePath));
    res.setHeader('Cache-Control', 'no-cache');
    fs.createReadStream(filePath).pipe(res);
  };

  return {
    name: 'serve-pagefind-from-build',
    configureServer(server) {
      server.middlewares.use(handler);
    },
    configurePreviewServer(server) {
      server.middlewares.use(handler);
    },
  };
}

function contentTypeFor(p: string): string {
  switch (path.extname(p).toLowerCase()) {
    case '.js':
      return 'application/javascript';
    case '.json':
      return 'application/json';
    case '.wasm':
      return 'application/wasm';
    case '.css':
      return 'text/css';
    case '.html':
      return 'text/html';
    default:
      return 'application/octet-stream';
  }
}

export default defineConfig({
  plugins: [sveltekit(), servePagefindFromBuild()],
  server: {
    fs: {
      // Allow serving files from the parent project (for example .ts files
      // imported via ?raw and the @office-kit/xlsx source via path alias).
      allow: ['..'],
    },
  },
});

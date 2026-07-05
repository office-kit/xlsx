// Server-only: collect every docs page's raw Markdown source so we can
// (a) serve <route>.md endpoints for LLMs and (b) build /llms.txt and
// /llms-full.txt indexes.
//
// import.meta.glob with eager + ?raw inlines every .svx source string at
// build time, no runtime fs lookup.

import { examples, type ExampleKey } from '$lib/examples';
import { recipeGroups } from '$lib/examples/recipes';
import { allDocLinks, type DocLink } from '$lib/docs-nav';

const RAW_SVX = import.meta.glob('/src/routes/**/+page.svx', {
  eager: true,
  query: '?raw',
  import: 'default',
}) as Record<string, string>;

/** Map filesystem key (`/src/routes/docs/install/+page.svx`) -> route (`/docs/install`). */
function fsKeyToRoute(fsKey: string): string {
  const stripped = fsKey.replace(/^\/src\/routes/, '').replace(/\/\+page\.svx$/, '');
  return stripped || '/';
}

/** Convert an .svx source into clean Markdown that an LLM can read. */
function svxToMarkdown(source: string): string {
  // 1. Drop the leading <script>...</script> block (if any).
  let body = source.replace(/^\s*<script\b[\s\S]*?<\/script>\s*/m, '');

  // 2. Replace <CodeBlock html={ex.<key>.html} title={...} /> placeholders
  //    with the actual example source as a fenced ts block. The .svx files
  //    use ex = data.examples (see docs/+layout.server.ts) so the only
  //    information we need from the tag is the example key.
  body = body.replace(
    /<CodeBlock\s+html=\{ex\.(\w+)\.html\}[^/]*\/>/g,
    (_, key: string) => {
      const ex = examples[key as ExampleKey];
      if (!ex) return `\`\`\`text\n[unknown example: ${key}]\n\`\`\``;
      return [
        `\`\`\`ts title="${ex.path}"`,
        ex.source.trimEnd(),
        '```',
      ].join('\n');
    },
  );

  return body.trim() + '\n';
}

export type MarkdownDoc = {
  /** Site path, e.g. `/docs/install`. */
  route: string;
  /** Markdown source as an LLM sees it (script-stripped, components inlined). */
  markdown: string;
  /** Sidebar / index metadata, if this route is a known docs page. */
  link: DocLink | null;
};

function buildRecipesMarkdown(): string {
  const intro = `# Recipes

Working code for the things people actually want to do with @office-kit/xlsx — open a workbook, build one from scratch, style cells, add a chart, stream millions of rows. Every snippet below is a real \`.ts\` file in the repo, type-checked against \`@office-kit/xlsx\` on every build.
`;
  const sections = recipeGroups.map((g) => {
    const recipes = g.recipes.map((r) => {
      const notes = r.notes?.length
        ? '\n' + r.notes.map((n) => `- ${n}`).join('\n')
        : '';
      const related = r.relatedApi?.length
        ? `\n\n*Related API: ${r.relatedApi.map((n) => `\`${n}\``).join(', ')}*`
        : '';
      return [
        `### ${r.title}`,
        '',
        r.teaser,
        '',
        `\`\`\`ts title="${r.path}"`,
        r.source.trimEnd(),
        '```',
        notes,
        related,
      ]
        .filter(Boolean)
        .join('\n');
    });
    return `## ${g.title}\n\n${recipes.join('\n\n')}`;
  });
  return `${intro}\n${sections.join('\n\n')}\n`;
}

const ROUTE_TO_DOC: Record<string, MarkdownDoc> = (() => {
  const out: Record<string, MarkdownDoc> = {};
  for (const [fsKey, source] of Object.entries(RAW_SVX)) {
    const route = fsKeyToRoute(fsKey);
    out[route] = {
      route,
      markdown: svxToMarkdown(source),
      link: allDocLinks.find((l) => l.href === route) ?? null,
    };
  }
  // Synthetic docs that aren't backed by a +page.svx.
  out['/docs/recipes'] = {
    route: '/docs/recipes',
    markdown: buildRecipesMarkdown(),
    link: allDocLinks.find((l) => l.href === '/docs/recipes') ?? null,
  };
  return out;
})();

export function getMarkdownDoc(route: string): MarkdownDoc | null {
  return ROUTE_TO_DOC[route] ?? null;
}

export function listMarkdownDocs(): MarkdownDoc[] {
  // Return docs in the order declared in docs-nav (so /llms.txt reads in a
  // sensible learning order), with any unlisted .svx pages appended.
  const ordered: MarkdownDoc[] = [];
  for (const link of allDocLinks) {
    const doc = ROUTE_TO_DOC[link.href];
    if (doc) ordered.push(doc);
  }
  for (const doc of Object.values(ROUTE_TO_DOC)) {
    if (!ordered.includes(doc)) ordered.push(doc);
  }
  return ordered;
}

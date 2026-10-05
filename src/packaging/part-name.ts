// OPC part-name resolution shared by the workbook reader, the streaming reader
// and the pivot-table lifter.

/**
 * Resolve an OPC relationship target against its source part path.
 *
 * - Targets starting with `/` are package-absolute.
 * - Otherwise the target is relative to the source part's parent directory.
 * - `..` segments collapse normally.
 */
export function resolveRelTarget(sourcePartPath: string, target: string): string {
  if (target.startsWith('/')) return target.slice(1);
  const lastSlash = sourcePartPath.lastIndexOf('/');
  const parentDir = lastSlash >= 0 ? sourcePartPath.slice(0, lastSlash + 1) : '';
  const joined = parentDir + target;
  return normalizePath(joined);
}

function normalizePath(path: string): string {
  const segments = path.split('/');
  const out: string[] = [];
  for (const seg of segments) {
    if (seg === '' || seg === '.') continue;
    if (seg === '..') {
      // Pop the last accumulated segment when one exists; when `out` is
      // empty (a relative target with more `..` than ancestors) the pop is
      // a no-op so the climb is silently absorbed at the package root. The
      // archive.has() check on the resolved path catches any escape attempt
      // because the entry simply won't exist outside the package — there's
      // no filesystem traversal to worry about, only a missing-entry error.
      out.pop();
      continue;
    }
    out.push(seg);
  }
  return out.join('/');
}

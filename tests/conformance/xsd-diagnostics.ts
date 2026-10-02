// libxml builds differ: some emit an unprefixed error followed by a file summary.
export function parseXsdDiagnostics(stderr: string, files: ReadonlyMap<string, string>): Array<{ part: string; message: string }> {
  const normalize = (path: string) => path.replaceAll('\\', '/');
  const parts = new Map([...files].map(([path, part]) => [normalize(path), part]));
  const issues: Array<{ part: string; message: string }> = [];
  let pending: string[] = [];
  for (const message of stderr.split(/\r?\n/).map(line => line.trim()).filter(Boolean)) {
    const file = /^(.+?\.xml)(?::| )/.exec(message)?.[1];
    const part = file ? parts.get(normalize(file)) : undefined;
    // Success summaries in a mixed valid/invalid batch are not errors.
    if (part && /\.xml validates$/.test(message)) continue;
    if (part && /\.xml fails to validate$/.test(message)) {
      for (const detail of pending) issues.push({ part, message: detail });
      pending = [];
    }
    if (part) issues.push({ part, message });
    else pending.push(message);
  }
  // A diagnostic without a failed-file summary cannot be safely attributed.
  for (const message of pending) issues.push({ part: '<runner>', message });
  return issues;
}

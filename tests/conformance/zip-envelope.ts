/** Independent bounded ZIP32 preflight, before fflate can collapse duplicate names
 * or allocate inflated data. ZIP64 is explicitly incomplete, never a pass. */
export function checkZipEnvelope(bytes: Uint8Array): 'zip32' | 'zip64' {
  const v = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  let end = -1;
  for (let i = bytes.length - 22; i >= Math.max(0, bytes.length - 65_557); i--) {
    if (v.getUint32(i, true) === 0x06054b50 && i + 22 + v.getUint16(i + 20, true) === bytes.length) { end = i; break; }
  }
  if (end < 0) throw new Error('Missing ZIP end record');
  const count = v.getUint16(end + 10, true);
  const size = v.getUint32(end + 12, true);
  const start = v.getUint32(end + 16, true);
  if (count === 0xffff || size === 0xffffffff || start === 0xffffffff) return 'zip64';
  if (v.getUint16(end + 4, true) || v.getUint16(end + 6, true) || v.getUint16(end + 8, true) !== count) throw new Error('Unsupported split ZIP');
  if (count > 10_000 || start + size !== end) throw new Error('Invalid or oversized central directory');
  const names = new Set<string>();
  let cursor = start;
  let total = 0;
  for (let i = 0; i < count; i++) {
    if (cursor + 46 > end || v.getUint32(cursor, true) !== 0x02014b50) throw new Error('Invalid central entry');
    const inflated = v.getUint32(cursor + 24, true);
    if (inflated === 0xffffffff || v.getUint32(cursor + 42, true) === 0xffffffff) return 'zip64';
    total += inflated;
    if (inflated > 16 * 1024 * 1024 || total > 64 * 1024 * 1024) throw new Error('ZIP exceeds oracle inflation budget');
    const length = v.getUint16(cursor + 28, true);
    const next = cursor + 46 + length + v.getUint16(cursor + 30, true) + v.getUint16(cursor + 32, true);
    if (next > end) throw new Error('Truncated central entry');
    const name = new TextDecoder('utf-8', { fatal: true }).decode(bytes.subarray(cursor + 46, cursor + 46 + length));
    if (!name || name.startsWith('/') || name.includes('\\') || name.split('/').some(p => p === '.' || p === '..') || name.includes('\0')) throw new Error('Invalid package entry name');
    if (names.has(name)) throw new Error(`Duplicate ZIP entry ${name}`);
    names.add(name);
    const offset = v.getUint32(cursor + 42, true);
    if (offset + 30 > start || v.getUint32(offset, true) !== 0x04034b50) throw new Error('Invalid local header');
    const localLength = v.getUint16(offset + 26, true);
    const dataStart = offset + 30 + localLength + v.getUint16(offset + 28, true);
    if (dataStart + v.getUint32(cursor + 20, true) > start) throw new Error('ZIP data exceeds envelope');
    const localName = new TextDecoder('utf-8', { fatal: true }).decode(bytes.subarray(offset + 30, offset + 30 + localLength));
    if (localName !== name) throw new Error('Local and central names disagree');
    cursor = next;
  }
  if (cursor !== end) throw new Error('Central directory count disagrees');
  return 'zip32';
}

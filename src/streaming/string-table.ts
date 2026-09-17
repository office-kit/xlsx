import { serializeRichString } from '../workbook/shared-strings.js';
import type { CellStringWriter } from '../worksheet/writer.js';
import { SHEET_MAIN_NS } from '../xml/namespaces.js';
import type { StreamingEntryWriter } from '../zip/writer.js';

// Entry count bounds Map/array overhead; payload bounds long strings and rich
// text markup. UTF-16 accounting is deterministic, not an exact V8 heap metric.
const MAX_ENTRIES = 100_000;
const MAX_PAYLOAD_BYTES = 8 * 1024 * 1024;
const UTF16_BYTES_PER_UNIT = 2;
const XML_CHUNK_UNITS = 16 * 1024;

export const createWriteOnlyStringTable = () => {
  const entries: string[] = [];
  const plainIndex = new Map<string, number>();
  const richIndex = new Map<string, number>();
  let payloadBytes = 0;
  let full = false;

  const serialize: CellStringWriter = (value) => {
    const plain = typeof value === 'string';
    // Plain strings can hit the cache without escaping. Rich text is keyed by
    // its emitted XML, which snapshots caller-owned runs and font objects.
    const key = plain ? value : serializeRichString(value);
    const index = plain ? plainIndex : richIndex;
    const cached = index.get(key);
    if (cached !== undefined) return { type: 's', xml: `<v>${cached}</v>` };
    const body = plain ? serializeRichString(value) : key;
    // Count both keys and XML, conservatively even when they share storage.
    const bytes = (key.length + body.length) * UTF16_BYTES_PER_UNIT;
    if (!full && entries.length < MAX_ENTRIES && payloadBytes + bytes <= MAX_PAYLOAD_BYTES) {
      const id = entries.length;
      entries.push(body);
      index.set(key, id);
      payloadBytes += bytes;
      return { type: 's', xml: `<v>${id}</v>` };
    }
    // Never evict: earlier worksheet cells already reference these indices.
    // Once a new entry does not fit, all subsequent new values stay inline.
    full = true;
    return { type: 'inlineStr', xml: `<is>${body}</is>` };
  };

  const write = async (stream: StreamingEntryWriter): Promise<void> => {
    const encoder = new TextEncoder();
    let pending = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><sst xmlns="${SHEET_MAIN_NS}" count="${entries.length}" uniqueCount="${entries.length}">`;
    const append = (text: string): void => {
      pending += text;
      while (pending.length >= XML_CHUNK_UNITS) {
        // TextEncoder must see surrogate pairs together, including at a chunk
        // boundary. Otherwise an astral character becomes two replacements.
        let end = XML_CHUNK_UNITS;
        const last = pending.charCodeAt(end - 1);
        if (last >= 0xd800 && last <= 0xdbff) end--;
        stream.write(encoder.encode(pending.slice(0, end)));
        pending = pending.slice(end);
      }
    };
    for (const body of entries) append(`<si>${body}</si>`);
    append('</sst>');
    if (pending.length > 0) stream.write(encoder.encode(pending));
    await stream.end();
  };

  return { serialize, write, get size() { return entries.length; } };
};

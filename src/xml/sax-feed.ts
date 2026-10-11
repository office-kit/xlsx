// The saxes driver behind `iterParse` and the streaming row walk: input
// decoding in bounded feeds, the DTD prescan, and saxes errors rewrapped as
// `OpenXmlSchemaError`.
//
// Internal on purpose. Its signatures name saxes types, and a public
// declaration that imports them hands consumers saxes' own `.d.ts`, which
// does not type-check under `skipLibCheck: false`. Keep it out of every
// subpath's declaration graph.

import type { SaxesAttributeNS, SaxesParser } from 'saxes';
import { OpenXmlSchemaError } from '../utils/exceptions.js';
import type { SaxInput } from './iterparse.js';
import { qname } from './namespaces.js';

/**
 * Recognize declarations only in the prologue, skipping comments and processing
 * instructions. State spans chunks, but retained text is at most one keyword or
 * terminator. Once a root tag begins, saxes handles all remaining syntax checks.
 */
const makeDtdScanner = (): ((chunk: string) => void) => {
  let state: 'prologue' | 'markup' | 'comment' | 'pi' | 'done' = 'prologue';
  let token = '';
  return (chunk) => {
    for (const char of chunk) {
      if (state === 'done') return;
      if (state === 'comment' || state === 'pi') {
        const terminator = state === 'comment' ? '-->' : '?>';
        token = (token + char).slice(-terminator.length);
        if (token === terminator) {
          state = 'prologue';
          token = '';
        }
      } else if (state === 'prologue') {
        if (char === '<') {
          state = 'markup';
          token = char;
        }
      } else {
        token += char;
        if (token === '<!--' || token === '<?') {
          state = token === '<!--' ? 'comment' : 'pi';
          token = '';
        } else if (/^<!DOCTYPE[ \t\r\n]$/.test(token)) {
          throw new OpenXmlSchemaError('DTD declarations are not permitted in OOXML payloads');
        } else if (/^<!ENTITY[ \t\r\n]$/.test(token)) {
          throw new OpenXmlSchemaError('Entity declarations are not permitted in OOXML payloads');
        } else if (!'<!--'.startsWith(token) && !'<!DOCTYPE'.startsWith(token) && !'<!ENTITY'.startsWith(token)) {
          // A root name starts here (or invalid markup, which saxes rejects).
          state = 'done';
          token = '';
        }
      }
    }
  };
};

const isReadableStream = (v: unknown): v is ReadableStream<Uint8Array> => {
  return typeof v === 'object' && v !== null && typeof (v as ReadableStream).getReader === 'function';
};

const decoder = (): TextDecoder => new TextDecoder('utf-8', { fatal: false });

/**
 * Feed size, in code units. saxes runs its handlers synchronously inside
 * `write()`, so one `write()` queues every event its argument produces before
 * the consumer sees any of them. Capping the argument caps the queue.
 *
 * This bounds every input shape, including a `ReadableStream`, whose upstream
 * chunk size is not ours to choose: the zip reader pushes 64 KB of *compressed*
 * bytes per pull, and worksheet XML inflates roughly sevenfold, so its chunks
 * arrive around 450 KB.
 */
const FEED_CHUNK_SIZE = 64 * 1024;

/**
 * Split text that exceeds the feed size. Boundaries fall on arbitrary offsets,
 * which is safe because saxes holds back a lone high surrogate at the end of a
 * `write()` and rejoins it with the next one, so a split codepoint still
 * surfaces as a single text event.
 */
function* atFeedSize(text: string): IterableIterator<string> {
  if (text.length === 0) return;
  if (text.length <= FEED_CHUNK_SIZE) {
    yield text;
    return;
  }
  for (let i = 0; i < text.length; i += FEED_CHUNK_SIZE) {
    yield text.slice(i, i + FEED_CHUNK_SIZE);
  }
}

/**
 * Decode bytes in feed-sized slices. Slicing before the decode rather than
 * after it also bounds the decoder's output string, which matters when the
 * producer hands over a whole part at once. UTF-8 never expands, so a slice of
 * FEED_CHUNK_SIZE bytes decodes to at most that many code units and needs no
 * second split.
 *
 * `stream: true` holds back a codepoint split across a slice boundary instead
 * of emitting a replacement character for each half.
 */
function* decodeAtFeedSize(bytes: Uint8Array, td: TextDecoder): IterableIterator<string> {
  for (let i = 0; i < bytes.byteLength; i += FEED_CHUNK_SIZE) {
    const text = td.decode(bytes.subarray(i, Math.min(i + FEED_CHUNK_SIZE, bytes.byteLength)), { stream: true });
    if (text.length > 0) yield text;
  }
}

/** Decode any supported input into a sequence of feed-sized text chunks. */
async function* decodedChunks(input: SaxInput): AsyncIterableIterator<string> {
  if (typeof input === 'string') {
    yield* atFeedSize(input);
    return;
  }
  const td = decoder();
  if (input instanceof Uint8Array) {
    yield* decodeAtFeedSize(input, td);
  } else if (isReadableStream(input)) {
    const reader = input.getReader();
    try {
      for (;;) {
        const { done, value } = await reader.read();
        if (done) break;
        yield* decodeAtFeedSize(value, td);
      }
    } finally {
      // Consumers abandon the iteration routinely (`iterRows({ maxRow })`
      // returns as soon as the band ends). Without this the source stays
      // locked and its `cancel()` never runs, which is what releases the zip
      // reader's inflate state. A cancel that rejects is cleanup on a stream
      // nobody will read again, and must not mask why we left the loop.
      await reader.cancel().catch(() => {});
      reader.releaseLock();
    }
  } else {
    throw new OpenXmlSchemaError('iterParse: unsupported input type');
  }
  const tail = td.decode();
  if (tail.length > 0) yield tail;
}

export const buildAttrsClark = (attrs: Record<string, SaxesAttributeNS>): Record<string, string> => {
  const out: Record<string, string> = {};
  for (const [, info] of Object.entries(attrs)) {
    // saxes already resolved the namespace when xmlns: true is set; raw xmlns /
    // xmlns:* declarations have prefix='xmlns' (or local==='xmlns' when
    // default) and we drop those — they're rebuilt by the serializer.
    if (info.prefix === 'xmlns' || (info.prefix === '' && info.local === 'xmlns')) continue;
    const key = qname(info.uri, info.local);
    out[key] = info.value;
  }
  return out;
};

/**
 * Drive `parser` over `input`, pausing after every feed so the caller can hand
 * on whatever its handlers collected from that feed. Handlers registered on
 * the parser run synchronously inside each feed, which is what lets a hot
 * consumer skip the per-event allocation and async step of `iterParse`.
 *
 * Installs the `error` and `doctype` handlers itself: every caller needs the
 * same refusal of DTDs and the same `OpenXmlSchemaError` wrapping of saxes'
 * plain `Error`.
 */
export async function* feedSax(input: SaxInput, parser: SaxesParser): AsyncIterableIterator<void> {
  let pending: Error | undefined;
  parser.on('error', (err: Error) => {
    // saxes reports syntax errors as plain Error; consumers of this library
    // only ever see OpenXmlError subclasses.
    pending = new OpenXmlSchemaError(`Malformed XML: ${err.message}`, { cause: err });
  });
  parser.on('doctype', () => {
    pending = new OpenXmlSchemaError('DTD declarations are not permitted in OOXML payloads');
  });

  const scanForDtd = makeDtdScanner();
  for await (const chunk of decodedChunks(input)) {
    scanForDtd(chunk);
    parser.write(chunk);
    if (pending !== undefined) throw pending;
    yield;
  }

  parser.close();
  if (pending !== undefined) throw pending;
  yield;
}

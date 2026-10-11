// SAX iterator over OOXML XML payloads. Wraps `saxes` (XMLNS-aware) and yields
// a flat stream of {start | end | text} events with names already converted to
// Clark notation (`{ns}local`) — same shape as the DOM parser produces for
// static XmlNode trees, so consumers can switch between bulk and streaming
// reads without retouching name comparison.
//
// Phase 1 §3 acceptance: 1 k–row sheetData walked end-to-end with cell counts
// matching the source. The phase-4 read-only worksheet drives real-world use;
// this layer just produces the events.
//
// DOCTYPE / external entity declarations are forbidden. saxes does not expand
// external entities, and a prescan rejects a declaration before the parser
// reaches it. A declaration is only legal in the prologue, so that is how far
// the prescan reads.

import { SaxesParser } from 'saxes';
import { qname } from './namespaces.js';
import { buildAttrsClark, feedSax } from './sax-feed.js';

export type SaxEvent =
  | { kind: 'start'; name: string; attrs: Record<string, string> }
  | { kind: 'end'; name: string }
  | { kind: 'text'; text: string };

/**
 * Streamable input: `Uint8Array`, plain string, or a Web `ReadableStream` of
 * `Uint8Array` chunks (produced by xlsx zip entries via fflate, fetch, file
 * streams, etc.).
 */
export type SaxInput = Uint8Array | string | ReadableStream<Uint8Array>;

/**
 * Parse the input as a stream of SAX events. Element / attribute names are
 * returned in Clark notation (`{ns}local`).
 */
export async function* iterParse(input: SaxInput): AsyncIterableIterator<SaxEvent> {
  // xmlns: true gives us resolved {uri, local, prefix} on every open / close
  // tag and on every attribute.
  const parser = new SaxesParser({ xmlns: true, fragment: false });

  // Head-pointer ring instead of Array#shift: each saxes write() can produce
  // hundreds of events in a single synchronous batch (a `<row>` with dozens of
  // cells flushes one opentag + one text + one closetag per cell). `shift()`
  // is O(n) per element in V8, so a single-batch drain of N events would be
  // O(N²) before iteration. The head advances on yield; the queue is reset
  // (head + length) once it drains. Peak depth is one chunk's worth of events,
  // which is what keeps the queue bounded on a multi-GB sheet.
  let queue: SaxEvent[] = [];
  let head = 0;

  parser.on('opentag', (node) => {
    queue.push({ kind: 'start', name: qname(node.uri, node.local), attrs: buildAttrsClark(node.attributes) });
  });
  parser.on('closetag', (node) => {
    queue.push({ kind: 'end', name: qname(node.uri, node.local) });
  });
  const pushText = (text: string): void => {
    if (text.length > 0) queue.push({ kind: 'text', text });
  };
  parser.on('text', pushText);
  // A CDATA section is character data like any other. saxes reports it on its
  // own event, so without this `<t><![CDATA[a<b]]></t>` would read as empty.
  parser.on('cdata', pushText);

  const drain = function* (): IterableIterator<SaxEvent> {
    for (;;) {
      const ev = head < queue.length ? queue[head] : undefined;
      if (ev === undefined) {
        // Reset rather than grow forever; the next batch starts at index 0.
        queue = [];
        head = 0;
        return;
      }
      head++;
      yield ev;
    }
  };

  for await (const _ of feedSax(input, parser)) yield* drain();
}

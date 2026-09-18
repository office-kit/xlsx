// Error hierarchy for @office-kit/xlsx. Public APIs throw subclasses of OpenXmlError;
// internals chain via the `cause` option (Node 18+ / modern browsers all
// support Error.cause).
//
// Class is the one explicitly allowed exception to the no-class rule because
// Error subclasses are how `instanceof` discrimination is expressed in JS.

export interface OpenXmlErrorOptions {
  /** Underlying cause; preserved on the standard `cause` property. */
  cause?: unknown;
}

/**
 * Base class for every error this library throws. Catching it is how a caller
 * separates "the file was not what it claimed to be" from a bug.
 *
 * The contract, for a service that loads files it did not produce:
 *
 * - An `OpenXmlError` means the input was rejected. With the same bytes the
 *   load fails the same way, so reject the file and tell the user; retrying
 *   changes nothing. The one exception is a source that failed to hand over
 *   its bytes at all, described on {@link OpenXmlIoError}.
 * - Anything that is **not** an `OpenXmlError` is a bug in this library.
 *   Nothing here is meant to escape as a `TypeError` or a raw `fflate` /
 *   `saxes` error, so those are worth reporting.
 *
 * Which subclass arrives says where the input broke, not how badly, and the
 * subclass is stable for a given kind of damage. Message text is not: it names
 * parts, offsets and cell references to make a failure diagnosable, so it
 * changes freely between releases. Branch on the class, not on the message.
 */
export class OpenXmlError extends Error {
  override readonly name: string = 'OpenXmlError';

  constructor(message: string, options?: OpenXmlErrorOptions) {
    super(message, options as ErrorOptions);
  }
}

/**
 * Thrown for ZIP, file system, network or stream-level failures: the bytes
 * never arrived, or they did and are not a readable zip archive.
 *
 * This is the class a file that is not an xlsx at all lands on, because the
 * failure happens before any OOXML is parsed. Where a magic number identifies
 * the input (a PDF, a byte-order mark, a zip with no central directory) the
 * message names it, so a CSV renamed to `.xlsx` says so rather than only
 * "not a valid zip".
 *
 * It is also the one class that is not always permanent. `openZip: failed to
 * read source bytes` comes from the {@link XlsxSource}, not from the payload,
 * and wraps the underlying fs / fetch / stream error as its `cause`; a retry
 * can succeed. Every other `OpenXmlIoError` is a verdict on the bytes. When
 * the bytes are already in memory (`fromBuffer`) the source cannot fail, so
 * every `OpenXmlIoError` from that load is permanent.
 */
export class OpenXmlIoError extends OpenXmlError {
  override readonly name: string = 'OpenXmlIoError';
}

/**
 * Thrown when an OOXML payload violates structural / schema invariants: the
 * archive opened, and a part inside it does not parse or contradicts the spec
 * (unreadable XML, a missing required relationship, a cell whose declared type
 * does not match its value).
 *
 * On load this means a file that is a zip but not a usable xlsx, and it is
 * permanent. The same class also guards the write-side model, where it reports
 * the calling code's mistake rather than a file's: a duplicate sheet title, a
 * merge overlapping an existing one, a style id belonging to another
 * workbook's pool.
 */
export class OpenXmlSchemaError extends OpenXmlError {
  override readonly name = 'OpenXmlSchemaError';
}

/**
 * Thrown when a workbook is structurally valid OOXML but semantically broken.
 *
 * No path in the library throws this today; the semantic checks that exist all
 * report {@link OpenXmlSchemaError}. Catching it is therefore a dead branch.
 */
export class OpenXmlInvalidWorkbookError extends OpenXmlError {
  override readonly name = 'OpenXmlInvalidWorkbookError';
}

/**
 * Thrown for features the port has chosen not to implement (yet), including
 * input that is a real Office format this library does not read, such as an
 * encrypted xlsx or a legacy `.xls`.
 *
 * Permanent, and the only class where the user can act on it directly: the
 * message says what to do (decrypt the file, re-save it as `.xlsx`).
 */
export class OpenXmlNotImplementedError extends OpenXmlError {
  override readonly name = 'OpenXmlNotImplementedError';
}

/**
 * Thrown when an archive trips the decompression-bomb safeguards configured on
 * {@link openZip} / {@link loadWorkbook} / {@link loadWorkbookStream}. Subclass
 * of {@link OpenXmlIoError} so existing `catch (OpenXmlIoError)` paths still
 * see it, while letting callers branch on bomb-specific recovery (reject the
 * upload, log a security event, etc.).
 *
 * Permanent, and the one class that says the input may be hostile rather than
 * merely broken.
 */
export class OpenXmlDecompressionBombError extends OpenXmlIoError {
  override readonly name = 'OpenXmlDecompressionBombError';
}

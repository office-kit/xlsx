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

export class OpenXmlError extends Error {
  override readonly name: string = 'OpenXmlError';

  constructor(message: string, options?: OpenXmlErrorOptions) {
    super(message, options as ErrorOptions);
  }
}

/** Thrown for ZIP, file system, network or stream-level failures. */
export class OpenXmlIoError extends OpenXmlError {
  override readonly name: string = 'OpenXmlIoError';
}

/** Thrown when an OOXML payload violates structural / schema invariants. */
export class OpenXmlSchemaError extends OpenXmlError {
  override readonly name = 'OpenXmlSchemaError';
}

/** Thrown when a workbook is structurally valid OOXML but semantically broken. */
export class OpenXmlInvalidWorkbookError extends OpenXmlError {
  override readonly name = 'OpenXmlInvalidWorkbookError';
}

/** Thrown for features the port has chosen not to implement (yet). */
export class OpenXmlNotImplementedError extends OpenXmlError {
  override readonly name = 'OpenXmlNotImplementedError';
}

/**
 * Thrown when an archive trips the decompression-bomb safeguards configured on
 * {@link openZip} / {@link loadWorkbook} / {@link loadWorkbookStream}. Subclass
 * of {@link OpenXmlIoError} so existing `catch (OpenXmlIoError)` paths still
 * see it, while letting callers branch on bomb-specific recovery (reject the
 * upload, log a security event, etc.).
 */
export class OpenXmlDecompressionBombError extends OpenXmlIoError {
  override readonly name = 'OpenXmlDecompressionBombError';
}

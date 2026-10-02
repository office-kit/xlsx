import { OpenXmlSchemaError } from '../../src/utils/exceptions.js';

/** Missing oracle data is a failure, never a default that hides missing coverage. */
export function required<T>(value: T | undefined | null): T {
  if (value === undefined || value === null) throw new OpenXmlSchemaError('Missing conformance fixture data');
  return value;
}

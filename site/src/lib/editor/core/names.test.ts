import { expect, it } from 'vitest';
import { validateName } from './names.ts';

it('accepts names Excel accepts, including non-ASCII ones', () => {
  for (const name of ['Price', '_tax', '価格', 'Rate.2024', 'Row1Total', 'RCX']) expect(validateName(name), name).toBeUndefined();
});

it('rejects names that read as A1 or R1C1 references', () => {
  for (const name of ['A1', 'xfd1048576', 'R', 'c', 'R1', 'C12', 'RC', 'R1C1', 'rc5', 'R2C', '1abc', 'a b', '']) expect(validateName(name), name).toBe('invalidName');
});

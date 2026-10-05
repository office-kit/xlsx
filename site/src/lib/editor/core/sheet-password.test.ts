import { createHash } from 'node:crypto';
import { describe, expect, it } from 'vitest';
import { hashSheetPassword } from './sheet-password.ts';

// Independent reference built on node:crypto, following MS-OFFCRYPTO §2.3.7.1.
function reference(password: string, saltBase64: string, spinCount: number): string {
  let hash = createHash('sha512').update(Buffer.from(saltBase64, 'base64')).update(Buffer.from(password, 'utf16le')).digest();
  for (let i = 0; i < spinCount; i++) {
    const counter = Buffer.alloc(4);
    counter.writeUInt32LE(i);
    hash = createHash('sha512').update(hash).update(counter).digest();
  }
  return hash.toString('base64');
}

describe('hashSheetPassword', () => {
  it('matches the salted, iterated SHA-512 Excel writes', async () => {
    const salt = Buffer.from('0123456789abcdef').toString('base64');
    expect(await hashSheetPassword('secret', salt, 50)).toBe(reference('secret', salt, 50));
    expect(await hashSheetPassword('パスワード', salt, 3)).toBe(reference('パスワード', salt, 3));
  });
});

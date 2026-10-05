// Sheet-protection password hashing as Excel 2010+ writes it
// (`algorithmName="SHA-512"` with `hashValue` / `saltValue` / `spinCount`,
// ECMA-376 Part 1 §18.2.29 and MS-OFFCRYPTO §2.3.7.1): hash the salt followed
// by the UTF-16LE password, then re-hash the digest with a little-endian
// 32-bit iteration counter appended, `spinCount` times.

export const SHEET_HASH_ALGORITHM = 'SHA-512';
/** Excel's own iteration count. */
export const SHEET_SPIN_COUNT = 100_000;
const SALT_BYTES = 16;

function utf16le(text: string): Uint8Array {
  const out = new Uint8Array(text.length * 2);
  for (let i = 0; i < text.length; i++) {
    const code = text.charCodeAt(i);
    out[i * 2] = code & 0xff;
    out[i * 2 + 1] = code >> 8;
  }
  return out;
}

function concat(a: Uint8Array, b: Uint8Array): Uint8Array<ArrayBuffer> {
  const out = new Uint8Array(a.length + b.length);
  out.set(a, 0);
  out.set(b, a.length);
  return out;
}

function toBase64(bytes: Uint8Array): string {
  let binary = '';
  for (const b of bytes) binary += String.fromCharCode(b);
  return btoa(binary);
}

function fromBase64(text: string): Uint8Array {
  return Uint8Array.from(atob(text), (c) => c.charCodeAt(0));
}

export async function hashSheetPassword(password: string, saltBase64: string, spinCount: number): Promise<string> {
  let hash = new Uint8Array(await crypto.subtle.digest(SHEET_HASH_ALGORITHM, concat(fromBase64(saltBase64), utf16le(password))));
  const counter = new Uint8Array(4);
  const view = new DataView(counter.buffer);
  for (let i = 0; i < spinCount; i++) {
    view.setUint32(0, i, true);
    hash = new Uint8Array(await crypto.subtle.digest(SHEET_HASH_ALGORITHM, concat(hash, counter)));
  }
  return toBase64(hash);
}

/** A fresh salt and the matching hash, ready for `SheetProtection`. */
export async function protectionHash(password: string): Promise<{ algorithmName: string; hashValue: string; saltValue: string; spinCount: number }> {
  const saltValue = toBase64(crypto.getRandomValues(new Uint8Array(SALT_BYTES)));
  const hashValue = await hashSheetPassword(password, saltValue, SHEET_SPIN_COUNT);
  return { algorithmName: SHEET_HASH_ALGORITHM, hashValue, saltValue, spinCount: SHEET_SPIN_COUNT };
}

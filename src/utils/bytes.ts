// `instanceof Uint8Array` is false for arrays built in another realm — jsdom test environments, iframes,
// Node `vm` contexts — where TextEncoder and Buffer return instances of a different global constructor.

/** Realm-independent `value instanceof Uint8Array` (Buffer counts, as it subclasses Uint8Array). */
export function isUint8Array(value: unknown): value is Uint8Array {
  return ArrayBuffer.isView(value) && Object.prototype.toString.call(value) === '[object Uint8Array]';
}

/** Realm-independent `value instanceof ArrayBuffer`. */
export function isArrayBuffer(value: unknown): value is ArrayBuffer {
  return Object.prototype.toString.call(value) === '[object ArrayBuffer]';
}

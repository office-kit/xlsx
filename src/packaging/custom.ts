// `docProps/custom.xml` — user-defined document properties. Each <property>
// carries fmtid/pid/name attributes plus a single typed value child from the
// vt: namespace (vt:lpwstr, vt:i4, vt:bool, …).
//
// The Schema layer drives the <property> element attribute set; the typed-value
// child is stored as a raw XmlNode and the `make*Value` / `read*Value` helpers
// below cover the most common conversions.

import { OpenXmlSchemaError } from '../utils/exceptions.js';
import { CPROPS_FMTID, CUSTPROPS_NS, localNameOf, VTYPES_NS } from '../xml/namespaces.js';
import { parseXml } from '../xml/parser.js';
import { serializeXml } from '../xml/serializer.js';
import { el, type XmlNode } from '../xml/tree.js';

export interface CustomProperty {
  /** User-visible property name (must be unique within the workbook). */
  name: string;
  /** OOXML "Property Identifier"; ≥ 2. Auto-allocated if absent on append. */
  pid: number;
  /** Format ID GUID. Defaults to the well-known {D5CDD505-…} on append. */
  fmtid?: string;
  /** Typed value as a raw vt: element. Use the `make*Value` helpers. */
  value: XmlNode;
}

export interface CustomProperties {
  properties: CustomProperty[];
}

export function makeCustomProperties(): CustomProperties {
  return { properties: [] };
}

const PROPERTY_NAME = `{${CUSTPROPS_NS}}property`;
const PROPERTIES_NAME = `{${CUSTPROPS_NS}}Properties`;

// ---- typed-value constructors ----------------------------------------------

const vt = (local: string, text: string): XmlNode => el(`{${VTYPES_NS}}${local}`, {}, [], text);

export function makeStringValue(s: string): XmlNode {
  return vt('lpwstr', s);
}
export function makeAsciiStringValue(s: string): XmlNode {
  return vt('lpstr', s);
}
export function makeIntValue(n: number): XmlNode {
  if (!Number.isInteger(n)) throw new OpenXmlSchemaError(`makeIntValue: ${n} is not an integer`);
  return vt('i4', String(n));
}
export function makeDoubleValue(n: number): XmlNode {
  if (!Number.isFinite(n)) throw new OpenXmlSchemaError(`makeDoubleValue: ${n} is not finite`);
  return vt('r8', String(n));
}
export function makeBoolValue(b: boolean): XmlNode {
  return vt('bool', b ? '1' : '0');
}
export function makeFiletimeValue(iso: string): XmlNode {
  return vt('filetime', iso);
}
export function makeDateValue(iso: string): XmlNode {
  return vt('date', iso);
}

// ---- typed-value readers ---------------------------------------------------

export function readStringValue(v: XmlNode): string | undefined {
  const ln = localNameOf(v.name);
  if (ln === 'lpwstr' || ln === 'lpstr' || ln === 'bstr') return v.text ?? '';
  return undefined;
}
export function readIntValue(v: XmlNode): number | undefined {
  const ln = localNameOf(v.name);
  if (
    ln === 'i4' ||
    ln === 'i2' ||
    ln === 'i1' ||
    ln === 'int' ||
    ln === 'uint' ||
    ln === 'ui4' ||
    ln === 'ui2' ||
    ln === 'ui1'
  ) {
    const n = Number.parseInt(v.text ?? '', 10);
    return Number.isFinite(n) ? n : undefined;
  }
  return undefined;
}
export function readDoubleValue(v: XmlNode): number | undefined {
  const ln = localNameOf(v.name);
  if (ln === 'r4' || ln === 'r8' || ln === 'decimal' || ln === 'cy') {
    const n = Number.parseFloat(v.text ?? '');
    return Number.isFinite(n) ? n : undefined;
  }
  return undefined;
}
export function readBoolValue(v: XmlNode): boolean | undefined {
  if (localNameOf(v.name) !== 'bool') return undefined;
  const t = (v.text ?? '').toLowerCase();
  if (t === '1' || t === 'true' || t === 't') return true;
  if (t === '0' || t === 'false' || t === 'f') return false;
  return undefined;
}
export function readFiletimeValue(v: XmlNode): string | undefined {
  if (localNameOf(v.name) === 'filetime') return v.text ?? '';
  return undefined;
}

// ---- collection ops --------------------------------------------------------

const allocatePid = (props: CustomProperties): number => {
  // pid 0 / 1 are reserved by the OPC spec; user pids start at 2.
  const used = new Set<number>();
  for (const p of props.properties) used.add(p.pid);
  let n = 2;
  while (used.has(n)) n++;
  return n;
};

export function appendCustomProperty(
  props: CustomProperties,
  name: string,
  value: XmlNode,
  opts?: { pid?: number; fmtid?: string },
): CustomProperty {
  const pid = opts?.pid ?? allocatePid(props);
  const out: CustomProperty = { name, pid, value };
  if (opts?.fmtid !== undefined) out.fmtid = opts.fmtid;
  props.properties.push(out);
  return out;
}

export function findCustomPropertyByName(props: CustomProperties, name: string): CustomProperty | undefined {
  for (const p of props.properties) if (p.name === name) return p;
  return undefined;
}

// ---- bytes round-trip ------------------------------------------------------

export function customPropsToBytes(p: CustomProperties): Uint8Array {
  const root = el(PROPERTIES_NAME);
  for (const prop of p.properties) {
    const propEl = el(PROPERTY_NAME, {
      fmtid: prop.fmtid ?? CPROPS_FMTID,
      pid: String(prop.pid),
      name: prop.name,
    });
    propEl.children.push(prop.value);
    root.children.push(propEl);
  }
  return serializeXml(root);
}

export function customPropsFromBytes(bytes: Uint8Array | string): CustomProperties {
  const root = parseXml(bytes);
  if (root.name !== PROPERTIES_NAME) {
    throw new OpenXmlSchemaError(`customPropsFromBytes: expected <Properties>, got "${root.name}"`);
  }
  const properties: CustomProperty[] = [];
  for (const propEl of root.children) {
    if (propEl.name !== PROPERTY_NAME) continue;
    const fmtid = propEl.attrs['fmtid'];
    const pidRaw = propEl.attrs['pid'];
    const name = propEl.attrs['name'];
    if (name === undefined || pidRaw === undefined) {
      throw new OpenXmlSchemaError('custom.xml: <property> requires name and pid attributes');
    }
    const pid = Number.parseInt(pidRaw, 10);
    if (!Number.isFinite(pid)) {
      throw new OpenXmlSchemaError(`custom.xml: <property> pid is not an integer (got "${pidRaw}")`);
    }
    const value = propEl.children[0];
    if (value === undefined) {
      throw new OpenXmlSchemaError(`custom.xml: <property name="${name}"> has no typed-value child`);
    }
    const out: CustomProperty = { name, pid, value };
    if (fmtid !== undefined) out.fmtid = fmtid;
    properties.push(out);
  }
  return { properties };
}

// ---- Workbook ergonomic helpers ----------------------------------------

import type { Workbook } from '../workbook/workbook.js';

const ensureCustomProperties = (wb: Workbook): CustomProperties => {
  if (!wb.customProperties) wb.customProperties = makeCustomProperties();
  return wb.customProperties;
};

const replaceOrAppend = (
  wb: Workbook,
  name: string,
  value: XmlNode,
  opts?: { fmtid?: string },
): CustomProperty => {
  const props = ensureCustomProperties(wb);
  const existing = findCustomPropertyByName(props, name);
  if (existing) {
    existing.value = value;
    if (opts?.fmtid !== undefined) existing.fmtid = opts.fmtid;
    return existing;
  }
  return appendCustomProperty(props, name, value, opts);
};

/** Set (or replace) a custom string property. */
export const setCustomStringProperty = (wb: Workbook, name: string, value: string): CustomProperty =>
  replaceOrAppend(wb, name, makeStringValue(value));

/**
 * Set (or replace) a custom numeric property. Integers (within Int32 range) are
 * stored as `vt:i4`; non-integer / out-of-range numbers as `vt:r8` doubles.
 */
export const setCustomNumberProperty = (wb: Workbook, name: string, value: number): CustomProperty => {
  if (!Number.isFinite(value)) {
    throw new OpenXmlSchemaError(`setCustomNumberProperty: value "${value}" is not finite`);
  }
  const isInt32 = Number.isInteger(value) && value >= -2147483648 && value <= 2147483647;
  const node = isInt32 ? makeIntValue(value) : makeDoubleValue(value);
  return replaceOrAppend(wb, name, node);
};

/** Set (or replace) a custom boolean property. */
export const setCustomBoolProperty = (wb: Workbook, name: string, value: boolean): CustomProperty =>
  replaceOrAppend(wb, name, makeBoolValue(value));

/**
 * Set (or replace) a custom date property. Accepts a `Date` (converted to ISO
 * via `toISOString()`) or a pre-formatted W3C-DTF string.
 */
export const setCustomDateProperty = (
  wb: Workbook,
  name: string,
  value: Date | string,
): CustomProperty => {
  const iso = value instanceof Date ? value.toISOString() : value;
  return replaceOrAppend(wb, name, makeFiletimeValue(iso));
};

/**
 * Read the typed value of a custom property by name. Tries each decoder in turn
 * — string, int, double, bool, filetime — and returns the first hit. Returns
 * `undefined` for unknown names or unsupported types.
 */
export const getCustomPropertyValue = (
  wb: Workbook,
  name: string,
): string | number | boolean | undefined => {
  if (!wb.customProperties) return undefined;
  const prop = findCustomPropertyByName(wb.customProperties, name);
  if (!prop) return undefined;
  const s = readStringValue(prop.value);
  if (s !== undefined) return s;
  const i = readIntValue(prop.value);
  if (i !== undefined) return i;
  const d = readDoubleValue(prop.value);
  if (d !== undefined) return d;
  const b = readBoolValue(prop.value);
  if (b !== undefined) return b;
  const ft = readFiletimeValue(prop.value);
  if (ft !== undefined) return ft;
  return undefined;
};

/**
 * Remove a custom property by name. Returns `true` when one was removed,
 * `false` when the name wasn't found.
 */
export const removeCustomProperty = (wb: Workbook, name: string): boolean => {
  if (!wb.customProperties) return false;
  const arr = wb.customProperties.properties;
  const i = arr.findIndex((p) => p.name === name);
  if (i < 0) return false;
  arr.splice(i, 1);
  return true;
};

/** Every custom property, as a read-only view of the live array, not a copy. Empty array when the workbook carries none. */
export const listCustomProperties = (wb: Workbook): ReadonlyArray<CustomProperty> => {
  return wb.customProperties?.properties ?? [];
};

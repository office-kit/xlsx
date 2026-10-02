// Schema layer types — pure data, zero classes.
//
// A Schema<T> describes how a plain `T` is mapped to and from an XmlNode. The
// runtime serialiser (src/schema/serialize.ts) is a single switch-on-kind walk
// that needs nothing beyond these tables; bundlers can drop schemas a build
// path never imports because each lives in its own const.
//
// Lazy element schemas (`schema: () => SideSchema`) are how circular references
// — Border has Sides, Side has its own Color, etc. — stay expressible without
// hoisting hazards at module load.

/** Primitive coercion kinds attribute values may declare. */
export type Primitive = 'string' | 'int' | 'float' | 'bool';

export interface AttrDef {
  kind: Primitive | 'enum';
  /** Allowed values when kind === 'enum'. */
  values?: readonly string[];
  /** Treat the attribute as optional; missing → undefined. */
  optional?: boolean;
  /**
   * Default applied during fromTree when the attribute is absent. Never
   * auto-stripped during toTree — round-trip equivalence is the primary
   * concern; explicit emitters decide when to omit.
   */
  default?: unknown;
  /** Numeric range bounds for kind in {'int', 'float'}. */
  min?: number;
  max?: number;
  /** XML attribute name when it differs from the property key on T. */
  xmlName?: string;
  /** Namespace URI for the XML attribute. Empty / undefined = no namespace. */
  xmlNs?: string;
}

/**
 * Discriminated union of element shapes a schema can declare. Names are
 * unprefixed (`xmlNs` carries the namespace); the serialiser pairs them into
 * Clark names internally.
 */
export type ElementDef =
  | {
      kind: 'text';
      /** Property key on T — the value lands at `T[key]`. */
      key: string;
      /** XML element local name. Defaults to `key` when omitted. */
      name?: string;
      xmlNs?: string;
      primitive: Primitive;
      optional?: boolean;
      default?: unknown;
      /**
       * Fixed attributes always emitted on this element. Keys in Clark
       * notation. On parse, attributes other than these are ignored; keep the
       * schema strict only for the value (text content).
       *
       * Used for the `xsi:type="dcterms:W3CDTF"` marker that docProps/core.xml
       * emits on its <dcterms:created> / <dcterms:modified> children, and
       * similar fixed-marker patterns.
       */
      attrs?: Record<string, string>;
    }
  | {
      kind: 'object';
      key: string;
      name?: string;
      xmlNs?: string;
      // Schema is contravariant in T at the element boundary; named types
      // come back via T anyway. The `any` is intentional — `unknown` would
      // force a cast in every concrete schema definition with no extra safety.
      // oxlint-disable-next-line typescript/no-explicit-any
      schema: () => Schema<any>;
      optional?: boolean;
    }
  | {
      kind: 'sequence';
      key: string;
      /** Local name of the repeated child element. */
      itemName: string;
      itemNs?: string;
      // See ElementDef.kind === 'object' for why `any` survives here.
      // oxlint-disable-next-line typescript/no-explicit-any
      itemSchema: () => Schema<any>;
      /** Optional wrapper element holding the items. */
      container?: { name: string; xmlNs?: string; count?: boolean };
    }
  | {
      /**
       * Empty marker element — `<key/>` whose presence sets `T[key] = true`,
       * absence leaves it `undefined`. Elements with a boolean `val` attribute
       * use `nested` instead so explicit false values survive.
       */
      kind: 'empty';
      key: string;
      name?: string;
      xmlNs?: string;
    }
  | {
      /**
       * `<key val="value"/>` style child carrying a single primitive value via
       * an attribute (default `val`). This is openpyxl's NestedString /
       * NestedFloat / NestedInteger / NestedBool / NestedNoneSet pattern; it
       * shows up across Font, NumberFormat children and several chart
       * sub-elements.
       */
      kind: 'nested';
      key: string;
      name?: string;
      xmlNs?: string;
      primitive: Primitive | 'enum';
      /** Allowed values when primitive === 'enum'. */
      values?: readonly string[];
      /** Attribute name carrying the value. Defaults to "val". */
      valAttr?: string;
      /** Numeric range bounds (int / float). */
      min?: number;
      max?: number;
      optional?: boolean;
      default?: unknown;
    }
  | {
      /**
       * Opaque round-trip slot — fromTree stores the matched child as a raw
       * `XmlNode`, toTree splices it back verbatim. Used for subtrees we don't
       * want to model in detail (e.g. the vt:vector content under app.xml's
       * HeadingPairs / TitlesOfParts) but still need to preserve byte-for-byte
       * through edits.
       */
      kind: 'raw';
      key: string;
      name?: string;
      xmlNs?: string;
      optional?: boolean;
    };

export interface Schema<T> {
  /** Local element name. */
  tagname: string;
  /** Namespace URI of this element. */
  xmlNs?: string;
  /** Attribute table keyed by property name on T. */
  attrs: { readonly [K in keyof T & string]?: AttrDef };
  /** Child elements in declaration / emission order. */
  elements: ReadonlyArray<ElementDef>;
  /** Hook to normalise the parsed object (e.g., infer derived fields). */
  postParse?: (value: T, node: import('../xml/tree.js').XmlNode) => T;
  /** Hook to normalise the value before serialisation. */
  preSerialize?: (value: T) => T;
}

/**
 * Identity helper that pins inference to the supplied `T`. The runtime is just
 * `s => s`; the value matters only for the call-site type.
 */
export function defineSchema<T>(s: Schema<T>): Schema<T> {
  return s;
}

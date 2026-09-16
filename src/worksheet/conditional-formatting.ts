// Conditional formatting.
//
// Stage-1 covers the **value-based** rule kinds (cellIs / expression / top10 /
// aboveAverage / containsText family / containsBlanks family / duplicateValues
// / uniqueValues / timePeriod). The visual rule kinds — colorScale / dataBar /
// iconSet — round-trip as opaque inner XML (`innerXml` field) so the data
// survives a save / load cycle without our needing to model cfvo / colors /
// iconSets fully.

import { escapeXmlAttr } from '../utils/escape.js';
import { OpenXmlSchemaError } from '../utils/exceptions.js';
import { normalizeFormulaText } from '../utils/formula-text.js';
import { type MultiCellRange, parseMultiCellRange } from './cell-range.js';

export type ConditionalFormattingRuleType =
  | 'expression'
  | 'cellIs'
  | 'colorScale'
  | 'dataBar'
  | 'iconSet'
  | 'top10'
  | 'aboveAverage'
  | 'uniqueValues'
  | 'duplicateValues'
  | 'containsText'
  | 'notContainsText'
  | 'beginsWith'
  | 'endsWith'
  | 'containsBlanks'
  | 'notContainsBlanks'
  | 'containsErrors'
  | 'notContainsErrors'
  | 'timePeriod';

export type CellIsOperator =
  | 'lessThan'
  | 'lessThanOrEqual'
  | 'equal'
  | 'notEqual'
  | 'greaterThanOrEqual'
  | 'greaterThan'
  | 'between'
  | 'notBetween';

export type TextOperator = 'containsText' | 'notContains' | 'beginsWith' | 'endsWith';

export type TimePeriod =
  | 'today'
  | 'yesterday'
  | 'tomorrow'
  | 'last7Days'
  | 'thisMonth'
  | 'lastMonth'
  | 'nextMonth'
  | 'thisWeek'
  | 'lastWeek'
  | 'nextWeek';

export interface ConditionalFormattingRule {
  /** Wire-level rule kind. */
  type: ConditionalFormattingRuleType;
  /** 1-based priority — Excel evaluates lower priority first. */
  priority: number;
  /** Index into Stylesheet.dxfs for the cell-format applied when the rule fires. */
  dxfId?: number;
  /** Stop evaluating subsequent rules on the same cell when this rule matches. */
  stopIfTrue?: boolean;
  /** cellIs operator. */
  operator?: CellIsOperator | TextOperator | string;
  /** Comparison string for the contains-text family. */
  text?: string;
  /** top10: rank percentage flag. */
  percent?: boolean;
  /** top10: bottom-N flag (false = top-N, true = bottom-N). */
  bottom?: boolean;
  /** top10: rank value (defaults to 10). */
  rank?: number;
  /** aboveAverage: aboveAverage="0" → below average. */
  aboveAverage?: boolean;
  /** aboveAverage: equalAverage flag. */
  equalAverage?: boolean;
  /** aboveAverage: stdDev. */
  stdDev?: number;
  /** timePeriod token. */
  timePeriod?: TimePeriod;
  /** 0..3 formula strings — varies by rule type. */
  formulas: string[];
  /**
   * Raw inner XML for colorScale / dataBar / iconSet rules. Stage-1 stores the
   * verbatim child markup so saves round-trip without our needing to model cfvo
   * / colors / iconSets fully.
   */
  innerXml?: string;
}

export interface ConditionalFormatting {
  sqref: MultiCellRange;
  rules: ConditionalFormattingRule[];
  pivot?: boolean;
}

export function makeConditionalFormatting(opts: {
  sqref: MultiCellRange | string;
  rules?: ConditionalFormattingRule[];
  pivot?: boolean;
}): ConditionalFormatting {
  return {
    sqref: typeof opts.sqref === 'string' ? parseMultiCellRange(opts.sqref) : opts.sqref,
    rules: opts.rules ?? [],
    ...(opts.pivot !== undefined ? { pivot: opts.pivot } : {}),
  };
}

export function makeCfRule(
  opts: Partial<ConditionalFormattingRule> & {
    type: ConditionalFormattingRuleType;
    priority: number;
  },
): ConditionalFormattingRule {
  return {
    type: opts.type,
    priority: opts.priority,
    formulas: (opts.formulas ?? []).map((f) => normalizeFormulaText(f)),
    ...(opts.dxfId !== undefined ? { dxfId: opts.dxfId } : {}),
    ...(opts.stopIfTrue !== undefined ? { stopIfTrue: opts.stopIfTrue } : {}),
    ...(opts.operator !== undefined ? { operator: opts.operator } : {}),
    ...(opts.text !== undefined ? { text: opts.text } : {}),
    ...(opts.percent !== undefined ? { percent: opts.percent } : {}),
    ...(opts.bottom !== undefined ? { bottom: opts.bottom } : {}),
    ...(opts.rank !== undefined ? { rank: opts.rank } : {}),
    ...(opts.aboveAverage !== undefined ? { aboveAverage: opts.aboveAverage } : {}),
    ...(opts.equalAverage !== undefined ? { equalAverage: opts.equalAverage } : {}),
    ...(opts.stdDev !== undefined ? { stdDev: opts.stdDev } : {}),
    ...(opts.timePeriod !== undefined ? { timePeriod: opts.timePeriod } : {}),
    ...(opts.innerXml !== undefined ? { innerXml: opts.innerXml } : {}),
  };
}

// ---- Worksheet ergonomic builders ---------------------------------------

import type { Worksheet } from './worksheet.js';

const resolveCfSqref = (sqref: MultiCellRange | string): MultiCellRange =>
  typeof sqref === 'string' ? parseMultiCellRange(sqref) : sqref;

const nextCfPriority = (ws: Worksheet): number => {
  let max = 0;
  for (const cf of ws.conditionalFormatting) {
    for (const r of cf.rules) {
      if (r.priority > max) max = r.priority;
    }
  }
  return max + 1;
};

/** Push one CF rule onto the worksheet, wrapping it in a ConditionalFormatting block keyed off `sqref`. */
const pushRule = (
  ws: Worksheet,
  sqref: MultiCellRange | string,
  rule: ConditionalFormattingRule,
): ConditionalFormattingRule => {
  ws.conditionalFormatting.push(makeConditionalFormatting({ sqref: resolveCfSqref(sqref), rules: [rule] }));
  return rule;
};

/**
 * "If cell value [op] formula → apply dxf". Mirrors Excel's "Highlight Cell
 * Rules → ..." UI.
 */
export const addCellIsRule = (
  ws: Worksheet,
  sqref: MultiCellRange | string,
  opts: {
    operator: CellIsOperator;
    formula1: string;
    formula2?: string;
    dxfId?: number;
    stopIfTrue?: boolean;
    priority?: number;
  },
): ConditionalFormattingRule => {
  const formulas = opts.formula2 !== undefined ? [opts.formula1, opts.formula2] : [opts.formula1];
  return pushRule(
    ws,
    sqref,
    makeCfRule({
      type: 'cellIs',
      priority: opts.priority ?? nextCfPriority(ws),
      operator: opts.operator,
      formulas,
      ...(opts.dxfId !== undefined ? { dxfId: opts.dxfId } : {}),
      ...(opts.stopIfTrue !== undefined ? { stopIfTrue: opts.stopIfTrue } : {}),
    }),
  );
};

/** Top-N or bottom-N rule. `bottom: true` → bottom-N. `percent: true` → percentile. */
export const addTopNRule = (
  ws: Worksheet,
  sqref: MultiCellRange | string,
  opts: { rank?: number; bottom?: boolean; percent?: boolean; dxfId?: number; priority?: number },
): ConditionalFormattingRule => {
  return pushRule(
    ws,
    sqref,
    makeCfRule({
      type: 'top10',
      priority: opts.priority ?? nextCfPriority(ws),
      formulas: [],
      ...(opts.rank !== undefined ? { rank: opts.rank } : {}),
      ...(opts.bottom !== undefined ? { bottom: opts.bottom } : {}),
      ...(opts.percent !== undefined ? { percent: opts.percent } : {}),
      ...(opts.dxfId !== undefined ? { dxfId: opts.dxfId } : {}),
    }),
  );
};

/** Above/below-average rule. `aboveAverage: false` → below-average; provide `stdDev` for ±N stddev. */
export const addAverageRule = (
  ws: Worksheet,
  sqref: MultiCellRange | string,
  opts: {
    aboveAverage?: boolean;
    equalAverage?: boolean;
    stdDev?: number;
    dxfId?: number;
    priority?: number;
  },
): ConditionalFormattingRule => {
  return pushRule(
    ws,
    sqref,
    makeCfRule({
      type: 'aboveAverage',
      priority: opts.priority ?? nextCfPriority(ws),
      formulas: [],
      ...(opts.aboveAverage !== undefined ? { aboveAverage: opts.aboveAverage } : {}),
      ...(opts.equalAverage !== undefined ? { equalAverage: opts.equalAverage } : {}),
      ...(opts.stdDev !== undefined ? { stdDev: opts.stdDev } : {}),
      ...(opts.dxfId !== undefined ? { dxfId: opts.dxfId } : {}),
    }),
  );
};

/** Duplicate-values rule (each cell whose value appears more than once gets the dxf). */
export const addDuplicateValuesRule = (
  ws: Worksheet,
  sqref: MultiCellRange | string,
  opts: { dxfId?: number; priority?: number; unique?: boolean } = {},
): ConditionalFormattingRule => {
  return pushRule(
    ws,
    sqref,
    makeCfRule({
      type: opts.unique ? 'uniqueValues' : 'duplicateValues',
      priority: opts.priority ?? nextCfPriority(ws),
      formulas: [],
      ...(opts.dxfId !== undefined ? { dxfId: opts.dxfId } : {}),
    }),
  );
};

/** Free-form formula rule — any `=ISNUMBER(A1)`-style boolean expression. */
export const addFormulaRule = (
  ws: Worksheet,
  sqref: MultiCellRange | string,
  opts: { formula: string; dxfId?: number; stopIfTrue?: boolean; priority?: number },
): ConditionalFormattingRule => {
  return pushRule(
    ws,
    sqref,
    makeCfRule({
      type: 'expression',
      priority: opts.priority ?? nextCfPriority(ws),
      formulas: [opts.formula],
      ...(opts.dxfId !== undefined ? { dxfId: opts.dxfId } : {}),
      ...(opts.stopIfTrue !== undefined ? { stopIfTrue: opts.stopIfTrue } : {}),
    }),
  );
};

/** Text-based rule (containsText / notContains / beginsWith / endsWith). */
export const addTextRule = (
  ws: Worksheet,
  sqref: MultiCellRange | string,
  opts: { operator: TextOperator; text: string; dxfId?: number; priority?: number },
): ConditionalFormattingRule => {
  // Excel uses different `type` tokens for the four text operators.
  const typeMap: Record<TextOperator, ConditionalFormattingRuleType> = {
    containsText: 'containsText',
    notContains: 'notContainsText',
    beginsWith: 'beginsWith',
    endsWith: 'endsWith',
  };
  return pushRule(
    ws,
    sqref,
    makeCfRule({
      type: typeMap[opts.operator],
      priority: opts.priority ?? nextCfPriority(ws),
      operator: opts.operator,
      text: opts.text,
      formulas: [],
      ...(opts.dxfId !== undefined ? { dxfId: opts.dxfId } : {}),
    }),
  );
};

// ---- Visual rule builders (colorScale / dataBar / iconSet) --------------

export type CfvoType = 'min' | 'max' | 'num' | 'percent' | 'percentile' | 'formula';

export interface Cfvo {
  type: CfvoType;
  /** Required for num / percent / percentile / formula; ignored for min / max. */
  val?: string;
}

export type IconSetStyle =
  | '3Arrows'
  | '3ArrowsGray'
  | '3Flags'
  | '3Signs'
  | '3Symbols'
  | '3Symbols2'
  | '3TrafficLights1'
  | '3TrafficLights2'
  | '4Arrows'
  | '4ArrowsGray'
  | '4Rating'
  | '4RedToBlack'
  | '4TrafficLights'
  | '5Arrows'
  | '5ArrowsGray'
  | '5Quarters'
  | '5Rating';

const escapeAttr = escapeXmlAttr;

const renderCfvo = (c: Cfvo): string => {
  if (c.type === 'min' || c.type === 'max') {
    if (c.val === undefined) return `<cfvo type="${c.type}"/>`;
    return `<cfvo type="${c.type}" val="${escapeAttr(c.val)}"/>`;
  }
  if (c.val === undefined) {
    throw new OpenXmlSchemaError(`cfvo type "${c.type}" requires val`);
  }
  return `<cfvo type="${c.type}" val="${escapeAttr(c.val)}"/>`;
};

const renderColor = (hex: string): string => `<color rgb="${escapeAttr(hex)}"/>`;

/**
 * Color-scale rule — gradient between 2 or 3 reference points. Each cfvo pairs
 * with one color.
 */
export const addColorScaleRule = (
  ws: Worksheet,
  sqref: MultiCellRange | string,
  opts: {
    cfvos: ReadonlyArray<Cfvo>;
    /** Hex strings, e.g. `'FFFF0000'`; one per cfvo. */
    colors: ReadonlyArray<string>;
    priority?: number;
    stopIfTrue?: boolean;
  },
): ConditionalFormattingRule => {
  if (opts.cfvos.length !== 2 && opts.cfvos.length !== 3) {
    throw new OpenXmlSchemaError(`addColorScaleRule: cfvos must be length 2 or 3; got ${opts.cfvos.length}`);
  }
  if (opts.colors.length !== opts.cfvos.length) {
    throw new OpenXmlSchemaError(
      `addColorScaleRule: colors length (${opts.colors.length}) must match cfvos length (${opts.cfvos.length})`,
    );
  }
  const inner = `<colorScale>${opts.cfvos.map(renderCfvo).join('')}${opts.colors.map(renderColor).join('')}</colorScale>`;
  return pushRule(
    ws,
    sqref,
    makeCfRule({
      type: 'colorScale',
      priority: opts.priority ?? nextCfPriority(ws),
      formulas: [],
      innerXml: inner,
      ...(opts.stopIfTrue !== undefined ? { stopIfTrue: opts.stopIfTrue } : {}),
    }),
  );
};

/**
 * Data-bar rule — a gradient bar inside each cell sized to the value. Defaults
 * to `min`/`max` cfvos so the bar spans the visible range.
 */
export const addDataBarRule = (
  ws: Worksheet,
  sqref: MultiCellRange | string,
  opts: {
    color: string;
    minCfvo?: Cfvo;
    maxCfvo?: Cfvo;
    minLength?: number;
    maxLength?: number;
    showValue?: boolean;
    priority?: number;
    stopIfTrue?: boolean;
  },
): ConditionalFormattingRule => {
  const min = opts.minCfvo ?? { type: 'min' };
  const max = opts.maxCfvo ?? { type: 'max' };
  let attrs = '';
  if (opts.minLength !== undefined) attrs += ` minLength="${opts.minLength}"`;
  if (opts.maxLength !== undefined) attrs += ` maxLength="${opts.maxLength}"`;
  if (opts.showValue !== undefined) attrs += ` showValue="${opts.showValue ? '1' : '0'}"`;
  const inner = `<dataBar${attrs}>${renderCfvo(min)}${renderCfvo(max)}${renderColor(opts.color)}</dataBar>`;
  return pushRule(
    ws,
    sqref,
    makeCfRule({
      type: 'dataBar',
      priority: opts.priority ?? nextCfPriority(ws),
      formulas: [],
      innerXml: inner,
      ...(opts.stopIfTrue !== undefined ? { stopIfTrue: opts.stopIfTrue } : {}),
    }),
  );
};

/**
 * Icon-set rule — visual icons (arrows / lights / flags) drawn next to each
 * cell value. The number of cfvos depends on the icon set (3 for `3Arrows`, 4
 * for `4Arrows`, 5 for `5Arrows`, etc).
 */
export const addIconSetRule = (
  ws: Worksheet,
  sqref: MultiCellRange | string,
  opts: {
    iconSet: IconSetStyle | string;
    cfvos: ReadonlyArray<Cfvo>;
    reverse?: boolean;
    showValue?: boolean;
    percent?: boolean;
    priority?: number;
    stopIfTrue?: boolean;
  },
): ConditionalFormattingRule => {
  if (opts.cfvos.length < 3 || opts.cfvos.length > 5) {
    throw new OpenXmlSchemaError(`addIconSetRule: cfvos must be length 3..5; got ${opts.cfvos.length}`);
  }
  let attrs = ` iconSet="${escapeAttr(opts.iconSet)}"`;
  if (opts.reverse !== undefined) attrs += ` reverse="${opts.reverse ? '1' : '0'}"`;
  if (opts.showValue !== undefined) attrs += ` showValue="${opts.showValue ? '1' : '0'}"`;
  if (opts.percent !== undefined) attrs += ` percent="${opts.percent ? '1' : '0'}"`;
  const inner = `<iconSet${attrs}>${opts.cfvos.map(renderCfvo).join('')}</iconSet>`;
  return pushRule(
    ws,
    sqref,
    makeCfRule({
      type: 'iconSet',
      priority: opts.priority ?? nextCfPriority(ws),
      formulas: [],
      innerXml: inner,
      ...(opts.stopIfTrue !== undefined ? { stopIfTrue: opts.stopIfTrue } : {}),
    }),
  );
};

// Function registry: one table drives evaluation, volatility and the
// catalog the formula-bar autocomplete shows.

import type { FunctionCategory, FunctionSpec } from '../function-spec.ts';
import { normalizeFunctionName } from '../parser.ts';
import { CONDITIONAL_FUNCTIONS } from './conditional.ts';
import { DATABASE_FUNCTIONS } from './database.ts';
import { DATETIME_FUNCTIONS } from './datetime.ts';
import { ENGINEERING_FUNCTIONS } from './engineering.ts';
import { FINANCIAL_FUNCTIONS } from './financial.ts';
import { INFO_FUNCTIONS } from './info.ts';
import { LOGICAL_FUNCTIONS } from './logical.ts';
import { LOOKUP_FUNCTIONS } from './lookup.ts';
import { MATH_FUNCTIONS } from './math.ts';
import { STATISTICAL_FUNCTIONS } from './stats.ts';
import { TEXT_FUNCTIONS } from './text.ts';

const ALL: readonly FunctionSpec[] = [
  ...MATH_FUNCTIONS,
  ...CONDITIONAL_FUNCTIONS,
  ...STATISTICAL_FUNCTIONS,
  ...LOGICAL_FUNCTIONS,
  ...LOOKUP_FUNCTIONS,
  ...TEXT_FUNCTIONS,
  ...DATETIME_FUNCTIONS,
  ...INFO_FUNCTIONS,
  ...FINANCIAL_FUNCTIONS,
  ...ENGINEERING_FUNCTIONS,
  ...DATABASE_FUNCTIONS,
];

export const FUNCTIONS: ReadonlyMap<string, FunctionSpec> = new Map(ALL.map((f) => [f.name, f]));

export interface FunctionInfo {
  readonly name: string;
  readonly category: FunctionCategory;
  /** Signature as Excel's tooltip shows it, e.g. `SUM(number1, [number2], ...)`. */
  readonly syntax: string;
  /** One English sentence. */
  readonly description: string;
}

export const FUNCTION_CATALOG: ReadonlyArray<FunctionInfo> = [...FUNCTIONS.values()]
  .map(({ name, category, syntax, description }) => ({ name, category, syntax, description }))
  .sort((a, b) => (a.name < b.name ? -1 : a.name > b.name ? 1 : 0));

/** Functions whose result can change without any precedent changing (RAND, NOW, OFFSET, INDIRECT…). */
export function isVolatileFunction(name: string): boolean {
  return FUNCTIONS.get(normalizeFunctionName(name))?.volatile === true;
}

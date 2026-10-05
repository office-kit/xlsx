// Formula text as Excel stores it in a file versus as a user types it.
//
// Functions added after Excel 2007 are written with a `_xlfn.` prefix
// (`_xlfn.XLOOKUP`, `_xlfn._xlws.SORT`), and LET / LAMBDA parameters with
// `_xlpm.`. Excel shows #NAME? for a bare `XLOOKUP` in a file, so formulas the
// editor writes must go through `toStorageFormula`; text read from a file goes
// through `fromStorageFormula` before it is shown in the formula bar.

import { renderArea, renderPrefix } from './address.ts';
import type { RefArea } from './ast.ts';
import { FUNCTIONS } from './functions/index.ts';
import { type Token, tokenize } from './lexer.ts';
import { normalizeFunctionName } from './parser.ts';

const XLWS = new Set(['FILTER', 'SORT']);

const FUTURE = new Set([
  'ACOT', 'ACOTH', 'AGGREGATE', 'ARABIC', 'ARRAYTOTEXT', 'BASE', 'BETA.DIST', 'BETA.INV', 'BINOM.DIST', 'BINOM.DIST.RANGE', 'BINOM.INV',
  'BITAND', 'BITLSHIFT', 'BITOR', 'BITRSHIFT', 'BITXOR', 'BYCOL', 'BYROW', 'CEILING.MATH', 'CEILING.PRECISE', 'CHISQ.DIST',
  'CHISQ.DIST.RT', 'CHISQ.INV', 'CHISQ.INV.RT', 'CHISQ.TEST', 'CHOOSECOLS', 'CHOOSEROWS', 'COMBINA', 'CONCAT', 'CONFIDENCE.NORM',
  'CONFIDENCE.T', 'COT', 'COTH', 'COVARIANCE.P', 'COVARIANCE.S', 'CSC', 'CSCH', 'DAYS', 'DECIMAL', 'DROP', 'ERF.PRECISE',
  'ERFC.PRECISE', 'EXPAND', 'EXPON.DIST', 'F.DIST', 'F.DIST.RT', 'F.INV', 'F.INV.RT', 'F.TEST', 'FILTER', 'FLOOR.MATH',
  'FLOOR.PRECISE', 'FORECAST.LINEAR', 'FORMULATEXT', 'GAMMA', 'GAMMA.DIST', 'GAMMA.INV', 'GAMMALN.PRECISE', 'GAUSS', 'HSTACK',
  'IFNA', 'IFS', 'ISFORMULA', 'ISOMITTED', 'ISOWEEKNUM', 'LAMBDA', 'LET', 'LOGNORM.DIST', 'LOGNORM.INV', 'MAKEARRAY', 'MAP',
  'MAXIFS', 'MINIFS', 'MODE.MULT', 'MODE.SNGL', 'MUNIT', 'NEGBINOM.DIST', 'NETWORKDAYS.INTL', 'NORM.DIST', 'NORM.INV',
  'NORM.S.DIST', 'NORM.S.INV', 'NUMBERVALUE', 'PDURATION', 'PERCENTILE.EXC', 'PERCENTILE.INC', 'PERCENTRANK.EXC',
  'PERCENTRANK.INC', 'PERMUTATIONA', 'PHI', 'POISSON.DIST', 'QUARTILE.EXC', 'QUARTILE.INC', 'RANDARRAY', 'RANK.AVG', 'RANK.EQ',
  'REDUCE', 'RRI', 'SCAN', 'SEC', 'SECH', 'SEQUENCE', 'SHEET', 'SHEETS', 'SKEW.P', 'SORT', 'SORTBY', 'STDEV.P', 'STDEV.S', 'SWITCH',
  'T.DIST', 'T.DIST.2T', 'T.DIST.RT', 'T.INV', 'T.INV.2T', 'T.TEST', 'TAKE', 'TEXTAFTER', 'TEXTBEFORE', 'TEXTJOIN', 'TEXTSPLIT',
  'TOCOL', 'TOROW', 'UNICHAR', 'UNICODE', 'UNIQUE', 'VALUETOTEXT', 'VAR.P', 'VAR.S', 'VSTACK', 'WEIBULL.DIST', 'WORKDAY.INTL',
  'WRAPCOLS', 'WRAPROWS', 'XLOOKUP', 'XMATCH', 'XOR', 'Z.TEST',
]);

const PARAM_PREFIX = '_xlpm.';

type Edit = { start: number; end: number; text: string };

const apply = (text: string, edits: Edit[]): string => {
  let out = '';
  let at = 0;
  for (const e of edits.sort((a, b) => a.start - b.start)) {
    out += text.slice(at, e.start) + e.text;
    at = e.end;
  }
  return out + text.slice(at);
};

/** Upper-cased names declared as LET / LAMBDA parameters anywhere in the formula. */
const parameterNames = (tokens: readonly Token[]): Set<string> => {
  const params = new Set<string>();
  for (let i = 0; i < tokens.length; i++) {
    const t = tokens[i];
    if (t?.kind !== 'func') continue;
    const fn = normalizeFunctionName(t.name);
    if (fn !== 'LET' && fn !== 'LAMBDA') continue;
    // Walk the call's top-level arguments; a parameter is an argument made of one bare name.
    let depth = 0;
    let argTokens: Token[] = [];
    const args: Token[][] = [];
    for (let j = i + 1; j < tokens.length; j++) {
      const u = tokens[j];
      if (u === undefined) break;
      if (u.kind === 'lparen' || u.kind === 'lbrace') {
        depth++;
        if (depth === 1) continue;
      } else if (u.kind === 'rparen' || u.kind === 'rbrace') {
        depth--;
        if (depth === 0) {
          args.push(argTokens);
          break;
        }
      } else if (u.kind === 'comma' && depth === 1) {
        args.push(argTokens);
        argTokens = [];
        continue;
      }
      argTokens.push(u);
    }
    args.forEach((a, k) => {
      const isParam = fn === 'LET' ? k % 2 === 0 && k < args.length - 1 : k < args.length - 1;
      const only = a[0];
      if (isParam && a.length === 1 && only?.kind === 'name' && only.prefix === undefined) {
        params.add(only.name.replace(/^_xlpm\./i, '').toUpperCase());
      }
    });
  }
  return params;
};

/** Add the `_xlfn.` / `_xlws.` / `_xlpm.` prefixes Excel expects in a file. */
/** Each axis's smaller end first; a `$` stays with the row or column it was written on. */
function orderArea(a: RefArea): RefArea {
  const flipRows = a.r1 > a.r2;
  const flipCols = a.c1 > a.c2;
  if (!flipRows && !flipCols) return a;
  return {
    kind: a.kind,
    r1: flipRows ? a.r2 : a.r1,
    r1Abs: flipRows ? a.r2Abs : a.r1Abs,
    r2: flipRows ? a.r1 : a.r2,
    r2Abs: flipRows ? a.r1Abs : a.r2Abs,
    c1: flipCols ? a.c2 : a.c1,
    c1Abs: flipCols ? a.c2Abs : a.c1Abs,
    c2: flipCols ? a.c1 : a.c2,
    c2Abs: flipCols ? a.c1Abs : a.c2Abs,
  };
}

/** Longest plain-decimal literal Excel keeps; past it (1E+21, 1E-20) it writes an exponent. */
const MAX_PLAIN_LITERAL = 21;

/** A number literal the way Excel stores it: 15 significant digits, =1E3 → =1000, =.5 → =0.5. */
function numberLiteral(value: number): string {
  const v = Number(value.toPrecision(15));
  if (v === 0) return '0';
  const [mantissa = '', expText = '0'] = v.toExponential().split('e');
  const exp = Number(expText);
  const digits = mantissa.replace('.', '');
  let plain: string;
  if (exp < 0) plain = `0.${'0'.repeat(-exp - 1)}${digits}`;
  else if (digits.length <= exp + 1) plain = digits + '0'.repeat(exp + 1 - digits.length);
  else plain = `${digits.slice(0, exp + 1)}.${digits.slice(exp + 1)}`;
  if (plain.length <= MAX_PLAIN_LITERAL) return plain;
  return `${mantissa}E${exp < 0 ? '-' : '+'}${Math.abs(exp)}`;
}

/**
 * `sheetTitles`, when given, corrects a typed sheet name's case to the real
 * sheet's (`=sheet1!a1` → `=Sheet1!A1`), as Excel does.
 */
export function toStorageFormula(text: string, sheetTitles: readonly string[] = []): string {
  const titles = new Map(sheetTitles.map((t) => [t.toLowerCase(), t]));
  const realTitle = (name: string): string => titles.get(name.toLowerCase()) ?? name;
  const tokens = tokenize(text, true);
  const params = parameterNames(tokens);
  const edits: Edit[] = [];
  for (const t of tokens) {
    if (t.kind === 'func') {
      const name = normalizeFunctionName(t.name);
      const upper = t.name.toUpperCase();
      const typedPrefix = t.name.slice(0, t.name.length - name.length).toLowerCase();
      const prefix = FUTURE.has(name) ? (XLWS.has(name) ? '_xlfn._xlws.' : '_xlfn.') : typedPrefix;
      if (FUNCTIONS.has(name)) {
        // Excel upper-cases the functions it knows (`=sum(a1)` → `=SUM(A1)`)
        // and leaves a misspelt one as typed.
        edits.push({ start: t.start, end: t.start + t.name.length, text: prefix + name });
      } else if (params.has(upper)) {
        edits.push({ start: t.start, end: t.start, text: PARAM_PREFIX });
      } else if (name === upper && FUTURE.has(name)) {
        edits.push({ start: t.start, end: t.start, text: prefix });
      }
    } else if (t.kind === 'name' && t.prefix === undefined && params.has(t.name.toUpperCase())) {
      edits.push({ start: t.start, end: t.start, text: PARAM_PREFIX });
    } else if (t.kind === 'ref' && t.area !== undefined) {
      // Excel stores a reference upper-cased, with a sheet name quoted only
      // when it must be, and its corners ordered: =SUM(A10:A3) → =SUM(A3:A10).
      const p = t.prefix;
      const prefix = p
        ? renderPrefix(p.external !== undefined ? p : { ...p, sheet: realTitle(p.sheet), ...(p.sheet2 !== undefined ? { sheet2: realTitle(p.sheet2) } : {}) })
        : '';
      const ref = prefix + renderArea(orderArea(t.area));
      if (ref !== text.slice(t.start, t.end)) edits.push({ start: t.start, end: t.end, text: ref });
    } else if (t.kind === 'bool') {
      const word = t.value ? 'TRUE' : 'FALSE';
      if (text.slice(t.start, t.end) !== word) edits.push({ start: t.start, end: t.end, text: word });
    } else if (t.kind === 'number') {
      const literal = numberLiteral(t.value);
      if (text.slice(t.start, t.end) !== literal) edits.push({ start: t.start, end: t.end, text: literal });
    }
  }
  return apply(text, edits);
}

/** Strip storage prefixes so the formula reads the way Excel's formula bar shows it. */
export function fromStorageFormula(text: string): string {
  const edits: Edit[] = [];
  for (const t of tokenize(text, true)) {
    if (t.kind !== 'func' && t.kind !== 'name') continue;
    const raw = t.kind === 'func' ? t.name : t.prefix === undefined ? t.name : '';
    const m = /^(?:_xlfn\.|_xlws\.|_xlpm\.)+/i.exec(raw);
    if (m !== null) edits.push({ start: t.start, end: t.start + m[0].length, text: '' });
  }
  return apply(text, edits);
}

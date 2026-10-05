// Formula text as Excel stores it in a file versus as a user types it.
//
// Functions added after Excel 2007 are written with a `_xlfn.` prefix
// (`_xlfn.XLOOKUP`, `_xlfn._xlws.SORT`), and LET / LAMBDA parameters with
// `_xlpm.`. Excel shows #NAME? for a bare `XLOOKUP` in a file, so formulas the
// editor writes must go through `toStorageFormula`; text read from a file goes
// through `fromStorageFormula` before it is shown in the formula bar.

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
export function toStorageFormula(text: string): string {
  const tokens = tokenize(text, true);
  const params = parameterNames(tokens);
  const edits: Edit[] = [];
  for (const t of tokens) {
    if (t.kind === 'func') {
      const name = normalizeFunctionName(t.name);
      if (name === t.name.toUpperCase() && FUTURE.has(name)) {
        edits.push({ start: t.start, end: t.start, text: XLWS.has(name) ? '_xlfn._xlws.' : '_xlfn.' });
      } else if (params.has(t.name.toUpperCase())) {
        edits.push({ start: t.start, end: t.start, text: PARAM_PREFIX });
      }
    } else if (t.kind === 'name' && t.prefix === undefined && params.has(t.name.toUpperCase())) {
      edits.push({ start: t.start, end: t.start, text: PARAM_PREFIX });
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

#!/usr/bin/env node
// Dev-only tool: ask the real Microsoft Excel (macOS) what each formula in
// CASES evaluates to, and record the answers in excel-oracle.fixture.json so
// the engine can be checked against Excel in CI without Excel.
//
//   pnpm build:js && node site/src/lib/editor/calc/excel-oracle.mjs
//
// How: build a workbook with a shared "Data" sheet and one formula per row of
// a "Cases" sheet (entered as single-cell array formulas, so array
// arithmetic evaluates the way a dynamic-array formula's top-left cell does),
// write it into Excel's sandbox container (no file-access prompt there), have
// Excel open it, `calculate full`, and save a copy. The copy's cached values
// carry full double precision and error types, which AppleScript's own value
// marshalling does not.
//
// Excel gotcha: any modal dialog blocks AppleEvents (error -1712). Recover
// with `pkill -x "Microsoft Excel"; open -a "Microsoft Excel"` and retry.

import { execFileSync } from 'node:child_process';
import { mkdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { homedir } from 'node:os';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const here = dirname(fileURLToPath(import.meta.url));
const root = join(here, '../../../../..');
const { makeArrayFormula } = await import(join(root, 'dist/cell.mjs'));
const { addWorksheet, createWorkbook } = await import(join(root, 'dist/workbook.mjs'));
const { setCell } = await import(join(root, 'dist/worksheet.mjs'));
const { loadWorkbook, workbookToBytes } = await import(join(root, 'dist/io.mjs'));
// Node strips the types from the engine's .ts sources directly.
const { fromBuffer } = await import(join(root, 'dist/node.mjs'));
const { toStorageFormula } = await import('./storage.ts');

/** Shared inputs. `[row, col, value]`; `{ error }` marks an error cell. */
export const DATA = [];
const put = (col, values) => values.forEach((v, i) => v !== undefined && DATA.push([i + 1, col, v]));
put(1, [10, 20, 30, 40, 50, -5, 0, 3.5, 100, 7]);
put(2, ['apple', 'Banana', 'cherry', 'apple', 'date', '', '3', true, undefined, 'Apple pie']);
put(3, [45292, 45323, 45351, 45382, 45412, 44927, 45657, 45000, 45016, 45306]);
put(4, ['x', 'y', 'x', 'z', 'y', 'x', 'z', 'x', 'y', 'x']);
put(5, [1, 2, 3, 4, 5, 6, 7, 8, 9, 10]);
put(6, [1, 3, 5, 7, 9, 11, 13, 15, 17, 19]);
put(7, ['one', 'three', 'five', 'seven', 'nine', 'eleven', 'thirteen', 'fifteen', 'seventeen', 'nineteen']);
put(8, [{ error: '#N/A' }, 5, { error: '#DIV/0!' }, 'text', 2.5]);
put(9, [-100, 30, 35, 40, 45, 50]);
put(10, [45292, 45323, 45383, 45444, 45505, 45566]);

export const CASES = [
  // arithmetic, precedence, coercion
  '1+2*3', '(1+2)*3', '-2^2', '2^3^2', '2^-1', '10/4', '7-10', '5%', '50%*200', '1/0', '0/0', '0^0', '0^-1', '(-8)^(1/3)', '2^0.5',
  '"3"+1', '"3"*"4"', 'TRUE+1', 'TRUE*TRUE', '"abc"+1', '" 12 "+0', '"1,234"+0', '"$12"+0', '"50%"+0', '"1e3"+0', '"(5)"+0', '"2024-01-15"+0', '"1/15/2024"+0', '"12:30"+0', '"12:30 PM"+0',
  '1&2', '"a"&TRUE', '0.1+0.2', '0.1+0.2=0.3', '1/3', '1/3&""', '2/3&""', '1E15&""', '123456789012345&""', '1234567890123456&""', '0.0001&""', '1E-9&""', '1E-10&""', '1E16&""', '1E19&""', '1E20&""', '1E21&""', '1.5E21&""', '-1E21&""', '2^70&""', '123456789012345678901&""', '1E-11&""', '1E-15&""', '1E-19&""', '1E-20&""', '1.234E-12&""', '-1.5E-20&""', '(0.1+0.2)&""', '(1/3)&""', '(2/3*1E-5)&""', '(2/3*1E20)&""', '1E100&""', '1.23456789012345E-7&""', '0.000000123&""', '1E20&""', '-1.5E-7&""', '100000&""', '12345678.9&""', '(0.1+0.2)&""',
  '1=1', '"a"="A"', '"a"<"b"', '"B">"a"', '1<"a"', '"a"<TRUE', 'TRUE>1', 'Data!I1=0', 'Data!I1=""', 'Data!I1=FALSE', '"10"<"9"', '10<9', '"abc"<>"ABC"', '1=1.0000000000000002', '0.3=0.1*3',
  'Data!A1+Data!A2', 'Data!A1:A3*2', 'SUM(Data!A1:A3*Data!E1:E3)', 'Data!A1:A3=10', '{1,2,3}+{10;20}', '{1,2}+{1,2,3}', 'SUM({1,2;3,4})', 'Data!A1&"x"', '-Data!A6', '--"5"',
  // math
  'SUM(Data!A1:A10)', 'SUM(Data!A:A)', 'SUM(Data!B1:B10)', 'SUM("3",TRUE)', 'SUM(1,"x")', 'SUM(Data!H1:H5)', 'SUM(Data!H2,Data!H5)', 'SUM((Data!A1,Data!A3))', 'SUM(Data!A1:B2 Data!B1:C3)', 'SUM(Data!A1:A3:Data!B5)',
  'PRODUCT(Data!E1:E5)', 'SUMSQ(1,2,3)', 'SUMPRODUCT(Data!A1:A5,Data!E1:E5)', 'SUMPRODUCT((Data!D1:D10="x")*Data!A1:A10)', 'SUMPRODUCT(Data!A1:A3,Data!E1:E4)', 'SUMPRODUCT(--(Data!A1:A10>20))',
  'ROUND(2.675,2)', 'ROUND(1.005,2)', 'ROUND(-2.5,0)', 'ROUND(1234.567,-2)', 'ROUND(0.5,0)', 'ROUND(1.45,1)', 'ROUNDUP(3.2,0)', 'ROUNDUP(-3.2,0)', 'ROUNDUP(0.1+0.2,1)', 'ROUNDDOWN(3.99,1)', 'ROUNDDOWN(-3.99,0)', 'MROUND(10,3)', 'MROUND(-10,-3)', 'MROUND(10,-3)', 'MROUND(1.3,0.2)',
  'CEILING(2.5,1)', 'CEILING(-2.5,2)', 'CEILING(-2.5,-2)', 'CEILING(2.5,-2)', 'CEILING(0.234,0.01)', 'CEILING.MATH(-2.5)', 'CEILING.MATH(-2.5,2,1)', 'CEILING.MATH(6.7,2)', 'FLOOR(2.5,1)', 'FLOOR(-2.5,2)', 'FLOOR(-2.5,-2)', 'FLOOR(2.5,0)', 'FLOOR.MATH(-2.5)', 'FLOOR.MATH(-2.5,2,1)', 'FLOOR.MATH(7.9,3)',
  'INT(-1.5)', 'INT(2.9999999999999996)', 'INT(0.7*10)', 'TRUNC(-1.57,1)', 'TRUNC(123.456,-1)', 'ABS(-3)', 'SIGN(-0.5)', 'MOD(10,3)', 'MOD(-3,2)', 'MOD(3,-2)', 'MOD(10.1,1)', 'MOD(5,0)', 'QUOTIENT(-7,2)', 'POWER(2,10)', 'POWER(-8,1/3)', 'SQRT(16)', 'SQRT(-1)',
  'EXP(1)', 'LN(10)', 'LN(0)', 'LOG(1000)', 'LOG(8,2)', 'LOG(10,1)', 'LOG10(0.001)', 'PI()', 'GCD(24,36,48)', 'GCD(7.9,3.1)', 'LCM(4,6,10)', 'FACT(5)', 'FACT(3.9)', 'FACT(-1)', 'FACTDOUBLE(7)', 'COMBIN(10,3)', 'COMBIN(3,5)', 'COMBINA(4,3)', 'PERMUT(10,3)', 'MULTINOMIAL(2,3,4)',
  'EVEN(1.5)', 'EVEN(-1.5)', 'ODD(2)', 'ODD(-2.1)', 'SIN(PI()/6)', 'COS(0)', 'TAN(PI()/4)', 'ASIN(1)', 'ACOS(2)', 'ATAN(1)', 'ATAN2(1,1)', 'ATAN2(0,0)', 'SINH(1)', 'COSH(1)', 'TANH(0.5)', 'ASINH(1)', 'ACOSH(1)', 'ATANH(0.5)', 'DEGREES(PI())', 'RADIANS(180)', 'COT(1)', 'SEC(1)', 'CSC(1)',
  'SUM(SEQUENCE(4))', 'SUM(SEQUENCE(2,3,10,5))', 'ROWS(SEQUENCE(5,2))', 'INDEX(SEQUENCE(3,3),2,3)', 'SUBTOTAL(9,Data!A1:A10)', 'SUBTOTAL(1,Data!A1:A10)', 'SUBTOTAL(3,Data!B1:B10)', 'SUBTOTAL(109,Data!A1:A10)', 'AGGREGATE(9,6,Data!H1:H5)', 'AGGREGATE(14,6,Data!H1:H5,1)', 'AGGREGATE(4,0,Data!A1:A10)',
  'SUMIF(Data!A1:A10,">20")', 'SUMIF(Data!D1:D10,"x",Data!A1:A10)', 'SUMIF(Data!B1:B10,"apple",Data!A1:A10)', 'SUMIF(Data!B1:B10,"a*",Data!A1:A10)', 'SUMIF(Data!B1:B10,"?????",Data!A1:A10)', 'SUMIF(Data!B1:B10,"<>apple",Data!A1:A10)', 'SUMIF(Data!A1:A10,"<>0")', 'SUMIF(Data!A:A,">=40")', 'SUMIF(Data!D1:D10,"x",Data!A1)',
  'SUMIFS(Data!A1:A10,Data!D1:D10,"x",Data!A1:A10,">5")', 'SUMIFS(Data!A1:A10,Data!D1:D10,"x",Data!B1:B10,"a*")', 'SUMIFS(Data!A:A,Data!D:D,"y")', 'MMULT({1,2;3,4},{5;6})', 'INDEX(MMULT({1,2;3,4},{5;6}),2)', 'MDETERM({1,2;3,4})', 'INDEX(MINVERSE({4,7;2,6}),1,1)', 'ROMAN(1999)', 'ARABIC("MCMXCIX")', 'BASE(255,16,4)', 'DECIMAL("FF",16)', 'SUMX2MY2({1,2},{3,4})', 'SQRTPI(2)',
  // statistics
  'AVERAGE(Data!A1:A10)', 'AVERAGE(Data!B1:B10)', 'AVERAGEA(Data!B1:B10)', 'AVERAGE(1,"2",TRUE)', 'COUNT(Data!A1:A10)', 'COUNT(Data!B1:B10)', 'COUNT(1,"2","x",TRUE)', 'COUNTA(Data!B1:B10)', 'COUNTBLANK(Data!B1:B10)', 'COUNTBLANK(Data!A1:J20)',
  'MAX(Data!A1:A10)', 'MIN(Data!A1:A10)', 'MAX(Data!B1:B10)', 'MAXA(Data!B1:B10)', 'MEDIAN(Data!A1:A10)', 'MEDIAN(1,2,3,4)', 'MODE.SNGL(1,2,2,3,3)', 'MODE({1,2,3})', 'STDEV(Data!A1:A10)', 'STDEV.P(Data!A1:A10)', 'VAR.S(Data!A1:A10)', 'VARP(Data!A1:A10)', 'STDEV(1)',
  'LARGE(Data!A1:A10,2)', 'LARGE(Data!A1:A10,1.5)', 'SMALL(Data!A1:A10,3)', 'SMALL(Data!A1:A10,11)', 'RANK(30,Data!A1:A10)', 'RANK.EQ(30,Data!A1:A10,1)', 'RANK.AVG(30,Data!A1:A10)', 'RANK(31,Data!A1:A10)',
  'PERCENTILE(Data!A1:A10,0.3)', 'PERCENTILE.EXC(Data!A1:A10,0.3)', 'PERCENTILE.EXC(Data!A1:A10,0.05)', 'QUARTILE(Data!A1:A10,1)', 'QUARTILE.EXC(Data!A1:A10,3)', 'PERCENTRANK.INC(Data!A1:A10,30)', 'PERCENTRANK.EXC(Data!A1:A10,30)',
  'CORREL(Data!A1:A10,Data!E1:E10)', 'PEARSON(Data!A1:A5,Data!E1:E5)', 'RSQ(Data!A1:A10,Data!E1:E10)', 'SLOPE(Data!A1:A10,Data!E1:E10)', 'INTERCEPT(Data!A1:A10,Data!E1:E10)', 'FORECAST(11,Data!A1:A10,Data!E1:E10)', 'FORECAST.LINEAR(11,Data!A1:A10,Data!E1:E10)', 'STEYX(Data!A1:A10,Data!E1:E10)', 'COVARIANCE.S(Data!A1:A10,Data!E1:E10)', 'COVAR(Data!A1:A10,Data!E1:E10)',
  'INDEX(TREND(Data!A1:A5,Data!E1:E5,{6;7}),2)', 'INDEX(LINEST(Data!A1:A5,Data!E1:E5),1)', 'GEOMEAN(1,2,4)', 'HARMEAN(1,2,4)', 'AVEDEV(1,2,3,4)', 'DEVSQ(1,2,3,4)', 'SKEW(1,2,3,10)', 'KURT(1,2,3,10,4)', 'TRIMMEAN(Data!A1:A10,0.2)',
  'COUNTIF(Data!A1:A10,">20")', 'COUNTIF(Data!B1:B10,"apple")', 'COUNTIF(Data!B1:B10,"APPLE")', 'COUNTIF(Data!B1:B10,"a*")', 'COUNTIF(Data!B1:B10,"*")', 'COUNTIF(Data!B1:B10,"")', 'COUNTIF(Data!B1:B10,"<>")', 'COUNTIF(Data!B1:B10,"=")', 'COUNTIF(Data!B1:B10,3)', 'COUNTIF(Data!B1:B10,"3")', 'COUNTIF(Data!B1:B10,TRUE)', 'COUNTIF(Data!B1:B10,"<c")', 'COUNTIF(Data!A1:A10,10)', 'COUNTIF(Data!H1:H5,"#N/A")', 'COUNTIF(Data!C1:C10,">"&DATE(2024,3,1))', 'COUNTIF(Data!C1:C10,">3/1/2024")', 'COUNTIF(Data!A:A,"")', 'COUNTIF(Data!B1:B10,"~*")',
  'COUNTIFS(Data!D1:D10,"x",Data!A1:A10,">5")', 'COUNTIFS(Data!D1:D10,"x",Data!E1:E10,"<5")', 'AVERAGEIF(Data!D1:D10,"x",Data!A1:A10)', 'AVERAGEIF(Data!D1:D10,"q",Data!A1:A10)', 'AVERAGEIFS(Data!A1:A10,Data!D1:D10,"y",Data!E1:E10,">2")', 'MAXIFS(Data!A1:A10,Data!D1:D10,"x")', 'MINIFS(Data!A1:A10,Data!D1:D10,"y")', 'MAXIFS(Data!A1:A10,Data!D1:D10,"q")',
  'NORM.DIST(1,0,1,TRUE)', 'NORM.DIST(1,0,1,FALSE)', 'NORM.S.DIST(-1.5,TRUE)', 'NORM.INV(0.9,10,2)', 'NORM.S.INV(0.975)', 'NORMSDIST(2)', 'EXPON.DIST(1,2,TRUE)', 'BINOM.DIST(3,10,0.3,FALSE)', 'BINOM.DIST(3,10,0.3,TRUE)', 'POISSON.DIST(2,3,TRUE)', 'GAMMA(5.5)', 'GAMMALN(10)', 'FISHER(0.5)', 'FISHERINV(0.5)', 'CONFIDENCE.NORM(0.05,2.5,50)', 'STANDARDIZE(42,40,1.5)', 'INDEX(FREQUENCY(Data!A1:A10,{10,40}),2)',
  // logical
  'IF(1,"y","n")', 'IF(0,"y")', 'IF("TRUE",1,2)', 'IF("x",1,2)', 'IF(Data!H1,1,2)', 'IF(TRUE,,2)', 'IFS(1>2,"a",2>1,"b")', 'IFS(FALSE,1)', 'IFERROR(1/0,"err")', 'IFERROR(5,"err")', 'IFNA(Data!H1,"na")', 'IFNA(1/0,"na")', 'AND(TRUE,1)', 'AND(Data!B1:B10)', 'OR(0,FALSE)', 'OR(Data!A1:A3>25)', 'XOR(TRUE,TRUE,TRUE)', 'NOT(0)', 'NOT("a")',
  'SWITCH(2,1,"a",2,"b","z")', 'SWITCH(9,1,"a","z")', 'SWITCH(9,1,"a")', 'CHOOSE(2,"a","b","c")', 'CHOOSE(4,"a","b")', 'SUM(CHOOSE(2,Data!A1:A2,Data!E1:E2))', 'LET(x,2,y,x*3,x+y)', 'LAMBDA(x,x*2)(21)', 'LET(f,LAMBDA(a,b,a-b),f(10,3))', 'SUM(MAP({1,2,3},LAMBDA(v,v*v)))', 'REDUCE(0,{1,2,3},LAMBDA(a,v,a+v))', 'INDEX(SCAN(0,{1,2,3},LAMBDA(a,v,a+v)),3)', 'SUM(BYROW({1,2;3,4},LAMBDA(r,MAX(r))))', 'SUM(MAKEARRAY(2,2,LAMBDA(r,c,r*c)))',
  // lookup
  'VLOOKUP(7,Data!F1:G10,2,FALSE)', 'VLOOKUP(8,Data!F1:G10,2)', 'VLOOKUP(8,Data!F1:G10,2,FALSE)', 'VLOOKUP(0,Data!F1:G10,2)', 'VLOOKUP(100,Data!F1:G10,2,TRUE)', 'VLOOKUP("APPLE",Data!B1:C10,2,FALSE)', 'VLOOKUP("ch*",Data!B1:C10,2,FALSE)', 'VLOOKUP(7,Data!F1:G10,3,FALSE)', 'VLOOKUP(7,Data!F1:G10,0,FALSE)', 'HLOOKUP(3,{1,2,3;"a","b","c"},2,FALSE)', 'HLOOKUP(2.5,{1,2,3;"a","b","c"},2)',
  'LOOKUP(8,Data!F1:F10,Data!G1:G10)', 'LOOKUP(2,1/(Data!D1:D10="y"),Data!A1:A10)', 'LOOKUP("c",{"a","b","d"},{1,2,3})', 'MATCH(9,Data!F1:F10,0)', 'MATCH(10,Data!F1:F10)', 'MATCH(10,Data!F1:F10,0)', 'MATCH("b*",Data!B1:B10,0)', 'MATCH(25,{50,40,30,20},-1)', 'MATCH("APPLE",Data!B1:B10,0)',
  'XLOOKUP(7,Data!F1:F10,Data!G1:G10)', 'XLOOKUP(8,Data!F1:F10,Data!G1:G10,"none")', 'XLOOKUP(8,Data!F1:F10,Data!G1:G10,,-1)', 'XLOOKUP(8,Data!F1:F10,Data!G1:G10,,1)', 'XLOOKUP("x",Data!D1:D10,Data!A1:A10,,0,-1)', 'XLOOKUP("a?ple",Data!B1:B10,Data!A1:A10,,2)', 'XLOOKUP(9,Data!F1:F10,Data!F1:F10,,0,2)', 'XMATCH(9,Data!F1:F10)', 'XMATCH(10,Data!F1:F10,1)', 'XMATCH("x",Data!D1:D10,0,-1)', 'SUM(XLOOKUP(5,Data!F1:F10,Data!E1:F10))',
  'INDEX(Data!A1:C10,3,2)', 'INDEX(Data!A1:A10,4)', 'INDEX(Data!A1:J1,3)', 'SUM(INDEX(Data!A1:C10,0,1))', 'SUM(INDEX(Data!A1:C10,2,0))', 'INDEX(Data!A1:C10,11,1)', 'INDEX({1,2;3,4},2,1)', 'INDEX((Data!A1:A3,Data!E1:E3),2,1,2)', 'SUM(Data!A1:INDEX(Data!A1:A10,3))',
  'SUM(OFFSET(Data!A1,1,0,3,1))', 'OFFSET(Data!A1,2,1)', 'ROWS(OFFSET(Data!A1,0,0,4,2))', 'INDIRECT("Data!A3")', 'SUM(INDIRECT("Data!A1:A3"))', 'INDIRECT("Data!R2C1",FALSE)', 'INDIRECT("nope!A1")', 'ROW(Data!C5)', 'COLUMN(Data!C5)', 'ROW()', 'COLUMN()', 'ROWS(Data!A1:C10)', 'COLUMNS(Data!A1:C10)', 'ROWS({1,2;3,4;5,6})', 'COLUMNS(Data!A:C)', 'ROWS(Data!1:3)', 'AREAS((Data!A1,Data!B2:C3))',
  'ADDRESS(2,3)', 'ADDRESS(2,3,2)', 'ADDRESS(2,3,4)', 'ADDRESS(2,3,1,FALSE)', 'ADDRESS(2,3,1,TRUE,"My Sheet")', 'ADDRESS(2,28)', 'SUM(TRANSPOSE({1,2,3}))', 'INDEX(TRANSPOSE({1,2,3}),2,1)', 'HYPERLINK("http://x.com","Go")',
  'TEXTJOIN(",",TRUE,FILTER(Data!A1:A10,Data!D1:D10="x"))', 'FILTER(Data!A1:A10,Data!D1:D10="q","none")', 'ROWS(FILTER(Data!A1:A10,Data!A1:A10>15))', 'TEXTJOIN(",",,SORT(Data!A1:A10))', 'TEXTJOIN(",",,SORT(Data!A1:A10,1,-1))', 'TEXTJOIN(",",,SORT(Data!B1:B5))', 'TEXTJOIN(",",,SORTBY(Data!G1:G5,Data!A1:A5,-1))', 'TEXTJOIN(",",,UNIQUE(Data!D1:D10))', 'TEXTJOIN(",",,UNIQUE(Data!D1:D10,,TRUE))', 'UNIQUE(Data!D1:D10,,TRUE)', 'ERROR.TYPE(UNIQUE(Data!D1:D10,,TRUE))', 'FILTER(Data!A1:A10,Data!A1:A10>1000)', 'TEXTJOIN(",",,FILTER(Data!A1:A10,Data!A1:A10>1000))', 'IFERROR(UNIQUE(Data!D1:D10,,TRUE),"none")', 'ROWS(UNIQUE(Data!D1:D10,,TRUE))', 'TEXTJOIN(",",,UNIQUE({"a";"A";"b"}))', 'ROWS(UNIQUE(Data!B1:B10))',
  'TEXTJOIN(",",,TAKE(Data!A1:A10,3))', 'TEXTJOIN(",",,TAKE(Data!A1:A10,-2))', 'TEXTJOIN(",",,DROP(Data!A1:A10,8))', 'TEXTJOIN(",",,CHOOSECOLS({1,2,3;4,5,6},3,1))', 'TEXTJOIN(",",,CHOOSEROWS({1,2;3,4;5,6},-1))', 'TEXTJOIN(",",,VSTACK({1,2},{3}))', 'TEXTJOIN(",",,HSTACK({1;2},{3;4}))', 'TEXTJOIN(",",,TOCOL({1,2;3,4}))', 'TEXTJOIN(",",,TOROW({1,2;3,4},,TRUE))', 'TEXTJOIN(",",,WRAPROWS({1,2,3,4,5},2,0))', 'ROWS(EXPAND({1,2},3,3))',
  // text
  'LEFT("hello",2)', 'LEFT("hello")', 'LEFT("hi",10)', 'RIGHT("hello",3)', 'MID("hello",2,3)', 'MID("hello",10,2)', 'MID("hello",0,2)', 'LEN("héllo")', 'LEN(123.45)', 'LEN(1/3)', 'LEN(TRUE)', 'FIND("l","hello")', 'FIND("L","hello")', 'FIND("l","hello",4)', 'FIND("","abc")', 'SEARCH("L","hello")', 'SEARCH("l?o","hello")', 'SEARCH("*l","hello")', 'SEARCH("z","hello")',
  'SUBSTITUTE("a-b-c","-","+")', 'SUBSTITUTE("a-b-c","-","+",2)', 'SUBSTITUTE("aaa","a","bb")', 'REPLACE("abcdef",2,3,"X")', 'UPPER("abc")', 'LOWER("ABC")', 'PROPER("hello wORLD o\'neil")', 'TRIM("  a   b  ")', 'CLEAN("a"&CHAR(7)&"b")', 'CONCAT(Data!G1:G3)', 'CONCAT("a",1,TRUE)', 'CONCATENATE("a","b",1)', 'TEXTJOIN("-",TRUE,"a","","b")', 'TEXTJOIN("-",FALSE,"a","","b")', 'TEXTJOIN({",",";"},TRUE,"a","b","c","d")',
  'TEXT(1234.567,"#,##0.00")', 'TEXT(0.256,"0.0%")', 'TEXT(45292,"yyyy-mm-dd")', 'TEXT(45292,"dddd, mmmm d, yyyy")', 'TEXT(45292.75,"h:mm AM/PM")', 'TEXT(-5,"0.00;(0.00)")', 'TEXT(0,"0;-0;""zero""")', 'TEXT(1234567,"0.00E+00")', 'TEXT(3.5,"0")', 'TEXT(2.5,"0")', 'TEXT("12","0.00")', 'TEXT("abc","0")', 'TEXT(1.5,"[h]:mm")', 'TEXT(0.333333,"# ?/?")', 'TEXT(1234.5,"$#,##0")', 'TEXT(42,"000000")',
  'VALUE("1,234.5")', 'VALUE("abc")', 'VALUE("10%")', 'VALUE("")', 'VALUE(TRUE)', 'NUMBERVALUE("1.234,5",",",".")', 'NUMBERVALUE("50%")', 'REPT("ab",3)', 'REPT("x",0)', 'EXACT("a","A")', 'EXACT("a","a")', 'CHAR(65)', 'CHAR(128)', 'CHAR(0)', 'CODE("A")', 'CODE("€")', 'UNICHAR(9731)', 'UNICODE("☃")', 'T("a")', 'T(1)', 'N("a")', 'N(TRUE)', 'N(5)',
  'DOLLAR(1234.567)', 'DOLLAR(-1234.567,1)', 'DOLLAR(1234.567,-2)', 'FIXED(1234.567,1)', 'FIXED(1234.567,1,TRUE)', 'FIXED(-0.5,0)', 'TEXTBEFORE("a-b-c","-")', 'TEXTBEFORE("a-b-c","-",-1)', 'TEXTAFTER("a-b-c","-",2)', 'TEXTAFTER("a-b-c","x")', 'TEXTAFTER("a-b-c","x",,,,"nf")', 'TEXTBEFORE("aXbxc","x",1,1)', 'TEXTJOIN("|",,TEXTSPLIT("a,b;c,d",",",";"))', 'COLUMNS(TEXTSPLIT("a,,b",","))', 'COLUMNS(TEXTSPLIT("a,,b",",",,TRUE))', 'VALUETOTEXT("a",1)',
  // date & time
  'DATE(2024,1,15)', 'DATE(2024,14,1)', 'DATE(2024,1,0)', 'DATE(1900,2,29)', 'DATE(1900,3,0)', 'DATE(99,1,1)', 'DATE(-1,1,1)', 'DATE(10000,1,1)', 'DATE(2024,-1,1)', 'TIME(12,30,0)', 'TIME(25,0,0)', 'TIME(0,-1,0)', 'TIME(1,90,30)',
  'YEAR(45292)', 'MONTH(45292)', 'DAY(45292)', 'YEAR(0)', 'MONTH(0)', 'DAY(0)', 'DAY(60)', 'MONTH(60)', 'DAY(61)', 'YEAR(-1)', 'YEAR("2024-03-05")', 'HOUR(0.75)', 'MINUTE(0.7555)', 'SECOND(0.7555)', 'HOUR(1.5)', 'SECOND(0.999999999)', 'HOUR("3:45 PM")',
  'WEEKDAY(45292)', 'WEEKDAY(45292,2)', 'WEEKDAY(45292,3)', 'WEEKDAY(45292,11)', 'WEEKDAY(45292,16)', 'WEEKDAY(1)', 'WEEKDAY(0)', 'WEEKDAY(45292,4)', 'WEEKNUM(45292)', 'WEEKNUM(45300,2)', 'WEEKNUM(45657)', 'WEEKNUM(45657,21)', 'ISOWEEKNUM(45292)', 'ISOWEEKNUM(DATE(2021,1,3))', 'ISOWEEKNUM(DATE(2020,12,31))',
  'EDATE(DATE(2024,1,31),1)', 'EDATE(DATE(2024,3,31),-1)', 'EDATE(45292,12)', 'EOMONTH(DATE(2024,1,15),1)', 'EOMONTH(DATE(2024,1,15),-1)', 'EOMONTH(DATE(2023,11,30),3)', 'DATEDIF(DATE(2020,1,15),DATE(2024,3,10),"Y")', 'DATEDIF(DATE(2020,1,15),DATE(2024,3,10),"M")', 'DATEDIF(DATE(2020,1,15),DATE(2024,3,10),"D")', 'DATEDIF(DATE(2020,1,15),DATE(2024,3,10),"MD")', 'DATEDIF(DATE(2020,1,15),DATE(2024,3,10),"YM")', 'DATEDIF(DATE(2020,1,15),DATE(2024,3,10),"YD")', 'DATEDIF(DATE(2024,1,31),DATE(2024,3,1),"MD")', 'DATEDIF(DATE(2024,3,10),DATE(2020,1,15),"Y")',
  'DAYS(DATE(2024,3,1),DATE(2024,1,1))', 'DAYS360(DATE(2024,1,31),DATE(2024,3,31))', 'DAYS360(DATE(2024,1,30),DATE(2024,2,29))', 'DAYS360(DATE(2024,2,29),DATE(2024,3,31))', 'DAYS360(DATE(2024,1,31),DATE(2024,3,31),TRUE)', 'DAYS360(DATE(2023,2,28),DATE(2023,3,31))',
  'NETWORKDAYS(DATE(2024,1,1),DATE(2024,1,31))', 'NETWORKDAYS(DATE(2024,1,31),DATE(2024,1,1))', 'NETWORKDAYS(DATE(2024,1,1),DATE(2024,1,31),{45292,45306})', 'NETWORKDAYS.INTL(DATE(2024,1,1),DATE(2024,1,31),11)', 'NETWORKDAYS.INTL(DATE(2024,1,1),DATE(2024,1,31),"0000110")', 'WORKDAY(DATE(2024,1,5),1)', 'WORKDAY(DATE(2024,1,5),-5)', 'WORKDAY(DATE(2024,1,1),10,{45293})', 'WORKDAY.INTL(DATE(2024,1,5),3,7)', 'WORKDAY.INTL(DATE(2024,1,5),3,"1111100")',
  'YEARFRAC(DATE(2024,1,1),DATE(2024,7,1))', 'YEARFRAC(DATE(2024,1,1),DATE(2024,7,1),1)', 'YEARFRAC(DATE(2023,1,1),DATE(2024,7,1),1)', 'YEARFRAC(DATE(2023,6,1),DATE(2024,3,1),1)', 'YEARFRAC(DATE(2024,1,31),DATE(2024,2,29),0)', 'YEARFRAC(DATE(2024,1,1),DATE(2024,7,1),2)', 'YEARFRAC(DATE(2024,1,1),DATE(2024,7,1),3)', 'YEARFRAC(DATE(2024,1,31),DATE(2024,3,31),4)', 'YEARFRAC(DATE(2023,2,28),DATE(2024,2,29),0)',
  'DATEVALUE("2024-01-15")', 'DATEVALUE("1/15/2024")', 'DATEVALUE("15-Jan-2024")', 'DATEVALUE("Jan 15, 2024")', 'DATEVALUE("January 15 2024")', 'DATEVALUE("2024-01-15 10:30")', 'DATEVALUE("abc")', 'DATEVALUE("2/30/2024")', 'DATEVALUE("1/1/29")', 'DATEVALUE("1/1/30")', 'TIMEVALUE("12:30")', 'TIMEVALUE("6:45:30 PM")', 'TIMEVALUE("2024-01-15 06:00")', 'TIMEVALUE("25:00")',
  // information
  'ISBLANK(Data!B9)', 'ISBLANK(Data!B6)', 'ISNUMBER(Data!A1)', 'ISNUMBER("1")', 'ISTEXT(Data!B1)', 'ISNONTEXT(1)', 'ISLOGICAL(Data!B8)', 'ISERROR(Data!H1)', 'ISERR(Data!H1)', 'ISERR(Data!H3)', 'ISNA(Data!H1)', 'ISEVEN(-2.5)', 'ISODD(3.9)', 'ISREF(Data!A1)', 'ISREF(1)', 'ISFORMULA(Data!A1)', 'NA()', 'ERROR.TYPE(Data!H3)', 'ERROR.TYPE(1)', 'ERROR.TYPE(#NULL!)',
  'TYPE(1)', 'TYPE("a")', 'TYPE(TRUE)', 'TYPE(#N/A)', 'TYPE({1,2})', 'TYPE(Data!I9)', 'CELL("row",Data!B7)', 'CELL("col",Data!B7)', 'CELL("address",Data!B7)', 'CELL("contents",Data!A2)', 'CELL("type",Data!B1)', 'CELL("type",Data!A1)', 'CELL("type",Data!I9)',
  // financial
  'PMT(0.05/12,360,200000)', 'PMT(0,12,1200)', 'PMT(0.06/12,60,-20000,0,1)', 'FV(0.06/12,10,-200,-500,1)', 'FV(0,10,-100)', 'PV(0.08/12,240,500)', 'PV(0.05,10,0,10000)', 'NPER(0.12/12,-100,-1000,10000,1)', 'NPER(0,-100,1000)', 'IPMT(0.1/12,1,36,8000)', 'IPMT(0.1,3,3,8000)', 'PPMT(0.1/12,1,24,2000)', 'IPMT(0.1/12,1,36,8000,0,1)', 'CUMIPMT(0.09/12,360,125000,13,24,0)', 'CUMPRINC(0.09/12,360,125000,13,24,0)',
  'RATE(48,-200,8000)', 'RATE(10,-100,1000,0,0,0.1)', 'NPV(0.1,-10000,3000,4200,6800)', 'NPV(0.08,Data!I2:I6)+Data!I1', 'IRR(Data!I1:I6)', 'IRR({-100,10,10})', 'IRR({100,10})', 'MIRR(Data!I1:I6,0.1,0.12)', 'XNPV(0.09,Data!I1:I6,Data!J1:J6)', 'XIRR(Data!I1:I6,Data!J1:J6)', 'SLN(30000,7500,10)', 'SYD(30000,7500,10,1)', 'DB(1000000,100000,6,1,7)', 'DB(1000000,100000,6,7,7)', 'DB(1000000,100000,6,3)', 'DDB(2400,300,10,1)', 'DDB(2400,300,10,10)', 'DDB(2400,300,120,1,1.5)', 'EFFECT(0.0525,4)', 'NOMINAL(0.053543,4)',
  // engineering
  'DEC2BIN(9)', 'DEC2BIN(-1)', 'DEC2BIN(9,8)', 'DEC2BIN(512)', 'DEC2HEX(255)', 'DEC2HEX(-1)', 'DEC2OCT(58,4)', 'BIN2DEC("1111111111")', 'BIN2DEC("1010")', 'HEX2DEC("FFFFFFFFFF")', 'HEX2DEC("a5")', 'OCT2DEC("777")', 'BIN2HEX("1110")', 'HEX2BIN("F",8)', 'HEX2OCT("1F")', 'OCT2BIN("7")', 'BIN2DEC("2")',
  'CONVERT(1,"mi","km")', 'CONVERT(68,"F","C")', 'CONVERT(1,"lbm","kg")', 'CONVERT(2.5,"ft","m")', 'CONVERT(1,"hr","mn")', 'CONVERT(1,"gal","l")', 'CONVERT(100,"C","K")', 'CONVERT(1,"m","lbm")', 'CONVERT(1,"m2","ft2")', 'DELTA(5,5)', 'GESTEP(5,4)', 'BITAND(13,25)', 'BITOR(13,25)', 'BITXOR(13,25)', 'BITLSHIFT(4,2)', 'BITRSHIFT(13,2)',
];
/** Database table (A1:D7) and criteria ranges (A9:B11) for the D* functions, on sheet Db. */
export const DB_DATA = [
  [1, 1, 'Tree'], [1, 2, 'Height'], [1, 3, 'Age'], [1, 4, 'Yield'],
  [2, 1, 'Apple'], [2, 2, 18], [2, 3, 20], [2, 4, 14],
  [3, 1, 'Pear'], [3, 2, 12], [3, 3, 12], [3, 4, 10],
  [4, 1, 'Cherry'], [4, 2, 13], [4, 3, 14], [4, 4, 9],
  [5, 1, 'Apple'], [5, 2, 14], [5, 3, 15], [5, 4, 10],
  [6, 1, 'Pear'], [6, 2, 9], [6, 3, 8], [6, 4, 8],
  [7, 1, 'Apple'], [7, 2, 8], [7, 3, 9], [7, 4, 6],
  [9, 1, 'Tree'], [9, 2, 'Height'], [10, 1, 'Apple'], [10, 2, '>10'], [11, 1, 'Pear'],
];
// Database functions need headed ranges, laid out on their own sheet above.
CASES.push(
  'DSUM(Db!A1:D7,"Yield",Db!A9:B11)', 'DSUM(Db!A1:D7,4,Db!A9:A10)', 'DCOUNT(Db!A1:D7,"Age",Db!A9:B11)', 'DCOUNTA(Db!A1:D7,,Db!A9:A10)', 'DAVERAGE(Db!A1:D7,"Yield",Db!A9:B10)', 'DMAX(Db!A1:D7,"Height",Db!A9:A11)', 'DMIN(Db!A1:D7,"Height",Db!A9:A11)', 'DGET(Db!A1:D7,"Age",Db!A9:B10)', 'DPRODUCT(Db!A1:D7,"Yield",Db!A9:A10)', 'DSTDEV(Db!A1:D7,"Yield",Db!A9:A10)', 'DVARP(Db!A1:D7,"Yield",Db!A9:A10)',
);

const toModel = (v) => (v !== null && typeof v === 'object' && 'error' in v ? { kind: 'error', code: v.error } : v);

function buildWorkbook() {
  const wb = createWorkbook();
  const data = addWorksheet(wb, 'Data');
  for (const [r, c, v] of DATA) setCell(data, r, c, toModel(v));
  const db = addWorksheet(wb, 'Db');
  for (const [r, c, v] of DB_DATA) setCell(db, r, c, v);
  const cases = addWorksheet(wb, 'Cases');
  CASES.forEach((f, i) => setCell(cases, i + 1, 1, makeArrayFormula(`A${i + 1}`, toStorageFormula(f))));
  return wb;
}

const osa = (script) => execFileSync('osascript', ['-e', script], { encoding: 'utf8', timeout: 120_000 });

async function main() {
  const dir = join(homedir(), 'Library/Containers/com.microsoft.Excel/Data/tmp/calc-oracle');
  mkdirSync(dir, { recursive: true });
  const input = join(dir, 'oracle-in.xlsx');
  const output = join(dir, 'oracle-out.xlsx');
  rmSync(output, { force: true });
  writeFileSync(input, await workbookToBytes(buildWorkbook()));
  // AppleScript's own `open` returns before the workbook exists, so the
  // following `active workbook` raced it (-1728); LaunchServices plus polling
  // for the name is reliable.
  execFileSync('open', ['-a', 'Microsoft Excel', input]);
  const name = 'oracle-in.xlsx';
  for (let i = 0; !osa('tell application "Microsoft Excel" to get name of every workbook').includes(name); i++) {
    if (i === 60) throw new Error('Excel did not open the oracle workbook (repair dialog?)');
    await new Promise((r) => {
      setTimeout(r, 1000);
    });
  }
  osa(`with timeout of 120 seconds
tell application "Microsoft Excel"
  set wb to workbook "${name}"
  calculate full
  save workbook as wb filename "${output}" file format workbook normal file format
  close workbook "oracle-out.xlsx" saving no
end tell
end timeout`);
  const out = await loadWorkbook(fromBuffer(readFileSync(output)));
  const sheet = out.sheets.find((s) => s.sheet.title === 'Cases').sheet;
  const cases = CASES.map((formula, i) => {
    const v = sheet.rows.get(i + 1)?.get(1)?.value;
    let expected;
    if (v === undefined || v === null || v.cachedValue === undefined) expected = { type: 'blank' };
    else if (v.cachedValueType === 'error') expected = { type: 'error', value: String(v.cachedValue) };
    else expected = { type: typeof v.cachedValue, value: v.cachedValue };
    return { formula, expected };
  });
  const fixture = { excelVersion: osa('tell application "Microsoft Excel" to get version').trim(), data: DATA, db: DB_DATA, cases };
  writeFileSync(join(here, 'excel-oracle.fixture.json'), `${JSON.stringify(fixture, null, 1)}\n`);
  console.log(`recorded ${cases.length} cases`);
}

if (process.argv[1] === fileURLToPath(import.meta.url)) await main();

// load → edit → save → reopen through the published subpaths only, against the
// built dist/ (run `pnpm build` first). Each subpath is resolved through this
// package's own `exports` map, the way a consumer's `@office-kit/xlsx/<name>`
// import is, without packing or installing anything.
//
// Checks, with the zip mtime and the document dates pinned:
//   - f(input) twice gives identical bytes;
//   - f(f(input)): the first output fed back in gives the same bytes again,
//     and the threaded comment and its person are kept, not duplicated or
//     re-keyed;
//   - reload + resave of the output changes nothing;
//   - `_xlfn.XLOOKUP`, the edited A3, the comment and the person survive.
// The person and comment ids are pinned too: makePerson / makeThreadedComment
// draw random GUIDs, so without that two runs of this script would start from
// different input bytes. Run it twice to compare the printed hashes across
// processes.

import assert from 'node:assert/strict';
import { createHash } from 'node:crypto';
import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath, pathToFileURL } from 'node:url';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const { exports: exportsMap } = JSON.parse(readFileSync(join(root, 'package.json'), 'utf8'));
const subpath = (name) => import(pathToFileURL(join(root, exportsMap[`./${name}`].import)).href);

const { makeFormula } = await subpath('cell');
const { loadWorkbook, workbookToBytes } = await subpath('io');
const { fromBuffer } = await subpath('node');
const { addWorksheet, createWorkbook, makePerson } = await subpath('workbook');
const { getCell, makeThreadedComment, setCell } = await subpath('worksheet');

const STAMP = new Date(Date.UTC(2026, 0, 2, 3, 4, 0));
const PERSON_ID = '{11111111-2222-3333-4444-555555555555}';
const COMMENT_ID = '{66666666-7777-8888-9999-AAAAAAAAAAAA}';
const XLOOKUP = '_xlfn.XLOOKUP("a",A1:A1,B1:B1)';
const A3 = 'SUM(B1:B2)';

const sha = (bytes) => createHash('sha256').update(bytes).digest('hex').slice(0, 16);

async function source() {
  const wb = createWorkbook();
  wb.properties = { creator: 'qa', created: STAMP.toISOString(), modified: STAMP.toISOString() };
  const ws = addWorksheet(wb, 'Data');
  setCell(ws, 1, 1, 'a');
  setCell(ws, 1, 2, 1);
  setCell(ws, 2, 1, makeFormula(XLOOKUP));
  wb.persons.push({ ...makePerson({ displayName: 'Reviewer' }), id: PERSON_ID });
  ws.threadedComments.push({ ...makeThreadedComment({ ref: 'B1', personId: PERSON_ID, text: 'check', created: STAMP }), id: COMMENT_ID });
  return workbookToBytes(wb, { mtime: STAMP });
}

/** The "user script": load, set A3, save. */
async function edit(input) {
  const wb = await loadWorkbook(fromBuffer(input));
  const ref = wb.sheets[0];
  assert.equal(ref?.kind, 'worksheet');
  setCell(ref.sheet, 3, 1, makeFormula(A3));
  return workbookToBytes(wb, { mtime: STAMP });
}

async function inspect(bytes) {
  const wb = await loadWorkbook(fromBuffer(bytes));
  const ws = wb.sheets[0].sheet;
  return {
    a2: getCell(ws, 2, 1)?.value?.formula,
    a3: getCell(ws, 3, 1)?.value?.formula,
    comments: ws.threadedComments.map((c) => [c.id, c.personId, c.ref, c.text]),
    persons: wb.persons.map((p) => [p.id, p.displayName]),
    resaved: await workbookToBytes(wb, { mtime: STAMP }),
  };
}

const input = await source();
const once = await edit(input);
const onceAgain = await edit(input);
const twice = await edit(once);

assert.deepEqual(onceAgain, once, 'f(input) is not deterministic');
assert.deepEqual(twice, once, 'f(f(input)) differs from f(input)');
for (const [label, bytes] of [['f(input)', once], ['f(f(input))', twice]]) {
  const got = await inspect(bytes);
  assert.equal(got.a2, XLOOKUP, `${label}: A2`);
  assert.equal(got.a3, A3, `${label}: A3`);
  assert.deepEqual(got.comments, [[COMMENT_ID, PERSON_ID, 'B1', 'check']], `${label}: threaded comments`);
  assert.deepEqual(got.persons, [[PERSON_ID, 'Reviewer']], `${label}: persons`);
  assert.deepEqual(got.resaved, bytes, `${label}: reload + resave changed the bytes`);
}

console.log(JSON.stringify({ node: process.version, input: sha(input), once: sha(once), twice: sha(twice) }));

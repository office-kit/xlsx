import { makeCfRule, makeConditionalFormatting, makeDataValidation } from '@office-kit/xlsx/worksheet';
import { expect, it } from 'vitest';
import * as A from './actions.ts';
import { EditorController } from './controller.svelte.ts';

function setup(): EditorController {
  const ctl = new EditorController();
  const ws = ctl.doc.ws;
  ws.conditionalFormatting.push(makeConditionalFormatting({ sqref: 'B1:B10', rules: [makeCfRule({ type: 'expression', priority: 1, formulas: ['$A$5>3'] })] }));
  ws.dataValidations.push(makeDataValidation({ sqref: 'D1:D10', type: 'list', formula1: '$A$1:$A$5' }));
  return ctl;
}

const ruleFormulas = (ctl: EditorController) => [ctl.doc.ws.conditionalFormatting[0]?.rules[0]?.formulas[0], ctl.doc.ws.dataValidations[0]?.formula1];

it('inserting rows moves the cells conditional formatting and validation formulas point at', () => {
  const ctl = setup();
  ctl.selectRange({ r1: 2, c1: 1, r2: 3, c2: 16384 });
  A.insertLines(ctl, 'row');
  expect(ruleFormulas(ctl)).toEqual(['$A$7>3', '$A$1:$A$7']);
  ctl.doc.undo();
  expect(ruleFormulas(ctl)).toEqual(['$A$5>3', '$A$1:$A$5']);
});

it('deleting the column they point at turns the reference into #REF!', () => {
  const ctl = setup();
  ctl.selectRange({ r1: 1, c1: 1, r2: 1048576, c2: 1 });
  A.deleteLines(ctl, 'col');
  expect(ruleFormulas(ctl)).toEqual(['#REF!>3', '#REF!']);
});

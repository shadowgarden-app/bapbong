import { describe, expect, it } from 'vitest';
import { schema } from './model.js';
import { effectiveFootnotes, footnoteNumbers } from './footnotes.js';

const ref = (num: number, shown: string) =>
  schema.text(shown, [
    schema.marks['vertAlign'].create({ value: 'super' }),
    schema.marks['footnote'].create({ num, id: null }),
  ]);
const body = (text: string) =>
  schema.node('doc', null, [
    schema.node('paragraph', null, [schema.text(text)]),
  ]);

describe('footnotes written in bapbong', () => {
  it('each body comes with its reference number, next to the imported ones', () => {
    const imported = { 1: body('1 Imported note') };
    const doc = schema.node(
      'doc',
      { footnoteBodies: { '2': body('New note').toJSON() } },
      [
        schema.node('paragraph', null, [
          schema.text('A'),
          ref(1, '1'),
          schema.text(' B'),
          ref(2, '2'),
        ]),
      ],
    );
    expect([...footnoteNumbers(doc)]).toEqual([
      [1, '1'],
      [2, '2'],
    ]);
    const all = effectiveFootnotes(doc, imported)!;
    expect(all[1]).toBe(imported[1]);
    expect(all[2].textContent).toBe('2 New note');
    expect(
      all[2].firstChild!.firstChild!.marks.map((m) => m.type.name),
    ).toEqual(['vertAlign']);
  });

  it('without any, the imported set is what it was', () => {
    const imported = { 1: body('x') };
    const doc = schema.node('doc', null, [schema.node('paragraph')]);
    expect(effectiveFootnotes(doc, imported)).toBe(imported);
  });
});

import { describe, expect, it } from 'vitest';
import type { Node as PMNode } from 'prosemirror-model';
import {
  exportDocx,
  importDocx,
  schema,
} from '@shadow-garden/bapbong-headless';
import { contentToNodes } from './blocks.js';
import { HeadlessSession } from './headless-session.js';

/** name → attrs of every mark on the first text node containing `text`. */
function marksOn(doc: PMNode, text: string): Record<string, unknown> {
  let out: Record<string, unknown> | null = null;
  doc.descendants((n) => {
    if (out) return false;
    if (n.isText && n.text?.includes(text)) {
      out = Object.fromEntries(n.marks.map((m) => [m.type.name, m.attrs]));
    }
    return true;
  });
  if (!out) throw new Error(`no text ${text}`);
  return out;
}

const docOf = (s: HeadlessSession): PMNode =>
  (s as unknown as { state: { doc: PMNode } }).state.doc;

describe('character formatting in blocks', () => {
  it('inlines carry strike, super/subscript, colour, highlight, font and size; the run wins over its paragraph', () => {
    const [p] = contentToNodes(
      [
        {
          paragraph: [
            'H',
            { text: '2', subscript: true },
            'O ',
            { text: 'new', color: '#C00000', strike: true },
            { text: ' note', superscript: true, size: 8, font: 'Arial' },
          ],
          color: '#1F4E79',
          highlight: '#FFFF00',
        },
      ],
      schema,
      { contentWidth: 602 },
    );
    const doc = schema.node('doc', null, [p]);
    expect(marksOn(doc, 'H')).toEqual({
      textColor: { color: '#1F4E79' },
      highlight: { color: '#FFFF00' },
    });
    expect(marksOn(doc, '2')).toMatchObject({ vertAlign: { value: 'sub' } });
    expect(marksOn(doc, 'new')).toMatchObject({
      strike: {},
      textColor: { color: '#C00000' },
    });
    expect(marksOn(doc, 'note')).toMatchObject({
      vertAlign: { value: 'super' },
      fontSize: { size: 8 },
      fontFamily: { family: 'Arial' },
    });
  });
});

describe('apply_formatting: font, colour, highlight, raise/lower, clear', () => {
  async function open(): Promise<HeadlessSession> {
    const link = schema.marks['link'].create({ href: 'https://example.com' });
    return HeadlessSession.open(
      await exportDocx(
        schema.node('doc', null, [
          schema.node('paragraph', null, [
            schema.text('Read '),
            schema.text('the guide', [link]),
            schema.text(' today.'),
          ]),
        ]),
      ),
    );
  }

  it('sets them on the text, and null (or "auto" / "none" at the tool) takes each away', async () => {
    const s = await open();
    await s.applyFormatting('today', {
      fontFamily: 'Georgia',
      color: '#C00000',
      highlight: '#FFFF00',
      verticalAlign: 'superscript',
    });
    expect(marksOn(docOf(s), 'today')).toMatchObject({
      fontFamily: { family: 'Georgia' },
      textColor: { color: '#C00000' },
      highlight: { color: '#FFFF00' },
      vertAlign: { value: 'super' },
    });
    await s.applyFormatting('today', {
      color: null,
      highlight: null,
      verticalAlign: null,
    });
    const left = marksOn(docOf(s), 'today');
    expect(left).toHaveProperty('fontFamily');
    expect(left).not.toHaveProperty('textColor');
    expect(left).not.toHaveProperty('highlight');
    expect(left).not.toHaveProperty('vertAlign');
  });

  it('clear strips formatting from a whole block but keeps its link', async () => {
    const s = await open();
    await s.applyFormatting(
      { blockIndex: 0 },
      { bold: true, color: '#00B050' },
    );
    await s.applyFormatting({ blockIndex: 0 }, { clear: true });
    expect(marksOn(docOf(s), 'guide')).toEqual({
      link: expect.objectContaining({ href: 'https://example.com' }),
    });
    expect(marksOn(docOf(s), 'Read')).toEqual({});
  });

  it('round-trips through .docx', async () => {
    const s = await open();
    await s.applyFormatting('today', {
      fontFamily: 'Georgia',
      color: '#C00000',
      highlight: '#FFFF00',
      verticalAlign: 'subscript',
    });
    let bytes: Uint8Array | undefined;
    (
      s as unknown as { opts: { onSave?: (b: Uint8Array) => void } }
    ).opts.onSave = (b) => {
      bytes = b;
    };
    await s.save();
    const { doc } = await importDocx(bytes!.slice().buffer as ArrayBuffer);
    expect(marksOn(doc, 'today')).toMatchObject({
      fontFamily: { family: 'Georgia' },
      textColor: { color: '#C00000' },
      highlight: { color: '#FFFF00' },
      vertAlign: { value: 'sub' },
    });
  });
});

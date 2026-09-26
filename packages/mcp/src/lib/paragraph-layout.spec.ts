import { describe, expect, it } from 'vitest';
import type { Node as PMNode } from 'prosemirror-model';
import {
  exportDocx,
  importDocx,
  schema,
} from '@shadow-garden/bapbong-headless';
import { contentToNodes } from './blocks.js';
import { ContentError, singleDocumentProvider } from './contract.js';
import { applyFormatting } from './document-commands.js';
import { HeadlessSession } from './headless-session.js';

const docOf = (s: HeadlessSession): PMNode =>
  (s as unknown as { state: { doc: PMNode } }).state.doc;

async function open(lines: string[]): Promise<HeadlessSession> {
  return HeadlessSession.open(
    await exportDocx(
      schema.node(
        'doc',
        null,
        lines.map((l) => schema.node('paragraph', null, [schema.text(l)])),
      ),
    ),
  );
}

describe('paragraph spacing and indents', () => {
  it('blocks: points, multiples, exact heights and lengths become the model attrs; a list level indent is the base', () => {
    const [a, b, c] = contentToNodes(
      [
        {
          paragraph: 'Spaced',
          spaceBefore: 6,
          spaceAfter: 12,
          lineSpacing: 1.5,
          indent: { left: 1, firstLine: '0.5in' },
        },
        { paragraph: 'Exact', lineSpacing: { exact: 18 } },
        {
          paragraph: 'Item',
          list: 'bullet',
          level: 2,
          indent: { hanging: '12pt' },
        },
      ],
      schema,
      { contentWidth: 602 },
    );
    expect(a.attrs['spacing']).toEqual({
      before: 8,
      after: 16,
      line: 1.5,
      lineRule: 'auto',
    });
    expect(a.attrs['indent']).toEqual({ left: 38, firstLine: 48 });
    expect(b.attrs['spacing']).toEqual({ line: 24, lineRule: 'exact' });
    expect(c.attrs['indent']).toEqual({ left: 24, hanging: 16 });
  });

  it('apply_formatting changes one side at a time; first line and hanging exclude each other; 0 removes', async () => {
    const s = await open(['One', 'Two']);
    await s.applyFormatting(
      { blockIndex: 0 },
      { spaceAfter: 6, indent: { left: '2cm', firstLine: '1cm' } },
    );
    await s.applyFormatting(
      { blockIndex: 0 },
      { lineSpacing: { atLeast: 14 }, indent: { hanging: '1cm' } },
    );
    const p = docOf(s).child(0);
    expect(p.attrs['spacing']).toMatchObject({
      after: 8,
      line: 19,
      lineRule: 'atLeast',
    });
    expect(p.attrs['indent']).toEqual({ left: 76, hanging: 38 });
    await s.applyFormatting(
      { blockIndex: 0 },
      { indent: { left: 0, hanging: 0 } },
    );
    expect(docOf(s).child(0).attrs['indent']).toBeNull();
    await expect(
      s.applyFormatting({ blockIndex: 1 }, { lineSpacing: 40 }),
    ).rejects.toThrow(ContentError);
  });

  it('the tool reads "18pt" as exact and "at least 12pt" as a minimum', async () => {
    const s = await open(['One']);
    const provider = singleDocumentProvider(s);
    await applyFormatting.run(provider, {
      block_index: 0,
      line_spacing: 'at least 12pt',
    } as never);
    expect(docOf(s).child(0).attrs['spacing']).toMatchObject({
      line: 16,
      lineRule: 'atLeast',
    });
    await applyFormatting.run(provider, {
      block_index: 0,
      line_spacing: '18pt',
    } as never);
    expect(docOf(s).child(0).attrs['spacing']).toMatchObject({
      line: 24,
      lineRule: 'exact',
    });
    const bad = await applyFormatting.run(provider, {
      block_index: 0,
      line_spacing: 'wide',
    } as never);
    expect(bad.isError).toBe(true);
  });

  it('round-trips through .docx', async () => {
    const s = await open(['One']);
    await s.applyFormatting(
      { blockIndex: 0 },
      {
        spaceBefore: 12,
        lineSpacing: 2,
        indent: { left: '1in', right: '0.5in', firstLine: '0.25in' },
      },
    );
    let bytes: Uint8Array | undefined;
    (
      s as unknown as { opts: { onSave?: (b: Uint8Array) => void } }
    ).opts.onSave = (b) => {
      bytes = b;
    };
    await s.save();
    const { doc } = await importDocx(bytes!.slice().buffer as ArrayBuffer);
    const p = doc.child(0);
    expect(p.attrs['spacing']).toMatchObject({
      before: 16,
      line: 2,
      lineRule: 'auto',
    });
    expect(p.attrs['indent']).toEqual({ left: 96, right: 48, firstLine: 24 });
  });
});

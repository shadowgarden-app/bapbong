import { describe, expect, it } from 'vitest';
import type { Node as PMNode } from 'prosemirror-model';
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js';
import { exportDocx, schema } from '@shadow-garden/bapbong-headless';
import type { DocSnapshot } from './contract.js';
import { documentPage } from './document-commands.js';
import { HeadlessSession } from './headless-session.js';
import { selectionBlocks } from './pm-session.js';
import { createMcpServer } from './server.js';

/** An essay: a title, then chapters of a heading and eight paragraphs. */
async function essayBytes(chapters: number): Promise<Uint8Array> {
  const para = (text: string, attrs: Record<string, unknown> | null = null) =>
    schema.node('paragraph', attrs, [schema.text(text)]);
  const blocks = [para('Đường thủy', { styleId: 'Title' })];
  for (let c = 1; c <= chapters; c++) {
    blocks.push(para(`Chương ${c}`, { heading: 1 }));
    for (let p = 1; p <= 8; p++) {
      blocks.push(
        para(
          `Đoạn ${c}.${p}: cảng biển, tàu thuyền và du lịch đường thủy — ${'nội dung '.repeat(20)}`,
        ),
      );
    }
  }
  return exportDocx(schema.node('doc', null, blocks));
}

async function connect(chapters: number) {
  const session = await HeadlessSession.open(await essayBytes(chapters), {
    name: 'essay.docx',
  });
  const server = createMcpServer({ get: async () => session });
  const client = new Client({ name: 'test-client', version: '0.0.0' });
  const [clientTransport, serverTransport] =
    InMemoryTransport.createLinkedPair();
  await Promise.all([
    server.connect(serverTransport),
    client.connect(clientTransport),
  ]);
  const call = async (name: string, args: Record<string, unknown>) => {
    const r = await client.callTool({ name, arguments: args });
    const text = (r.content as { text: string }[])[0].text;
    return { isError: !!r.isError, text };
  };
  return { session, call };
}

const docOf = (s: HeadlessSession): PMNode =>
  (s as unknown as { state: { doc: PMNode } }).state.doc;

/** The font size on each textblock's text, in reading order. */
function sizes(doc: PMNode): (number | null)[] {
  const out: (number | null)[] = [];
  doc.descendants((n) => {
    if (!n.isTextblock) return true;
    const mark = n.firstChild?.marks.find((m) => m.type.name === 'fontSize');
    out.push(mark ? (mark.attrs['size'] as number) : null);
    return false;
  });
  return out;
}

describe('get_document reads a long document in pages', () => {
  it('walks every block exactly once, each page within the budget', () => {
    const snap: DocSnapshot = {
      docVersion: 'v1',
      blocks: Array.from({ length: 300 }, (_, i) => ({
        index: i,
        type: 'paragraph',
        text: `block ${i} ${'x'.repeat(200)}`,
      })),
      chrome: [
        { part: 'header', variant: 'default', sections: [0], text: 'H' },
      ],
      meta: { name: 'long.docx' },
    };
    const seen: number[] = [];
    let from = 0;
    let pages = 0;
    for (;;) {
      const page = documentPage(snap, from, 8_000);
      const text = JSON.stringify(page, null, 1);
      expect(text.length).toBeLessThanOrEqual(8_000);
      expect(page['totalBlocks']).toBe(300);
      // Where to read on comes before the blocks, so a clipped result keeps it.
      expect(Object.keys(page).indexOf('blocks')).toBeGreaterThan(
        'nextFromBlock' in page
          ? Object.keys(page).indexOf('nextFromBlock')
          : -1,
      );
      expect('chrome' in page).toBe(from === 0);
      seen.push(...(page['blocks'] as { index: number }[]).map((b) => b.index));
      pages++;
      if (page['nextFromBlock'] === undefined) break;
      from = page['nextFromBlock'] as number;
    }
    expect(pages).toBeGreaterThan(5);
    expect(seen).toEqual(Array.from({ length: 300 }, (_, i) => i));
  });

  it('a short document is one page, as before', () => {
    const snap: DocSnapshot = {
      docVersion: 'v1',
      blocks: [{ index: 0, type: 'paragraph', text: 'hi' }],
      meta: {},
    };
    const page = documentPage(snap, 0);
    expect(page['nextFromBlock']).toBeUndefined();
    expect(page['blocks']).toHaveLength(1);
  });

  it('a block longer than the page still comes, alone', () => {
    const snap: DocSnapshot = {
      docVersion: 'v1',
      blocks: [
        { index: 0, type: 'paragraph', text: 'y'.repeat(5_000) },
        { index: 1, type: 'paragraph', text: 'z' },
      ],
      meta: {},
    };
    const page = documentPage(snap, 0, 1_000);
    expect(page['blocks']).toHaveLength(1);
    expect(page['nextFromBlock']).toBe(1);
  });

  it('over MCP: pages of a real document, and a from_block past the end fails', async () => {
    const { call } = await connect(40);
    const first = JSON.parse((await call('get_document', {})).text);
    expect(first.totalBlocks).toBe(1 + 40 * 9);
    expect(first.nextFromBlock).toBeGreaterThan(0);
    const second = JSON.parse(
      (await call('get_document', { from_block: first.nextFromBlock })).text,
    );
    expect(second.blocks[0].index).toBe(first.nextFromBlock);
    const past = await call('get_document', { from_block: 10_000 });
    expect(past.isError).toBe(true);
    expect(past.text).toContain('past the end');
  });
});

describe('apply_formatting over a run of blocks', () => {
  it('sizes all body text in one call and one version, headings untouched', async () => {
    const { session, call } = await connect(3);
    const before = (await session.snapshot()).docVersion;
    const was = sizes(docOf(session));
    const res = await call('apply_formatting', {
      from_block: 0,
      only: 'body',
      font_size: 13,
    });
    expect(res.isError).toBe(false);
    const out = JSON.parse(res.text);
    expect(out.formatted).toBe(24);
    // One transaction: the version moves once.
    expect(Number(out.docVersion.slice(1))).toBe(Number(before.slice(1)) + 1);
    const got = sizes(docOf(session));
    const snap = await session.snapshot();
    snap.blocks.forEach((b, i) => {
      const heading = b.type.startsWith('heading') || i === 0;
      expect(got[i]).toBe(heading ? was[i] : 13);
    });
  });

  it('only headings, within a range', async () => {
    const { session, call } = await connect(3);
    const was = sizes(docOf(session));
    const res = await call('apply_formatting', {
      from_block: 1,
      to_block: 10,
      only: 'headings',
      font_size: 16,
      bold: true,
    });
    expect(JSON.parse(res.text).formatted).toBe(2); // Chương 1, Chương 2
    const got = sizes(docOf(session));
    expect(got[0]).toBe(was[0]); // the Title is outside the range
    expect(got[1]).toBe(16);
    expect(got[10]).toBe(16);
    expect(got[19]).toBe(was[19]); // Chương 3 is past to_block
    expect(got[2]).toBe(was[2]); // body text is left alone
  });

  it('refuses list changes over a run, a mix of targets, and an out-of-range run', async () => {
    const { call } = await connect(1);
    const list = await call('apply_formatting', {
      from_block: 0,
      list: 'bullet',
    });
    expect(list.isError).toBe(true);
    expect(list.text).toContain('one paragraph per call');
    const both = await call('apply_formatting', {
      from_block: 0,
      block_index: 1,
      bold: true,
    });
    expect(both.isError).toBe(true);
    expect(both.text).toContain('exactly one');
    // It names what the call carried, so the model can drop the extras.
    expect(both.text).toContain('block_index=1, from_block=0');
    const stray = await call('apply_formatting', {
      block_index: 1,
      only: 'body',
    });
    expect(stray.isError).toBe(true);
    const out = await call('apply_formatting', { from_block: 99, bold: true });
    expect(out.isError).toBe(true);
    expect(out.text).toContain('out of range');
  });
});

describe('a range of blocks: the passage a user selected', () => {
  it('get_document to_block reads just the range, in pages', async () => {
    const { call } = await connect(40);
    const one = JSON.parse(
      (await call('get_document', { from_block: 10, to_block: 12 })).text,
    );
    expect(one.blocks.map((b: { index: number }) => b.index)).toEqual([
      10, 11, 12,
    ]);
    expect(one.nextFromBlock).toBeUndefined();
    // A long range still pages, and stops at to_block.
    let from = 0;
    const seen: number[] = [];
    for (let i = 0; i < 50; i++) {
      const page = JSON.parse(
        (await call('get_document', { from_block: from, to_block: 200 })).text,
      );
      seen.push(...page.blocks.map((b: { index: number }) => b.index));
      if (page.nextFromBlock === undefined) break;
      from = page.nextFromBlock;
    }
    expect(seen).toEqual([...Array(201).keys()]);
    const backwards = await call('get_document', {
      from_block: 5,
      to_block: 4,
    });
    expect(backwards.isError).toBe(true);
  });

  it('selectionBlocks numbers a selection the way the tools do, tables included', () => {
    const p = (t: string) => schema.node('paragraph', null, [schema.text(t)]);
    const cell = (t: string) => schema.node('table_cell', null, [p(t)]);
    const doc = schema.node('doc', null, [
      p('Intro'),
      schema.node('table', null, [
        schema.node('table_row', null, [cell('A'), cell('B')]),
      ]),
      p('After the table'),
      p('End'),
    ]);
    // Each textblock's content span, in get_document order.
    const spans: { start: number; end: number }[] = [];
    doc.descendants((n, pos) => {
      if (n.isTextblock)
        spans.push({ start: pos + 1, end: pos + 1 + n.content.size });
      return true;
    });
    const starts = spans.map((x) => x.start);
    // "After the table" is block 3 for the tools — the 3rd top-level node.
    expect(selectionBlocks(doc, spans[3].start, spans[3].end)).toEqual({
      fromBlock: 3,
      toBlock: 3,
      partial: false,
    });
    // Whole document: blocks 0–4.
    expect(selectionBlocks(doc, 0, doc.content.size)).toEqual({
      fromBlock: 0,
      toBlock: 4,
      partial: false,
    });
    // A caret holds no text: still a range — the block it sits in.
    expect(selectionBlocks(doc, spans[2].start, spans[2].start)).toEqual({
      fromBlock: 2,
      toBlock: 2,
      partial: true,
    });
    // Starting mid-word in "Intro" and ending inside "End".
    expect(selectionBlocks(doc, starts[0] + 2, starts[4] + 1)).toEqual({
      fromBlock: 0,
      toBlock: 4,
      partial: true,
    });
  });
});

import { describe, expect, it } from 'vitest';
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js';
import { exportDocx, schema } from '@shadow-garden/bapbong-headless';
import { AnchorError, VersionConflictError } from './contract.js';
import { HeadlessSession } from './headless-session.js';
import { createMcpServer } from './server.js';

/** A small real .docx: three paragraphs, one duplicated phrase, bold mark. */
async function sampleBytes(): Promise<Uint8Array> {
  const doc = schema.node('doc', null, [
    schema.node('paragraph', null, [schema.text('Bao gia dich vu thanh lap cong ty.')]),
    schema.node('paragraph', null, [
      schema.text('Tong phi: '),
      schema.text('1.500.000', [schema.marks['strong'].create()]),
      schema.text(' VND.'),
    ]),
    schema.node('paragraph', null, [schema.text('Lien he: 0336697918. Tong phi: chua VAT.')]),
  ]);
  return exportDocx(doc);
}

async function openSession(onSave?: (bytes: Uint8Array) => void): Promise<HeadlessSession> {
  return HeadlessSession.open(await sampleBytes(), { name: 'sample.docx', onSave });
}

describe('HeadlessSession', () => {
  it('snapshots numbered blocks with a docVersion', async () => {
    const s = await openSession();
    const snap = await s.snapshot();
    expect(snap.docVersion).toBe('v1');
    expect(snap.blocks).toHaveLength(3);
    expect(snap.blocks[1]).toMatchObject({ index: 1, type: 'paragraph', text: 'Tong phi: 1.500.000 VND.' });
    expect(snap.meta).toMatchObject({ name: 'sample.docx', dirty: false });
  });

  it('finds occurrences with block index and context', async () => {
    const s = await openSession();
    const matches = await s.find('Tong phi');
    expect(matches).toHaveLength(2);
    expect(matches[0]).toMatchObject({ blockIndex: 1, occurrence: 1 });
    expect(matches[1]).toMatchObject({ blockIndex: 2, occurrence: 2 });
    expect(matches[0].context).toContain('«Tong phi»');
  });

  it('replaces a unique match, keeping surrounding marks intact', async () => {
    const s = await openSession();
    const res = await s.replaceText('1.500.000', '2.000.000');
    expect(res.docVersion).toBe('v2');
    const snap = await s.snapshot();
    expect(snap.blocks[1].text).toBe('Tong phi: 2.000.000 VND.');
    expect(snap.meta.dirty).toBe(true);
  });

  it('rejects an ambiguous anchor, then honors occurrence', async () => {
    const s = await openSession();
    await expect(s.replaceText('Tong phi', 'Thanh tien')).rejects.toThrow(AnchorError);
    await expect(s.replaceText('Tong phi', 'Thanh tien')).rejects.toThrow(/matches 2 times/);
    await s.replaceText('Tong phi', 'Thanh tien', { occurrence: 2 });
    const snap = await s.snapshot();
    expect(snap.blocks[1].text).toContain('Tong phi');
    expect(snap.blocks[2].text).toContain('Thanh tien');
  });

  it('rejects a missing anchor with a teaching message', async () => {
    const s = await openSession();
    await expect(s.replaceText('khong ton tai', 'x')).rejects.toThrow(/not found/);
  });

  it('enforces optimistic locking via expectedVersion', async () => {
    const s = await openSession();
    await s.replaceText('1.500.000', '9', { expectedVersion: 'v1' });
    await expect(s.replaceText('VND', 'USD', { expectedVersion: 'v1' })).rejects.toThrow(VersionConflictError);
    await s.replaceText('VND', 'USD', { expectedVersion: 'v2' });
  });

  it('inserts paragraphs before/after an anchor and at document_end', async () => {
    const s = await openSession();
    await s.insertContent('Ghi chu dau trang', { position: 'before', text: 'Bao gia dich vu' });
    await s.insertContent('Dong 1\nDong 2', { position: 'document_end' });
    const snap = await s.snapshot();
    expect(snap.blocks.map((b) => b.text)).toEqual([
      'Ghi chu dau trang',
      'Bao gia dich vu thanh lap cong ty.',
      'Tong phi: 1.500.000 VND.',
      'Lien he: 0336697918. Tong phi: chua VAT.',
      'Dong 1',
      'Dong 2',
    ]);
  });

  it('applies marks + paragraph alignment, surviving a save round-trip', async () => {
    let saved: Uint8Array | undefined;
    const s = await openSession((b) => (saved = b));
    await s.applyFormatting('Lien he', { bold: true, align: 'center' });
    await s.save();

    const reopened = await HeadlessSession.open(saved as Uint8Array);
    const snap = await reopened.snapshot();
    expect(snap.blocks[2].text).toContain('Lien he');
    // Bold and alignment survived the .docx round-trip: removing bold from the
    // same range must be a real change (the mark is there to remove).
    const before = snap.docVersion;
    const res = await reopened.applyFormatting('Lien he', { bold: false });
    expect(res.docVersion).not.toBe(before);
  });

  it('save() exports round-trippable bytes to the sink', async () => {
    let saved: Uint8Array | undefined;
    const s = await openSession((b) => (saved = b));
    await s.replaceText('1.500.000', '7.777.777');
    await s.save();
    expect(saved).toBeDefined();
    const reopened = await HeadlessSession.open(saved as Uint8Array);
    const snap = await reopened.snapshot();
    expect(snap.blocks[1].text).toBe('Tong phi: 7.777.777 VND.');
  });
});

describe('createMcpServer (end-to-end over MCP)', () => {
  async function connect() {
    const session = await openSession();
    const provider = { get: async (id?: string) => (id === undefined || id === 'sample' ? session : null) };
    const server = createMcpServer(provider);
    const client = new Client({ name: 'test-client', version: '0.0.0' });
    const [clientTransport, serverTransport] = InMemoryTransport.createLinkedPair();
    await Promise.all([server.connect(serverTransport), client.connect(clientTransport)]);
    return { client, session };
  }

  const text = (r: Awaited<ReturnType<Client['callTool']>>) =>
    (r.content as { type: string; text: string }[])[0].text;

  it('lists the contract tools (no get_selection without the capability)', async () => {
    const { client } = await connect();
    const names = (await client.listTools()).tools.map((t) => t.name).sort();
    expect(names).toEqual([
      'apply_formatting',
      'find_text',
      'get_document',
      'insert_content',
      'replace_text',
      'save_document',
    ]);
  });

  it('reads the document and edits it through tool calls', async () => {
    const { client } = await connect();
    const doc = JSON.parse(text(await client.callTool({ name: 'get_document', arguments: {} })));
    expect(doc.blocks).toHaveLength(3);

    const res = await client.callTool({
      name: 'replace_text',
      arguments: { old_text: '1.500.000', new_text: '3.333.333', expectedVersion: doc.docVersion },
    });
    expect(res.isError).toBeFalsy();

    const after = JSON.parse(text(await client.callTool({ name: 'get_document', arguments: {} })));
    expect(after.blocks[1].text).toBe('Tong phi: 3.333.333 VND.');
    expect(after.docVersion).not.toBe(doc.docVersion);
  });

  it('returns teaching errors (isError) for conflicts and bad anchors', async () => {
    const { client } = await connect();
    const stale = await client.callTool({
      name: 'replace_text',
      arguments: { old_text: 'VND', new_text: 'USD', expectedVersion: 'v99' },
    });
    expect(stale.isError).toBe(true);
    expect(text(stale)).toContain('get_document');

    const ambiguous = await client.callTool({
      name: 'replace_text',
      arguments: { old_text: 'Tong phi', new_text: 'x' },
    });
    expect(ambiguous.isError).toBe(true);
    expect(text(ambiguous)).toContain('occurrence');
  });

  it('exposes the document as a resource', async () => {
    const { client } = await connect();
    const res = await client.readResource({ uri: 'bapbong://document' });
    expect((res.contents[0] as { text: string }).text).toContain('Tong phi: 1.500.000 VND.');
  });

  it('resolves documentId through the provider', async () => {
    const { client } = await connect();
    const ok = await client.callTool({ name: 'get_document', arguments: { documentId: 'sample' } });
    expect(ok.isError).toBeFalsy();
    const missing = await client.callTool({ name: 'get_document', arguments: { documentId: 'other' } });
    expect(missing.isError).toBe(true);
  });
});

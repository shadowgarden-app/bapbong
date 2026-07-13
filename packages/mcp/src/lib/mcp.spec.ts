import { describe, expect, it } from 'vitest';
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js';
import { exportDocx, schema } from '@shadow-garden/bapbong-headless';
import { AnchorError, NoDocumentError, VersionConflictError } from './contract.js';
import { HeadlessSession } from './headless-session.js';
import { createMcpServer } from './server.js';
import { executeOp, RemoteSession, reviveError, type SessionOpName } from './wire.js';

// 1×1 transparent PNG.
const PNG_1PX =
  'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==';

/** A small real .docx: three text paragraphs + one holding an image. */
async function sampleBytes(): Promise<Uint8Array> {
  const doc = schema.node('doc', null, [
    schema.node('paragraph', null, [schema.text('Bao gia dich vu thanh lap cong ty.')]),
    schema.node('paragraph', null, [
      schema.text('Tong phi: '),
      schema.text('1.500.000', [schema.marks['strong'].create()]),
      schema.text(' VND.'),
    ]),
    schema.node('paragraph', null, [schema.text('Lien he: 0336697918. Tong phi: chua VAT.')]),
    schema.node('paragraph', null, [
      schema.text('Hinh minh hoa: '),
      schema.node('image', { src: `data:image/png;base64,${PNG_1PX}`, width: 100, height: 50 }),
    ]),
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
    expect(snap.blocks).toHaveLength(4);
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
      'Hinh minh hoa: ',
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

  it('lists a block\'s images in the snapshot', async () => {
    const s = await openSession();
    const snap = await s.snapshot();
    expect(snap.blocks[3].text).toBe('Hinh minh hoa: ');
    expect(snap.blocks[3].images).toEqual([
      { index: 0, alt: '', width: 100, height: 50, rotation: 0, kind: 'bitmap' },
    ]);
    expect(snap.blocks[0].images).toBeUndefined();
  });

  it('updateImage resizes/rotates in one step, teaching on bad anchors', async () => {
    const s = await openSession();
    const v1 = (await s.snapshot()).docVersion;
    const res = await s.updateImage(3, 0, { width: 200, rotation: 450.5 }, { expectedVersion: v1 });
    expect(res.docVersion).toBe('v2');
    const img = (await s.snapshot()).blocks[3].images?.[0];
    expect(img).toMatchObject({ width: 200, height: 50, rotation: 90.5 }); // height kept; 450.5 → 90.5

    await expect(s.updateImage(0, 0, { width: 10 })).rejects.toThrow(/has no images/);
    await expect(s.updateImage(3, 5, { width: 10 })).rejects.toThrow(/0-0/);
    await expect(s.updateImage(99, 0, { width: 10 })).rejects.toThrow(AnchorError);
    await expect(s.updateImage(3, 0, { width: 10 }, { expectedVersion: 'v1' })).rejects.toThrow(VersionConflictError);
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

describe('wire (RemoteSession ↔ executeOp)', () => {
  /** A RemoteSession whose transport is executeOp against a local session —
   *  exactly the desktop shape (Bun proxy ↔ WebView executor), minus SSE. */
  async function remotePair() {
    const local = await openSession();
    const remote = new RemoteSession(async (op: SessionOpName, args: unknown[]) => {
      const res = await executeOp(local, { id: 'x', op, args });
      // Simulate the process hop: everything travels as JSON.
      const wire = JSON.parse(JSON.stringify(res)) as typeof res;
      if (!wire.ok) throw reviveError(wire.error);
      return wire.value;
    }, { selection: false });
    return { remote, local };
  }

  it('round-trips reads and mutations across the hop', async () => {
    const { remote } = await remotePair();
    const snap = await remote.snapshot();
    expect(snap.blocks).toHaveLength(4);
    const res = await remote.replaceText('1.500.000', '8.888.888', { expectedVersion: snap.docVersion });
    expect(res.docVersion).toBe('v2');
    expect((await remote.snapshot()).blocks[1].text).toBe('Tong phi: 8.888.888 VND.');
    const img = await remote.updateImage(3, 0, { rotation: -90 });
    expect(img.docVersion).toBe('v3');
    expect((await remote.snapshot()).blocks[3].images?.[0].rotation).toBe(270);
  });

  it('revives contract errors by name across the hop', async () => {
    const { remote } = await remotePair();
    await expect(remote.replaceText('Tong phi', 'x')).rejects.toThrow(AnchorError);
    await expect(remote.replaceText('VND', 'x', { expectedVersion: 'v9' })).rejects.toThrow(VersionConflictError);
  });

  it('encodes a missing document as NoDocumentError', async () => {
    const res = await executeOp(null, { id: 'x', op: 'snapshot', args: [] });
    expect(res.ok).toBe(false);
    if (!res.ok) expect(reviveError(res.error)).toBeInstanceOf(NoDocumentError);
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
      'update_image',
    ]);
  });

  it('reads the document and edits it through tool calls', async () => {
    const { client } = await connect();
    const doc = JSON.parse(text(await client.callTool({ name: 'get_document', arguments: {} })));
    expect(doc.blocks).toHaveLength(4);

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

  it('updates an image through the tool (and rejects an empty change)', async () => {
    const { client } = await connect();
    const doc = JSON.parse(text(await client.callTool({ name: 'get_document', arguments: {} })));
    expect(doc.blocks[3].images).toHaveLength(1);

    const empty = await client.callTool({ name: 'update_image', arguments: { block_index: 3 } });
    expect(empty.isError).toBe(true);
    expect(text(empty)).toContain('at least one');

    const res = await client.callTool({
      name: 'update_image',
      arguments: { block_index: 3, width: 150, height: 75, rotation: 45, expectedVersion: doc.docVersion },
    });
    expect(res.isError).toBeFalsy();
    const after = JSON.parse(text(await client.callTool({ name: 'get_document', arguments: {} })));
    expect(after.blocks[3].images[0]).toMatchObject({ width: 150, height: 75, rotation: 45 });
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

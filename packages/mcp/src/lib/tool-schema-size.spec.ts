import { describe, expect, it } from 'vitest';
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js';
import { exportDocx, schema } from '@shadow-garden/bapbong-headless';
import { HeadlessSession } from './headless-session.js';
import { createMcpServer } from './server.js';

async function connect() {
  const bytes = await exportDocx(
    schema.node('doc', null, [
      schema.node('paragraph', null, [schema.text('Cảng biển loại 1.')]),
    ]),
  );
  const session = await HeadlessSession.open(bytes, { name: 'a.docx' });
  const server = createMcpServer(
    { get: async () => session },
    {
      selection: true,
      images: true,
      pageSetup: true,
      headerFooter: true,
      styles: true,
    },
  );
  const client = new Client({ name: 'test-client', version: '0.0.0' });
  const [c, s] = InMemoryTransport.createLinkedPair();
  await Promise.all([server.connect(s), client.connect(c)]);
  return client;
}

describe('the content grammar is spelled out once in the tool list', () => {
  it('insert_content carries the full block schema; the other content tools only its outline', async () => {
    const client = await connect();
    const { tools } = await client.listTools();
    const size = (name: string) =>
      JSON.stringify(tools.find((t) => t.name === name)?.inputSchema).length;
    // The full grammar is some 7K characters of schema.
    expect(size('insert_content')).toBeGreaterThan(5_000);
    expect(size('insert_footnote')).toBeLessThan(1_500);
    expect(size('edit_header_footer')).toBeLessThan(2_500);
  });

  it('a malformed block is still refused, in words that point at the grammar', async () => {
    const client = await connect();
    const bad = await client.callTool({
      name: 'insert_footnote',
      arguments: {
        anchor_text: 'Cảng biển',
        content: [{ paragraph: 'ok', heading: 9 }],
      },
    });
    expect(bad.isError).toBe(true);
    const text = (bad.content as { text: string }[])[0].text;
    expect(text).toContain('content.0');
    expect(text).toContain('insert_content');

    const good = await client.callTool({
      name: 'insert_footnote',
      arguments: { anchor_text: 'Cảng biển', content: 'Theo Luật Hàng hải.' },
    });
    expect(good.isError).toBeFalsy();
  });
});

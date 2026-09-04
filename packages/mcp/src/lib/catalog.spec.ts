import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { InMemoryTransport } from '@modelcontextprotocol/sdk/inMemory.js';
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { z } from 'zod';
import { describe, expect, it } from 'vitest';
import {
  defineCommand,
  isOffered,
  json,
  registerCommands,
  type AgentCommand,
} from './catalog.js';
import { documentCommands } from './document-commands.js';

const byName = (name: string) => {
  const c = documentCommands.find((c) => c.name === name);
  if (!c) throw new Error(`no command ${name}`);
  return c;
};

describe('agent catalog', () => {
  it('ships the eight document commands as data, in reading order', () => {
    expect(documentCommands.map((c) => c.name)).toEqual([
      'get_document',
      'find_text',
      'replace_text',
      'insert_content',
      'apply_formatting',
      'update_image',
      'save_document',
      'get_selection',
    ]);
  });

  it('labels every command with the effect a gate decides on', () => {
    const effects = Object.fromEntries(
      documentCommands.map((c) => [c.name, c.effect]),
    );
    expect(effects).toEqual({
      get_document: 'read',
      find_text: 'read',
      replace_text: 'edit',
      insert_content: 'edit',
      apply_formatting: 'edit',
      update_image: 'edit',
      save_document: 'save',
      get_selection: 'read',
    });
  });

  it('targets the documentId argument (undefined = the open document)', () => {
    for (const c of documentCommands) {
      expect(c.targets({ documentId: 'a.docx' })).toEqual(['a.docx']);
      expect(c.targets({})).toEqual([undefined]);
    }
  });

  it('offers get_selection only to hosts with a selection', () => {
    const sel = byName('get_selection');
    expect(isOffered(sel, {})).toBe(false);
    expect(isOffered(sel, { selection: true })).toBe(true);
    const doc = byName('get_document');
    expect(isOffered(doc, {})).toBe(true);
  });

  it('registers a host command on MCP the same way, through an interposer', async () => {
    // A host command over its OWN ports — nothing document-shaped about it.
    const echo = defineCommand<{ text: z.ZodString }, { prefix: string }>({
      name: 'echo',
      title: 'Echo',
      description: 'Echoes.',
      input: { text: z.string() },
      effect: 'read',
      targets: () => [],
      run: async (ports, { text }) => json({ out: ports.prefix + text }),
    });
    const seen: string[] = [];
    const server = new McpServer({ name: 't', version: '0' });
    registerCommands(
      server,
      [echo as AgentCommand<z.ZodRawShape, { prefix: string }>],
      { prefix: '> ' },
      {},
      // The seam a permission gate sits in: every call passes through here.
      async (command, args) => {
        seen.push(`${command.name}:${command.effect}`);
        return command.run({ prefix: '> ' }, args);
      },
    );
    const client = new Client({ name: 'c', version: '0' });
    const [ct, st] = InMemoryTransport.createLinkedPair();
    await Promise.all([server.connect(st), client.connect(ct)]);
    const tools = (await client.listTools()).tools;
    expect(tools.map((t) => t.name)).toEqual(['echo']);
    expect(tools[0].description).toBe('Echoes.');
    const res = await client.callTool({
      name: 'echo',
      arguments: { text: 'hi' },
    });
    expect(JSON.parse((res.content as { text: string }[])[0].text)).toEqual({
      out: '> hi',
    });
    expect(seen).toEqual(['echo:read']);
  });
});

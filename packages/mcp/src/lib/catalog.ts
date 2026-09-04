/**
 * Agent catalog — the shape of ONE command an AI agent can run, as data.
 *
 * Whatever door the agent comes through (MCP, a CLI, an in-app assistant),
 * it runs the same commands: a command is a record — name, description,
 * input schema, effect, and a `run` over host-supplied ports — and each
 * transport is a thin adapter that reads the records and translates syntax.
 * Nothing in a record knows which transport invoked it, so the host's
 * permission gate can sit in ONE place, between the adapters and the ports,
 * and no adapter can route around it.
 *
 * This package ships the engine half: the type, the helpers, and the
 * document commands ({@link documentCommands} in ./document-commands). A
 * host adds its own records (folder-level tools, UI actions) beside them.
 */
import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import type { z } from 'zod';
import {
  AnchorError,
  NoDocumentError,
  ReadOnlyError,
  VersionConflictError,
  type DocumentSession,
  type SessionProvider,
} from './contract.js';

/** What a document command does — the axis a permission gate decides on.
 *  `read` touches nothing; `edit` changes content; `save` persists it. This
 *  package works on ONE document, so these are the only effects it knows; a
 *  host with more (a folder to create or move files in, a window to drive)
 *  widens the type for its own commands via {@link AgentCommand}'s `Effect`. */
export type DocumentEffect = 'read' | 'edit' | 'save';

/** What a command hands back — the MCP tool-result shape, so an MCP adapter
 *  forwards it untouched and every other adapter unwraps the same fields.
 *  (A type literal, not an interface: the SDK's result type carries an index
 *  signature, which only type literals satisfy implicitly.) */
export type CommandResult = {
  content: { type: 'text'; text: string }[];
  isError?: true;
};

/** The arguments a command receives: its zod shape, parsed. */
export type CommandArgs<Shape extends z.ZodRawShape> = z.objectOutputType<
  Shape,
  z.ZodTypeAny
>;

/** A host capability a command needs; hosts without it leave it out. */
export type CommandRequirement = 'selection';

/**
 * One command, as data. `Ports` is whatever the host gives commands to act
 * through — for the document commands it is a {@link SessionProvider}; a host
 * with a file system or a window supplies a richer object for its own
 * commands, and a wider `Effect` to classify them.
 */
export interface AgentCommand<
  Shape extends z.ZodRawShape = z.ZodRawShape,
  Ports = SessionProvider,
  Effect extends string = DocumentEffect,
> {
  name: string;
  title: string;
  /** Written once, for the agent — every adapter shows this same text. */
  description: string;
  /** Zod raw shape: MCP takes it as-is; a CLI derives flags from it. */
  input: Shape;
  effect: Effect;
  requires?: CommandRequirement;
  /** The document ids (or paths) the call touches — what a gate resolves
   *  permission for. `undefined` means "the document the user has open". */
  targets(args: CommandArgs<Shape>): (string | undefined)[];
  run(ports: Ports, args: CommandArgs<Shape>): Promise<CommandResult>;
}

/** Identity with inference: `run`/`targets` see the typed args of `input`. */
export function defineCommand<
  Shape extends z.ZodRawShape,
  Ports = SessionProvider,
  Effect extends string = DocumentEffect,
>(
  command: AgentCommand<Shape, Ports, Effect>,
): AgentCommand<Shape, Ports, Effect> {
  return command;
}

/** JSON text content for a command result. */
export function json(value: unknown): CommandResult {
  return { content: [{ type: 'text', text: JSON.stringify(value, null, 1) }] };
}

/** A failed result whose text teaches the agent the retry. */
export function errorText(message: string): CommandResult {
  return { content: [{ type: 'text', text: message }], isError: true };
}

/** Resolve a session through the provider and translate contract errors
 *  into teaching text; anything else is a host bug and propagates. */
export async function withSession(
  provider: SessionProvider,
  id: string | undefined,
  fn: (session: DocumentSession) => Promise<CommandResult>,
): Promise<CommandResult> {
  const session = await provider.get(id);
  if (!session) {
    return errorText(
      id === undefined
        ? 'No document is open. Ask the user to open a document first.'
        : `No document with id ${JSON.stringify(id)}.`,
    );
  }
  try {
    return await fn(session);
  } catch (err) {
    if (
      err instanceof AnchorError ||
      err instanceof VersionConflictError ||
      err instanceof NoDocumentError ||
      err instanceof ReadOnlyError
    ) {
      return errorText(err.message);
    }
    throw err;
  }
}

/** Which optional capabilities a host offers (see {@link CommandRequirement}). */
export interface HostCapabilities {
  selection?: boolean;
}

/** Does the host offer everything this command needs? */
export function isOffered(
  command: Pick<AgentCommand, 'requires'>,
  caps: HostCapabilities,
): boolean {
  return command.requires === undefined || caps[command.requires] === true;
}

/**
 * The MCP adapter for a set of commands: one `registerTool` per record,
 * title/description/schema straight from the data. A host uses this for its
 * own records too, so its tools and the engine's are registered the same way.
 *
 * `execute` lets a host interpose (permission gate, logging) between the
 * transport and `run`; the default runs the command directly.
 */
export function registerCommands<Ports, Effect extends string>(
  server: McpServer,
  commands: readonly AgentCommand<z.ZodRawShape, Ports, Effect>[],
  ports: Ports,
  caps: HostCapabilities = {},
  execute: (
    command: AgentCommand<z.ZodRawShape, Ports, Effect>,
    args: CommandArgs<z.ZodRawShape>,
  ) => Promise<CommandResult> = (command, args) => command.run(ports, args),
): void {
  for (const command of commands) {
    if (!isOffered(command, caps)) continue;
    server.registerTool(
      command.name,
      {
        title: command.title,
        description: command.description,
        inputSchema: command.input,
      },
      async (args) => execute(command, args as CommandArgs<z.ZodRawShape>),
    );
  }
}

/** The `bapbong://document` resource: the open document as plain text, one
 *  line per block. Registered by {@link createMcpServer}; a host that builds
 *  its own server from the records registers it the same way. */
export function registerDocumentResource(
  server: McpServer,
  provider: SessionProvider,
): void {
  server.registerResource(
    'document',
    'bapbong://document',
    {
      title: 'Open document',
      description:
        'The currently open document as plain text (one line per block).',
      mimeType: 'text/plain',
    },
    async (uri) => {
      const session = await provider.get(undefined);
      if (!session)
        return { contents: [{ uri: uri.href, text: '(no document open)' }] };
      const snap = await session.snapshot();
      return {
        contents: [
          { uri: uri.href, text: snap.blocks.map((b) => b.text).join('\n') },
        ],
      };
    },
  );
}

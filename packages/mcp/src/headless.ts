/**
 * Opt-in headless-document surface: {@link HeadlessSession} runs a real
 * ProseMirror document with no DOM, importing/exporting .docx itself.
 *
 * It is a SEPARATE entry from ./host on purpose. A host that only proxies to a
 * live editor (the desktop shell's original shape) must not drag the docx
 * pipeline into its process; a host that also serves documents straight from
 * storage — the workspace folder, a server's object store — imports this and
 * takes that dependency deliberately.
 */
export {
  HeadlessSession,
  type HeadlessSessionOptions,
} from './lib/headless-session.js';

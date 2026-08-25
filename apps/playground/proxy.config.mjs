/**
 * Dev-server API for the playground.
 *
 * Angular's dev server hands this file straight to Vite's `server.proxy`, and
 * Vite calls `bypass` before it ever touches `target`: answering there and
 * returning a string ends the request locally (Vite bails out as soon as it
 * sees `res.writableEnded`). So `/api/samples` is a real endpoint without a
 * second process to run.
 *
 * Dev only — a built playground (`serve-static`) has no such route, and the UI
 * falls back to the handful of samples it knows by name.
 */
import { readdir, stat } from 'node:fs/promises';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const PUBLIC_DIR = join(dirname(fileURLToPath(import.meta.url)), 'public');

/** Every .docx sitting in `public/`, name-sorted, with its size. */
async function listSamples() {
  const entries = await readdir(PUBLIC_DIR, { withFileTypes: true });
  const names = entries
    .filter(
      (e) => e.isFile() && !e.name.startsWith('.') && /\.docx$/i.test(e.name),
    )
    .map((e) => e.name);

  const files = await Promise.all(
    names.map(async (name) => ({
      name,
      size: (await stat(join(PUBLIC_DIR, name))).size,
    })),
  );

  return files.sort((a, b) =>
    a.name.localeCompare(b.name, 'en', { sensitivity: 'base' }),
  );
}

export default {
  '/api/samples': {
    // Never contacted: `bypass` always answers the request itself.
    target: 'http://localhost',
    bypass: async (req, res) => {
      try {
        const body = JSON.stringify({ files: await listSamples() });
        res.setHeader('Content-Type', 'application/json; charset=utf-8');
        res.setHeader('Cache-Control', 'no-store');
        res.end(body);
      } catch (err) {
        res.statusCode = 500;
        res.setHeader('Content-Type', 'application/json; charset=utf-8');
        res.end(JSON.stringify({ error: String(err?.message ?? err) }));
      }
      // A string + a finished response is Vite's "handled, stop here" signal.
      return req.url;
    },
  },
};

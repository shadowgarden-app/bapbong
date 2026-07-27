// EditorView via input-bridge's re-export (not straight from prosemirror-view)
// so the view a caller gets from `bridge.view` is the same identity these
// handlers expect — see the note in input-bridge.
import type { EditorView } from '@shadow-garden/bapbong-input-bridge';

/**
 * Image paste support.
 *
 * Pasted bitmaps are embedded as data URLs — the same representation the DOCX
 * importer uses for `image.src` — so layout, painting (painter's image cache)
 * and DOCX export (media extraction) all work unchanged. Only the acquisition
 * is new: clipboard blob → data URL → inline image node.
 */

/** Widest a pasted image may display, in CSS px. Roughly the text width of an
 *  A4/Letter page with default margins; larger bitmaps keep their pixels (the
 *  data URL is the original) but display scaled down, preserving aspect. */
const MAX_DISPLAY_WIDTH = 600;

/** Image blobs in a DataTransfer (clipboard or drop), in item order. */
export function imageFilesIn(dt: DataTransfer | null): File[] {
  const files: File[] = [];
  for (const item of Array.from(dt?.items ?? [])) {
    if (item.kind === 'file' && item.type.startsWith('image/')) {
      const f = item.getAsFile();
      if (f) files.push(f);
    }
  }
  return files;
}

/** Read intrinsic size without touching the DOM (no <img> element churn). */
async function bitmapSize(
  blob: Blob,
): Promise<{ width: number; height: number } | null> {
  try {
    const bmp = await createImageBitmap(blob);
    const size = { width: bmp.width, height: bmp.height };
    bmp.close();
    return size;
  } catch {
    return null; // undecodable blob (corrupt, or an unsupported format)
  }
}

const toDataURL = (blob: Blob) =>
  new Promise<string>((resolve, reject) => {
    const r = new FileReader();
    r.onload = () => resolve(r.result as string);
    r.onerror = () => reject(r.error);
    r.readAsDataURL(blob);
  });

/** Insert image blobs at the selection as inline image nodes (one
 *  transaction). Resolves true if anything was inserted. */
export async function insertImageBlobs(
  view: EditorView,
  blobs: Blob[],
): Promise<boolean> {
  const imageType = view.state.schema.nodes['image'];
  if (!imageType) return false; // schema without images (e.g. comment composer)

  const nodes = [];
  for (const blob of blobs) {
    const size = await bitmapSize(blob);
    if (!size) continue;
    const scale = Math.min(1, MAX_DISPLAY_WIDTH / size.width);
    nodes.push(
      imageType.create({
        src: await toDataURL(blob),
        width: Math.round(size.width * scale),
        height: Math.round(size.height * scale),
      }),
    );
  }
  if (!nodes.length) return false;

  let tr = view.state.tr;
  for (const node of nodes) tr = tr.replaceSelectionWith(node);
  view.dispatch(tr.scrollIntoView());
  return true;
}

/** editorProps.handlePaste: claim the event only when the clipboard's best
 *  payload is an image. Rich HTML (real text content) wins over the blob —
 *  ProseMirror's default paste parses it through the schema; but HTML that is
 *  merely an <img> wrapper (how browsers copy a picture) would lose the image
 *  there (remote srcs are rejected by the schema), so the blob is used. */
export function imagePasteHandler(
  view: EditorView,
  event: ClipboardEvent,
): boolean {
  const dt = event.clipboardData;
  const files = imageFilesIn(dt);
  if (!files.length) return false;

  const html = dt?.getData('text/html') ?? '';
  if (html) {
    const body = new DOMParser().parseFromString(html, 'text/html').body;
    if ((body.textContent ?? '').trim()) return false; // real text → PM default
  }

  void insertImageBlobs(view, files);
  return true; // claimed: the async insert lands when the blobs are decoded
}

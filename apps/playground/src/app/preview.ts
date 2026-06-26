import { Component, ElementRef, OnDestroy, inject, signal, viewChild } from '@angular/core';
import { ActivatedRoute } from '@angular/router';
import { BapbongView } from '@shadow-garden/bapbong-view';

type PreviewState = 'idle' | 'loading' | 'rendered' | 'error';

const ZOOM_MIN = 0.5;
const ZOOM_MAX = 2;
const ZOOM_STEP = 0.1;

/**
 * Read-only document preview route (`/preview`). With a `?link=<url>` query it
 * fetches and renders that `.docx`; without one it shows a drag-and-drop upload
 * box. Non-`.docx` input falls into a "can't preview" state. Rendering uses the
 * M9 {@link BapbongView} (the read-only tier — no editing, no input-bridge).
 */
@Component({
  selector: 'app-preview',
  templateUrl: './preview.html',
  styleUrl: './preview.css',
})
export class Preview implements OnDestroy {
  private readonly route = inject(ActivatedRoute);
  private readonly host = viewChild<ElementRef<HTMLElement>>('host');

  protected readonly state = signal<PreviewState>('idle');
  protected readonly fileName = signal<string | null>(null);
  protected readonly errorMsg = signal<string | null>(null);
  protected readonly pageCount = signal(0);
  protected readonly zoomPct = signal(100);
  protected readonly dragOver = signal(false);
  /** The source for the loading state's caption (host of the `?link` URL). */
  protected readonly source = signal<string | null>(null);

  private view: BapbongView | null = null;

  constructor() {
    const link = this.route.snapshot.queryParamMap.get('link');
    if (link) this.loadFromUrl(link);
  }

  ngOnDestroy(): void {
    this.view?.destroy();
    this.view = null;
  }

  // ── Loading ─────────────────────────────────────────────────────────

  /** Fetch a `.docx` from `url` and render it (or surface a failure). */
  private async loadFromUrl(url: string): Promise<void> {
    const name = fileNameFromUrl(url);
    this.fileName.set(name);
    try {
      this.source.set(new URL(url, location.href).host);
    } catch {
      this.source.set(null);
    }
    this.state.set('loading');
    try {
      const res = await fetch(url);
      if (!res.ok) throw new Error(`HTTP ${res.status}`);
      const bytes = await res.arrayBuffer();
      await this.validateAndRender(bytes, name);
    } catch (e) {
      this.fail(name, `Couldn't load the document from this link (${(e as Error).message}). ` +
        `The server may block cross-origin requests, or the link may be wrong.`);
    }
  }

  /** Handle a file chosen via the picker. */
  protected onFileInput(ev: Event): void {
    const file = (ev.target as HTMLInputElement).files?.[0];
    if (file) void this.loadFile(file);
  }

  /** Handle a file dropped on the upload zone. */
  protected onDrop(ev: DragEvent): void {
    ev.preventDefault();
    this.dragOver.set(false);
    const file = ev.dataTransfer?.files?.[0];
    if (file) void this.loadFile(file);
  }

  protected onDragOver(ev: DragEvent): void {
    ev.preventDefault();
    this.dragOver.set(true);
  }

  protected onDragLeave(): void {
    this.dragOver.set(false);
  }

  /** Reset to the upload box (e.g. the error state's "choose another"). */
  protected reset(): void {
    this.state.set('idle');
    this.errorMsg.set(null);
    this.fileName.set(null);
  }

  /** Open a file picker to load a different document (toolbar "open" button). */
  protected openFile(): void {
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = '.docx';
    input.addEventListener('change', () => {
      const file = input.files?.[0];
      if (file) void this.loadFile(file);
    });
    input.click();
  }

  /** Print the rendered document (toolbar "print" button). */
  protected print(): void {
    void this.view?.print();
  }

  private async loadFile(file: File): Promise<void> {
    this.fileName.set(file.name);
    this.state.set('loading');
    this.source.set(null);
    try {
      const bytes = await file.arrayBuffer();
      await this.validateAndRender(bytes, file.name);
    } catch (e) {
      this.fail(file.name, (e as Error).message);
    }
  }

  // ── Validate + render ───────────────────────────────────────────────

  private async validateAndRender(bytes: ArrayBuffer, name: string): Promise<void> {
    const reason = whyNotDocx(bytes, name);
    if (reason) {
      this.fail(name, reason);
      return;
    }
    this.fileName.set(name);
    this.state.set('rendered');
    // Wait a frame so the (now-visible) host has real dimensions before the
    // viewer measures the viewport for page virtualization.
    await nextFrame();
    const host = this.host()?.nativeElement;
    if (!host) return;
    try {
      if (!this.view) this.view = new BapbongView(host, { viewport: host, a11yLabel: name });
      await this.view.loadDocx(bytes);
      this.pageCount.set(this.view.pageCount);
      this.zoomPct.set(Math.round(this.view.getZoom() * 100));
    } catch (e) {
      this.fail(name, `This file looks like a .docx but couldn't be read (${(e as Error).message}).`);
    }
  }

  private fail(name: string | null, msg: string): void {
    this.fileName.set(name);
    this.errorMsg.set(msg);
    this.state.set('error');
  }

  // ── Zoom ────────────────────────────────────────────────────────────

  protected zoomIn(): void {
    this.applyZoom((this.view?.getZoom() ?? 1) + ZOOM_STEP);
  }

  protected zoomOut(): void {
    this.applyZoom((this.view?.getZoom() ?? 1) - ZOOM_STEP);
  }

  private applyZoom(z: number): void {
    if (!this.view) return;
    const clamped = Math.min(ZOOM_MAX, Math.max(ZOOM_MIN, z));
    this.view.setZoom(clamped);
    this.zoomPct.set(Math.round(clamped * 100));
  }
}

/** A frame's delay (lets a state change reflect in the DOM before measuring). */
function nextFrame(): Promise<void> {
  return new Promise((resolve) => requestAnimationFrame(() => resolve()));
}

/** Derive a display filename from a URL's path (or a sensible default). */
function fileNameFromUrl(url: string): string {
  try {
    const path = new URL(url, location.href).pathname;
    const last = decodeURIComponent(path.split('/').filter(Boolean).pop() ?? '');
    return last || 'document.docx';
  } catch {
    return 'document.docx';
  }
}

/** Why `bytes`/`name` can't be previewed as a .docx, or null if it can. A .docx
 *  is an OOXML ZIP, so the bytes must start with the ZIP local-file signature. */
function whyNotDocx(bytes: ArrayBuffer, name: string): string | null {
  const ext = name.toLowerCase().match(/\.[a-z0-9]+$/)?.[0];
  if (ext && ext !== '.docx') {
    return `Only .docx files can be previewed (this is ${ext}).`;
  }
  const b = new Uint8Array(bytes);
  const isZip = b[0] === 0x50 && b[1] === 0x4b && (b[2] === 0x03 || b[2] === 0x05 || b[2] === 0x07);
  if (!isZip) return 'This file is not a valid .docx (Word) document.';
  return null;
}

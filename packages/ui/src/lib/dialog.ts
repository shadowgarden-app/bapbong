import { injectStyle } from './internal.js';

export interface DialogOptions {
  title?: string;
  /** Modal (backdrop + focus trap) vs non-modal floating panel. Default false. */
  modal?: boolean;
  /** Extra class on the `<dialog>` element. */
  className?: string;
  /**
   * Pin a non-modal dialog's top-right corner just inside this rect (e.g. the
   * editor's canvas viewport) instead of the screen corner. Re-evaluated on
   * open + scroll/resize. Return null to fall back to the default position.
   */
  anchor?: () => DOMRect | null;
}

const STYLE = `
.bb-dialog{position:fixed;inset:auto;z-index:1100;margin:0;padding:0;border:1px solid var(--bb-ui-pop-border,var(--bb-ui-border,#e3e3e0));border-radius:10px;background:var(--bb-ui-dialog-bg,var(--bb-ui-menu-bg,#fff));-webkit-backdrop-filter:var(--bb-ui-dialog-filter,none);backdrop-filter:var(--bb-ui-dialog-filter,none);color:var(--bb-ui-fg,#2c2c2a);box-shadow:0 12px 40px rgba(0,0,0,.18);font-family:var(--bb-ui-font,system-ui,-apple-system,sans-serif);min-width:280px;max-width:min(92vw,440px)}
.bb-dialog *{box-sizing:border-box}
.bb-dialog::backdrop{background:rgba(0,0,0,.28)}
.bb-dialog-modal{top:50%;left:50%;transform:translate(-50%,-50%)}
.bb-dialog-float{top:62px;right:24px}
.bb-dialog-header{display:flex;align-items:center;justify-content:space-between;gap:12px;padding:9px 10px 9px 12px;border-bottom:1px solid var(--bb-ui-border,#e3e3e0)}
.bb-dialog-title{font-size:13px;font-weight:600}
.bb-dialog-close{border:0;background:transparent;color:inherit;opacity:.55;cursor:pointer;font-size:14px;width:24px;height:24px;border-radius:5px;line-height:1}
.bb-dialog-close:hover{opacity:1;background:var(--bb-ui-hover,#f1efe8)}
.bb-dialog-body{padding:12px}
.bb-prompt{display:flex;flex-direction:column;gap:10px;min-width:280px}
.bb-prompt-input{height:30px;padding:0 9px;border:1px solid var(--bb-ui-control-border,var(--bb-ui-border,#d8d6cf));border-radius:6px;font:inherit;font-size:13px}
.bb-prompt-actions{display:flex;justify-content:flex-end;gap:8px}
.bb-prompt-btn{height:30px;padding:0 14px;border:1px solid var(--bb-ui-control-border,var(--bb-ui-border,#d8d6cf));border-radius:6px;background:var(--bb-ui-control-bg,var(--bb-ui-bg,#fff));color:inherit;font:inherit;font-size:13px;cursor:pointer}
.bb-prompt-btn[data-primary]{background:var(--bb-ui-active-bg,#e6f1fb);border-color:var(--bb-ui-active-border,#b5d4f4);color:var(--bb-ui-active-fg,#0c447c)}
.bb-prompt-btn:hover{filter:brightness(.97)}
`;

/**
 * A generic, framework-agnostic dialog — a reusable overlay primitive built on
 * the native `<dialog>` element (so modal mode gets focus-trapping, `Esc` and a
 * backdrop for free). Drop any element into `body`; call `open()` / `close()`.
 * Used by the find panel, the prompt helper, the keyboard-shortcuts popup, and
 * (later) the cell-properties dialog — one primitive, not N bespoke widgets.
 */
export class Dialog {
  readonly el: HTMLDialogElement;
  /** Content slot — put your element(s) here. */
  readonly body: HTMLElement;
  private readonly title: HTMLElement;
  private readonly modal: boolean;
  private readonly anchor?: () => DOMRect | null;
  private readonly closeListeners = new Set<() => void>();
  private readonly reposition = (): void => {
    if (this.modal || !this.anchor) return;
    const r = this.anchor();
    if (!r) return;
    this.el.style.top = `${Math.max(8, r.top + 8)}px`;
    this.el.style.left = 'auto';
    this.el.style.right = `${Math.max(8, window.innerWidth - r.right + 8)}px`;
  };

  constructor(options: DialogOptions = {}) {
    injectStyle('bb-ui-dialog-styles', STYLE);
    this.modal = options.modal ?? false;
    this.anchor = options.anchor;

    const el = document.createElement('dialog');
    // Position via an explicit modifier class (avoids relying on the :modal
    // pseudo-class, which isn't matched by every engine).
    el.className =
      `bb-dialog ${this.modal ? 'bb-dialog-modal' : 'bb-dialog-float'}` +
      (options.className ? ` ${options.className}` : '');

    const header = document.createElement('div');
    header.className = 'bb-dialog-header';
    this.title = document.createElement('span');
    this.title.className = 'bb-dialog-title';
    this.title.textContent = options.title ?? '';
    const close = document.createElement('button');
    close.type = 'button';
    close.className = 'bb-dialog-close';
    close.textContent = '✕';
    close.setAttribute('aria-label', 'Close');
    close.addEventListener('click', () => this.close());
    header.append(this.title, close);

    this.body = document.createElement('div');
    this.body.className = 'bb-dialog-body';
    el.append(header, this.body);
    document.body.appendChild(el);

    el.addEventListener('close', () => {
      this.detachReposition();
      this.closeListeners.forEach((cb) => cb());
    });
    // Modal dialogs close on Esc natively; add it for non-modal panels too.
    if (!this.modal) {
      el.addEventListener('keydown', (e) => {
        if (e.key === 'Escape') this.close();
      });
    }

    this.el = el;
  }

  private attachReposition(): void {
    window.addEventListener('scroll', this.reposition, true); // capture: catch inner scrolls
    window.addEventListener('resize', this.reposition);
  }
  private detachReposition(): void {
    window.removeEventListener('scroll', this.reposition, true);
    window.removeEventListener('resize', this.reposition);
  }

  setTitle(text: string): void {
    this.title.textContent = text;
  }

  /** Replace the body content. */
  setContent(node: HTMLElement): void {
    this.body.replaceChildren(node);
  }

  open(): void {
    if (this.el.open) return;
    if (this.modal) this.el.showModal();
    else this.el.show();
    if (!this.modal && this.anchor) {
      this.reposition();
      this.attachReposition();
    }
  }

  close(): void {
    if (this.el.open) this.el.close();
  }

  get isOpen(): boolean {
    return this.el.open;
  }

  /** Subscribe to close (Esc, ✕, backdrop, or `close()`); returns unsubscribe. */
  onClose(cb: () => void): () => void {
    this.closeListeners.add(cb);
    return () => this.closeListeners.delete(cb);
  }

  destroy(): void {
    this.detachReposition();
    this.closeListeners.clear();
    this.el.remove();
  }
}

export interface PromptOptions {
  title: string;
  placeholder?: string;
  initial?: string;
  confirm?: string;
  cancel?: string;
}

/**
 * A one-line input modal built on {@link Dialog} — the vanilla replacement for
 * `window.prompt`. Resolves to the trimmed value, or null if cancelled/closed.
 */
export function promptDialog(options: PromptOptions): Promise<string | null> {
  return new Promise((resolve) => {
    const dialog = new Dialog({ title: options.title, modal: true });
    const form = document.createElement('form');
    form.className = 'bb-prompt';

    const input = document.createElement('input');
    input.type = 'text';
    input.className = 'bb-prompt-input';
    input.placeholder = options.placeholder ?? '';
    input.value = options.initial ?? '';

    const actions = document.createElement('div');
    actions.className = 'bb-prompt-actions';
    const cancel = document.createElement('button');
    cancel.type = 'button';
    cancel.className = 'bb-prompt-btn';
    cancel.textContent = options.cancel ?? 'Cancel';
    const ok = document.createElement('button');
    ok.type = 'submit';
    ok.className = 'bb-prompt-btn';
    ok.dataset['primary'] = 'true';
    ok.textContent = options.confirm ?? 'OK';
    actions.append(cancel, ok);
    form.append(input, actions);

    let settled = false;
    const done = (value: string | null) => {
      if (settled) return;
      settled = true;
      resolve(value);
      dialog.destroy();
    };
    form.addEventListener('submit', (e) => {
      e.preventDefault();
      done(input.value.trim() || null);
    });
    cancel.addEventListener('click', () => done(null));
    dialog.onClose(() => done(null));

    dialog.setContent(form);
    dialog.open();
    input.focus();
  });
}

import { injectStyle, placeFloating } from './internal.js';

/**
 * Floating link panel, anchored at the caret / selection (Google Docs-style).
 *
 * Three modes:
 * - **view** — the caret sits inside an existing link: the href plus three
 *   icon actions (copy / edit / remove). Edit switches to the form.
 * - **form with selection** — one URL field; Apply wraps the selection.
 * - **form at a bare caret** — an optional "Text" field above the URL; left
 *   empty, the URL itself becomes the inserted text.
 *
 * The panel is one-shot DOM (no framework), closes on Escape / outside click /
 * scroll, and positions with connected pairs: below the anchor first, flipped
 * above when the viewport runs out.
 */

export interface LinkPanelAnchor {
  /** Caret top-left in viewport CSS px + the caret's rendered height. */
  x: number;
  y: number;
  height: number;
}

export interface LinkPanelOptions {
  anchor: LinkPanelAnchor;
  /** Existing link under the caret → open in view mode (prefills the form). */
  existing?: { href: string; text: string } | null;
  /** The existing link points inside the document (`#bookmark`). Word never
   *  shows the raw anchor — it shows where the link GOES — so the panel
   *  displays `label` and offers a jump instead of a copyable URL.
   *  `generated` marks a link that is field output (a TOC entry): Word won't
   *  let you edit or unlink those one by one, and neither do we — updating
   *  the field is the way to change them. */
  internal?: {
    label: string | null;
    generated: boolean;
    onGo(): void;
  } | null;
  /** A text range is selected → the Text field is hidden (the selection is
   *  the text). */
  hasSelection?: boolean;
  /** Identity for keep-alive: with a key, an outside click defers the close
   *  briefly, and re-showing the SAME key within that window keeps the
   *  existing panel untouched — so a caret moving within one link never
   *  re-creates it. Hosts pass e.g. the link's doc range. */
  key?: string;
  /** Apply: `text` is present only when the Text field was shown. */
  onApply(href: string, text?: string): void;
  onUnlink?(): void;
}

export interface LinkPanelHandle {
  close(): void;
  /** Whether this panel is still on screen (hosts skip redundant reopens —
   *  e.g. the caret moving WITHIN the same link should not repaint). */
  isOpen(): boolean;
}

const STYLE = `
.bb-linkpanel{position:fixed;z-index:1200;width:300px;padding:8px;background:var(--bb-ui-menu-bg,#fff);-webkit-backdrop-filter:var(--bb-ui-pop-filter,none);backdrop-filter:var(--bb-ui-pop-filter,none);color:var(--bb-ui-fg,#2c2c2a);border:1px solid var(--bb-ui-pop-border,var(--bb-ui-border,#e3e3e0));border-radius:10px;box-shadow:0 8px 28px rgba(0,0,0,.16);font-family:var(--bb-ui-font,system-ui,-apple-system,sans-serif);box-sizing:border-box}
.bb-linkpanel *{box-sizing:border-box}
.bb-linkpanel-row{display:flex;align-items:center;gap:8px}
.bb-linkpanel-form{display:flex;flex-direction:column;gap:8px}
.bb-linkpanel-field{display:flex;align-items:center;gap:7px;border:1px solid var(--bb-ui-control-border,var(--bb-ui-border,#e3e3e0));border-radius:7px;padding:0 9px;background:var(--bb-ui-control-bg,var(--bb-ui-bg,#fff))}
.bb-linkpanel-field:focus-within{border-color:var(--bb-ui-accent,#d85a30)}
.bb-linkpanel-field svg{flex:none;opacity:.55}
.bb-linkpanel-input{flex:1;min-width:0;height:30px;border:0;background:transparent;color:inherit;font:inherit;font-size:13px;outline:none}
.bb-linkpanel-href{flex:1;min-width:0;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;font-size:13px;color:var(--bb-ui-accent,#d85a30)}
.bb-linkpanel-icon{display:inline-flex;flex:none;opacity:.55}
.bb-linkpanel-stale{opacity:.6;font-style:italic;color:inherit}
.bb-linkpanel-btn{display:inline-flex;align-items:center;justify-content:center;width:28px;height:28px;padding:0;border:0;border-radius:6px;background:transparent;color:inherit;cursor:pointer;flex:none}
.bb-linkpanel-btn:hover{background:var(--bb-ui-hover,#f1efe8)}
.bb-linkpanel-apply{font:inherit;font-size:13px;font-weight:600;color:var(--bb-ui-accent,#d85a30);background:transparent;border:0;border-radius:6px;padding:6px 8px;cursor:pointer;flex:none}
.bb-linkpanel-apply:hover{background:var(--bb-ui-hover,#f1efe8)}
.bb-linkpanel-apply:disabled{opacity:.4;cursor:default}
`;

const ICONS = {
  link: '<svg viewBox="0 0 16 16" width="15" height="15" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M6.5 9.5 9.5 6.5"/><path d="M7.5 4.5 9 3a2.47 2.47 0 0 1 3.5 0L13 3.5a2.47 2.47 0 0 1 0 3.5l-1.5 1.5"/><path d="M8.5 11.5 7 13a2.47 2.47 0 0 1-3.5 0L3 12.5a2.47 2.47 0 0 1 0-3.5l1.5-1.5"/></svg>',
  text: '<svg viewBox="0 0 16 16" width="15" height="15" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M3 4V2.5h10V4M8 2.5v11M6 13.5h4"/></svg>',
  copy: '<svg viewBox="0 0 16 16" width="15" height="15" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><rect x="5.5" y="5.5" width="8" height="8" rx="1.5"/><path d="M10.5 5.5v-2a1 1 0 0 0-1-1h-6a1 1 0 0 0-1 1v6a1 1 0 0 0 1 1h2"/></svg>',
  edit: '<svg viewBox="0 0 16 16" width="15" height="15" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="m9.5 3.5 3 3L6 13l-3.5.5L3 10z"/><path d="m11 2 .9-.9a1.55 1.55 0 0 1 2.2 0l.8.8a1.55 1.55 0 0 1 0 2.2L14 5"/></svg>',
  unlink:
    '<svg viewBox="0 0 16 16" width="15" height="15" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M7.5 4.5 9 3a2.47 2.47 0 0 1 3.5 0L13 3.5a2.47 2.47 0 0 1 0 3.5l-1.5 1.5"/><path d="M8.5 11.5 7 13a2.47 2.47 0 0 1-3.5 0L3 12.5a2.47 2.47 0 0 1 0-3.5l1.5-1.5"/><path d="m3 3 10 10"/></svg>',
  bookmark:
    '<svg viewBox="0 0 16 16" width="15" height="15" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M4 2.5h8v11l-4-3-4 3z"/></svg>',
  goTo: '<svg viewBox="0 0 16 16" width="15" height="15" fill="none" stroke="currentColor" stroke-width="1.4" stroke-linecap="round" stroke-linejoin="round" aria-hidden="true"><path d="M3 8h9"/><path d="m8.5 4.5 4 3.5-4 3.5"/></svg>',
};

let current: (() => void) | null = null;
let currentKey: string | null = null;
let currentHandle: LinkPanelHandle | null = null;
/** A deferred outside-click close (keep-alive window for keyed panels). */
let pendingClose: ReturnType<typeof setTimeout> | null = null;

function cancelPendingClose(): void {
  if (pendingClose !== null) {
    clearTimeout(pendingClose);
    pendingClose = null;
  }
}

function closeCurrent(): void {
  cancelPendingClose();
  if (current) {
    const dispose = current;
    current = null;
    currentKey = null;
    currentHandle = null;
    dispose();
  }
}

export function showLinkPanel(opts: LinkPanelOptions): LinkPanelHandle {
  injectStyle('bb-ui-linkpanel-styles', STYLE);
  // Keep-alive: the same keyed panel (still open, or in its deferred-close
  // window after an outside click) is reused untouched — no re-create, no
  // repaint when the caret merely moves within the same link.
  if (opts.key && current && currentKey === opts.key && currentHandle) {
    cancelPendingClose();
    return currentHandle;
  }
  closeCurrent();

  const el = document.createElement('div');
  el.className = 'bb-linkpanel';
  el.setAttribute('role', 'dialog');
  el.setAttribute('aria-label', 'Link');

  const iconBtn = (
    svg: string,
    label: string,
    onClick: () => void,
  ): HTMLButtonElement => {
    const b = document.createElement('button');
    b.type = 'button';
    b.className = 'bb-linkpanel-btn';
    b.title = label;
    b.setAttribute('aria-label', label);
    b.innerHTML = svg;
    b.addEventListener('mousedown', (e) => e.preventDefault());
    b.addEventListener('click', onClick);
    return b;
  };

  const field = (
    icon: string,
    placeholder: string,
    value: string,
  ): { wrap: HTMLDivElement; input: HTMLInputElement } => {
    const wrap = document.createElement('div');
    wrap.className = 'bb-linkpanel-field';
    wrap.innerHTML = icon;
    const input = document.createElement('input');
    input.className = 'bb-linkpanel-input';
    input.type = 'text';
    input.placeholder = placeholder;
    input.value = value;
    input.spellcheck = false;
    input.autocomplete = 'off';
    wrap.appendChild(input);
    return { wrap, input };
  };

  const renderView = (existing: { href: string; text: string }): void => {
    el.textContent = '';
    const row = document.createElement('div');
    row.className = 'bb-linkpanel-row';
    const internal = opts.internal;

    if (internal) {
      // In-document target: show WHERE it goes, never the bookmark id.
      const icon = document.createElement('span');
      icon.className = 'bb-linkpanel-icon';
      icon.innerHTML = ICONS.bookmark;
      row.appendChild(icon);
      const label = document.createElement('span');
      label.className = 'bb-linkpanel-href';
      label.textContent =
        internal.label ?? 'This place is no longer in the document';
      label.title = internal.label ?? existing.href;
      if (!internal.label) label.classList.add('bb-linkpanel-stale');
      row.appendChild(label);
      if (internal.label) {
        row.appendChild(
          iconBtn(ICONS.goTo, 'Go to this place', () => {
            closeCurrent();
            internal.onGo();
          }),
        );
      }
    } else {
      const href = document.createElement('span');
      href.className = 'bb-linkpanel-href';
      href.textContent = existing.href;
      href.title = existing.href;
      row.appendChild(href);
      row.appendChild(
        iconBtn(ICONS.copy, 'Copy link', () => {
          void navigator.clipboard
            ?.writeText(existing.href)
            .catch(() => undefined);
          closeCurrent();
        }),
      );
    }

    // Field output (a TOC entry) is regenerated wholesale — editing or
    // unlinking one entry would be undone by the next update, so Word
    // doesn't offer it here and neither do we.
    if (!internal?.generated) {
      row.appendChild(
        iconBtn(ICONS.edit, 'Edit link', () => renderForm(existing)),
      );
      row.appendChild(
        iconBtn(ICONS.unlink, 'Remove link', () => {
          closeCurrent();
          opts.onUnlink?.();
        }),
      );
    }
    el.appendChild(row);
    placeFloating(el, opts.anchor);
  };

  const renderForm = (
    prefill?: { href: string; text: string } | null,
  ): void => {
    el.textContent = '';
    const form = document.createElement('div');
    form.className = 'bb-linkpanel-form';
    // The Text field only exists when there is no selection: a selection IS
    // the text; editing an existing run always offers both fields.
    const withText = !opts.hasSelection;
    let textInput: HTMLInputElement | null = null;
    if (withText) {
      const t = field(ICONS.text, 'Text (optional)', prefill?.text ?? '');
      textInput = t.input;
      form.appendChild(t.wrap);
    }
    const u = field(ICONS.link, 'Paste or type a link', prefill?.href ?? '');
    const row = document.createElement('div');
    row.className = 'bb-linkpanel-row';
    row.appendChild(u.wrap);
    u.wrap.style.flex = '1';
    const apply = document.createElement('button');
    apply.type = 'button';
    apply.className = 'bb-linkpanel-apply';
    apply.textContent = 'Apply';
    apply.disabled = !u.input.value.trim();
    u.input.addEventListener('input', () => {
      apply.disabled = !u.input.value.trim();
    });
    const submit = (): void => {
      const href = u.input.value.trim();
      if (!href) return;
      const text = textInput?.value ?? undefined;
      closeCurrent();
      opts.onApply(href, text);
    };
    apply.addEventListener('click', submit);
    for (const input of [textInput, u.input]) {
      input?.addEventListener('keydown', (e) => {
        if (e.key === 'Enter') {
          e.preventDefault();
          submit();
        }
      });
    }
    row.appendChild(apply);
    form.appendChild(row);
    el.appendChild(form);
    placeFloating(el, opts.anchor);
    u.input.focus();
    u.input.select();
  };

  document.body.appendChild(el);
  if (opts.existing) renderView(opts.existing);
  else renderForm(null);

  const onDown = (e: Event): void => {
    if (el.contains(e.target as Node)) return;
    // Keyed panels defer the close one beat: if the click lands back in the
    // same link, the host re-shows the same key and cancels this.
    if (opts.key) {
      cancelPendingClose();
      pendingClose = setTimeout(closeCurrent, 80);
    } else {
      closeCurrent();
    }
  };
  const onKey = (e: KeyboardEvent): void => {
    if (e.key === 'Escape') closeCurrent();
  };
  const onScroll = (): void => closeCurrent();
  // Defer so the opening click doesn't immediately dismiss the panel.
  setTimeout(() => {
    document.addEventListener('pointerdown', onDown, true);
    document.addEventListener('keydown', onKey, true);
    window.addEventListener('scroll', onScroll, true);
  }, 0);

  const dispose = (): void => {
    document.removeEventListener('pointerdown', onDown, true);
    document.removeEventListener('keydown', onKey, true);
    window.removeEventListener('scroll', onScroll, true);
    el.remove();
  };
  current = dispose;
  currentKey = opts.key ?? null;
  const handle: LinkPanelHandle = {
    close: () => {
      // Only close THIS panel — a later panel may have replaced it.
      if (current === dispose) closeCurrent();
    },
    isOpen: () => current === dispose,
  };
  currentHandle = handle;
  return handle;
}

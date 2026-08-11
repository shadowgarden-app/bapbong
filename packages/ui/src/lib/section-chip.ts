import { injectStyle } from './internal.js';

/**
 * The section-break marker for a next-page break: a segmented chip that lives
 * in the GAP between two pages (the break *is* that boundary), always showing
 * its quick actions — no hover reveal, so touch devices get the same surface.
 *
 * Segments: a static title ("Next page"), the following section's page
 * numbering, its paper/orientation, and a delete button. The two action
 * segments hand their anchor rect back to the host, which opens the shared
 * menu components (`showMenu` + the page-setup pickers) — the chip itself
 * knows no commands and no document model.
 *
 * Styled entirely from `--bb-ui-*` tokens, so each shell (playground,
 * desktop-ui glass theme, dark mode) skins it like the rest of the UI kit.
 */

const STYLE = `
.bb-secchip{position:absolute;z-index:8;display:inline-flex;align-items:center;transform:translate(-50%,-50%);white-space:nowrap;background:var(--bb-ui-menu-bg,#fff);-webkit-backdrop-filter:var(--bb-ui-pop-filter,none);backdrop-filter:var(--bb-ui-pop-filter,none);border:1px solid var(--bb-ui-pop-border,var(--bb-ui-border,#e3e3e0));border-radius:10px;padding:1px 2px;box-shadow:0 2px 10px rgba(0,0,0,.12);font-family:var(--bb-ui-font,system-ui,-apple-system,sans-serif);font-size:11px;color:var(--bb-ui-fg,#2c2c2a)}
.bb-secchip *{box-sizing:border-box}
.bb-secchip-title{padding:0 7px;font-size:11px;font-weight:600;opacity:.8}
.bb-secchip-seg{display:inline-flex;align-items:center;gap:4px;height:16px;padding:0 7px;border:0;border-radius:6px;background:transparent;color:inherit;font:inherit;font-size:11px;cursor:pointer}
.bb-secchip-seg:hover,.bb-secchip-seg[aria-expanded="true"]{background:var(--bb-ui-hover,#f1efe8)}
.bb-secchip-caret{font-size:8px;opacity:.55}
.bb-secchip-seg[data-muted="true"] .bb-secchip-seglabel{font-style:italic;opacity:.55}
.bb-secchip-div{flex:none;width:1px;height:11px;margin:0 2px;background:var(--bb-ui-border,#e3e3e0)}
.bb-secchip-x{display:inline-flex;align-items:center;justify-content:center;width:20px;height:16px;padding:0;border:0;border-radius:6px;background:transparent;color:inherit;font:inherit;opacity:.65;cursor:pointer}
.bb-secchip-x:hover{background:var(--bb-ui-hover,#f1efe8);opacity:1}
`;

export interface SectionChipOptions {
  /** Static first segment, e.g. "Next page" / "Trang mới". */
  title: string;
  /** Accessible names for the action segments. */
  ariaPageNumbers?: string;
  ariaPaper?: string;
  ariaDelete?: string;
  /** Open the page-numbering menu; `anchor` is the clicked segment. */
  onPageNumbers(anchor: HTMLElement): void;
  /** Open the orientation / paper-size menu; `anchor` is the clicked segment. */
  onPaper(anchor: HTMLElement): void;
  onDelete(): void;
}

export interface SectionChipHandle {
  /** Position with `left`/`top` (the chip centers itself on that point). */
  el: HTMLElement;
  /** Refresh the two data segments' labels. `pageNumbersMuted` styles the
   *  numbering segment as inactive ("no page number") — still clickable. */
  update(data: {
    pageNumbers: string;
    pageNumbersMuted?: boolean;
    paper: string;
  }): void;
  destroy(): void;
}

export function createSectionChip(
  options: SectionChipOptions,
): SectionChipHandle {
  injectStyle('bb-ui-section-chip-styles', STYLE);

  const el = document.createElement('div');
  el.className = 'bb-secchip';

  const title = document.createElement('span');
  title.className = 'bb-secchip-title';
  title.textContent = options.title;

  const divider = (): HTMLElement => {
    const d = document.createElement('span');
    d.className = 'bb-secchip-div';
    return d;
  };

  const segment = (
    aria: string | undefined,
    onOpen: (anchor: HTMLElement) => void,
  ): { btn: HTMLButtonElement; label: HTMLSpanElement } => {
    const btn = document.createElement('button');
    btn.type = 'button';
    btn.className = 'bb-secchip-seg';
    btn.setAttribute('aria-haspopup', 'true');
    if (aria) btn.setAttribute('aria-label', aria);
    const label = document.createElement('span');
    label.className = 'bb-secchip-seglabel';
    const caret = document.createElement('span');
    caret.className = 'bb-secchip-caret';
    caret.textContent = '▾';
    btn.append(label, caret);
    btn.addEventListener('mousedown', (e) => e.preventDefault()); // keep the editor selection
    btn.addEventListener('click', () => onOpen(btn));
    return { btn, label };
  };

  const nums = segment(options.ariaPageNumbers, options.onPageNumbers);
  const paper = segment(options.ariaPaper, options.onPaper);

  const x = document.createElement('button');
  x.type = 'button';
  x.className = 'bb-secchip-x';
  x.setAttribute('aria-label', options.ariaDelete ?? 'Remove section break');
  x.innerHTML =
    '<svg viewBox="0 0 16 16" width="9" height="9" fill="none" stroke="currentColor" stroke-width="2.2" stroke-linecap="round"><path d="M3 3l10 10M13 3 3 13"/></svg>';
  x.addEventListener('mousedown', (e) => e.preventDefault());
  x.addEventListener('click', () => options.onDelete());

  el.append(title, divider(), nums.btn, divider(), paper.btn, divider(), x);

  return {
    el,
    update(data) {
      nums.label.textContent = data.pageNumbers;
      nums.btn.dataset['muted'] = String(!!data.pageNumbersMuted);
      paper.label.textContent = data.paper;
    },
    destroy() {
      el.remove();
    },
  };
}

import type { CharacterFormatting } from '@shadow-garden/bapbong-contracts';
import { Dialog } from './dialog.js';
import { injectStyle } from './internal.js';

/**
 * Format ▸ Font — the home for character formatting, in two tabs the way Word
 * splits it.
 *
 * It exists because five marks the importer has always produced — smallCaps,
 * dstrike, position, letterSpacing, charScale — had no way in at all: a
 * document carrying them rendered correctly and could not be edited. Neither
 * Word nor Google Docs puts these on a toolbar (Docs omits the last two
 * entirely), so a dialog behind Format ▸ Font is where people look for them.
 *
 * Presentation only, like the other widgets here: it takes the selection's
 * current values and hands back what the user chose, so this package keeps its
 * single contracts dependency and any shell can mount it.
 *
 * Units are the ones Word's dialog shows — points and percent — while the
 * model keeps the document's own twips and half-points. The conversion lives
 * here, at the only boundary that speaks points, exactly as page-setup.ts
 * converts px to centimetres.
 */

/** Twips (1/20 pt) ↔ points, and half-points ↔ points. */
const twipsToPt = (tw: number) => tw / 20;
const ptToTwips = (pt: number) => Math.round(pt * 20);
const halfToPt = (hp: number) => hp / 2;
const ptToHalf = (pt: number) => Math.round(pt * 2);

export interface FontDialogOptions {
  initial: CharacterFormatting;
  /** Families offered in the picker; the current one is added if missing. */
  families: readonly string[];
  /** Sizes offered in the picker (points). */
  sizes: readonly number[];
  onApply: (values: CharacterFormatting) => void;
}

const SCALE_PRESETS = [200, 150, 100, 90, 80, 66, 50, 33];

const TEXT_COLORS = [
  { label: 'Automatic', value: null },
  { label: 'Black', value: '#000000' },
  { label: 'Grey', value: '#5F5E5A' },
  { label: 'Red', value: '#C0392B' },
  { label: 'Blue', value: '#185FA5' },
  { label: 'Green', value: '#1D9E75' },
  { label: 'Orange', value: '#BA7517' },
];

const HIGHLIGHTS = [
  { label: 'None', value: null },
  { label: 'Yellow', value: '#FFFF00' },
  { label: 'Green', value: '#00FF00' },
  { label: 'Cyan', value: '#00FFFF' },
  { label: 'Magenta', value: '#FF00FF' },
  { label: 'Grey', value: '#C0C0C0' },
];

const STYLE = `
.bb-fd{display:flex;flex-direction:column;min-width:396px;max-width:430px;color:var(--bb-ui-fg,#2c2c2a)}
.bb-fd *{box-sizing:border-box}
/* A segmented control, not a "tab joined to its pane": that pattern needs the
   active tab's background to equal the pane's, and this dialog floats on
   frosted glass whose colour the widget cannot know. Same shape cell-properties
   already uses, so the two dialogs read as one product. */
.bb-fd-tabs{display:inline-flex;border:0.5px solid var(--bb-ui-border,#d8d6cf);border-radius:6px;overflow:hidden;margin-bottom:15px;align-self:flex-start}
.bb-fd-tab{height:30px;padding:0 18px;border:0;border-right:0.5px solid var(--bb-ui-border,#e3e3e0);background:transparent;color:inherit;opacity:.65;font:inherit;font-size:13px;cursor:pointer}
.bb-fd-tab:last-child{border-right:0}
.bb-fd-tab[aria-selected="true"]{background:var(--bb-ui-active-bg,#e6f1fb);color:var(--bb-ui-active-fg,#0c447c);opacity:1}
.bb-fd-pane{display:flex;flex-direction:column;gap:13px}
/* A class-level display beats the UA's [hidden]{display:none}; without this
   both tabs render at once. */
.bb-fd-pane[hidden]{display:none}
.bb-fd-lbl{font-size:12px;opacity:.7;margin-bottom:5px}
.bb-fd-sec{font-size:12px;opacity:.7;margin-bottom:9px}
.bb-fd-grid{display:grid;gap:10px}
.bb-fd-ctl{width:100%;height:32px;border:0.5px solid var(--bb-ui-border,#d8d6cf);border-radius:6px;background:var(--bb-ui-bg,#fff);color:inherit;font:inherit;font-size:13px;padding:0 8px}
.bb-fd-row{display:grid;grid-template-columns:96px 1fr 88px;gap:10px;align-items:center;margin-bottom:10px}
.bb-fd-row label{font-size:13px}
.bb-fd-state{font-size:12px;color:var(--bb-ui-active-fg,#0c447c)}
.bb-fd-state.off{color:inherit;opacity:.5}
.bb-fd-seg{display:inline-flex;border:0.5px solid var(--bb-ui-border,#d8d6cf);border-radius:6px;overflow:hidden;height:32px}
.bb-fd-segbtn{width:40px;border:0;border-right:0.5px solid var(--bb-ui-border,#e3e3e0);background:transparent;color:inherit;font:inherit;font-size:13px;cursor:pointer}
.bb-fd-segbtn:last-child{border-right:0}
.bb-fd-segbtn[aria-pressed="true"]{background:var(--bb-ui-active-bg,#e6f1fb);color:var(--bb-ui-active-fg,#0c447c)}
.bb-fd-segbtn[aria-pressed="mixed"]{background:repeating-linear-gradient(135deg,transparent 0 4px,var(--bb-ui-active-bg,#e6f1fb) 4px 8px)}
.bb-fd-eff{display:grid;grid-template-columns:1fr 1fr;gap:8px 14px}
.bb-fd-cb{display:flex;align-items:center;gap:7px;font-size:13px;cursor:pointer}
.bb-fd-prev{border:0.5px solid var(--bb-ui-border,#d8d6cf);border-radius:8px;background:var(--bb-ui-bg,#fff);padding:15px 10px;text-align:center;overflow:hidden}
.bb-fd-prev-cap{font-size:11px;opacity:.55;margin-bottom:9px}
.bb-fd-prev-txt{font-size:17px;display:inline-block}
.bb-fd-dim{opacity:.4}
.bb-fd-note{font-size:11px;opacity:.55;margin-top:5px}
.bb-fd-foot{display:flex;justify-content:flex-end;gap:8px;margin-top:16px}
.bb-fd-btn{height:31px;padding:0 17px;border:0.5px solid var(--bb-ui-border,#d8d6cf);border-radius:6px;background:var(--bb-ui-bg,#fff);color:inherit;font:inherit;font-size:13px;cursor:pointer}
.bb-fd-btn.primary{background:var(--bb-ui-active-bg,#e6f1fb);border-color:var(--bb-ui-active-border,#7fb2ec);color:var(--bb-ui-active-fg,#0c447c)}
`;

function el<K extends keyof HTMLElementTagNameMap>(
  tag: K,
  cls?: string,
  text?: string,
): HTMLElementTagNameMap[K] {
  const n = document.createElement(tag);
  if (cls) n.className = cls;
  if (text) n.textContent = text;
  return n;
}

/** A labelled <select>; `options` are [value, label] and '' means "unset". */
function select(
  options: readonly (readonly [string, string])[],
  value: string,
): HTMLSelectElement {
  const s = el('select', 'bb-fd-ctl');
  for (const [v, label] of options) {
    const o = document.createElement('option');
    o.value = v;
    o.textContent = label;
    s.append(o);
  }
  s.value = value;
  return s;
}

/** `undefined` renders indeterminate: the selection is mixed, and unless the
 *  user clicks it the property is left alone rather than flattened. */
function checkbox(label: string, on: boolean | undefined): HTMLInputElement {
  const wrap = el('label', 'bb-fd-cb');
  const input = document.createElement('input');
  input.type = 'checkbox';
  input.checked = on === true;
  input.indeterminate = on === undefined;
  // Clicking resolves the mixed state; only then does it get written.
  input.addEventListener('change', () => {
    input.indeterminate = false;
  });
  wrap.append(input, document.createTextNode(label));
  // The caller appends the WRAPPER; hand back the input it needs to read.
  (input as HTMLInputElement & { wrap?: HTMLElement }).wrap = wrap;
  return input;
}

const wrapOf = (i: HTMLInputElement) =>
  (i as HTMLInputElement & { wrap?: HTMLElement }).wrap ?? i;

/**
 * Open the Font dialog. Resolves nothing — the caller applies through
 * `onApply`, which fires once, on OK.
 */
export function openFontDialog({
  initial,
  families,
  sizes,
  onApply,
}: FontDialogOptions): void {
  injectStyle('bb-font-dialog', STYLE);
  const dialog = new Dialog({ title: 'Font', modal: true });
  const root = el('div', 'bb-fd');

  // ── tabs ────────────────────────────────────────────────────────
  const tabs = el('div', 'bb-fd-tabs');
  tabs.setAttribute('role', 'tablist');
  const fontTab = el('button', 'bb-fd-tab', 'Font');
  const advTab = el('button', 'bb-fd-tab', 'Advanced');
  for (const t of [fontTab, advTab]) {
    t.type = 'button';
    t.setAttribute('role', 'tab');
    tabs.append(t);
  }
  const fontPane = el('div', 'bb-fd-pane');
  const advPane = el('div', 'bb-fd-pane');
  const showTab = (advanced: boolean) => {
    fontTab.setAttribute('aria-selected', String(!advanced));
    advTab.setAttribute('aria-selected', String(advanced));
    fontPane.hidden = advanced;
    advPane.hidden = !advanced;
  };
  fontTab.addEventListener('click', () => showTab(false));
  advTab.addEventListener('click', () => showTab(true));

  // ── Font tab ────────────────────────────────────────────────────
  const famList = families.includes(initial.family ?? '')
    ? families
    : initial.family
      ? [initial.family, ...families]
      : families;
  // '(unchanged)', not '(mixed)': a null here means "no explicit mark OR a
  // mixed selection", and the two are indistinguishable at this layer. What
  // IS certain is what OK will do with it — leave the property alone — so the
  // label promises that instead of guessing which case it is.
  const familySel = select(
    [['', '(unchanged)'], ...famList.map((f) => [f, f] as const)],
    initial.family ?? '',
  );
  const sizeSel = select(
    [['', '(unchanged)'], ...sizes.map((n) => [String(n), String(n)] as const)],
    initial.sizePt == null ? '' : String(initial.sizePt),
  );

  const topGrid = el('div', 'bb-fd-grid');
  topGrid.style.gridTemplateColumns = '1fr 96px';
  const famCell = el('div');
  famCell.append(el('div', 'bb-fd-lbl', 'Font'), familySel);
  const sizeCell = el('div');
  sizeCell.append(el('div', 'bb-fd-lbl', 'Size'), sizeSel);
  topGrid.append(famCell, sizeCell);

  const seg = el('div', 'bb-fd-seg');
  // 'mixed' is a real aria-pressed value, and it survives until clicked —
  // same contract as the indeterminate checkboxes.
  const mkSeg = (text: string, on: boolean | undefined, style?: string) => {
    const b = el('button', 'bb-fd-segbtn', text);
    b.type = 'button';
    b.setAttribute('aria-pressed', on === undefined ? 'mixed' : String(on));
    if (style) b.setAttribute('style', style);
    b.addEventListener('click', () => {
      b.setAttribute(
        'aria-pressed',
        b.getAttribute('aria-pressed') === 'true' ? 'false' : 'true',
      );
      paint();
    });
    seg.append(b);
    return b;
  };
  const boldBtn = mkSeg('B', initial.bold, 'font-weight:600');
  const italicBtn = mkSeg('I', initial.italic, 'font-style:italic');
  const underBtn = mkSeg('U', initial.underline, 'text-decoration:underline');

  const colorSel = select(
    TEXT_COLORS.map((c) => [c.value ?? '', c.label] as const),
    initial.color ?? '',
  );
  const hlSel = select(
    HIGHLIGHTS.map((c) => [c.value ?? '', c.label] as const),
    initial.highlight ?? '',
  );

  const styleGrid = el('div', 'bb-fd-grid');
  styleGrid.style.gridTemplateColumns = 'auto 1fr 1fr';
  const segCell = el('div');
  segCell.append(el('div', 'bb-fd-lbl', 'Style'), seg);
  const colorCell = el('div');
  colorCell.append(el('div', 'bb-fd-lbl', 'Text color'), colorSel);
  const hlCell = el('div');
  hlCell.append(el('div', 'bb-fd-lbl', 'Highlight'), hlSel);
  styleGrid.append(segCell, colorCell, hlCell);

  const effWrap = el('div');
  effWrap.append(el('div', 'bb-fd-sec', 'Effects'));
  const effGrid = el('div', 'bb-fd-eff');
  const strikeCb = checkbox('Strikethrough', initial.strike);
  const dstrikeCb = checkbox('Double strikethrough', initial.doubleStrike);
  const smallCapsCb = checkbox('Small caps', initial.smallCaps);
  const supCb = checkbox('Superscript', initial.vertAlign === 'super');
  const subCb = checkbox('Subscript', initial.vertAlign === 'sub');
  // Word treats each pair as mutually exclusive and so does the UI. The
  // IMPORTER stays permissive: a file that carries both is read as it is,
  // because refusing to represent someone else's document is worse than a
  // state our own UI will not author.
  const exclusive = (a: HTMLInputElement, b: HTMLInputElement) => {
    a.addEventListener('change', () => {
      if (a.checked) b.checked = false;
      paint();
    });
  };
  exclusive(strikeCb, dstrikeCb);
  exclusive(dstrikeCb, strikeCb);
  exclusive(supCb, subCb);
  exclusive(subCb, supCb);
  smallCapsCb.addEventListener('change', () => paint());
  // Laid out so each mutually-exclusive pair sits on one row — the geometry
  // says "pick one of these two" before any of the behaviour does.
  effGrid.append(
    wrapOf(strikeCb),
    wrapOf(dstrikeCb),
    wrapOf(supCb),
    wrapOf(subCb),
    wrapOf(smallCapsCb),
  );
  effWrap.append(effGrid);
  fontPane.append(topGrid, styleGrid, effWrap);

  // ── Advanced tab ────────────────────────────────────────────────
  const advWrap = el('div');
  advWrap.append(el('div', 'bb-fd-sec', 'Character spacing'));

  const mkRow = (
    label: string,
    control: HTMLElement,
    stateText: string,
    dim = false,
  ) => {
    const row = el('div', 'bb-fd-row');
    if (dim) row.classList.add('bb-fd-dim');
    const l = el('label', undefined, label);
    const st = el('span', 'bb-fd-state off', stateText);
    row.append(l, control, st);
    advWrap.append(row);
    return st;
  };

  const scaleSel = select(
    SCALE_PRESETS.map((n) => [String(n), `${n}%`] as const),
    String(initial.scalePercent),
  );
  const scaleState = mkRow('Scale', scaleSel, '');

  const spacingInput = el('input', 'bb-fd-ctl') as HTMLInputElement;
  spacingInput.type = 'number';
  spacingInput.step = '0.05';
  spacingInput.value = String(twipsToPt(initial.letterSpacingTwips));
  const spacingState = mkRow('Spacing', spacingInput, '');

  const positionInput = el('input', 'bb-fd-ctl') as HTMLInputElement;
  positionInput.type = 'number';
  positionInput.step = '0.5';
  positionInput.value = String(halfToPt(initial.positionHalfPoints));
  const positionState = mkRow('Position', positionInput, '');

  const kernInput = el('input', 'bb-fd-ctl') as HTMLInputElement;
  kernInput.type = 'number';
  kernInput.value = '12';
  kernInput.disabled = true;
  mkRow('Kerning', kernInput, 'not supported', true);
  const kernNote = el(
    'div',
    'bb-fd-note',
    'Kerning is read from the file but not reproduced when drawing, so it is shown here rather than silently dropped.',
  );
  advWrap.append(kernNote);
  advPane.append(advWrap);

  // ── shared preview ──────────────────────────────────────────────
  // One sample for both tabs: the settings compose, so showing them apart
  // would misrepresent what the run will look like.
  const prev = el('div', 'bb-fd-prev');
  const prevTxt = el('span', 'bb-fd-prev-txt', 'The quick brown fox');
  prev.append(el('div', 'bb-fd-prev-cap', 'Preview'), prevTxt);

  const num = (i: HTMLInputElement) => {
    const v = Number(i.value);
    return Number.isFinite(v) ? v : 0;
  };
  const pressed = (b: HTMLButtonElement) =>
    b.getAttribute('aria-pressed') === 'true';
  /** undefined while still mixed — see mkSeg. */
  const pressedOrMixed = (b: HTMLButtonElement) =>
    b.getAttribute('aria-pressed') === 'mixed' ? undefined : pressed(b);
  const checkedOrMixed = (i: HTMLInputElement) =>
    i.indeterminate ? undefined : i.checked;

  function paint(): void {
    const spacingPt = num(spacingInput);
    const positionPt = num(positionInput);
    const scale = Number(scaleSel.value) || 100;

    scaleState.textContent = scale === 100 ? 'normal' : `${scale}%`;
    scaleState.classList.toggle('off', scale === 100);
    spacingState.textContent =
      spacingPt === 0 ? 'normal' : spacingPt > 0 ? 'expanded' : 'condensed';
    spacingState.classList.toggle('off', spacingPt === 0);
    positionState.textContent =
      positionPt === 0 ? 'on baseline' : positionPt > 0 ? 'raised' : 'lowered';
    positionState.classList.toggle('off', positionPt === 0);

    const deco = [
      pressed(underBtn) ? 'underline' : '',
      strikeCb.checked || dstrikeCb.checked ? 'line-through' : '',
    ]
      .filter(Boolean)
      .join(' ');
    const s = prevTxt.style;
    s.fontFamily = familySel.value || 'inherit';
    s.fontSize = sizeSel.value ? `${sizeSel.value}pt` : '17px';
    s.fontWeight = pressed(boldBtn) ? '700' : '400';
    s.fontStyle = pressed(italicBtn) ? 'italic' : 'normal';
    s.textDecoration = deco || 'none';
    s.textDecorationStyle = dstrikeCb.checked ? 'double' : 'solid';
    s.fontVariant = smallCapsCb.checked ? 'small-caps' : 'normal';
    s.color = colorSel.value || 'inherit';
    s.background = hlSel.value || 'transparent';
    s.letterSpacing = `${spacingPt}pt`;
    s.transform = scale === 100 ? 'none' : `scaleX(${scale / 100})`;
    s.verticalAlign = supCb.checked
      ? 'super'
      : subCb.checked
        ? 'sub'
        : `${positionPt}pt`;
  }
  for (const c of [familySel, sizeSel, colorSel, hlSel, scaleSel])
    c.addEventListener('change', paint);
  for (const i of [spacingInput, positionInput])
    i.addEventListener('input', paint);

  // ── footer ──────────────────────────────────────────────────────
  const foot = el('div', 'bb-fd-foot');
  const cancel = el('button', 'bb-fd-btn', 'Cancel');
  const ok = el('button', 'bb-fd-btn primary', 'OK');
  cancel.type = ok.type = 'button';
  cancel.addEventListener('click', () => dialog.close());
  ok.addEventListener('click', () => {
    const spacingPt = num(spacingInput);
    const positionPt = num(positionInput);
    const scale = Number(scaleSel.value) || 100;
    // '' on a select means the selection was mixed and stayed that way:
    // report undefined so the command layer leaves that property alone.
    onApply({
      family: familySel.value || undefined,
      sizePt: sizeSel.value ? Number(sizeSel.value) : undefined,
      bold: pressedOrMixed(boldBtn),
      italic: pressedOrMixed(italicBtn),
      underline: pressedOrMixed(underBtn),
      strike: checkedOrMixed(strikeCb),
      doubleStrike: checkedOrMixed(dstrikeCb),
      smallCaps: checkedOrMixed(smallCapsCb),
      vertAlign: supCb.checked ? 'super' : subCb.checked ? 'sub' : null,
      color: colorSel.value || null,
      highlight: hlSel.value || null,
      scalePercent: scale,
      letterSpacingTwips: ptToTwips(spacingPt),
      positionHalfPoints: ptToHalf(positionPt),
    });
    dialog.close();
  });
  foot.append(cancel, ok);

  root.append(tabs, fontPane, advPane, prev, foot);
  dialog.body.append(root);
  showTab(false);
  paint();
  dialog.open();
}

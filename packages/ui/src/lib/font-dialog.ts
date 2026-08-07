import type { CharacterFormatting } from '@shadow-garden/bapbong-contracts';
import { Dialog } from './dialog.js';
import { injectStyle, swatchRow } from './internal.js';

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

/** Must track the layout engine's own constant: the preview is only useful
 *  while it agrees with what gets drawn. */
const SUPERSUB_SCALE = 0.66;

const TEXT_COLORS = [
  '#000000',
  '#5F5E5A',
  '#B0B0B0',
  '#C0392B',
  '#BA7517',
  '#1D9E75',
  '#185FA5',
  '#7F77DD',
];

/** Word's own highlighter set, which is what a .docx round-trips exactly
 *  (w:highlight is a named enum; anything else has to travel as w:shd). */
const HIGHLIGHTS = [
  '#FFFF00',
  '#00FF00',
  '#00FFFF',
  '#FF00FF',
  '#0000FF',
  '#FF0000',
  '#C0C0C0',
];

const STYLE = `
.bb-fd{display:flex;flex-direction:column;gap:15px;min-width:396px;max-width:430px;color:var(--bb-ui-fg,#2c2c2a)}
.bb-fd *{box-sizing:border-box}
/* A segmented control, not a "tab joined to its pane": that pattern needs the
   active tab's background to equal the pane's, and this dialog floats on
   frosted glass whose colour the widget cannot know. Same shape cell-properties
   already uses, so the two dialogs read as one product. */
.bb-fd-tabs{display:inline-flex;border:0.5px solid var(--bb-ui-border,#d8d6cf);border-radius:6px;overflow:hidden;align-self:flex-start}
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
.bb-fd-swatches{display:flex;flex-wrap:wrap;gap:6px}
.bb-fd-swatch{width:26px;height:26px;border-radius:6px;border:0.5px solid var(--bb-ui-border,#d8d6cf);padding:0;cursor:pointer}
.bb-fd-swatch.on{box-shadow:0 0 0 2px var(--bb-ui-active-fg,#0c447c)}
.bb-fd-swatch-none{display:flex;align-items:center;justify-content:center;background:var(--bb-ui-bg,#fff);font-size:15px;color:inherit;opacity:.55}
/* The OS picker is a chip like the rest: strip the native chrome so it does
   not read as a different KIND of control from the presets beside it. */
.bb-fd-swatch-custom{background:conic-gradient(#c0392b,#ba7517,#1d9e75,#185fa5,#7f77dd,#c0392b)}
.bb-fd-swatch-custom::-webkit-color-swatch-wrapper{padding:0}
.bb-fd-swatch-custom::-webkit-color-swatch{border:0;border-radius:5px;opacity:0}
.bb-fd-eff{display:grid;grid-template-columns:1fr 1fr;gap:8px 14px}
.bb-fd-cb{display:flex;align-items:center;gap:7px;font-size:13px;cursor:pointer}
.bb-fd-prev{border:0.5px solid var(--bb-ui-border,#d8d6cf);border-radius:8px;background:var(--bb-ui-bg,#fff);padding:13px 14px;overflow:hidden}
.bb-fd-prev-cap{font-size:11px;opacity:.55;margin-bottom:9px;text-align:center}
/* The surrounding sentence stays at a fixed size and weight — it is the ruler
   the sample is read against, so it must not move when the sample does. */
.bb-fd-prev-line{font-size:15px;line-height:2;text-align:center}
.bb-fd-prev-txt{display:inline-block}
.bb-fd-dim{opacity:.4}
.bb-fd-note{font-size:11px;opacity:.55;margin-top:5px}
.bb-fd-foot{display:flex;justify-content:flex-end;gap:8px}
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
  // Wide enough for the '(unchanged)' option, which is longer than any size.
  topGrid.style.gridTemplateColumns = '1fr 118px';
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

  // Colours get swatches and the platform picker, not a dropdown of names:
  // a colour is chosen by looking at it.
  let color: string | null = initial.color ?? null;
  let highlight: string | null = initial.highlight ?? null;
  const colorRow = swatchRow({
    prefix: 'bb-fd',
    presets: TEXT_COLORS,
    clearLabel: 'Automatic',
    get: () => color,
    set: (c) => {
      color = c;
      paint();
    },
  });
  const hlRow = swatchRow({
    prefix: 'bb-fd',
    presets: HIGHLIGHTS,
    clearLabel: 'No highlight',
    get: () => highlight,
    set: (c) => {
      highlight = c;
      paint();
    },
  });

  const styleCell = el('div');
  styleCell.append(el('div', 'bb-fd-lbl', 'Style'), seg);
  const colorCell = el('div');
  colorCell.append(el('div', 'bb-fd-lbl', 'Text color'), colorRow.el);
  const hlCell = el('div');
  hlCell.append(el('div', 'bb-fd-lbl', 'Highlight'), hlRow.el);

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
  fontPane.append(topGrid, styleCell, colorCell, hlCell, effWrap);

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
  // The sample sits INSIDE a plain sentence, sharing its baseline. On its own
  // a centred line cannot show what superscript, a baseline shift, tracking or
  // a glyph scale actually do — each of those is only legible against
  // unstyled text beside it.
  const prevLine = el('div', 'bb-fd-prev-line');
  const prevTxt = el('span', 'bb-fd-prev-txt', 'The quick brown fox');
  prevLine.append(
    document.createTextNode('Text before '),
    prevTxt,
    document.createTextNode(' and text after.'),
  );
  prev.append(el('div', 'bb-fd-prev-cap', 'Preview'), prevLine);

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
    // Super/subscript render at a REDUCED size, so the preview has to reduce
    // too — the same factor the layout engine uses. Showing them at full size
    // would make the preview disagree with the page it is previewing.
    const shrink = supCb.checked || subCb.checked ? SUPERSUB_SCALE : 1;
    s.fontSize = sizeSel.value
      ? `${Number(sizeSel.value) * shrink}pt`
      : `${15 * shrink}px`;
    s.fontWeight = pressed(boldBtn) ? '700' : '400';
    s.fontStyle = pressed(italicBtn) ? 'italic' : 'normal';
    s.textDecoration = deco || 'none';
    s.textDecorationStyle = dstrikeCb.checked ? 'double' : 'solid';
    s.fontVariant = smallCapsCb.checked ? 'small-caps' : 'normal';
    s.color = color ?? 'inherit';
    s.background = highlight ?? 'transparent';
    s.letterSpacing = `${spacingPt}pt`;
    s.transform = scale === 100 ? 'none' : `scaleX(${scale / 100})`;
    s.verticalAlign = supCb.checked
      ? 'super'
      : subCb.checked
        ? 'sub'
        : `${positionPt}pt`;
  }
  for (const c of [familySel, sizeSel, scaleSel])
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
      color,
      highlight,
      scalePercent: scale,
      letterSpacingTwips: ptToTwips(spacingPt),
      positionHalfPoints: ptToHalf(positionPt),
    });
    dialog.close();
  });
  foot.append(cancel, ok);

  root.append(tabs, fontPane, advPane, prev, foot);
  dialog.body.append(root);
  // Dialog appends its element to the body on construction and this helper
  // builds a fresh one per open, so without this every visit leaves one behind.
  dialog.onClose(() => dialog.destroy());
  showTab(false);
  paint();
  dialog.open();
}

import { injectStyle } from './internal.js';

/**
 * One scrollbar for the whole app — the editor canvas, the sidebar, a
 * dialog's list, a picker's grid: an idle-hidden overlay rail that fades in
 * while its container scrolls (or the rail is hovered) and fades out a second
 * later, in the app's own tint. Lifted from the desktop shell so every host
 * and every runtime surface bapbong-ui injects looks the same; before this,
 * a dialog's list scrolled with whatever the engine draws by default.
 *
 * Call once per document. Colours come from `--bb-ui-scroll-thumb` /
 * `--bb-ui-scroll-thumb-hover` (a host maps its theme tokens; the defaults
 * suit a light page). The class is `bb-scrolling`, set on the scrolled
 * element from a capture-phase scroll listener, so a container needs no
 * markup of its own.
 */
const STYLE = `
@property --bb-scroll-thumb-now {
  syntax: '<color>';
  /* MUST inherit: the thumb pseudo-element never matches .bb-scrolling
     itself — it only ever INHERITS the value from its owner. */
  inherits: true;
  initial-value: transparent;
}
* { transition: --bb-scroll-thumb-now 0.35s ease; }
.bb-scrolling { --bb-scroll-thumb-now: var(--bb-ui-scroll-thumb, rgba(64,52,28,.28)); }
::-webkit-scrollbar { width: 13px; height: 13px; background: transparent; }
::-webkit-scrollbar-track { background: transparent; }
::-webkit-scrollbar-corner { background: transparent; }
::-webkit-scrollbar-thumb {
  background-color: var(--bb-scroll-thumb-now);
  background-clip: content-box;
  border: 3px solid transparent;
  border-radius: 99px;
  min-height: 40px; /* a 277-page doc must not shrink the thumb to a dot */
}
::-webkit-scrollbar-thumb:hover,
::-webkit-scrollbar-thumb:active {
  background-color: var(--bb-ui-scroll-thumb-hover, rgba(64,52,28,.45));
}
@supports not selector(::-webkit-scrollbar) {
  /* non-WebKit (Firefox): native thin overlay, same colours */
  * { scrollbar-width: thin; scrollbar-color: var(--bb-ui-scroll-thumb, rgba(64,52,28,.28)) transparent; }
}
`;

let installed: (() => void) | null = null;

/** Install the shared scrollbar (styles + scroll-activity flagging).
 *  Idempotent; returns a disposer for the listener. */
export function installScrollbars(root: Document = document): () => void {
  if (installed) return installed;
  injectStyle('bb-ui-scrollbars', STYLE);
  const quiet = new WeakMap<Element, number>();
  const onScroll = (ev: Event): void => {
    const el =
      ev.target instanceof HTMLElement ? ev.target : root.documentElement;
    el.classList.add('bb-scrolling');
    clearTimeout(quiet.get(el));
    quiet.set(
      el,
      window.setTimeout(() => el.classList.remove('bb-scrolling'), 1000),
    );
  };
  root.addEventListener('scroll', onScroll, { capture: true, passive: true });
  installed = () => {
    root.removeEventListener('scroll', onScroll, { capture: true });
    installed = null;
  };
  return installed;
}

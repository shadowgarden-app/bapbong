/**
 * Lightweight, flag-gated performance instrumentation shared across the bapbong
 * packages. This is a **measurement-only** tool for investigating hot paths
 * (typing latency, document switch, large-doc layout) — it adds no behavior and
 * is a near-zero-cost no-op unless explicitly enabled.
 *
 * Enable it at runtime (before or during a session) with either:
 *   - `globalThis.__BAPBONG_PERF__ = true`
 *   - `localStorage.setItem('bapbong.perf', '1')`  (browser / webview)
 *
 * Then every wrapped span logs its duration to the console, nested by call
 * depth so a single keystroke's breakdown reads as an indented tree:
 *
 *   [perf] layout: 42.7ms
 *   [perf]   build-items: 8.1ms
 *   [perf]   placeBlocks: 33.4ms
 *   [perf]   paras: 1 miss / 1240 hit · tables: 0 miss / 3 hit
 *
 * All entry points are safe to leave in place; they cost one boolean check when
 * the flag is off.
 */

/** Counters accumulated within the current top-level span (hits/misses etc.). */
type Counters = Record<string, number>;

interface PerfState {
  depth: number;
  counters: Counters;
}

// Kept on globalThis so every package instance (even if bundled separately)
// shares one depth/counter state and one enabled flag.
function state(): PerfState {
  const g = globalThis as unknown as { __BAPBONG_PERF_STATE__?: PerfState };
  return (g.__BAPBONG_PERF_STATE__ ??= { depth: 0, counters: {} });
}

function computeEnabled(): boolean {
  const g = globalThis as unknown as {
    __BAPBONG_PERF__?: unknown;
    localStorage?: { getItem(k: string): string | null };
  };
  if (g.__BAPBONG_PERF__ != null) return !!g.__BAPBONG_PERF__;
  try {
    return g.localStorage?.getItem('bapbong.perf') === '1';
  } catch {
    return false;
  }
}

function nowMs(): number {
  const p = (globalThis as { performance?: { now(): number } }).performance;
  return p ? p.now() : Date.now();
}

function indent(depth: number): string {
  return '  '.repeat(Math.max(0, depth));
}

// Optional extra destination for emitted lines (besides console). A host whose
// console is invisible (e.g. a packaged WKWebView) registers one to funnel the
// trace somewhere readable — see `perf.setSink`.
let sink: ((line: string) => void) | null = null;

function emit(depth: number, msg: string): void {
  const line = `[perf] ${indent(depth)}${msg}`;
  // eslint-disable-next-line no-console
  console.log(line);
  if (sink) {
    try {
      sink(line);
    } catch {
      // a broken sink must never break the traced code
    }
  }
}

export const perf = {
  /** Whether instrumentation is currently active. Re-checked each call to the
   *  public helpers so it can be toggled live from the console. */
  get enabled(): boolean {
    return computeEnabled();
  },

  /** Force the enabled flag on/off from code (equivalent to setting the global). */
  setEnabled(on: boolean): void {
    (globalThis as { __BAPBONG_PERF__?: boolean }).__BAPBONG_PERF__ = on;
  },

  /** Register an extra destination for each emitted `[perf] …` line (in addition
   *  to console.log). Pass null to clear. Used by hosts with an invisible
   *  console (e.g. the packaged desktop WKWebView) to funnel the trace to a
   *  readable place like the shell's terminal. */
  setSink(fn: ((line: string) => void) | null): void {
    sink = fn;
  },

  /** High-resolution timestamp (ms). */
  now: nowMs,

  /** Emit a single labeled duration (ms) at the current depth — for a timing
   *  that spans async ticks and so can't be a wrapped span (e.g. keydown → the
   *  paint it eventually triggers). No-op when off. */
  log(label: string, ms: number): void {
    if (!computeEnabled()) return;
    emit(state().depth, `${label}: ${ms.toFixed(1)}ms`);
  },

  /** Time a synchronous span, logging its duration nested by depth. Returns
   *  the wrapped function's result. No-op wrapper (just calls fn) when off. */
  span<T>(label: string, fn: () => T): T {
    if (!computeEnabled()) return fn();
    const s = state();
    const t0 = nowMs();
    s.depth++;
    try {
      return fn();
    } finally {
      s.depth--;
      emit(s.depth, `${label}: ${(nowMs() - t0).toFixed(1)}ms`);
    }
  },

  /** Async variant of {@link span}. */
  async spanAsync<T>(label: string, fn: () => Promise<T>): Promise<T> {
    if (!computeEnabled()) return fn();
    const s = state();
    const t0 = nowMs();
    s.depth++;
    try {
      return await fn();
    } finally {
      s.depth--;
      emit(s.depth, `${label}: ${(nowMs() - t0).toFixed(1)}ms`);
    }
  },

  /** Add `n` to a named counter (e.g. cache hits/misses) for the current span.
   *  Flush with {@link counters}. No-op when off. */
  bump(name: string, n = 1): void {
    if (!computeEnabled()) return;
    const c = state().counters;
    c[name] = (c[name] ?? 0) + n;
  },

  /** Log the given counters (formatted) at the current depth and clear them.
   *  `format` maps the raw counter bag to a one-line summary string. */
  counters(format: (c: Counters) => string): void {
    if (!computeEnabled()) return;
    const s = state();
    const line = format(s.counters);
    if (line) emit(s.depth, line);
    s.counters = {};
  },

  /** Manually open a span; returns an `end()` that logs its duration. For the
   *  rare path that can't be expressed as a single wrapped callback. */
  begin(label: string): () => void {
    if (!computeEnabled()) return () => undefined;
    const s = state();
    const t0 = nowMs();
    s.depth++;
    return () => {
      s.depth--;
      emit(s.depth, `${label}: ${(nowMs() - t0).toFixed(1)}ms`);
    };
  },
};

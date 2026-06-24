# @shadow-garden/bapbong-ui

**Framework-agnostic editor UI** (L3). Vanilla DOM widgets that mount into a
host element and bind to `editor.commands` — a host framework only supplies the
element and calls `mount`/`destroy`. The same lib works in Angular, React,
Svelte, or plain HTML.

- **Scope:** `scope:app`
- **Depends on:** `@shadow-garden/bapbong-contracts` only. The editor surface is
  a structural `EditorHandle` interface (which `BapbongEditor` satisfies), and
  the editor-state type is derived from the `Command` contract — so this package
  imports neither the editor nor ProseMirror (same decoupling as the plugin
  contract).

## Toolbar

```ts
import { mountToolbar } from '@shadow-garden/bapbong-ui';

const handle = mountToolbar(hostEl, editor);   // renders + wires everything
// …
handle.destroy();                              // removes DOM + listeners
```

- Renders a button per command from the registry; the lib owns labels/icons
  (the headless `Command` carries none), groups, styling and active state.
- Click → `command.run(state, dispatch)`; active/disabled tracked via
  `editor.onChange` + `command.isActive`/`isEnabled`.
- Customise with `{ groups, items }` — groups reference command names; default
  is marks then alignments, derived from the registry.
- Theme via CSS variables: `--bb-ui-bg`, `--bb-ui-fg`, `--bb-ui-border`,
  `--bb-ui-hover`, `--bb-ui-active-bg`, `--bb-ui-active-fg`.

## Menubar

```ts
import { mountMenubar } from '@shadow-garden/bapbong-ui';

const handle = mountMenubar(hostEl, editor);   // top-level titles → dropdowns
handle.destroy();
```

- Top-level titles open dropdowns of command rows; active toggles show a check.
- Click-outside / Escape close; arrow keys move within an open menu; hovering
  another title switches menus.
- Customise with `{ menus, labels }` — `menus` is a declarative tree of
  `{ label, entries }` where entries reference command names (or `'separator'`).
  Default is a "Format" menu (marks, separator, alignments) from the registry.

## Build / test

```sh
pnpm nx build @shadow-garden/bapbong-ui
pnpm nx test  @shadow-garden/bapbong-ui   # Node (DOM behaviour verified in-browser)
```

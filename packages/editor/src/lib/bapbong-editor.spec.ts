import { BapbongEditor, type EditorPlugin } from './bapbong-editor';

describe('BapbongEditor', () => {
  it('is a constructable class', () => {
    // DOM/canvas wiring is exercised end-to-end in the playground app; here we
    // only assert the public surface is exported as a class.
    expect(typeof BapbongEditor).toBe('function');
    expect(BapbongEditor.prototype.constructor).toBe(BapbongEditor);
  });

  it('accepts an EditorPlugin shape', () => {
    // Type-level proof the plugin contract is importable + implementable; the
    // hooks themselves are exercised in-browser (they need a canvas context).
    let setupCalled = false;
    const probe: EditorPlugin = {
      name: 'probe',
      setup: () => {
        setupCalled = true;
        return () => undefined; // teardown
      },
      onChange: (c) => void c.pageCount,
      onCaretPick: (pos) => void pos,
    };
    expect(probe.name).toBe('probe');
    expect(typeof probe.setup).toBe('function');
    expect(setupCalled).toBe(false); // setup runs only when the editor calls it
  });
});

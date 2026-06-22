import { BapbongEditor } from './bapbong-editor';

describe('BapbongEditor', () => {
  it('is a constructable class', () => {
    // DOM/canvas wiring is exercised end-to-end in the playground app; here we
    // only assert the public surface is exported as a class.
    expect(typeof BapbongEditor).toBe('function');
    expect(BapbongEditor.prototype.constructor).toBe(BapbongEditor);
  });
});

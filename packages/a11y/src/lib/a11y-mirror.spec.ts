import { A11yMirror } from '../index.js';

// A11yMirror builds DOM via DOMSerializer + a host element — exercised
// end-to-end in the playground (the node test env has no DOM). Here we only
// assert the public surface is exported as a class.
describe('bapbong-a11y exports', () => {
  it('exposes A11yMirror as a class', () => {
    expect(typeof A11yMirror).toBe('function');
    expect(A11yMirror.prototype.constructor).toBe(A11yMirror);
  });
});

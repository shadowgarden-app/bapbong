import type { FontSpec, PageConfig, ResolvedLayout } from './contracts.js';

describe('contracts', () => {
  it('shapes compose as expected', () => {
    const font: FontSpec = {
      family: 'Arial',
      sizePt: 11,
      bold: false,
      italic: false,
    };
    const page: PageConfig = {
      width: 816,
      height: 1056,
      margin: { top: 96, right: 96, bottom: 96, left: 96 },
    };
    const layout: ResolvedLayout = {
      pages: [{ index: 0, width: page.width, height: page.height, lines: [] }],
    };
    expect(font.sizePt).toBe(11);
    expect(page.margin.left).toBe(96);
    expect(layout.pages[0].index).toBe(0);
  });
});

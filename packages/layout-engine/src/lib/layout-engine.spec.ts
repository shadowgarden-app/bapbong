import { Schema } from 'prosemirror-model';
import type { FlowBlock, FontSpec, LayoutConfig, MeasureText } from '@shadow-garden/bapbong-contracts';
import { layoutBlocks, toFlowBlocks } from './layout-engine.js';

// 10px per character, font-agnostic — keeps wrapping math predictable.
const measure: MeasureText = (text) => text.length * 10;

const font = (over: Partial<FontSpec> = {}): FontSpec => ({
  family: 'Arial',
  sizePt: 10, // → 10 * 96/72 * 1.2 = 16px line height
  bold: false,
  italic: false,
  ...over,
});

const para = (text: string, marker?: string): FlowBlock => ({
  type: 'paragraph',
  runs: [{ text, font: font() }],
  marker,
});

const config = (over: Partial<LayoutConfig['page']> = {}): LayoutConfig => ({
  measureText: measure,
  defaultFont: { sizePt: 10 },
  page: { width: 240, height: 1000, margin: { top: 20, right: 20, bottom: 20, left: 20 }, ...over },
});

describe('layoutBlocks', () => {
  it('wraps a paragraph at the content width', () => {
    // content width = 240 - 20 - 20 = 200px → 5 words of "aaaa"(40) + spaces.
    const { pages } = layoutBlocks([para('aaaa bbbb cccc dddd eeee')], config());
    expect(pages).toHaveLength(1);
    expect(pages[0].lines).toHaveLength(2);
    expect(pages[0].lines[0].segments[0].text).toBe('aaaa');
    expect(pages[0].lines[0].y).toBe(20);
    expect(pages[0].lines[1].segments[0].text).toBe('eeee');
    expect(pages[0].lines[1].y).toBe(36); // 20 + 16
  });

  it('paginates when content exceeds the page height', () => {
    // page height 80, margins 20 → content 20..60 = 40px → 2 lines/page.
    const cfg = config({ height: 80 });
    const { pages } = layoutBlocks([para('one'), para('two'), para('three')], cfg);
    expect(pages).toHaveLength(2);
    expect(pages[0].lines).toHaveLength(2);
    expect(pages[1].lines).toHaveLength(1);
    expect(pages[1].lines[0].y).toBe(20); // reset to top on the new page
  });

  it('places a list marker before the indented content', () => {
    const { pages } = layoutBlocks([para('hi', '1.')], config());
    const [markerSeg, textSeg] = pages[0].lines[0].segments;
    expect(markerSeg.text).toBe('1.');
    expect(markerSeg.x).toBe(20);
    expect(textSeg.text).toBe('hi');
    expect(textSeg.x).toBe(50); // 20 + measure("1. ") = 20 + 30
  });

  it('always emits at least one page', () => {
    expect(layoutBlocks([], config()).pages).toHaveLength(1);
  });
});

describe('toFlowBlocks', () => {
  const schema = new Schema({
    nodes: {
      doc: { content: 'block+' },
      paragraph: { group: 'block', content: 'inline*', attrs: { list: { default: null } } },
      text: { group: 'inline' },
    },
    marks: { strong: {}, fontSize: { attrs: { size: {} } } },
  });

  it('resolves marks into run fonts', () => {
    const doc = schema.node('doc', null, [
      schema.node('paragraph', null, [
        schema.text('Hi', [schema.mark('strong')]),
        schema.text('Big', [schema.mark('fontSize', { size: 20 })]),
      ]),
    ]);
    const blocks = toFlowBlocks(doc, { sizePt: 10 });
    expect(blocks).toHaveLength(1);
    expect(blocks[0].runs[0].font.bold).toBe(true);
    expect(blocks[0].runs[1].font.sizePt).toBe(20);
  });

  it('carries the list marker', () => {
    const doc = schema.node('doc', null, [
      schema.node('paragraph', { list: { marker: '1.' } }, [schema.text('item')]),
    ]);
    expect(toFlowBlocks(doc)[0].marker).toBe('1.');
  });
});

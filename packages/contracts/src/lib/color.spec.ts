import { describe, expect, it } from 'vitest';
import { hexColor } from './color.js';

describe('hexColor', () => {
  it('keeps hex, with or without #, short or long, as upper-case #RRGGBB', () => {
    expect(hexColor('#1f4e79')).toBe('#1F4E79');
    expect(hexColor('1F4E79')).toBe('#1F4E79');
    expect(hexColor('#abc')).toBe('#AABBCC');
  });

  it('reads the rgb() a browser pastes (the report: text drawn blue)', () => {
    expect(hexColor('rgb(0, 0, 0)')).toBe('#000000');
    expect(hexColor('rgb(31 78 121)')).toBe('#1F4E79');
    expect(hexColor('rgba(255, 0, 0, 0.5)')).toBe('#FF0000');
    expect(hexColor('rgb(100%, 0%, 0%)')).toBe('#FF0000');
  });

  it('reads basic CSS names', () => {
    expect(hexColor('Black')).toBe('#000000');
    expect(hexColor('navy')).toBe('#000080');
  });

  it('is no colour for auto, transparent and anything unreadable', () => {
    for (const v of [
      'auto',
      'transparent',
      'rgba(0, 0, 0, 0)',
      'currentcolor',
      'text1',
      'rgb(1, 2)',
      '',
      null,
      undefined,
    ])
      expect(hexColor(v)).toBeUndefined();
  });
});

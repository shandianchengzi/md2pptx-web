import test from 'node:test';
import assert from 'node:assert/strict';
import { layoutMarkdown, parseMarkdown, wrapText } from '../src/pagination.mjs';

test('headings produce logical sections without losing content', () => {
  const doc = parseMarkdown('# Report\n\n## Collector\n\n测试内容。\n\n## Validator\n\nNext part.');
  assert.equal(doc.documentTitle, 'Report');
  assert.equal(doc.sections.length, 2);
  assert.equal(doc.sections[0].title, 'Collector');
  assert.equal(doc.sections[1].blocks[0].text, 'Next part.');
});
test('long Chinese text creates multiple pages and preserves all characters', () => {
  const input = '# Long\n\n' + '中文自动分页验证。'.repeat(400);
  const out = layoutMarkdown(input);
  assert.ok(out.pages.length >= 3);
  const original = '中文自动分页验证。'.repeat(400);
  const actual = out.pages.flatMap(p => p.rows).map(x => x.text).join('');
  assert.equal(actual, original);
});
test('explicit slide separators force a page break', () => {
  const out = layoutMarkdown('## Test\n\nFirst.\n\n---\n\nSecond.');
  assert.equal(out.pages.length, 2);
});
test('bullets, links, quotes and fenced code remain as editable text', () => {
  const out = layoutMarkdown('## Parts\n\n- **Bold** item\n> Quoted\n\nSee [Docs](https://example.com)\n\n~~~\nlet x = 1;\n~~~');
  const text = out.pages.flatMap(p=>p.rows).map(r=>r.text).join(' ');
  assert.match(text, /Bold item/);
  assert.match(text, /Quoted/);
  assert.match(text, /Docs/);
  assert.match(text, /let x = 1/);
});
test('wrapText never exceeds width under same measure function', () => {
  const measure = (text) => Array.from(text).length * 10;
  const wrapped = wrapText('中文混排英文字母ABC', 70, 23, measure);
  assert.ok(wrapped.every(line=>measure(line)<=70));
});

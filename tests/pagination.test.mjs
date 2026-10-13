import test from 'node:test';
import assert from 'node:assert/strict';
import { layoutMarkdown, parseMarkdown, wrapText, resolveBodyFontPt } from '../src/pagination.mjs';

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

test('mixed Chinese and English share one line when room is available', () => {
  const lines = wrapText('Collector 已经完成重构，LLM 可以继续抽取内容。', 350, 29);
  assert.match(lines[0], /Collector 已经/);
  assert.equal(lines.join('').replace(/\s/g, ''), 'Collector已经完成重构，LLM可以继续抽取内容。');
});
test('avoids orphan paragraph lines at bottom of slides', () => {
  const long = '我们需要根据参考手册检查外设规则是否完整。'.repeat(45);
  const doc = layoutMarkdown('## Collector\n\n' + long);
  const tail = doc.pages.map(p => p.rows.length);
  assert.ok(tail.every(count => count >= 2), JSON.stringify(tail));
});

test('editable body font sizes use actual PowerPoint points', () => {
  const content = '## Collector\n\n' + 'AgentEmu 包含寄存器和事件语义。'.repeat(250);
  const small = layoutMarkdown(content, {bodyFontPt: 16});
  const large = layoutMarkdown(content, {bodyFontPt: 30});
  assert.ok(small.pages.length < large.pages.length, 'larger body fonts should need more slides');
  const smallBody = small.pages.flatMap(p=>p.rows).find(r=>r.kind === 'paragraph');
  const largeBody = large.pages.flatMap(p=>p.rows).find(r=>r.kind === 'paragraph');
  assert.ok(Math.abs(smallBody.size * .75 - 16) < .1);
  assert.ok(Math.abs(largeBody.size * .75 - 30) < .1);
  for (const output of [small, large]) {
    const actual = output.pages.flatMap(p=>p.rows).map(r=>r.text).join('');
    assert.equal(actual, 'AgentEmu 包含寄存器和事件语义。'.repeat(250));
    assert.ok(output.pages.every(p=>p.rows.every(r=>r.y+r.height<=635)), 'rows must fit within page');
  }
});
test('font size defaults and invalid settings are controlled', () => {
  assert.equal(resolveBodyFontPt(undefined,'comfortable'),22);
  assert.equal(resolveBodyFontPt(undefined,'compact'),19);
  assert.equal(resolveBodyFontPt('20','compact'),20);
  assert.equal(resolveBodyFontPt(200),22);
  assert.equal(resolveBodyFontPt(''),22);
});

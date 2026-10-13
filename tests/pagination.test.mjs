import test from 'node:test';
import assert from 'node:assert/strict';
import { layoutMarkdown, parseMarkdown, wrapText, resolveBodyFontPt, resolveMarginsCm, parseBoldRuns, PX_PER_CM } from '../src/pagination.mjs';

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
    assert.equal(actual.replace(/\s/g, ''), 'AgentEmu 包含寄存器和事件语义。'.repeat(250).replace(/\s/g, ''));
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

test('four margins change exact header placement and usable area', () => {
  const md = '## Collector\n\n' + '根据固件测试分析生成规则。'.repeat(170);
  const a = layoutMarkdown(md, {marginsCm:{top:1.5,bottom:2.2,left:2.1,right:2.1}});
  const b = layoutMarkdown(md, {marginsCm:{top:3.0,bottom:4.0,left:4.0,right:4.0}});
  assert.ok(b.pages.length > a.pages.length, 'larger margins should produce more pages');
  assert.ok(Math.abs(b.pages[0].header[0].x - 4*PX_PER_CM) < .01);
  assert.ok(Math.abs(b.pages[0].header[0].y - 3*PX_PER_CM) < .01);
  assert.ok(b.pages.every(page=>page.rows.every(row => row.y+row.height <= page.layout.contentBottom + .001)));
  assert.ok(b.pages.every(page=>page.rows.every(row => row.x+row.width <= 1280-4*PX_PER_CM + .01)));
  assert.equal(a.pages.flatMap(p=>p.rows).map(r=>r.text).join('').replace(/\s/g,''),
    b.pages.flatMap(p=>p.rows).map(r=>r.text).join('').replace(/\s/g,''));
});
test('margin presets clamp invalid inputs without changing defaults', () => {
  assert.deepEqual(resolveMarginsCm(), {top:1.5,bottom:2.2,left:2.1,right:2.1});
  assert.deepEqual(resolveMarginsCm({top:-10,left:99,bottom:NaN,right:""}),
    {top:.5,bottom:2.2,left:5,right:2.1});
});
test('bold Markdown survives parsing, wrapping, and page breaks', () => {
  const runs = parseBoldRuns('前面 **重要事实** 和 __特殊处理__。');
  assert.equal(runs.map(x=>x.text).join(''), '前面 重要事实 和 特殊处理。');
  assert.deepEqual(runs.filter(x=>x.bold).map(x=>x.text), ['重要事实','特殊处理']);
  const result = layoutMarkdown('## Test\n\n' + '这段文字说明 **重要事实**，需要进行对照。'.repeat(150),
    {bodyFontPt:30, marginsCm:{top:4,bottom:4,left:4,right:4}});
  const bodyRows = result.pages.flatMap(p=>p.rows).filter(r=>r.kind==='paragraph');
  assert.ok(result.pages.length>2);
  assert.ok(bodyRows.some(r=>r.runs?.some(x=>x.bold && x.text.includes('重要事实'))));
  for (const row of bodyRows) {
    assert.equal(row.runs.map(r=>r.text).join(''),row.text,'rich runs must preserve line text');
  }
});
test('footer title and page numbers are optional and off by default', () => {
  const a = layoutMarkdown('## Test\n\nSome text.');
  assert.equal(a.layout.showFooterTitle,false);
  assert.equal(a.layout.showPageNumber,false);
  const b = layoutMarkdown('## Test\n\nSome text.',{showFooterTitle:true,showPageNumber:true});
  assert.equal(b.layout.showFooterTitle,true);
  assert.equal(b.layout.showPageNumber,true);
  assert.ok(b.layout.contentBottom<a.layout.contentBottom);
});

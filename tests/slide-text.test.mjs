import test from 'node:test';
import assert from 'node:assert/strict';
import {layoutMarkdown} from '../src/pagination.mjs';
import {buildSlideTextBox} from '../src/slide-text.mjs';

const UNIT=13.333333/1280;
test('one editable box contains heading, body, and inline bold styles',()=>{
  const deck=layoutMarkdown('## Collector\n\n- **加粗文本** normal text\n\n## Validator\n\n说明文字');
  for(const [index,page] of deck.pages.entries()){
    const data=buildSlideTextBox(page,deck.title,index,deck.pages.length,UNIT);
    assert.ok(data.text.length>=page.header.length+page.rows.length);
    assert.equal(data.text.filter(x=>x.options.breakLine).length,
      page.header.length+page.rows.length-1);
    if(page.title==='Collector') assert.ok(data.text.some(x=>x.options.bold&&x.text.includes('加粗文本')));
    assert.ok(data.box.x>=0&&data.box.y>=0&&data.box.w>0&&data.box.h>0);
    const joined=data.text.map(x=>x.text).join('');
    for(const line of [...page.header,...page.rows])
      assert.ok(joined.includes(line.text.slice(0,Math.min(6,line.text.length))));
  }
});

test('enabled footer title and page number are in the same rich text sequence',()=>{
  const deck=layoutMarkdown('## Title\n\nParagraph.',{showFooterTitle:true,showPageNumber:true});
  const data=buildSlideTextBox(deck.pages[0],'Doc heading',0,3,UNIT);
  assert.equal(data.text.at(-1).text,'Doc heading\t1 / 3');
  assert.equal(data.text.at(-1).options.tabStops[0].alignment,'r');
  assert.equal(data.text.at(-2).options.breakLine,true);
});

test('disabled footer emits neither title nor watermark',()=>{
  const deck=layoutMarkdown('## Example\n\nContent.');
  const data=buildSlideTextBox(deck.pages[0],'Custom title',0,1,UNIT);
  assert.equal(data.text.at(-1).text,'Content.');
  assert.ok(data.text.every(x=>!x.text.includes('MD2PPTX')&&!x.text.includes('Custom title')));
});

test('bold stays as run-level format when a paragraph spans slides',()=>{
  const deck=layoutMarkdown('## Example\n\n'+'前文 **强调这一点**，需要测试。'.repeat(150),{bodyFontPt:30});
  assert.ok(deck.pages.length>1);
  const runs=deck.pages.flatMap((p,i)=>buildSlideTextBox(p,deck.title,i,deck.pages.length,UNIT).text);
  assert.ok(runs.some(x=>x.text.includes('强调这一点')&&x.options.bold));
  assert.ok(runs.some(x=>x.text.includes('前文')&&!x.options.bold));
});

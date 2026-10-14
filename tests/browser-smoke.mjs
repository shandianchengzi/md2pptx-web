import { createServer } from 'node:http';
import { readFile, mkdtemp, stat, rm, access } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import assert from 'node:assert/strict';
import { chromium } from 'playwright-core';
import JSZip from 'jszip';
import { execFileSync } from 'node:child_process';


async function validateSlideStructuralIntegrity(zip) {
  const slideNames=Object.keys(zip.files).filter(x=>/^ppt\/slides\/slide\d+\.xml$/.test(x));
  for(const path of slideNames) {
    const slide=await zip.file(path).async('string');
    const ids=[...slide.matchAll(/<p:cNvPr\b[^>]*\bid="(\d+)"/g)].map(m=>m[1]);
    assert.equal(ids.length,new Set(ids).size,path+': all shape ids must be unique');
    for(const m of slide.matchAll(/<a:p>([\s\S]*?)<\/a:p>/g)) {
      const starts=(m[1].match(/<a:pPr(?:\s|>)/g)||[]).length;
      assert.ok(starts<=1,path+': each paragraph may have one a:pPr only');
    }
  }
}

const prefix = '/md2pptx-web/';

// A valid ZIP is not enough: Microsoft PowerPoint can reject a well-formed
// archive with invalid or incompatible presentation structures.
// Load the output with a real office engine and render it to PDF as a smoke test.
async function verifyCanOpenInOffice(pptxPath,label) {
  const root=await mkdtemp(join(tmpdir(),'md2pptx-office-'));
  try {
    const profile='file://'+join(root,'profile');
    const result=execFileSync('libreoffice',[
      '-env:UserInstallation='+profile,
      '--headless','--convert-to','pdf','--outdir',root,pptxPath
    ],{encoding:'utf8',timeout:90000});
    const expected=join(root,pptxPath.split('/').at(-1).replace(/\.pptx$/i,'.pdf'));
    const info=await stat(expected);
    assert.ok(info.size>500, 'LibreOffice produced an empty PDF: '+label+' '+result);
    console.log('Office compatibility verified:',label,info.size,'bytes PDF');
  } finally {
    await rm(root,{recursive:true,force:true});
  }
}

const allowed = new Set([
  'index.html', 'favicon.svg', 'src/app.js', 'src/pagination.mjs',
  'src/style.css', 'src/slide-text.mjs', 'src/office-compat.mjs', 'src/math-render.mjs', 'vendor/tex-svg-full.js', 'vendor/jszip.min.js', 'vendor/pptxgen.bundle.js'
]);
const types = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.mjs': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8',
  '.svg': 'image/svg+xml'
};
const server = createServer(async (req, res) => {
  try {
    const url = new URL(req.url, 'http://127.0.0.1');
    if (!url.pathname.startsWith(prefix)) {
      res.writeHead(404); res.end(); return;
    }
    const path = url.pathname.slice(prefix.length) || 'index.html';
    if (!allowed.has(path)) {
      res.writeHead(404); res.end(); return;
    }
    const content = await readFile(join('dist', path));
    const ext = path.slice(path.lastIndexOf('.'));
    res.writeHead(200, { 'Content-Type': types[ext] || 'application/octet-stream' });
    res.end(content);
  } catch (err) {
    res.writeHead(500); res.end(String(err));
  }
});
await new Promise(resolve => server.listen(0, '127.0.0.1', resolve));
const origin = 'http://127.0.0.1:' + server.address().port;
const choices = [
  process.env.CHROME_BIN,
  '/usr/bin/google-chrome',
  '/usr/bin/google-chrome-stable',
  '/usr/bin/chromium',
  '/usr/bin/chromium-browser'
].filter(Boolean);
let browser;
let downloadDir;
try {
  let executablePath;
  for (const candidate of choices) {
    try { await access(candidate); executablePath = candidate; break; } catch {}
  }
  assert.ok(executablePath, 'Chrome/Chromium is required for browser smoke test');
  browser = await chromium.launch({
    executablePath, headless: true,
    args: ['--no-sandbox', '--disable-dev-shm-usage']
  });
  const page = await browser.newPage({ acceptDownloads: true });
  const errors = [];
  page.on('pageerror', err => errors.push(err.message));
  page.on('console', msg => { if (msg.type() === 'error') errors.push(msg.text()); });
  const response = await page.goto(origin + prefix, { waitUntil: 'load' });
  assert.equal(response.status(), 200, 'static site must be served from project prefix');
  await page.waitForFunction(() => document.getElementById('slideCanvas'));
  const constructorName = await page.evaluate(() => typeof (window.PptxGenJS || window.pptxgen));
  assert.equal(constructorName, 'function', 'browser PPTX bundle must expose constructor');

  const body = '## Collector\n\n' +
    'Collector **已经**按照六类信息完成重构。LLM 与 EP-DSL 协同提取外设功能语义。'.repeat(60) +
    '\n\n## Validator\n\n' +
    '必须根据固件输入与输出检查正确性。'.repeat(35);
  await page.locator('#markdownInput').fill(body);
  await page.waitForFunction(() => Number(document.getElementById('pageCount').textContent) >= 4);
  const total = Number(await page.locator('#pageCount').textContent());
  assert.ok(total >= 4);
  const previewText = await page.locator('#slideCanvas').innerText();
  assert.match(previewText, /Collector 已经/);
  assert.ok(await page.locator('.slide-line.kind-paragraph span[style*="750"]').count() > 0,
    'preview must visually preserve Markdown bold');
  assert.equal(await page.locator('.thumb').count(), total);
  assert.equal(await page.locator('#status').evaluate(el=>el.classList.contains('error')), false);
  assert.equal(await page.locator('#bodyFontPt').inputValue(), '22');
  await page.locator('#bodyFontPt').fill('30');
  await page.waitForFunction(() => {
    const row = document.querySelector('.slide-line.kind-paragraph');
    return row && Number.parseFloat(row.style.fontSize) >= 39.9;
  });
  const largerCount = Number(await page.locator('#pageCount').textContent());
  assert.ok(largerCount >= total, 'font increase should not reduce page count');
  await page.reload({waitUntil:'load'});
  assert.equal(await page.locator('#bodyFontPt').inputValue(), '30', 'font size preference must persist');
  await page.locator('#markdownInput').fill(body);
  await page.waitForFunction(() => Number(document.querySelector('#pageCount').textContent) >= 4);
  const resizedCount = Number(await page.locator('#pageCount').textContent());
  assert.equal(resizedCount, largerCount, 'restored font size must preserve pagination');
  assert.equal(await page.locator('#showFooterTitle').isChecked(), false);
  assert.equal(await page.locator('#showPageNumber').isChecked(), false);
  await page.locator('#marginLeft').fill('3.5');
  await page.locator('#marginRight').fill('4');
  await page.locator('#marginTop').fill('2.5');
  await page.locator('#marginBottom').fill('3');
  // Commit the last input synchronously instead of reading the slide count
  // while the 90 ms live-preview debounce is still pending.
  await page.locator('#marginBottom').dispatchEvent('change');
  await page.waitForFunction(() => {
    const heading = document.querySelector('.slide-line.kind-title');
    return heading && Math.abs(parseFloat(heading.style.left) - 3.5*1280/33.8666666667)<1
      && Math.abs(parseFloat(heading.style.top) - 2.5*1280/33.8666666667)<1;
  });
  const customPageCount = Number(await page.locator('#pageCount').textContent());
  await page.reload({waitUntil:'load'});
  for (const [id,v] of [['marginLeft','3.5'],['marginRight','4.0'],['marginTop','2.5'],['marginBottom','3.0']])
    assert.equal(await page.locator('#'+id).inputValue(),v,'margin persistence '+id);
  assert.equal(await page.locator('#showFooterTitle').isChecked(),false);
  assert.equal(await page.locator('#showPageNumber').isChecked(),false);
  await page.locator('#markdownInput').fill(body);
  await page.waitForFunction(expected => Number(document.getElementById('pageCount').textContent)===expected, customPageCount);
  assert.equal(Number(await page.locator('#pageCount').textContent()),customPageCount);
  assert.equal(await page.locator('.slide-footer').count(), 0, 'footer must be hidden by default');
  const [download] = await Promise.all([
    page.waitForEvent('download', { timeout: 35000 }),
    page.locator('#downloadButton').click()
  ]);
  downloadDir = await mkdtemp(join(tmpdir(), 'md2pptx-browser-test-'));
  const output = join(downloadDir, 'example.pptx');
  await download.saveAs(output);
  assert.ok((await stat(output)).size > 4000, 'download must contain PPTX data');
  await verifyCanOpenInOffice(output,'plain text and inline bold');
  const zip = await JSZip.loadAsync(await readFile(output));
  await validateSlideStructuralIntegrity(zip);
  assert.ok(zip.file('ppt/presentation.xml'), 'download must be an editable PowerPoint package');
  const slideFiles = Object.keys(zip.files).filter(p => /^ppt\/slides\/slide\d+\.xml$/.test(p));
  assert.equal(slideFiles.length, customPageCount, 'preview and PPTX must have the same slide count');
  const xml = await zip.file(slideFiles[0]).async('string');
  assert.match(xml, /Collector/, 'PPTX should contain editable slide text');
  for (const slidePath of slideFiles) {
    const currentXml = await zip.file(slidePath).async('string');
    assert.equal((currentXml.match(/<p:txBody>/g)||[]).length, 1,
      slidePath + ': exactly one editable text box expected');
  }
  assert.match(xml, /sz="3000"/, 'exported slide must use the selected 30pt body font');
  assert.match(xml, /b="1"/, 'exported rich-text runs must preserve Markdown bold');
  assert.doesNotMatch(xml, /MD2PPTX/, 'slide content must not contain product watermark');
  assert.doesNotMatch(xml, /Markdown Presentation/, 'optional footer title defaults to hidden');
  await page.locator('#showFooterTitle').check();
  await page.locator('#showPageNumber').check();
  await page.waitForFunction(() => document.querySelector('.slide-footer')?.textContent?.includes(' / '));
  const [footerDownload] = await Promise.all([
    page.waitForEvent('download',{timeout:35000}),
    page.locator('#downloadButton').click()
  ]);
  const footerOutput = join(downloadDir,'footer.pptx');
  await footerDownload.saveAs(footerOutput);
  const footerZip = await JSZip.loadAsync(await readFile(footerOutput));
  const footerXml = await footerZip.file('ppt/slides/slide1.xml').async('string');
  assert.match(footerXml, /Markdown Presentation/, 'enabled footer should show document title only');
  assert.equal((footerXml.match(/<p:txBody>/g)||[]).length, 1,
    'enabled footer must share the only text box');
  assert.match(footerXml, /1 \/ /, 'enabled page numbers should show the page index');
  assert.doesNotMatch(footerXml, /MD2PPTX/, 'enabled footer also must not include product watermark');
  await page.reload({waitUntil:'load'});
  assert.equal(await page.locator('#showFooterTitle').isChecked(),true);
  assert.equal(await page.locator('#showPageNumber').isChecked(),true);

  // Native Markdown tables must be previewed as tables, not prose, and
  // must export as editable a:tbl structures (not images or flat text).
  await page.locator('#showFooterTitle').uncheck();
  await page.locator('#showPageNumber').uncheck();
  const tableDoc='## Table verification\n\n' +
    '| Peripheral | Condition | Result |\n' +
    '| :--- | :---: | ---: |\n' +
    Array.from({length:34},(_,i)=>
      '| UART'+i+' | Check **IRQ** and DMA states repeatedly | '+i+' |').join('\n') +
    '\n\nAfter the table.';
  await page.locator('#markdownInput').fill(tableDoc);
  await page.waitForFunction(() => document.querySelector('.slide-native-table'));
  const tablePreview=await page.locator('.slide-native-table').count();
  assert.ok(tablePreview===1,'active slide must show one visible table');
  assert.equal(await page.locator('.slide-table-head .slide-table-cell').count(),3);
  assert.ok(await page.locator('.slide-table-row').count()>=2);
  const tablePageCount=Number(await page.locator('#pageCount').textContent());
  assert.ok(tablePageCount>=3,'table and following paragraph should span pages');
  const [tableDownload]=await Promise.all([
    page.waitForEvent('download',{timeout:35000}),
    page.locator('#downloadButton').click()
  ]);
  const tableOutput=join(downloadDir,'table.pptx');
  await tableDownload.saveAs(tableOutput);
  await verifyCanOpenInOffice(tableOutput,'multi-page editable table');
  const tableZip=await JSZip.loadAsync(await readFile(tableOutput));
  await validateSlideStructuralIntegrity(tableZip);
  const tableSlides=Object.keys(tableZip.files).filter(path=>/^ppt\/slides\/slide\d+\.xml$/.test(path));
  assert.equal(tableSlides.length,tablePageCount);
  let tableSlideCount=0;
  const slideTexts=[];
  for(const path of tableSlides) {
    const slideXml=await tableZip.file(path).async('string');
    slideTexts.push(slideXml);
    if(slideXml.includes('<a:tbl>')) {
      tableSlideCount++;
      assert.equal((slideXml.match(/<a:tbl>/g)||[]).length,1,'one native table per table slide');
      assert.equal((slideXml.match(/<p:txBody>/g)||[]).length,1,
        'one heading/text frame per slide; native table is a separate editable object');
      assert.match(slideXml,/Peripheral/,'table header must repeat across continuation slides');
      assert.match(slideXml,/IRQ/,'table cells must retain data');
      assert.match(slideXml,/b="1"/,'cell emphasis must remain editable bold text');
    }
  }
  assert.ok(tableSlideCount>=2,'table should continue over multiple slides');
  assert.ok(slideTexts.some(xml=>xml.includes('After the table.')),'text after table must not disappear');


  // Mixed text and native table must share the first slide where there is room.
  const mixedSource='**当前规模：142 个固件数据集，85 个单元测试用例**\n\n' +
    '| 固件类别 | 固件数 | 单元测试 | 模糊测试 | 其他固件 |\n' +
    '|---|---:|---:|---:|---:|\n' +
    '| 基础外设 | 46 | 66 | 0 | 0 |\n' +
    '| 数据传输及复杂外设 | 13 | 15 | 0 | 2 |\n' +
    '| 网络通信 | 34 | 2 | 15 | 18 |\n' +
    '| 射频通信 | 24 | 2 | 5 | 17 |\n' +
    '| 实际应用 | 25 | 0 | 15 | 10 |\n' +
    '| **合计** | **142** | **85** | **35** | **47** |';
  await page.locator('#bodyFontPt').fill('14');
  await page.locator('#density').selectOption('compact');
  await page.locator('#marginTop').fill('1.1');
  await page.locator('#marginBottom').fill('0.9');
  await page.locator('#marginLeft').fill('1.4');
  await page.locator('#marginRight').fill('1.4');
  await page.locator('#markdownInput').fill(mixedSource);
  await page.waitForFunction(() =>
    !!document.querySelector('.slide-native-table') &&
    document.querySelector('.slide-canvas .kind-paragraph')?.textContent?.includes('142'));
  assert.ok(await page.locator('.slide-native-table').isVisible());
  assert.ok(await page.locator('.slide-canvas .kind-paragraph').count()>0);
  const [mixedDownload]=await Promise.all([
    page.waitForEvent('download',{timeout:35000}),
    page.locator('#downloadButton').click()
  ]);
  const mixedPath=join(downloadDir,'mixed.pptx');
  await mixedDownload.saveAs(mixedPath);
  await verifyCanOpenInOffice(mixedPath,'prose and native table on one slide');
  const mixedZip=await JSZip.loadAsync(await readFile(mixedPath));
  await validateSlideStructuralIntegrity(mixedZip);
  const mixedXml=await mixedZip.file('ppt/slides/slide1.xml').async('string');
  assert.match(mixedXml,/142 个固件数据集/,'first slide includes prose');
  assert.match(mixedXml,/<a:tbl>/,'first slide includes a native editable table');
  assert.equal((mixedXml.match(/<p:txBody>/g)||[]).length,1);


  // Verify real offline rendering of inline and display LaTeX formulas.
  // No raw $ delimiters or TeX commands should appear as slide prose.
  const mathDoc=[
    '## Mathematics',
    '',
    'Mass-energy relation: $E=mc^2$ and inline \\\\(\\\\alpha+\\\\beta\\\\).',
    '',
    '$\\\\frac{a}{b}=\\\\sum_{i=1}^{n}i$',
    '',
    '\\\\[\\\\int_0^1 x^2 dx=\\\\frac13\\\\]'
  ].join('\n').replace(/\\\\/g,'\\');
  await page.locator('#markdownInput').fill(mathDoc);
  await page.waitForFunction(() =>
    document.querySelectorAll('.rendered-math svg').length >= 3,
    null,{timeout:25000});
  assert.ok(await page.locator('.rendered-math svg').count()>=3,
    'inline and block math must render to SVG in live preview');
  const [mathDownload]=await Promise.all([
    page.waitForEvent('download',{timeout:35000}),
    page.locator('#downloadButton').click()
  ]);
  const mathPath=join(downloadDir,'equations.pptx');
  await mathDownload.saveAs(mathPath);
  await verifyCanOpenInOffice(mathPath,'inline and block MathJax equations');
  const mathZip=await JSZip.loadAsync(await readFile(mathPath));
  await validateSlideStructuralIntegrity(mathZip);
  const imageNames=Object.keys(mathZip.files).filter(x=>/^ppt\/media\/image.*\.png$/i.test(x));
  assert.ok(imageNames.length>=3,'math must export as compatible PNG formula images');
  const mathXml=await mathZip.file('ppt/slides/slide1.xml').async('string');
  assert.doesNotMatch(mathXml,/\$\$|\\frac/, 'raw LaTeX should not be printed in slides');
  assert.ok((mathXml.match(/<p:pic>/g)||[]).length>=3,'each math image should be an editable placement');


  // Exact screenshot regression: multiline display math followed immediately
  // by a table. Both formulas must be SVGs, never raw TeX text/placeholders.
  const screenshotDoc=[
    '**当前规模：142 个固件数据集，85 个单元测试用例**',
    '$',
    'a_b,a^b',
    '$',
    '$\\frac{a}{b}$',
    '| 固件类别 | 固件数 | 单元测试用例 | 模糊测试固件 | 其他固件 |',
    '|---|---:|---:|---:|---:|',
    '| 基础外设 | 46 | 66 | 0 | 0 |',
    '| 数据传输及复杂外设 | 13 | 15 | 0 | 2 |',
    '| 网络通信 | 34 | 2 | 15 | 18 |',
    '| 射频通信 | 24 | 2 | 5 | 17 |',
    '| 实际应用 | 25 | 0 | 15 | 10 |',
    '| **合计** | **142** | **85** | **35** | **47** |'
  ].join('\n');
  await page.locator('#bodyFontPt').fill('14');
  await page.locator('#density').selectOption('compact');
  await page.locator('#marginTop').fill('1.1');
  await page.locator('#marginBottom').fill('0.9');
  await page.locator('#marginLeft').fill('1.4');
  await page.locator('#marginRight').fill('1.4');
  await page.locator('#markdownInput').fill(screenshotDoc);
  try {
    await page.waitForFunction(() => document.querySelectorAll('.display-math svg').length===2 &&
      !!document.querySelector('.slide-native-table'),null,{timeout:16000});
  } catch (e) {
    const detail=await page.evaluate(async()=>{
      const renderer=await import('./src/math-render.mjs');
      return {
        body:document.querySelector('#markdownInput')?.value,
        count:document.querySelector('#pageCount')?.textContent,
        active:document.querySelector('#slidePosition')?.textContent,
        displayMath:[...document.querySelectorAll('.display-math')].map(x=>({txt:x.innerText,tex:x.dataset.tex,svg:x.querySelectorAll('svg').length,classes:x.className})),
        tableCount:document.querySelectorAll('.slide-native-table').length,
        err:document.querySelector('#status')?.innerText,
        renderer:renderer.mathStatus(),
        mathJax:typeof window.MathJax?.tex2svg
      };
    });
    console.log('SCREENSHOT_CASE_DIAGNOSTICS '+JSON.stringify(detail));
    throw e;
  }
  assert.equal(await page.locator('.display-math').count(),2);
  assert.equal(await page.locator('.math-loading').count(),0);
  assert.equal(await page.locator('.math-error').count(),0);
  const checkGeometry=await page.evaluate(()=>{
    const formulas=[...document.querySelectorAll('.display-math')].map(el=>({
      y:parseFloat(el.style.top),h:parseFloat(el.style.height)
    }));
    const table=document.querySelector('.slide-native-table');
    return {gap:parseFloat(table.style.top) - (formulas.at(-1).y+formulas.at(-1).h),
      dataRows:table.querySelectorAll('.slide-table-row').length};
  });
  assert.ok(checkGeometry.gap>=0&&checkGeometry.gap<=25,
    'display formula should directly precede table: '+JSON.stringify(checkGeometry));
  assert.ok(checkGeometry.dataRows>=4,'formula space must not strand table at bottom');
  const [screenshotDownload]=await Promise.all([
    page.waitForEvent('download',{timeout:35000}),
    page.locator('#downloadButton').click()
  ]);
  const screenshotPptx=join(downloadDir,'screenshot-case.pptx');
  await screenshotDownload.saveAs(screenshotPptx);
  await verifyCanOpenInOffice(screenshotPptx,'multiline display math and editable table');
  const screenshotZip=await JSZip.loadAsync(await readFile(screenshotPptx));
  await validateSlideStructuralIntegrity(screenshotZip);
  const screenshotXml=await screenshotZip.file('ppt/slides/slide1.xml').async('string');
  assert.ok((screenshotXml.match(/<p:pic>/g)||[]).length>=2,'both formulas must export as PNG');

  assert.deepEqual(errors, [], 'browser console should have no runtime errors');
  console.log('Browser E2E passed: ' + customPageCount + ' pages with custom margins, bold and footer options at 30pt, valid ' + (await stat(output)).size + '-byte PPTX');
} finally {
  if (browser) await browser.close();
  if (downloadDir) await rm(downloadDir, { recursive:true, force:true });
  await new Promise(resolve => server.close(resolve));
}

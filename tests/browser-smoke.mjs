import { createServer } from 'node:http';
import { readFile, mkdtemp, stat, rm, access } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import assert from 'node:assert/strict';
import { chromium } from 'playwright-core';
import JSZip from 'jszip';

const prefix = '/md2pptx-web/';
const allowed = new Set([
  'index.html', 'favicon.svg', 'src/app.js', 'src/pagination.mjs',
  'src/style.css', 'vendor/pptxgen.bundle.js'
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
  const zip = await JSZip.loadAsync(await readFile(output));
  assert.ok(zip.file('ppt/presentation.xml'), 'download must be an editable PowerPoint package');
  const slideFiles = Object.keys(zip.files).filter(p => /^ppt\/slides\/slide\d+\.xml$/.test(p));
  assert.equal(slideFiles.length, customPageCount, 'preview and PPTX must have the same slide count');
  const xml = await zip.file(slideFiles[0]).async('string');
  assert.match(xml, /Collector/, 'PPTX should contain editable slide text');
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
  assert.match(footerXml, /1 \/ /, 'enabled page numbers should show the page index');
  assert.doesNotMatch(footerXml, /MD2PPTX/, 'enabled footer also must not include product watermark');
  await page.reload({waitUntil:'load'});
  assert.equal(await page.locator('#showFooterTitle').isChecked(),true);
  assert.equal(await page.locator('#showPageNumber').isChecked(),true);
  assert.deepEqual(errors, [], 'browser console should have no runtime errors');
  console.log('Browser E2E passed: ' + customPageCount + ' pages with custom margins, bold and footer options at 30pt, valid ' + (await stat(output)).size + '-byte PPTX');
} finally {
  if (browser) await browser.close();
  if (downloadDir) await rm(downloadDir, { recursive:true, force:true });
  await new Promise(resolve => server.close(resolve));
}

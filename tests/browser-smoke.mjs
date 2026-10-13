import { createServer } from 'node:http';
import { readFile, mkdtemp, stat, rm, access } from 'node:fs/promises';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
import assert from 'node:assert/strict';
import { chromium } from 'playwright-core';
import JSZip from 'jszip';

const prefix = '/md2pptx-web/';
const allowed = new Set([
  'index.html', 'src/app.js', 'src/pagination.mjs',
  'src/style.css', 'vendor/pptxgen.bundle.js'
]);
const types = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.mjs': 'text/javascript; charset=utf-8',
  '.css': 'text/css; charset=utf-8'
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
    'Collector 已经按照六类信息完成重构。LLM 与 EP-DSL 协同提取外设功能语义。'.repeat(60) +
    '\n\n## Validator\n\n' +
    '必须根据固件输入与输出检查正确性。'.repeat(35);
  await page.locator('#markdownInput').fill(body);
  await page.waitForFunction(() => Number(document.getElementById('pageCount').textContent) >= 4);
  const total = Number(await page.locator('#pageCount').textContent());
  assert.ok(total >= 4);
  const previewText = await page.locator('#slideCanvas').innerText();
  assert.match(previewText, /Collector 已经/);
  assert.equal(await page.locator('.thumb').count(), total);
  assert.equal(await page.locator('#status').evaluate(el=>el.classList.contains('error')), false);
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
  assert.equal(slideFiles.length, total, 'preview and PPTX must have the same slide count');
  const xml = await zip.file(slideFiles[0]).async('string');
  assert.match(xml, /Collector/, 'PPTX should contain editable slide text');
  assert.deepEqual(errors, [], 'browser console should have no runtime errors');
  console.log('Browser E2E passed: ' + total + ' preview pages, valid ' + (await stat(output)).size + '-byte PPTX');
} finally {
  if (browser) await browser.close();
  if (downloadDir) await rm(downloadDir, { recursive:true, force:true });
  await new Promise(resolve => server.close(resolve));
}

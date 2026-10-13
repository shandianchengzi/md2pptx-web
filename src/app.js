import {
  SLIDE_W, SLIDE_H, EXAMPLE,
  layoutMarkdown
} from './pagination.mjs';

const $ = (id) => document.getElementById(id);
const editor = $('markdownInput');
const titleInput = $('documentTitle');
const density = $('density');
const preview = $('slideCanvas');
const frame = $('slideFrame');
const thumbs = $('thumbnails');
const download = $('downloadButton');
const status = $('status');
const canvas = document.createElement('canvas');
const context = canvas.getContext('2d');
const PPT_WIDTH = 13.333333;
const UNIT = PPT_WIDTH / SLIDE_W;

let deck = { title: '', pages: [], blockCount: 0 };
let activeSlide = 0;
let pending = null;
let lastLength = 0;

function measure(text, size) {
  if (!context) return Array.from(text).length * size;
  context.font = '500 ' + size + 'px "Microsoft YaHei","PingFang SC","Noto Sans CJK SC",Arial,sans-serif';
  // Extra room avoids edge-of-line overflow on Windows PowerPoint.
  return context.measureText(text).width * 1.045;
}

function message(value, isError = false) {
  status.textContent = value;
  status.classList.toggle('error', isError);
}

function element(tag, className, content) {
  const el = document.createElement(tag);
  if (className) el.className = className;
  if (content !== undefined) el.textContent = content;
  return el;
}

function resizeSlide() {
  const width = frame.clientWidth || SLIDE_W;
  const height = frame.clientHeight || SLIDE_H;
  const scale = Math.min(width / SLIDE_W, height / SLIDE_H);
  preview.style.transform = 'scale(' + scale + ')';
}

function paintSlide() {
  preview.replaceChildren();
  const page = deck.pages[activeSlide];
  if (!page) {
    const placeholder = element('div', 'empty-preview', '粘贴 Markdown，即刻预览');
    placeholder.style.cssText = 'position:absolute;top:315px;left:0;width:100%;text-align:center;font-size:30px;color:#9db0bd';
    preview.appendChild(placeholder);
    resizeSlide();
    return;
  }
  preview.appendChild(element('div', 'decor-bar'));
  preview.appendChild(element('div', 'decor-line'));
  for (const row of [...page.header, ...page.rows]) {
    const el = element('div', 'slide-line kind-' + row.kind, row.text);
    el.style.left = row.x + 'px';
    el.style.top = row.y + 'px';
    el.style.fontSize = row.size + 'px';
    el.style.fontWeight = row.bold ? '750' : '450';
    preview.appendChild(el);
  }
  const footer = element('div', 'slide-footer');
  footer.appendChild(element('span', '', 'MD2PPTX  /  ' + (deck.title || 'Presentation')));
  footer.appendChild(element('span', '', String(activeSlide + 1) + ' / ' + deck.pages.length));
  preview.appendChild(footer);
  resizeSlide();
}

function paintThumbnails() {
  const frag = document.createDocumentFragment();
  deck.pages.forEach((page, index) => {
    const button = element('button', 'thumb' + (index === activeSlide ? ' active' : ''));
    button.type = 'button';
    button.setAttribute('role', 'option');
    button.setAttribute('aria-selected', index === activeSlide ? 'true' : 'false');
    button.setAttribute('aria-label', '第 ' + (index+1) + ' 页：' + page.title);
    const inner = element('div', 'thumb-inner');
    inner.appendChild(element('div', 'thumb-title', page.title + (page.continued ? '（续）' : '')));
    inner.appendChild(element('div', 'thumb-copy', page.rows.slice(0,4).map(x => x.text).join(' ')));
    button.appendChild(inner);
    button.appendChild(element('div', 'thumb-count', String(index + 1).padStart(2, '0')));
    button.addEventListener('click', () => {
      activeSlide = index;
      refreshNavigation();
      paintSlide();
      paintThumbnails();
    });
    frag.appendChild(button);
  });
  thumbs.replaceChildren(frag);
}

function refreshNavigation() {
  $('pageCount').textContent = String(deck.pages.length);
  $('slidePosition').textContent = '第 ' + (deck.pages.length ? activeSlide+1 : 0) + ' / ' + deck.pages.length + ' 页';
  $('prevSlide').disabled = activeSlide <= 0;
  $('nextSlide').disabled = activeSlide >= deck.pages.length - 1;
  $('textStats').textContent = editor.value.length.toLocaleString() + ' 字 · ' + deck.blockCount + ' 段';
  download.disabled = !editor.value.trim() || !deck.pages.length;
}

function regenerate() {
  clearTimeout(pending);
  pending = null;
  const length = editor.value.length;
  if (Math.abs(length - lastLength) > 100) activeSlide = 0;
  lastLength = length;
  if (!editor.value.trim()) {
    deck = { title: '', pages: [], blockCount: 0 };
    activeSlide = 0;
    message('粘贴内容后将自动显示预览。');
  } else {
    try {
      deck = layoutMarkdown(editor.value, {
        title: titleInput.value.trim(),
        density: density.value
      }, measure);
      activeSlide = Math.max(0, Math.min(activeSlide, deck.pages.length - 1));
      message('已自动分页：' + deck.pages.length + ' 页。点击下方按钮下载可编辑 PPTX。');
    } catch (err) {
      console.error(err);
      deck = { title: '', pages: [], blockCount: 0 };
      activeSlide = 0;
      message('排版时发生错误，请尝试缩短输入或检查格式。', true);
    }
  }
  refreshNavigation();
  paintSlide();
  paintThumbnails();
}

function scheduleRegenerate() {
  clearTimeout(pending);
  // Brief debounce prevents layout work on every keystroke but paste is almost immediate.
  pending = setTimeout(regenerate, 90);
}

function gotoSlide(nextIndex) {
  if (nextIndex < 0 || nextIndex >= deck.pages.length) return;
  activeSlide = nextIndex;
  refreshNavigation();
  paintSlide();
  paintThumbnails();
}

function addTextRow(slide, row) {
  const availableWidth = SLIDE_W - row.x - 77;
  slide.addText(row.text, {
    x: row.x * UNIT,
    y: row.y * UNIT,
    w: availableWidth * UNIT,
    h: Math.max(row.height, row.size * 1.3) * UNIT,
    margin: 0,
    fontFace: row.kind === 'code' ? 'Consolas' : 'Microsoft YaHei',
    fontSize: row.size * .75,
    bold: Boolean(row.bold),
    color: row.kind === 'title' ? '19354A' :
      row.kind === 'subheading' ? '205C79' :
      row.kind === 'quote' ? '587487' :
      row.kind === 'code' ? '355B72' : '233E52',
    valign: 'mid',
    breakLine: false,
    transparency: 0
  });
}

function sanitizeFileName(filename) {
  let clean = filename.replace(/[\\/:*?"<>|]+/g, '_').replace(/^\.+/, '').trim();
  if (clean.length > 72) clean = Array.from(clean).slice(0,72).join('');
  return clean || 'markdown-presentation';
}

async function exportPptx() {
  if (!deck.pages.length) return;
  if (typeof window.PptxGenJS !== 'function') {
    message('PPTX 转换模块未加载。请刷新页面或检查网站构建结果。', true);
    return;
  }
  download.disabled = true;
  message('正在生成 PPTX，所有文字保持可编辑……');
  try {
    const pptx = new window.PptxGenJS();
    pptx.layout = 'LAYOUT_WIDE';
    pptx.author = 'MD2PPTX Web';
    pptx.subject = 'Editable slides generated from Markdown';
    pptx.title = deck.title;
    pptx.lang = 'zh-CN';
    deck.pages.forEach((page, index) => {
      const slide = pptx.addSlide();
      slide.background = { color: 'FBFCFE' };
      slide.addShape(pptx.ShapeType.rect, {
        x: 0, y: 0, w: PPT_WIDTH, h: 12*UNIT,
        line: { color: '224B65', transparency: 100 }, fill: { color: '224B65' }
      });
      slide.addShape(pptx.ShapeType.line, {
        x: 78*UNIT, y: 132*UNIT, w: 1125*UNIT, h: 0,
        line: { color: 'E0E9EE', width: 1.1 }
      });
      [...page.header, ...page.rows].forEach(row => addTextRow(slide,row));
      slide.addShape(pptx.ShapeType.line, {
        x: 78*UNIT, y: 666*UNIT, w: 1125*UNIT, h: 0,
        line: { color: 'E7EDF1', width: .7 }
      });
      slide.addText('MD2PPTX  /  ' + deck.title, {
        x: 78*UNIT, y: 673*UNIT, w: 915*UNIT, h: 24*UNIT,
        fontFace: 'Arial', fontSize: 10, color: '8DA2AF', margin: 0, valign: 'mid'
      });
      slide.addText(String(index+1) + ' / ' + deck.pages.length, {
        x: 1090*UNIT, y: 673*UNIT, w: 108*UNIT, h: 24*UNIT,
        fontFace: 'Arial', fontSize: 10, color: '8DA2AF', margin: 0,
        align: 'right', valign: 'mid'
      });
    });
    await pptx.writeFile({ fileName: sanitizeFileName(deck.title) + '.pptx' });
    message('PPTX 已生成：' + deck.pages.length + ' 页。可以在 PowerPoint 中逐字编辑。');
  } catch (err) {
    console.error(err);
    message('PPTX 生成失败：' + (err && err.message ? err.message : String(err)), true);
  } finally {
    download.disabled = !deck.pages.length;
  }
}

editor.addEventListener('input', scheduleRegenerate);
titleInput.addEventListener('input', scheduleRegenerate);
density.addEventListener('change', regenerate);
download.addEventListener('click', exportPptx);

$('loadSample').addEventListener('click', () => {
  editor.value = EXAMPLE;
  activeSlide = 0;
  regenerate();
  editor.focus();
});
$('clearInput').addEventListener('click', () => {
  editor.value = '';
  activeSlide = 0;
  regenerate();
  editor.focus();
});
$('fileInput').addEventListener('change', async (event) => {
  const file = event.target.files && event.target.files[0];
  if (!file) return;
  if (file.size > 3 * 1024 * 1024) {
    message('文件过大。请导入小于 3 MB 的 Markdown 或纯文本文件。', true);
    return;
  }
  try {
    editor.value = await file.text();
    activeSlide = 0;
    regenerate();
  } catch (e) {
    message('读取文件失败，请确认编码为 UTF-8。', true);
  }
  event.target.value = '';
});
$('prevSlide').addEventListener('click', () => gotoSlide(activeSlide - 1));
$('nextSlide').addEventListener('click', () => gotoSlide(activeSlide + 1));

document.addEventListener('keydown', event => {
  if ((event.ctrlKey || event.metaKey) && event.key === 'Enter') {
    event.preventDefault();
    exportPptx();
  }
  if (event.target === editor || event.target === titleInput) return;
  if (event.key === 'ArrowLeft') gotoSlide(activeSlide - 1);
  else if (event.key === 'ArrowRight') gotoSlide(activeSlide + 1);
});

const resize = typeof ResizeObserver !== 'undefined' ? new ResizeObserver(resizeSlide) : null;
if (resize) resize.observe(frame);
window.addEventListener('resize', resizeSlide);
regenerate();

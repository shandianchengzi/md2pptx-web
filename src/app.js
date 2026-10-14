import {
  SLIDE_W, SLIDE_H, EXAMPLE,
  layoutMarkdown, MIN_BODY_FONT_PT, MAX_BODY_FONT_PT,
  DENSITY_FONT_PRESETS, resolveBodyFontPt, DEFAULT_MARGINS_CM, resolveMarginsCm,
  MIN_MARGIN_CM, MAX_MARGIN_CM
} from './pagination.mjs';
import { buildSlideTextBox } from './slide-text.mjs';

const $ = (id) => document.getElementById(id);
const editor = $('markdownInput');
const titleInput = $('documentTitle');
const density = $('density');
const bodyFontInput = $('bodyFontPt');
const marginInputs = {
  top: $('marginTop'), bottom: $('marginBottom'),
  left: $('marginLeft'), right: $('marginRight')
};
const footerTitleCheckbox = $('showFooterTitle');
const pageNumberCheckbox = $('showPageNumber');
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
const SETTINGS_KEY = 'md2pptx-web:settings:v1';
let bodyFontPt = DENSITY_FONT_PRESETS.comfortable;
let previousDensity = density.value;
let marginsCm = {...DEFAULT_MARGINS_CM};
let showFooterTitle = false;
let showPageNumber = false;

function updateMarginInputs() {
  for (const [key, input] of Object.entries(marginInputs)) input.value = marginsCm[key].toFixed(1);
}
function updateFooterInputs() {
  footerTitleCheckbox.checked = showFooterTitle;
  pageNumberCheckbox.checked = showPageNumber;
}
function onMarginInput(key, immediate = false) {
  const raw = marginInputs[key].value.trim();
  if (!raw) return;
  const parsed = Number(raw);
  if (!Number.isFinite(parsed)) return;
  if (!immediate && (parsed < MIN_MARGIN_CM || parsed > MAX_MARGIN_CM)) return;
  marginsCm = resolveMarginsCm({...marginsCm, [key]: parsed});
  if (immediate) updateMarginInputs();
  savePreferences();
  if (immediate) regenerate(); else scheduleRegenerate();
}
function onFooterChange() {
  showFooterTitle = footerTitleCheckbox.checked;
  showPageNumber = pageNumberCheckbox.checked;
  savePreferences();
  regenerate();
}

function restorePreferences() {
  try {
    const saved = JSON.parse(localStorage.getItem(SETTINGS_KEY));
    if (saved && (saved.density === 'comfortable' || saved.density === 'compact')) density.value = saved.density;
    bodyFontPt = resolveBodyFontPt(saved?.bodyFontPt, density.value);
    marginsCm = resolveMarginsCm(saved?.marginsCm);
    showFooterTitle = saved?.showFooterTitle === true;
    showPageNumber = saved?.showPageNumber === true;
  } catch {
    bodyFontPt = DENSITY_FONT_PRESETS[density.value];
    marginsCm = {...DEFAULT_MARGINS_CM};
    showFooterTitle = false;
    showPageNumber = false;
  }
  previousDensity = density.value;
  bodyFontInput.value = String(bodyFontPt);
  updateMarginInputs();
  updateFooterInputs();
}

function savePreferences() {
  try {
    localStorage.setItem(SETTINGS_KEY, JSON.stringify({ density: density.value, bodyFontPt, marginsCm, showFooterTitle, showPageNumber }));
  } catch {
    // The editor and downloads remain usable when storage is unavailable.
  }
}

function onFontInput() {
  const raw = bodyFontInput.value.trim();
  if (!/^\d+$/.test(raw)) return;
  const pt = Number(raw);
  if (pt < MIN_BODY_FONT_PT || pt > MAX_BODY_FONT_PT) return;
  bodyFontPt = pt;
  savePreferences();
  scheduleRegenerate();
}

function onFontChange() {
  const pt = Number(bodyFontInput.value);
  bodyFontPt = Number.isFinite(pt)
    ? Math.max(MIN_BODY_FONT_PT, Math.min(MAX_BODY_FONT_PT, Math.round(pt)))
    : bodyFontPt;
  bodyFontInput.value = String(bodyFontPt);
  savePreferences();
  regenerate();
}

function onDensityChange() {
  const priorPreset = DENSITY_FONT_PRESETS[previousDensity];
  // Switching a preset should not overwrite a manually customized size.
  if (bodyFontPt === priorPreset) bodyFontPt = DENSITY_FONT_PRESETS[density.value];
  previousDensity = density.value;
  bodyFontInput.value = String(bodyFontPt);
  savePreferences();
  regenerate();
}


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
  const divider = element('div', 'decor-line');
  divider.style.left = page.layout.left + 'px';
  divider.style.top = page.layout.titleDividerY + 'px';
  divider.style.width = page.layout.width + 'px';
  preview.appendChild(divider);
  for (const row of [...page.header, ...page.rows]) {
    const el = element('div', 'slide-line kind-' + row.kind);
    el.style.left = row.x + 'px';
    el.style.top = row.y + 'px';
    el.style.fontSize = row.size + 'px';
    el.style.fontWeight = row.bold ? '750' : '450';
    if (row.runs?.length) {
      for (const run of row.runs) {
        const span = element('span','',run.text);
        if (run.bold) span.style.fontWeight = '750';
        el.appendChild(span);
      }
    } else {
      el.textContent = row.text;
    }
    preview.appendChild(el);
  }
  if (page.layout.showFooterTitle || page.layout.showPageNumber) {
    const footer = element('div', 'slide-footer');
    footer.style.left = page.layout.left + 'px';
    footer.style.right = page.layout.right + 'px';
    footer.style.top = page.layout.footerLineY + 'px';
    footer.style.bottom = 'auto';
    if (page.layout.showFooterTitle) footer.appendChild(element('span','',deck.title || 'Presentation'));
    else footer.appendChild(element('span'));
    if (page.layout.showPageNumber) footer.appendChild(element('span','',String(activeSlide+1)+' / '+deck.pages.length));
    preview.appendChild(footer);
  }
  if (page.table) {
    const tab=page.table;
    const tableEl=element('div','slide-native-table');
    tableEl.style.left=tab.x+'px';
    tableEl.style.top=tab.y+'px';
    tableEl.style.width=tab.width+'px';
    tableEl.style.fontSize=tab.fontSize+'px';
    const template=tab.columns.map(w=>w+'px').join(' ');
    for(const row of tab.rows) {
      const gridRow=element('div','slide-table-row'+(row.header?' slide-table-head':''));
      gridRow.style.display='grid';
      gridRow.style.gridTemplateColumns=template;
      gridRow.style.height=row.height+'px';
      for(const cell of row.cells) {
        const cellEl=element('div','slide-table-cell');
        cellEl.style.textAlign=cell.align;
        if(cell.runs?.length) {
          for(const run of cell.runs) {
            const span=element('span','',run.text);
            if(run.bold) span.style.fontWeight='750';
            cellEl.appendChild(span);
          }
        } else cellEl.textContent=cell.text;
        gridRow.appendChild(cellEl);
      }
      tableEl.appendChild(gridRow);
    }
    preview.appendChild(tableEl);
  }
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
    inner.appendChild(element('div', 'thumb-copy', page.table ? page.table.rows.slice(0,2).flatMap(r=>r.cells.map(c=>c.text)).join(' · ') : page.rows.slice(0,4).map(x => x.text).join(' ')));
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
        density: density.value, bodyFontPt, marginsCm,
        showFooterTitle, showPageNumber
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

function sanitizeFileName(filename) {
  let clean = filename.replace(/[\\/:*?"<>|]+/g, '_').replace(/^\.+/, '').trim();
  if (clean.length > 72) clean = Array.from(clean).slice(0,72).join('');
  return clean || 'markdown-presentation';
}

async function exportPptx() {
  if (!deck.pages.length) return;
  const PptxConstructor = window.PptxGenJS || window.pptxgen;
  if (typeof PptxConstructor !== 'function') {
    message('PPTX 转换模块未加载。请刷新页面或检查网站构建结果。', true);
    return;
  }
  download.disabled = true;
  message('正在生成 PPTX，所有文字保持可编辑……');
  try {
    const pptx = new PptxConstructor();
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
        x: page.layout.left*UNIT, y: page.layout.titleDividerY*UNIT,
        w: page.layout.width*UNIT, h: 0,
        line: { color: 'E0E9EE', width: 1.1 }
      });
      const combined = buildSlideTextBox(page, deck.title, index, deck.pages.length, UNIT);
      slide.addText(combined.text, combined.box);
      if(page.table) {
        const t=page.table;
        const tabRows=t.rows.map(row=>row.cells.map(cell=>({
          text:cell.runs?.some(run=>run.bold)
            ? cell.runs.map(run=>({text:run.text,options:{bold:Boolean(run.bold)}}))
            : cell.text,
          options:{
            align:cell.align,
            bold:Boolean(row.header),
            color:row.header?'19354A':'233E52',
            fill:{color:row.header?'EAF2F6':'FFFFFF'}
          }
        })));
        slide.addTable(tabRows,{
          x:t.x*UNIT, y:t.y*UNIT,
          w:t.width*UNIT,
          colW:t.columns.map(w=>w*UNIT),
          rowH:t.rows.map(row=>row.height*UNIT),
          autoPage:false,
          border:{type:'solid',pt:.7,color:'D7E3EB'},
          fontFace:'Microsoft YaHei',
          fontSize:t.fontSize*.75,
          color:'233E52',
          margin:0.06,
          valign:'mid',
          breakLine:false
        });
      }
      if (page.layout.showFooterTitle || page.layout.showPageNumber) {
        slide.addShape(pptx.ShapeType.line, {
          x: page.layout.left*UNIT, y: page.layout.footerLineY*UNIT,
          w: page.layout.width*UNIT, h: 0,
          line: { color: 'E7EDF1', width: .7 }
        });
      }
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
density.addEventListener('change', onDensityChange);
bodyFontInput.addEventListener('input', onFontInput);
bodyFontInput.addEventListener('change', onFontChange);
for (const key of Object.keys(marginInputs)) {
  marginInputs[key].addEventListener('input', () => onMarginInput(key));
  marginInputs[key].addEventListener('change', () => onMarginInput(key, true));
}
footerTitleCheckbox.addEventListener('change', onFooterChange);
pageNumberCheckbox.addEventListener('change', onFooterChange);
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
restorePreferences();
regenerate();

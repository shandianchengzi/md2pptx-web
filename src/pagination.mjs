// Shared Markdown parser and browser/PPTX pagination model.
// Units are CSS pixels on a fixed 1280 × 720 slide.

export const SLIDE_W = 1280;
export const SLIDE_H = 720;
export const BODY_LEFT = 78;
export const BODY_WIDTH = 1124;
export const BOTTOM = 635;
export const SLIDE_WIDTH_CM = 33.8666666667;
export const PX_PER_CM = SLIDE_W / SLIDE_WIDTH_CM;
export const DEFAULT_MARGINS_CM = Object.freeze({ top: 1.5, bottom: 2.2, left: 2.1, right: 2.1 });
export const MIN_MARGIN_CM = 0.5;
export const MAX_MARGIN_CM = 5;
export function resolveMarginsCm(value = {}) {
  const output = {};
  for (const [key, fallback] of Object.entries(DEFAULT_MARGINS_CM)) {
    const raw = value && typeof value === 'object' ? value[key] : undefined;
    const number = raw === '' || raw === null || raw === undefined ? fallback : Number(raw);
    output[key] = Number.isFinite(number)
      ? Math.round(Math.max(MIN_MARGIN_CM, Math.min(MAX_MARGIN_CM, number)) * 10) / 10 : fallback;
  }
  return output;
}
export function sliceStyledRuns(runs, from, to) {
  const result = [];
  let cursor = 0;
  for (const part of runs) {
    const end = cursor + part.text.length;
    const beginLocal = Math.max(from, cursor);
    const endLocal = Math.min(to, end);
    if (endLocal > beginLocal) result.push({
      text: part.text.slice(beginLocal - cursor, endLocal - cursor), bold: Boolean(part.bold)
    });
    cursor = end;
  }
  return result;
}
// Preserve Markdown strong emphasis as rich text runs through line wrapping.
export function parseBoldRuns(raw) {
  const cleaned = String(raw)
    .replace(/!\[([^\]]*)\]\([^)]*\)/g, (_m, label) => '[Image: ' + label + ']')
    .replace(/\[([^\]]+)\]\([^)]*\)/g, '$1')
    .replace(/<https?:\/\/[^>]+>/g, m => m.slice(1,-1))
    .replace(/~~([^~]+)~~/g, '$1')
    .replace(/<\/?[^>]+>/g, '');
  function cleanPlain(value) {
    return value.replace(/(?<!\*)\*([^*]+)\*(?!\*)/g,'$1')
      .replace(/(?<!_)_([^_]+)_(?!_)/g,'$1')
      .replace(/`([^`]+)`/g,'$1');
  }
  const pieces = [];
  const pattern = /\*\*([\s\S]+?)\*\*|__([\s\S]+?)__/g;
  let i=0,m;
  while ((m=pattern.exec(cleaned))) {
    if(m.index>i) pieces.push({text:cleanPlain(cleaned.slice(i,m.index)),bold:false});
    pieces.push({text:cleanPlain(m[1] ?? m[2]),bold:true});
    i=pattern.lastIndex;
  }
  if(i<cleaned.length) pieces.push({text:cleanPlain(cleaned.slice(i)),bold:false});
  const all=pieces.map(x=>x.text).join('');
  const trimStart=all.length-all.trimStart().length;
  return sliceStyledRuns(pieces,trimStart,trimStart+all.trim().length);
}


export const EXAMPLE = `# AgentEmu 项目进度汇报

## 总体情况

AgentEmu 的四阶段核心流程已经具备实现基础，但距离设计中预期的可靠自动化外设模型生成系统，仍存在信息提取完整性、行为规则正确性、代码生成一致性以及运行时功能验证方面的缺口。

当前工作的重点已从建立模块转向通过真实参考手册和固件测试，修复缺陷并形成可复现的验证证据。

## Collector：信息提取

- 已经按六类目标信息重构提取流程。
- 仍需完善 PDF 表格解析、跨章节信息关联和缺失信息检查。
- 计划保留现有架构，优先修复真实失败案例并在更多手册上复测。

## Extractor 与 Coder

Extractor 已经能够构建符号表并提取 EP-DSL 规则，Coder 可以确定性地将规则编译为 QEMU C 代码。

但格式正确不等于语义完整。我们将通过独立的规则验证集和跨芯片回归测试，检查必要行为是否遗漏，以及生成代码是否忠实于规则。

## Validator 与后续工作

运行时已经可以使用 GDB 和 fuzzing 收集证据。下一步需要明确的输入—预期输出测试，随后完善错误归因、自动修复和回归执行。最终目标是形成可靠、可追溯、能持续改进的端到端流程。`;

export function stripInline(value) {
  return String(value)
    .replace(/!\[([^\]]*)\]\([^)]*\)/g, (_m, label) => '[Image: ' + label + ']')
    .replace(/\[([^\]]+)\]\([^)]*\)/g, '$1')
    .replace(/<https?:\/\/[^>]+>/g, (m) => m.slice(1, -1))
    .replace(/\*\*([\s\S]*?)\*\*/g, '$1')
    .replace(/__([\s\S]*?)__/g, '$1')
    .replace(/(?<!\*)\*([^*]+)\*(?!\*)/g, '$1')
    .replace(/(?<!_)_([^_]+)_(?!_)/g, '$1')
    .replace(/~~([^~]+)~~/g, '$1')
    .replace(/`([^`]+)`/g, '$1')
    .replace(/<\/?[^>]+>/g, '')
    .trim();
}

function normalizeParagraph(lines) {
  // A single Markdown newline is a soft wrap. Preserve a visible boundary
  // when an English word is immediately followed by Chinese text.
  return lines.reduce((out, line) => {
    if (!out) return line;
    const prev = out.slice(-1);
    const next = line[0] || '';
    const needSpace = /[A-Za-z0-9]$/.test(prev) || (
      /[\u4e00-\u9fff]/.test(prev) && /^[A-Za-z0-9]/.test(next)
    );
    return out + (needSpace ? ' ' : '') + line;
  }, '');
}


// GitHub-flavored Markdown pipe table detection, including escaped pipes.
export function splitTableCells(raw) {
  let line = String(raw).trim();
  if (line.startsWith('|')) line = line.slice(1);
  if (line.endsWith('|') && !line.endsWith('\\|')) line = line.slice(0,-1);
  const cells = [];
  let value = '', code = false;
  for (let i=0;i<line.length;i++) {
    const c = line[i];
    if (c==='\\' && line[i+1]==='|') { value+='|'; i++; continue; }
    if (c===String.fromCharCode(96)) code=!code;
    if (c==='|' && !code) { cells.push(value.trim()); value=''; }
    else value+=c;
  }
  cells.push(value.trim());
  return cells;
}

export function readMarkdownTable(lines, startIndex) {
  if (startIndex+1 >= lines.length || !lines[startIndex].includes('|')) return null;
  const header=splitTableCells(lines[startIndex]);
  const divider=splitTableCells(lines[startIndex+1]);
  if(header.length<2 || divider.length!==header.length ||
    !divider.every(c=>/^:?-{3,}:?$/.test(c))) return null;
  const aligns=divider.map(c=>c.startsWith(':')&&c.endsWith(':')?'center':
    c.endsWith(':')?'right':'left');
  const cell=(raw)=>{const runs=parseBoldRuns(raw);return {text:runs.map(x=>x.text).join(''),runs};};
  const headers=header.map(cell);
  const rows=[];
  let i=startIndex+2;
  while (i<lines.length && lines[i].trim() && lines[i].includes('|')) {
    const pieces=splitTableCells(lines[i]);
    // A new heading/divider ends this table.
    if (/^\s*#{1,6}\s/.test(lines[i]) || pieces.every(x=>/^:?-{3,}:?$/.test(x))) break;
    rows.push(Array.from({length:headers.length},(_,j)=>cell(pieces[j]??'')));
    i++;
  }
  return {block:{type:'table',headers,aligns,rows}, nextIndex:i};
}

function tableCellLines(cell,width,fontSize,measure) {
  const content=cell.text || '';
  const wrapped=wrapText(content,width-18,fontSize,measure);
  let cursor=0;
  return wrapped.map(text=>{
    let at=content.indexOf(text,cursor);
    if(at<0) at=cursor;
    cursor=at+text.length;
    return {text,runs:sliceStyledRuns(cell.runs || [{text:content,bold:false}],at,at+text.length)};
  });
}

export function planMarkdownTable(table,layout,fontPx,measure=defaultMeasure) {
  const count=table.headers.length;
  const fontSize=Math.max(13,Math.min(fontPx*.84,24,layout.width/(count*6.0)));
  const lineHeight=Math.ceil(fontSize*1.42);
  const weights=Array.from({length:count},(_,j)=>{
    const examples=[table.headers[j],...table.rows.slice(0,32).map(row=>row[j])];
    return Math.max(5,Math.min(26,Math.max(...examples.map(c=>Array.from(c.text).length),5)));
  });
  // Avoid overly narrow columns but allow wider columns for verbose values.
  let colWidths=weights.map(w=>Math.sqrt(w));
  const total=colWidths.reduce((a,b)=>a+b,0);
  colWidths=colWidths.map(w=>layout.width*w/total);
  const minWidth=layout.width / count * .64;
  for(let n=0;n<3;n++){
    const low=colWidths.map(w=>w<minWidth);
    const shortage=colWidths.reduce((sum,w,j)=>sum+(low[j]?minWidth-w:0),0);
    const free=colWidths.reduce((sum,w,j)=>sum+(!low[j]?w-minWidth:0),0);
    if(!shortage||!free)break;
    colWidths=colWidths.map((w,j)=>low[j]?minWidth:w-shortage*(w-minWidth)/free);
  }
  const prepare=row=>row.map((cell,j)=>tableCellLines(cell,colWidths[j],fontSize,measure));
  const preparedHeader=prepare(table.headers);
  const headerLines=Math.max(1,...preparedHeader.map(c=>c.length));
  const headerHeight=Math.max(44,headerLines*lineHeight+14);
  const preparedRows=table.rows.map(prepare);
  return {fontSize,lineHeight,colWidths,preparedHeader,headerHeight,preparedRows,aligns:table.aligns};
}

function makeTableRow(prepared,aligns,offset,take,lineHeight,header=false) {
  const cells=prepared.map((cell,i)=>{
    const lines=cell.slice(offset,offset+take);
    const runs=[];
    lines.forEach((line,j)=>{
      if(j) runs.push({text:'\n',bold:false});
      runs.push(...line.runs);
    });
    return {text:lines.map(x=>x.text).join('\n'),runs,align:aligns[i]};
  });
  const shownLines=Math.max(1,...cells.map(c=>c.text.split('\n').length));
  return {header,cells,height:Math.max(44,shownLines*lineHeight+14)};
}

export function parseMarkdown(markdown, overrideTitle = '') {
  const lines = String(markdown || '').replace(/\r\n?/g, '\n').split('\n');
  const headings = lines.map((line, index) => {
    const m = /^(#{1,6})\s+(.+?)\s*#*\s*$/.exec(line);
    return m ? { index, level: m[1].length, title: stripInline(m[2]) } : null;
  }).filter(Boolean);
  const first = headings[0];
  const treatFirstAsDocumentTitle = Boolean(first && first.level === 1 && first.index <= 2 &&
    headings.some(h => h.index > first.index && h.level <= 2));
  const documentTitle = String(overrideTitle || '').trim() ||
    (treatFirstAsDocumentTitle ? first.title : 'Markdown Presentation');

  const sections = [];
  let current = { title: documentTitle === 'Markdown Presentation' ? '内容' : documentTitle, blocks: [] };
  let paragraphLines = [];
  let inCode = false;
  function flush() {
    if (!paragraphLines.length) return;
    const runs = parseBoldRuns(normalizeParagraph(paragraphLines));
    const text = runs.map(r=>r.text).join('');
    if (text) current.blocks.push({ type: 'paragraph', text, runs });
    paragraphLines = [];
  }
  function nextSection(title) {
    flush();
    if (current.blocks.length) sections.push(current);
    current = { title: title || '内容', blocks: [] };
  }

  for (let index=0; index<lines.length; index++) {
    const raw=lines[index];
    const line = raw.trim();
    if (/^(```|~~~)/.test(line)) {
      flush();
      inCode = !inCode;
      continue;
    }
    if (inCode) {
      current.blocks.push({ type: 'code', text: raw || ' ' });
      continue;
    }
    const h = /^(#{1,6})\s+(.+?)\s*#*\s*$/.exec(line);
    if (h) {
      if (treatFirstAsDocumentTitle && first.index === index) return;
      if (h[1].length <= 2) nextSection(stripInline(h[2]));
      else {
        flush();
        current.blocks.push({ type: 'subheading', text: stripInline(h[2]) });
      }
      continue;
    }
    if (/^(---+|\*\*\*+|___+)$/.test(line)) {
      flush();
      current.blocks.push({ type: 'break' });
      continue;
    }
    if (!line) { flush(); continue; }
    const table = readMarkdownTable(lines,index);
    if (table) {
      flush();
      current.blocks.push(table.block);
      index=table.nextIndex-1;
      continue;
    }
    const li = /^([-+*]|\d+[.)])\s+(.+)$/.exec(line);
    if (li) {
      flush();
      const marker = /^\d/.test(li[1]) ? li[1] : '•';
      const runs = parseBoldRuns(marker + '  ' + li[2]);
      current.blocks.push({ type: 'list', text: runs.map(r=>r.text).join(''), runs });
      continue;
    }
    if (/^>\s?/.test(line)) {
      flush();
      const runs = parseBoldRuns('❝  ' + line.replace(/^>\s?/, ''));
      current.blocks.push({ type: 'quote', text: runs.map(r=>r.text).join(''), runs });
      continue;
    }
    paragraphLines.push(line);
  }
  flush();
  if (current.blocks.length) sections.push(current);
  if (!sections.length) sections.push({ title: documentTitle, blocks: [{ type: 'paragraph', text: '粘贴 Markdown 或普通文本后，幻灯片会自动出现在这里。' }] });
  return { documentTitle, sections };
}

export function defaultMeasure(text, px) {
  return Array.from(text).reduce((total, char) => total + px * (
    /[\u3000-\u9FFF\uF900-\uFAFF]/.test(char) ? 1 :
    /[ilI1|.,:;'\s]/.test(char) ? .3 : .56), 0);
}

// Browser and Node support Intl.Segmenter: split Han words at readable
// boundaries while retaining complete Latin words and punctuation.
// Unlike splitting on whitespace, a long Chinese phrase no longer pushes
// a short leading English word onto its own line.
const wordSegmenter = typeof Intl !== 'undefined' && typeof Intl.Segmenter === 'function'
  ? new Intl.Segmenter('zh-CN', { granularity: 'word' })
  : null;

function wordPieces(text) {
  if (wordSegmenter) return Array.from(wordSegmenter.segment(text), x => x.segment);
  return text.match(/[A-Za-z0-9_-]+|[\u4e00-\u9fff]|\s+|./gu) || [];
}

export function wrapText(text, width, fontSize, measure = defaultMeasure) {
  const result = [];
  const availableWidth = Math.max(fontSize, width - fontSize * 0.35);
  for (const paragraph of String(text).split('\n')) {
    if (!paragraph) { result.push(' '); continue; }
    let line = '';
    function flush() {
      if (line.trim()) result.push(line.trimEnd());
      line = '';
    }
    for (const piece of wordPieces(paragraph)) {
      if (!line && !piece.trim()) continue;
      const candidate = line + piece;
      if (measure(candidate, fontSize) <= availableWidth) {
        line = candidate;
        continue;
      }
      // A whitespace token can itself overflow the line; flush the line
      // before dropping that whitespace, otherwise the existing text is lost.
      if (!piece.trim()) {
        flush();
        continue;
      }
      if (line.trim()) flush();
      if (measure(piece.trimStart(), fontSize) <= availableWidth) {
        line = piece.trimStart();
        continue;
      }
      // Extremely long URL / non-breaking word: split only as fallback.
      for (const char of Array.from(piece.trimStart())) {
        if (line && measure(line + char, fontSize) > availableWidth) flush();
        line += char;
      }
    }
    flush();
  }
  return result.length ? result : [''];
}

export const MIN_BODY_FONT_PT = 14;
export const MAX_BODY_FONT_PT = 36;
export const DENSITY_FONT_PRESETS = Object.freeze({comfortable: 22, compact: 19});

export function resolveBodyFontPt(value, density = 'comfortable') {
  const preset = DENSITY_FONT_PRESETS[density] ?? DENSITY_FONT_PRESETS.comfortable;
  if (value === undefined || value === null || value === '') return preset;
  const pt = Number(value);
  return Number.isInteger(pt) && pt >= MIN_BODY_FONT_PT && pt <= MAX_BODY_FONT_PT ? pt : preset;
}

export function layoutMarkdown(markdown, options = {}, measure = defaultMeasure) {
  const parsed = parseMarkdown(markdown, options.title);
  const compact = options.density === 'compact';
  const bodyFontPt = resolveBodyFontPt(options.bodyFontPt, options.density);
  const BODY_FONT = Math.round(bodyFontPt * (4 / 3) * 10) / 10;
  const LINE_HEIGHT = Math.ceil(BODY_FONT * (compact ? 1.48 : 1.52));
  const GAP = Math.round(BODY_FONT * (compact ? 0.48 : 0.59));
  const pages = [];
  const marginsCm = resolveMarginsCm(options.marginsCm);
  const left = marginsCm.left * PX_PER_CM;
  const right = marginsCm.right * PX_PER_CM;
  const top = marginsCm.top * PX_PER_CM;
  const bottom = marginsCm.bottom * PX_PER_CM;
  const width = SLIDE_W - left - right;
  const showFooterTitle = options.showFooterTitle === true;
  const showPageNumber = options.showPageNumber === true;
  const showFooter = showFooterTitle || showPageNumber;
  const footerLineY = SLIDE_H - Math.max(38, bottom * .65);
  const contentBottom = SLIDE_H - bottom - (showFooter ? 34 : 0);
  const layout = {marginsCm, left, right, top, bottom, width, contentBottom,
    footerLineY, footerTextY: footerLineY + 7, showFooterTitle, showPageNumber};

  function titleRows(title, continued) {
    const prefix = continued ? '（续）' : '';
    const t = title + prefix;
    const lines = wrapText(t, width - 15, 38, measure).slice(0,2);
    return lines.map((text,i) => ({
      text, x: left, y: top + i*48, width, height: 50,
      size: 38, bold: true, kind: 'title'
    }));
  }

  function makePage(title, continued = false) {
    const header = titleRows(title, continued);
    const page = {
      title, continued, header, rows: [], table: null,
      layout: {...layout, titleDividerY: top + (header.length === 1 ? 77 : 122)},
      y: top + (header.length === 1 ? 111 : 148)
    };
    pages.push(page);
    return page;
  }

  for (const section of parsed.sections) {
    if (!section.blocks.length) continue;
    let page = makePage(section.title);
    for (const block of section.blocks) {
      if (block.type === 'table') {
        if (page.rows.length || page.table) page=makePage(section.title,true);
        const plan=planMarkdownTable(block,layout,BODY_FONT,measure);
        const reserve=16;
        let currentTable=null;
        function startTablePage(newPage=false) {
          if(newPage) page=makePage(section.title,true);
          const y=page.y+4;
          if(y+plan.headerHeight+44 > contentBottom) {
            throw new Error('Table cannot fit on a slide with these font and margin settings');
          }
          currentTable={x:left,y,width,columns:plan.colWidths,fontSize:plan.fontSize,
            lineHeight:plan.lineHeight,rows:[
              makeTableRow(plan.preparedHeader,plan.aligns,0,Infinity,plan.lineHeight,true)
            ]};
          page.table=currentTable;
          page.y=y+plan.headerHeight;
        }
        startTablePage();
        for (const prepared of plan.preparedRows) {
          const totalLines=Math.max(1,...prepared.map(c=>c.length));
          let offset=0;
          while(offset<totalLines) {
            const free=contentBottom-page.y-reserve;
            const lines=Math.floor((free-14)/plan.lineHeight);
            if (lines<1 || free<44) { startTablePage(true); continue; }
            const take=Math.min(totalLines-offset,lines);
            const fragment=makeTableRow(prepared,plan.aligns,offset,take,plan.lineHeight);
            if(fragment.height>free+0.001) {startTablePage(true);continue;}
            currentTable.rows.push(fragment);
            page.y+=fragment.height;
            offset+=take;
          }
        }
        continue;
      }
      if(page.table) page=makePage(section.title,true);
      if (block.type === 'break') {
        if (page.rows.length) page = makePage(section.title, true);
        continue;
      }
      const sub = block.type === 'subheading';
      const code = block.type === 'code';
      const styleSize = sub ? Math.max(30, Math.round(BODY_FONT * 1.07))
        : code ? Math.max(18, Math.round(BODY_FONT * .8)) : BODY_FONT;
      const height = sub ? Math.ceil(styleSize * 1.5)
        : code ? Math.ceil(styleSize * 1.5) : LINE_HEIGHT;
      const rowLeft = left + (block.type === 'quote' ? 21 : code ? 18 : 0);
      const rowWidth = width - (rowLeft - left) - 9;
      const wrapped = wrapText(block.text, rowWidth, styleSize, measure);
      let textCursor = 0;
      const styledLines = wrapped.map(line => {
        if (!block.runs) return undefined;
        let at = block.text.indexOf(line, textCursor);
        if (at === -1) at = textCursor;
        textCursor = at + line.length;
        return sliceStyledRuns(block.runs, at, at + line.length);
      });
      let offset = 0;
      while (offset < wrapped.length) {
        let capacity = Math.floor((contentBottom - page.y) / height);
        const remaining = wrapped.length - offset;
        // Avoid leaving a single line of a long paragraph on either slide.
        if (capacity <= 0 || (capacity === 1 && remaining > 1 && page.rows.length)) {
          page = makePage(section.title, true);
          continue;
        }
        capacity = Math.min(capacity, remaining);
        if (remaining - capacity === 1 && capacity > 1) capacity--;
        for (let i = 0; i < capacity; i++) {
          page.rows.push({
            x: rowLeft, y: page.y, width: rowWidth,
            text: wrapped[offset + i], runs: styledLines[offset + i], height,
            size: styleSize, bold: sub, kind: block.type
          });
          page.y += height;
        }
        offset += capacity;
        if (offset < wrapped.length) page = makePage(section.title, true);
      }
      // Preserve a visible separation between paragraphs without wasting a whole line.
      if (page.y + GAP + LINE_HEIGHT <= contentBottom) page.y += sub ? 6 : GAP;
    }
  }
  return { title: parsed.documentTitle, pages, layout, blockCount: parsed.sections.reduce((n,s)=>n+s.blocks.length,0) };
}

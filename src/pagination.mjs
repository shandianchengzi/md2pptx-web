// Shared Markdown parser and browser/PPTX pagination model.
// Units are CSS pixels on a fixed 1280 × 720 slide.

export const SLIDE_W = 1280;
export const SLIDE_H = 720;
export const BODY_LEFT = 78;
export const BODY_WIDTH = 1124;
export const BOTTOM = 635;

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
    const text = normalizeParagraph(paragraphLines).trim();
    if (text) current.blocks.push({ type: 'paragraph', text });
    paragraphLines = [];
  }
  function nextSection(title) {
    flush();
    if (current.blocks.length) sections.push(current);
    current = { title: title || '内容', blocks: [] };
  }

  lines.forEach((raw, index) => {
    const line = raw.trim();
    if (/^(```|~~~)/.test(line)) {
      flush();
      inCode = !inCode;
      return;
    }
    if (inCode) {
      current.blocks.push({ type: 'code', text: raw || ' ' });
      return;
    }
    const h = /^(#{1,6})\s+(.+?)\s*#*\s*$/.exec(line);
    if (h) {
      if (treatFirstAsDocumentTitle && first.index === index) return;
      if (h[1].length <= 2) nextSection(stripInline(h[2]));
      else {
        flush();
        current.blocks.push({ type: 'subheading', text: stripInline(h[2]) });
      }
      return;
    }
    if (/^(---+|\*\*\*+|___+)$/.test(line)) {
      flush();
      current.blocks.push({ type: 'break' });
      return;
    }
    if (!line) { flush(); return; }
    const li = /^([-+*]|\d+[.)])\s+(.+)$/.exec(line);
    if (li) {
      flush();
      const marker = /^\d/.test(li[1]) ? li[1] : '•';
      current.blocks.push({ type: 'list', text: marker + '  ' + stripInline(li[2]) });
      return;
    }
    if (/^>\s?/.test(line)) {
      flush();
      current.blocks.push({ type: 'quote', text: '❝  ' + stripInline(line.replace(/^>\s?/, '')) });
      return;
    }
    // Textual fallback for Markdown tables: retain cell contents.
    if (/^\|.+\|$/.test(line)) {
      flush();
      if (!/^\|[\s:|-]+\|$/.test(line)) {
        const cells = line.slice(1,-1).split('|').map(x => stripInline(x)).filter(Boolean);
        if (cells.length) current.blocks.push({ type: 'paragraph', text: cells.join('  ·  ') });
      }
      return;
    }
    paragraphLines.push(stripInline(line));
  });
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
      if (piece.trim() && line.trim()) flush();
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

export function layoutMarkdown(markdown, options = {}, measure = defaultMeasure) {
  const parsed = parseMarkdown(markdown, options.title);
  const compact = options.density === 'compact';
  const BODY_FONT = compact ? 25 : 29;
  const LINE_HEIGHT = compact ? 37 : 44;
  const GAP = compact ? 12 : 17;
  const pages = [];

  function titleRows(title, continued) {
    const prefix = continued ? '（续）' : '';
    const t = title + prefix;
    const lines = wrapText(t, 1085, 38, measure).slice(0,2);
    return lines.map((text,i) => ({
      text, x: BODY_LEFT, y: 55 + i*48, height: 50,
      size: 38, bold: true, kind: 'title'
    }));
  }

  function makePage(title, continued = false) {
    const header = titleRows(title, continued);
    const page = {
      title, continued, header, rows: [],
      y: header.length === 1 ? 166 : 203
    };
    pages.push(page);
    return page;
  }

  for (const section of parsed.sections) {
    if (!section.blocks.length) continue;
    let page = makePage(section.title);
    for (const block of section.blocks) {
      if (block.type === 'break') {
        if (page.rows.length) page = makePage(section.title, true);
        continue;
      }
      const sub = block.type === 'subheading';
      const code = block.type === 'code';
      const styleSize = sub ? 30 : code ? (compact ? 20 : 23) : BODY_FONT;
      const height = sub ? 45 : code ? (compact ? 30 : 35) : LINE_HEIGHT;
      const left = BODY_LEFT + (block.type === 'quote' ? 21 : code ? 18 : 0);
      const width = BODY_WIDTH - (left - BODY_LEFT) - 9;
      const wrapped = wrapText(block.text, width, styleSize, measure);
      let offset = 0;
      while (offset < wrapped.length) {
        let capacity = Math.floor((BOTTOM - page.y) / height);
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
            x: left, y: page.y, text: wrapped[offset + i], height,
            size: styleSize, bold: sub, kind: block.type
          });
          page.y += height;
        }
        offset += capacity;
        if (offset < wrapped.length) page = makePage(section.title, true);
      }
      // Preserve a visible separation between paragraphs without wasting a whole line.
      if (page.y + GAP + LINE_HEIGHT <= BOTTOM) page.y += sub ? 6 : GAP;
    }
  }
  return { title: parsed.documentTitle, pages, blockCount: parsed.sections.reduce((n,s)=>n+s.blocks.length,0) };
}

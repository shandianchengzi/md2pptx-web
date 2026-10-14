// Normalize OOXML defects emitted by the pinned PptxGenJS browser bundle.
// In particular, mixed text/graphic/table slides can reuse shape IDs, and
// rich text may include multiple paragraph-properties elements in one a:p.
// PowerPoint repairs or rejects such slides even though LibreOffice opens them.
const P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';

function nodeChildren(parent, localName, ns=A) {
  return Array.from(parent?.childNodes ?? []).filter(node =>
    node.nodeType === 1 && node.localName === localName && node.namespaceURI === ns);
}

export function normalizeSlideXml(source, browserDom = globalThis) {
  const parser = new browserDom.DOMParser();
  const xml = parser.parseFromString(source,'application/xml');
  if(xml.getElementsByTagName('parsererror').length) throw Error('Invalid slide XML before normalization');
  let reindexed = 0, splitParagraphs = 0, resizedTables = 0;
  const shapeIds = Array.from(xml.getElementsByTagNameNS(P,'cNvPr'));
  const used = new Set();
  let freshId = Math.max(1,...shapeIds.map(x=>Number(x.getAttribute('id'))||1))+1;
  for(const node of shapeIds) {
    const id = Number(node.getAttribute('id')) || 0;
    if (!id || used.has(id)) {
      while (used.has(freshId)) freshId++;
      node.setAttribute('id',String(freshId));
      used.add(freshId++);
      reindexed++;
    } else used.add(id);
  }
  // Only direct a:pPr child nodes are relevant; a:p can also contain a:r/a:br.
  for(const paragraph of Array.from(xml.getElementsByTagNameNS(A,'p'))) {
    const properties = nodeChildren(paragraph,'pPr');
    if(properties.length<2) continue;
    const parent=paragraph.parentNode;
    let current=paragraph;
    let afterContent=false;
    for(const child of Array.from(paragraph.childNodes)) {
      if (child.nodeType===1 && child.localName==='pPr' && child.namespaceURI===A && afterContent) {
        const next=xml.createElementNS(A,'a:p');
        parent.insertBefore(next,current.nextSibling);
        current=next;
        splitParagraphs++;
      }
      if(current!==paragraph) current.appendChild(child);
      if(child.nodeType===1 && child.localName!=='pPr') afterContent=true;
      if(child.nodeType===1 && child.localName==='pPr') afterContent=false;
    }
  }
  // PptxGenJS 4.0 sometimes sets a 1-inch graphicFrame extent regardless of
  // the table's explicit row heights. Align extent with the real row layout.
  for(const frame of Array.from(xml.getElementsByTagNameNS(P,'graphicFrame'))) {
    const tbl=frame.getElementsByTagNameNS(A,'tbl')[0];
    const xfrm=nodeChildren(frame,'xfrm',P)[0];
    const ext=nodeChildren(xfrm,'ext')[0];
    if(!tbl || !ext) continue;
    const rowHeights=Array.from(tbl.getElementsByTagNameNS(A,'tr'))
      .map(row=>Number(row.getAttribute('h')||0));
    const total=rowHeights.reduce((a,b)=>a+b,0);
    if(total>0 && total!==Number(ext.getAttribute('cy'))) {
      ext.setAttribute('cy',String(total));
      resizedTables++;
    }
  }
  const output = new browserDom.XMLSerializer().serializeToString(xml);
  return {xml:output, stats:{reindexed,splitParagraphs,resizedTables}};
}

export function checkSlideXml(source,browserDom=globalThis) {
  const root=new browserDom.DOMParser().parseFromString(source,'application/xml');
  if(root.getElementsByTagName('parsererror').length)throw Error('Slide has invalid XML');
  const ids=Array.from(root.getElementsByTagNameNS(P,'cNvPr'))
    .map(node=>node.getAttribute('id'));
  if(ids.length!==new Set(ids).size) throw Error('Duplicate PowerPoint shape ID');
  for(const paragraph of Array.from(root.getElementsByTagNameNS(A,'p'))) {
    const properties=nodeChildren(paragraph,'pPr');
    if(properties.length>1)throw Error('Duplicate a:pPr in PowerPoint paragraph');
    const children=Array.from(paragraph.children);
    if(properties.length===1 && children[0]!==properties[0])
      throw Error('Misplaced a:pPr in PowerPoint paragraph');
  }
  return true;
}

export async function sanitizePptx(blob,JSZipCtor,browserDom=globalThis) {
  if(!JSZipCtor) throw Error('JSZip is unavailable; cannot safely export PPTX');
  const zip=await JSZipCtor.loadAsync(blob);
  const entries=Object.keys(zip.files).filter(name=>/^ppt\/slides\/slide\d+\.xml$/.test(name));
  if(!entries.length)throw Error('PPTX has no slides');
  const totals={reindexed:0,splitParagraphs:0,resizedTables:0};
  for(const path of entries) {
    const before=await zip.file(path).async('string');
    const normalized=normalizeSlideXml(before,browserDom);
    checkSlideXml(normalized.xml,browserDom);
    for(const key of Object.keys(totals))totals[key]+=normalized.stats[key];
    zip.file(path,normalized.xml);
  }
  const clean=await zip.generateAsync({type:'blob',compression:'DEFLATE'});
  return {blob:clean,stats:totals,slides:entries.length};
}

export function triggerDownload(blob,filename) {
  const address=URL.createObjectURL(blob);
  const link=document.createElement('a');
  link.href=address;
  link.download=filename;
  document.body.append(link);
  link.click();
  link.remove();
  setTimeout(()=>URL.revokeObjectURL(address),60000);
}

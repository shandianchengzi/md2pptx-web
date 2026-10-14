// Offline MathJax SVG renderer. PowerPoint export uses high-DPI PNGs of the
// MathJax result (not SVG media parts) for maximum Office compatibility.
// Formula images are intentionally separate from the one editable text box.
let readyPromise;
let engineState='loading';
let engineError=null;
export function mathStatus() { return {state:engineState,error:engineError}; }
export function mathReady() {
  if(!readyPromise) {
    readyPromise=(async()=>{
      const m=window.MathJax;
      if(!m || !m.startup) throw Error('MathJax library is unavailable');
      await m.startup.promise;
      if(typeof m.tex2svg!=='function')throw Error('MathJax TeX renderer unavailable');
      engineState='ready';
      engineError=null;
      return m;
    })().catch(err=>{
      engineState='error';
      engineError=err?.message || String(err);
      readyPromise=null;
      throw err;
    });
  }
  return readyPromise;
}

function mathSvg(tex,display=false) {
  if(engineState!=='ready' || !window.MathJax || typeof window.MathJax.tex2svg!=='function')return null;
  const container=window.MathJax.tex2svg(String(tex),{display,em:16,ex:8,containerWidth:1200});
  return container.querySelector('svg');
}
function aspect(svg) {
  const vb=(svg?.getAttribute('viewBox')||'0 0 1 1').trim().split(/[\s,]+/).map(Number);
  const x=vb.length===4?vb[2]:1,y=vb.length===4?vb[3]:1;
  return x>0&&y>0?x/y:1;
}
export function mathDimensions(tex,maxWidth,maxHeight,display=false) {
  const svg=mathSvg(tex,display);
  if(!svg)return {width:Math.min(maxWidth,maxHeight*2),height:maxHeight,svg:null};
  const ratio=aspect(svg);
  let width=Math.min(maxWidth,maxHeight*ratio);
  let height=Math.min(maxHeight,maxWidth/ratio);
  return {width,height,svg};
}

export function previewMath(tex,box,display=false) {
  const node=document.createElement('div');
  node.className='rendered-math'+(display?' display-math':' inline-math');
  node.style.display=display?'flex':'inline-flex';
  node.style.justifyContent='center';
  node.style.alignItems='center';
  node.style.width=box.width+'px';
  node.style.height=box.height+'px';
  node.dataset.tex=tex;
  const measure=mathDimensions(tex,box.width,box.height,display);
  if(measure.svg) {
    const svg=measure.svg.cloneNode(true);
    svg.removeAttribute('focusable');
    svg.setAttribute('width',String(measure.width));
    svg.setAttribute('height',String(measure.height));
    svg.style.width=measure.width+'px';
    svg.style.height=measure.height+'px';
    node.append(svg);
    node.title=tex;
  } else {
    node.textContent=engineState==='error' ? '公式渲染失败' : '公式加载中…';
    node.classList.add(engineState==='error' ? 'math-error':'math-loading');
    node.title=engineState==='error' ? String(engineError) : 'MathJax is initializing';
  }
  return node;
}

export async function mathPng(tex,maxWidth,maxHeight,display=false) {
  await mathReady();
  const measured=mathDimensions(tex,maxWidth,maxHeight,display);
  if(!measured.svg)throw Error('Math formula could not be rendered: '+tex);
  const ratio=3;
  const pxWidth=Math.max(1,Math.round(measured.width*ratio));
  const pxHeight=Math.max(1,Math.round(measured.height*ratio));
  const clone=measured.svg.cloneNode(true);
  clone.setAttribute('xmlns','http://www.w3.org/2000/svg');
  clone.setAttribute('width',pxWidth+'px');
  clone.setAttribute('height',pxHeight+'px');
  // A contained SVG built from MathJax path glyphs has no webfont dependency.
  const serial=new XMLSerializer().serializeToString(clone);
  const url=URL.createObjectURL(new Blob([serial],{type:'image/svg+xml;charset=utf-8'}));
  let img;
  try {
    img=await new Promise((resolve,reject)=>{
      const image=new Image();
      image.onload=()=>resolve(image);
      image.onerror=()=>reject(Error('Unable to rasterize math formula'));
      image.src=url;
    });
  } finally {
    URL.revokeObjectURL(url);
  }
  const canvas=document.createElement('canvas');
  canvas.width=pxWidth;
  canvas.height=pxHeight;
  const ctx=canvas.getContext('2d');
  if(!ctx)throw Error('Canvas rendering context unavailable');
  ctx.clearRect(0,0,pxWidth,pxHeight);
  ctx.drawImage(img,0,0,pxWidth,pxHeight);
  return {data:canvas.toDataURL('image/png'),width:measured.width,height:measured.height};
}

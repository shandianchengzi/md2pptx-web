// A single PowerPoint text box per slide containing rich editable paragraphs.
const PX_TO_PT = .75;
export function buildSlideTextBox(page,documentTitle,slideIndex,slideCount,unit) {
  const rows=[...page.header,...page.rows];
  if(!rows.length) throw new Error('Cannot export an empty slide');
  const layout=page.layout;
  const footer=layout.showFooterTitle || layout.showPageNumber;
  const runs=[];
  const colors={title:'19354A',subheading:'205C79',quote:'587487',code:'355B72'};
  for(let i=0;i<rows.length;i++){
    const row=rows[i];
    const next=rows[i+1];
    const nextY=next?next.y:(footer?layout.footerTextY:null);
    const gap=nextY===null?0:Math.max(0,nextY-(row.y+row.height));
    const parts=row.runs?.length?row.runs:[{text:row.text,bold:false}];
    parts.forEach((part,j)=>{
      const lastPart=j===parts.length-1;
      const lastRow=i===rows.length-1;
      const opts={
        fontFace:row.kind==='code'?'Consolas':'Microsoft YaHei',
        fontSize:row.size*PX_TO_PT,
        bold:Boolean(row.bold||part.bold),
        color:colors[row.kind]||'233E52',
        lineSpacing:row.height*PX_TO_PT,
        ...(lastPart&&gap>0?{paraSpaceAfter:gap*PX_TO_PT}:{}),
        ...(lastPart&&(!lastRow||footer)?{breakLine:true}:{})
      };
      runs.push({text:part.mathTex ? '\u3000'.repeat(part.text.length) : part.text,
        options:part.mathTex?{...opts,color:'FBFCFE',bold:false}:opts});
    });
  }
  if(footer){
    const title=layout.showFooterTitle?documentTitle:'';
    const pageNum=layout.showPageNumber?(String(slideIndex+1)+' / '+String(slideCount)):'';
    const text=title&&pageNum?(title+'\t'+pageNum):(title||pageNum);
    runs.push({text,options:{fontFace:'Microsoft YaHei',fontSize:10,color:'8DA2AF',lineSpacing:16,
      ...(pageNum&&!title?{align:'right'}:{}),
      ...(pageNum&&title?{tabStops:[{position:Math.max(.5,layout.width*unit-.12),alignment:'r'}]}:{})
    }});
  }
  const top=rows[0].y;
  return {text:runs,box:{x:layout.left*unit,y:top*unit,w:layout.width*unit,
    h:(720-top-6)*unit,margin:0,valign:'top',fontFace:'Microsoft YaHei',breakLine:false,
    fit:'none',color:'233E52'}};
}

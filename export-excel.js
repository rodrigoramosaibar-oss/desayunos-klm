/* ============================================================
   Informe Excel corporativo · Escandallo de desayunos
   Se carga bajo demanda desde index.html (window.KDExport.build)
   Usa ExcelJS (estilos, formatos condicionales, imágenes) y dibuja los
   gráficos con canvas → PNG incrustado (ExcelJS no crea gráficos nativos).
============================================================ */
(function(root){
'use strict';
const C={navy:'0F2747',navy2:'1D3A63',gold:'B8934F',goldL:'F3ECDC',teal:'2F7D6B',red:'C2563A',bg:'F4F6FA',line:'D9DEE7',txt:'17202E',mut:'6B7686',white:'FFFFFF',ok:'1F8A5B',bad:'C0392B'};
let _pr=0;const P=()=>++_pr;
const FAM=['0F2747','B8934F','2F7D6B','C2563A','5B6FB5','8A5A9E','3A9BB5','9AA33B','9AA3B0'];
const FONT='Calibri';
const argb=h=>'FF'+h;
const fill=h=>({type:'pattern',pattern:'solid',fgColor:{argb:argb(h)}});
const thin=h=>({style:'thin',color:{argb:argb(h)}});
const EUR='#,##0.00" €"',EUR0='#,##0" €"',INT='#,##0',PCT='0.0%',QTY='#,##0.###',DTF='ddd dd/mm/yyyy';
const nf=(n,d=2)=>new Intl.NumberFormat('es-ES',{minimumFractionDigits:d,maximumFractionDigits:d}).format(+n||0);
const q3=n=>new Intl.NumberFormat('es-ES',{maximumFractionDigits:3}).format(+n||0);
const ddmm=f=>f.slice(5).split('-').reverse().join('/');
const ddmmyyyy=f=>f.split('-').reverse().join('/');
const dt=f=>{const[y,m,d]=f.split('-').map(Number);return new Date(Date.UTC(y,m-1,d));};
const WD=['do','lu','ma','mi','ju','vi','sá'];
const wd=f=>WD[dt(f).getUTCDay()];
const px=w=>Math.trunc(w*7+5);               // ancho de columna (caracteres) → píxeles
const ptpx=p=>p*4/3;
const UNI={BTL:['botella','botellas'],KG:['kilo','kilos'],LTR:['litro','litros'],L:['litro','litros'],UT:['unidad','unidades'],EA:['unidad','unidades'],CS:['caja','cajas'],PK:['pack','packs'],CAN:['bote','botes'],BAG:['bolsa','bolsas'],BKT:['cubo','cubos'],SK:['saco','sacos'],RL:['rollo','rollos'],TRY:['bandeja','bandejas'],DZ:['docena','docenas']};
const uk=u=>{const k=String(u||'').trim().toUpperCase();return k==='EA'?'UT':k==='L'?'LTR':k;};
const uname=(u,n)=>{const s=UNI[String(u||'').toUpperCase()];return s?(Math.abs(+n)===1?s[0]:s[1]):String(u||'').toLowerCase();};
const cap=s=>s.charAt(0).toUpperCase()+s.slice(1);
const dpct=(a,b)=>b?(a-b)/Math.abs(b):null;
function delta(a,b,good){ // → {t,c}
  const p=dpct(a,b);if(p==null)return null;
  if(Math.abs(p)<0.0005)return{t:'= 0 %',c:C.mut};
  const up=p>0,bad=good==='down'?up:good==='up'?!up:null;
  return{t:(up?'▲ ':'▼ ')+nf(Math.abs(p)*100,1)+' %',c:bad==null?C.mut:bad?C.bad:C.ok,p};
}

/* ---------------- GRÁFICOS (canvas 2D) ---------------- */
const FF='Calibri, Carlito, "Segoe UI", Arial, sans-serif';
function mk(ctx,w,h,title,sub){
  const S=2,cv=ctx.makeCanvas(w*S,h*S),g=cv.getContext('2d');g.scale(S,S);
  g.fillStyle='#fff';g.fillRect(0,0,w,h);
  g.strokeStyle='#'+C.line;g.lineWidth=1;g.strokeRect(.5,.5,w-1,h-1);
  g.fillStyle='#'+C.navy;g.font=`700 16px ${FF}`;g.textBaseline='alphabetic';g.textAlign='left';g.fillText(title,18,28);
  if(sub){g.fillStyle='#'+C.mut;g.font=`400 11.5px ${FF}`;g.fillText(sub,18,46);}
  return{cv,g,w,h};
}
const out=c=>c.cv.toDataURL('image/png');
const lum=h=>{const n=parseInt(h,16);return .299*(n>>16)+.587*((n>>8)&255)+.114*(n&255);};
function niceMax(m){if(m<=0)return 1;const e=Math.pow(10,Math.floor(Math.log10(m))),f=m/e;const n=f<=1?1:f<=2?2:f<=2.5?2.5:f<=5?5:10;return n*e;}
function fit(g,t,maxW){t=String(t);if(g.measureText(t).width<=maxW)return t;while(t.length>2&&g.measureText(t+'…').width>maxW)t=t.slice(0,-1);return t+'…';}
function shortNum(v){return v>=1000?nf(v/1000,1)+'k':nf(v,v<10?2:0);}
function legend(g,items,x,y,maxW){
  g.font=`400 11.5px ${FF}`;g.textBaseline='middle';let cx=x,cy=y;
  for(const it of items){const wTxt=g.measureText(it.name).width,wTot=wTxt+22;if(cx+wTot>x+maxW&&cx>x){cx=x;cy+=17;}
    g.fillStyle='#'+it.color;if(it.line){g.fillRect(cx,cy-1.5,12,3);}else{g.beginPath();g.rect(cx,cy-5,10,10);g.fill();}
    g.fillStyle='#'+C.txt;g.textAlign='left';g.fillText(it.name,cx+16,cy);cx+=wTot+8;}
  return cy;
}
// barras agrupadas (+ línea opcional en eje derecho)
function combo(ctx,{w,h,title,sub,labels,sublabels,bars,line,fmt}){
  const c=mk(ctx,w,h,title,sub),g=c.g,n=labels.length;
  const L=52,R=line?56:18,T=78,B=sublabels?50:36,pw=w-L-R,ph=h-T-B;
  const all=bars.flatMap(s=>s.vals.filter(v=>v!=null));const mx=niceMax(Math.max(...all,0));
  legend(g,[...bars.map(s=>({name:s.name,color:s.color})),...(line?[{name:line.name,color:line.color,line:true}]:[])],18,60,w-36);
  g.font=`400 11px ${FF}`;g.textAlign='right';g.textBaseline='middle';
  for(let i=0;i<=4;i++){const y=T+ph-ph*i/4;g.strokeStyle='#'+(i?'E8ECF2':'C9D0DB');g.lineWidth=1;g.beginPath();g.moveTo(L,y);g.lineTo(L+pw,y);g.stroke();g.fillStyle='#'+C.mut;g.fillText(shortNum(mx*i/4),L-8,y);}
  const gw=pw/n,k=bars.length,bw=Math.min(46,gw*0.7/k);
  const lmx=line?niceMax(Math.max(...line.vals.filter(v=>v!=null),0)):0;
  if(line){g.textAlign='left';for(let i=0;i<=4;i++){const y=T+ph-ph*i/4;g.fillStyle='#'+line.color;g.fillText(line.axis(lmx*i/4),L+pw+8,y);}}
  const pts=[];
  for(let i=0;i<n;i++){
    const cx=L+gw*(i+.5),x0=cx-bw*k/2;
    bars.forEach((s,j)=>{const v=s.vals[i];if(v==null)return;const bh=ph*v/mx;g.fillStyle='#'+s.color;g.beginPath();const x=x0+j*bw,y=T+ph-bh,r=Math.min(3,bh/2);
      g.moveTo(x,y+bh);g.lineTo(x,y+r);g.quadraticCurveTo(x,y,x+r,y);g.lineTo(x+bw-2-r,y);g.quadraticCurveTo(x+bw-2,y,x+bw-2,y+r);g.lineTo(x+bw-2,y+bh);g.closePath();g.fill();
      if(n<=16&&bw>=18){g.font=`600 ${k>1?10:11}px ${FF}`;g.textAlign='center';g.textBaseline='alphabetic';const t=(fmt||shortNum)(v);
        if(bh>=24&&lum(s.color)<150){g.fillStyle='#fff';g.fillText(t,x+(bw-2)/2,y+15);}else{g.fillStyle='#'+C.txt;g.fillText(t,x+(bw-2)/2,y-5);}}});
    if(line&&line.vals[i]!=null)pts.push([cx,T+ph-ph*line.vals[i]/lmx,line.vals[i]]);
    g.fillStyle='#'+C.mut;g.font=`400 11px ${FF}`;g.textAlign='center';g.textBaseline='alphabetic';
    if(n<=20||i%Math.ceil(n/20)===0){g.fillText(labels[i],cx,T+ph+16);if(sublabels&&sublabels[i]){g.fillStyle='#9AA3B0';g.font=`400 9.5px ${FF}`;g.fillText(sublabels[i],cx,T+ph+30);}}
  }
  if(line&&pts.length){g.strokeStyle='#'+line.color;g.lineWidth=2.4;g.lineJoin='round';g.beginPath();pts.forEach(([x,y],i)=>i?g.lineTo(x,y):g.moveTo(x,y));g.stroke();
    pts.forEach(([x,y,v])=>{g.fillStyle='#fff';g.strokeStyle='#'+line.color;g.lineWidth=2;g.beginPath();g.arc(x,y,3.6,0,7);g.fill();g.stroke();
      if(n<=16){g.font=`700 10.5px ${FF}`;g.textAlign='center';g.textBaseline='alphabetic';g.lineWidth=3.5;g.strokeStyle='#fff';g.strokeText(line.fmt(v),x,y-9);g.fillStyle='#'+line.color;g.fillText(line.fmt(v),x,y-9);}});}
  return out(c);
}
function stack(ctx,{w,h,title,sub,labels,stacks,keys,colorOf}){
  const c=mk(ctx,w,h,title,sub),g=c.g,n=labels.length;
  const L=52,R=18,T=86,B=36,pw=w-L-R,ph=h-T-B;
  const tots=stacks.map(s=>keys.reduce((a,k)=>a+(s[k]||0),0)),mx=niceMax(Math.max(...tots,0));
  legend(g,keys.filter(k=>stacks.some(s=>s[k]>0)).map(k=>({name:k,color:colorOf(k)})),18,60,w-36);
  g.font=`400 11px ${FF}`;g.textAlign='right';g.textBaseline='middle';
  for(let i=0;i<=4;i++){const y=T+ph-ph*i/4;g.strokeStyle='#'+(i?'E8ECF2':'C9D0DB');g.beginPath();g.moveTo(L,y);g.lineTo(L+pw,y);g.stroke();g.fillStyle='#'+C.mut;g.fillText(shortNum(mx*i/4),L-8,y);}
  const gw=pw/n,bw=Math.min(46,gw*0.7);
  for(let i=0;i<n;i++){const x=L+gw*(i+.5)-bw/2;let y=T+ph;
    for(const k of keys){const v=stacks[i][k]||0;if(v<=0)continue;const bh=ph*v/mx;y-=bh;g.fillStyle='#'+colorOf(k);g.fillRect(x,y,bw-2,bh);}
    if(n<=16){g.fillStyle='#'+C.txt;g.font=`600 10.5px ${FF}`;g.textAlign='center';g.textBaseline='alphabetic';g.fillText(shortNum(tots[i]),x+(bw-2)/2,y-5);}
    g.fillStyle='#'+C.mut;g.font=`400 11px ${FF}`;g.textAlign='center';g.textBaseline='alphabetic';if(n<=20||i%Math.ceil(n/20)===0)g.fillText(labels[i],x+(bw-2)/2,T+ph+16);}
  return out(c);
}
function donut(ctx,{w,h,title,sub,items,total,centerLabel}){
  const c=mk(ctx,w,h,title,sub),g=c.g;const cx=h*0.52+8,cy=h/2+22,R=h*0.34,r=R*0.62;
  let a=-Math.PI/2;const tot=items.reduce((s,i)=>s+i.v,0)||1;
  for(const it of items){const da=it.v/tot*Math.PI*2;g.fillStyle='#'+it.color;g.beginPath();g.arc(cx,cy,R,a,a+da);g.arc(cx,cy,r,a+da,a,true);g.closePath();g.fill();a+=da;}
  g.strokeStyle='#fff';g.lineWidth=2;a=-Math.PI/2;for(const it of items){g.beginPath();g.moveTo(cx+Math.cos(a)*r,cy+Math.sin(a)*r);g.lineTo(cx+Math.cos(a)*R,cy+Math.sin(a)*R);g.stroke();a+=it.v/tot*Math.PI*2;}
  g.fillStyle='#'+C.navy;g.textAlign='center';g.textBaseline='alphabetic';g.font=`700 19px ${FF}`;g.fillText(total,cx,cy+4);g.fillStyle='#'+C.mut;g.font=`400 11px ${FF}`;g.fillText(centerLabel,cx,cy+20);
  const lx=cx+R+34;let y=66;const rows=items.length,step=Math.min(26,(h-84)/rows);
  g.textBaseline='middle';
  for(const it of items){g.fillStyle='#'+it.color;g.beginPath();g.rect(lx,y-5,11,11);g.fill();g.fillStyle='#'+C.txt;g.font=`400 12px ${FF}`;g.textAlign='left';
    g.fillText(fit(g,it.name,w-lx-130),lx+18,y);g.textAlign='right';g.font=`600 12px ${FF}`;g.fillText(it.right,w-18,y);y+=step;}
  return out(c);
}
function hbar(ctx,{w,h,title,sub,items,color,fmt}){
  const c=mk(ctx,w,h,title,sub),g=c.g;const L=224,R=92,T=62,n=items.length,rowH=(h-T-14)/Math.max(n,1),bh=Math.min(20,rowH*0.62);
  const mx=Math.max(...items.map(i=>i.v),1e-9);
  items.forEach((it,i)=>{const y=T+rowH*i+rowH/2;
    g.fillStyle='#'+C.mut;g.font=`600 11px ${FF}`;g.textAlign='right';g.textBaseline='middle';g.fillText(String(i+1),24,y);
    g.fillStyle='#'+C.txt;g.font=`400 11.5px ${FF}`;g.fillText(fit(g,it.name,L-60),L-8,y);
    const bw=(w-L-R)*it.v/mx;g.fillStyle='#EEF1F6';g.fillRect(L,y-bh/2,w-L-R,bh);g.fillStyle='#'+color;g.fillRect(L,y-bh/2,Math.max(bw,2),bh);
    g.fillStyle='#'+C.txt;g.font=`700 11.5px ${FF}`;g.textAlign='left';g.fillText(fmt(it.v),L+(w-L-R)+8,y);});
  return out(c);
}

/* ---------------- EXCEL ---------------- */
async function build(ctx){
  const{ExcelJS,A,B,meta,det,movs}=ctx;
  const nd=A.P.length,hasB=!!B&&B.P.length>0;
  const wb=new ExcelJS.Workbook();wb.creator='Escandallo Desayunos · Kimpton Los Monteros';wb.created=new Date();wb.title='Escandallo de desayunos '+meta.desde+' a '+meta.hasta;
  const sPanel=wb.addWorksheet('Panel',{properties:{tabColor:{argb:argb(C.navy)}},views:[{showGridLines:false}]});
  const sEvo=(meta.evo&&meta.evo.length)?wb.addWorksheet('Evolución',{properties:{tabColor:{argb:argb(C.navy)}},views:[{showGridLines:false}]}):null;
  const sDia=wb.addWorksheet('Por día',{properties:{tabColor:{argb:argb(C.gold)}},views:[{showGridLines:false}]});
  const sFam=wb.addWorksheet('Familias',{properties:{tabColor:{argb:argb(C.gold)}},views:[{showGridLines:false}]});
  const sRk=wb.addWorksheet('Rankings',{properties:{tabColor:{argb:argb(C.gold)}},views:[{showGridLines:false}]});
  const sDet=wb.addWorksheet('Detalle',{properties:{tabColor:{argb:argb('9AA3B0')}},views:[{state:'frozen',xSplit:3,ySplit:1,showGridLines:false}]});
  const sMov=wb.addWorksheet('Movimientos',{properties:{tabColor:{argb:argb('9AA3B0')}},views:[{state:'frozen',xSplit:0,ySplit:1,showGridLines:false}]});
  const page=(ws,land=true)=>{ws.pageSetup={orientation:land?'landscape':'portrait',paperSize:9,fitToPage:true,fitToWidth:1,fitToHeight:0,margins:{left:.4,right:.4,top:.5,bottom:.6,header:.3,footer:.3}};
    ws.headerFooter.oddFooter='&L&8Escandallo de desayunos · Kimpton Los Monteros Marbella&C&8&A&R&8Página &P de &N';};
  [sPanel,sDia,sFam,sRk,sDet,sMov].forEach(w=>page(w));
  if(sEvo){page(sEvo);sEvo.pageSetup.orientation='portrait';sEvo.pageSetup.fitToHeight=1;}
  if(nd<=16)sDia.pageSetup.fitToHeight=1;

  const famNames=A.fams.filter(f=>f.eur>0||f.merE>0).map(f=>f.cat);
  const topF=A.fams.slice(0,8).map(f=>f.cat);
  const colorOf=c=>{const i=topF.indexOf(c);return i<0?FAM[8]:FAM[i];};
  const days=A.P;
  const last=det.length+1;                      // última fila de Detalle
  // posiciones fijas de 'Por día'
  const D0=6,DL=D0+nd-1,DT=DL+1;                // filas de datos y total
  const dref=(col,row)=>`'Por día'!$${col}$${row}`;

  /* ===== helpers de estilo ===== */
  const setC=(ws,addr,val,o={})=>{const c=ws.getCell(addr);if(val!==undefined&&val!==null)c.value=val;
    if(o.font)c.font={name:FONT,size:11,color:{argb:argb(C.txt)},...o.font};
    if(o.fill)c.fill=fill(o.fill);if(o.al)c.alignment=o.al;if(o.nf)c.numFmt=o.nf;if(o.b)c.border=o.b;return c;};
  const banner=(ws,r1,c1,c2,title,subtitle,info)=>{
    ws.mergeCells(r1,c1,r1,c2);ws.mergeCells(r1+1,c1,r1+1,c2);
    const a=ws.getCell(r1,c1);a.value=title;a.font={name:FONT,size:22,bold:true,color:{argb:argb(C.white)}};a.alignment={vertical:'middle',horizontal:'left',indent:1};
    const b=ws.getCell(r1+1,c1);b.value=subtitle;b.font={name:FONT,size:11,color:{argb:argb('E3CC9B')}};b.alignment={vertical:'middle',horizontal:'left',indent:1};
    for(let c=c1;c<=c2;c++){ws.getCell(r1,c).fill=fill(C.navy);ws.getCell(r1+1,c).fill=fill(C.navy);}
    ws.getRow(r1).height=36;ws.getRow(r1+1).height=20;
    if(info){ws.mergeCells(r1+2,c1,r1+2,c2);const i=ws.getCell(r1+2,c1);i.value=info;i.font={name:FONT,size:10,italic:true,color:{argb:argb(C.mut)}};i.alignment={vertical:'middle',horizontal:'left',indent:1,wrapText:true};
      for(let c=c1;c<=c2;c++){ws.getCell(r1+2,c).fill=fill(C.bg);}ws.getRow(r1+2).height=22;}
    for(let c=c1;c<=c2;c++){ws.getCell(r1+1,c).border={bottom:{style:'medium',color:{argb:argb(C.gold)}}};}
  };
  const section=(ws,r,c1,c2,text)=>{ws.mergeCells(r,c1,r,c2);const c=ws.getCell(r,c1);c.value=text.toUpperCase();c.font={name:FONT,size:11,bold:true,color:{argb:argb(C.gold)}};c.alignment={vertical:'bottom',indent:0};
    for(let i=c1;i<=c2;i++)ws.getCell(r,i).border={bottom:{style:'thin',color:{argb:argb(C.gold)}}};ws.getRow(r).height=24;};
  const headRow=(ws,r,c1,labels,opts={})=>{labels.forEach((t,i)=>{const c=ws.getCell(r,c1+i);c.value=t;c.font={name:FONT,size:10.5,bold:true,color:{argb:argb(C.white)}};c.fill=fill(C.navy);
    c.alignment={vertical:'middle',horizontal:(opts.left&&opts.left.includes(i))?'left':'center',wrapText:true,indent:(opts.left&&opts.left.includes(i))?1:0};c.border={bottom:{style:'medium',color:{argb:argb(C.gold)}}};});ws.getRow(r).height=opts.h||32;};
  const bodyCell=(c,{nf:f,al,bold,fillc,color}={})=>{c.font={name:FONT,size:11,bold:!!bold,color:{argb:argb(color||C.txt)}};if(f)c.numFmt=f;c.alignment={vertical:'middle',horizontal:al||'right',indent:al==='left'?1:0};
    c.border={bottom:thin('E8ECF2')};if(fillc)c.fill=fill(fillc);};
  const totalCell=(c,{nf:f,al}={})=>{c.font={name:FONT,size:11,bold:true,color:{argb:argb(C.navy)}};if(f)c.numFmt=f;c.fill=fill(C.goldL);c.alignment={vertical:'middle',horizontal:al||'right',indent:al==='left'?1:0};c.border={top:{style:'thin',color:{argb:argb(C.gold)}},bottom:{style:'medium',color:{argb:argb(C.gold)}}};};
  const bar=(ws,ref,color)=>ws.addConditionalFormatting({ref,rules:[{type:'dataBar',priority:P(),minLength:0,maxLength:100,gradient:false,showValue:true,cfvo:[{type:'num',value:0},{type:'max'}],color:{argb:argb(color)}}]});
  const info=`Periodo ${ddmmyyyy(meta.desde)} – ${ddmmyyyy(meta.hasta)}  ·  ${nd} días con datos${meta.soloCerrados?' (solo días cerrados)':''}${hasB?`  ·  Comparado con ${ddmmyyyy(meta.cdesde)} – ${ddmmyyyy(meta.chasta)} (${B.P.length} días${B.P.length!==nd?' · distinto nº de días: compara con el coste por pax':''})`:''}  ·  Generado el ${meta.generado} por ${meta.usuario}`;
  const TITLE='Escandallo de desayunos',SUB='Kimpton Los Monteros Marbella  ·  Food & Beverage';

  /* ===== DETALLE (datos con fórmulas) ===== */
  {
    const ws=sDet,H=['Fecha','Código','Artículo','Categoría','Unidad','Expuesto','Reposición','Aprovechamiento','Merma','Sacado','Consumo','Coste unidad (€)','Coste consumo (€)','Coste merma (€)','Coste bruto (€)','¿Reutiliza mañana?','Cerrado por'];
    headRow(ws,1,1,H,{left:[2,3],h:34});
    const W=[13,17,52,24,9,10,11,15,9,9,10,12,14,13,13,13,14];W.forEach((w,i)=>ws.getColumn(i+1).width=w);
    det.forEach((d,i)=>{const R=i+2,row=ws.getRow(R);
      const vals=[dt(d.fecha),d.codigo,d.desc,d.cat,d.uni,d.E,d.R,d.A,d.M,{formula:`F${R}+G${R}`,result:d.E+d.R},{formula:`F${R}+G${R}-H${R}-I${R}`,result:d.E+d.R-d.A-d.M},d.cu,
        {formula:`K${R}*L${R}`,result:(d.E+d.R-d.A-d.M)*d.cu},{formula:`I${R}*L${R}`,result:d.M*d.cu},{formula:`(F${R}+G${R}-H${R})*L${R}`,result:(d.E+d.R-d.A)*d.cu},d.reut,d.cerrado];
      vals.forEach((v,j)=>{const c=row.getCell(j+1);c.value=v;const nfm=j===0?'dd/mm/yyyy':[5,6,7,8,9,10].includes(j)?QTY:j===11?'#,##0.000" €"':[12,13,14].includes(j)?EUR:null;
        bodyCell(c,{nf:nfm,al:[1,2,3,4,15,16].includes(j)?'left':j===0?'center':'right'});if(j===1)c.numFmt='@';});
    });
    ws.autoFilter={from:{row:1,column:1},to:{row:Math.max(last,2),column:17}};
    ws.addConditionalFormatting({ref:`A2:Q${Math.max(last,2)}`,rules:[{type:'expression',priority:P(),formulae:['MOD(ROW(),2)=0'],style:{fill:{type:'pattern',pattern:'solid',bgColor:{argb:argb('F7F9FC')}}}}]});
    ws.addConditionalFormatting({ref:`I2:I${Math.max(last,2)}`,rules:[{type:'cellIs',priority:P(),operator:'greaterThan',formulae:[0],style:{font:{color:{argb:argb(C.bad)},bold:true}}}]});
    ws.pageSetup.printTitlesRow='1:1';
  }

  /* ===== POR DÍA ===== */
  {
    const ws=sDia;[2,18,9,16,14,14,14,16,14,12,14,14,12].forEach((w,i)=>ws.getColumn(i+1).width=w);
    ws.getColumn(1).width=2;
    const lastCol=hasB?13:10;
    banner(ws,1,2,lastCol,'Resumen por día',SUB,info);
    ws.getRow(4).height=10;
    const heads=['Día','PAX','Coste consumo (€)','Coste / pax (€)','Merma (€)','Merma % s/ bruto','Coste bruto (€)','Coste bruto / pax (€)','Estado'];
    headRow(ws,5,2,heads,{left:[0],h:36});ws.getRow(5).height=36;
    if(hasB)headRow(ws,5,11,['Día comparado','Coste / pax comp. (€)','Δ % coste / pax'],{h:36});
    days.forEach((d,i)=>{const R=D0+i;
      const dateC=ws.getCell(R,2);dateC.value=dt(d.fecha);bodyCell(dateC,{nf:DTF,al:'left'});
      const cons=d.consumo,mer=d.merma,br=d.bruto;
      const f=(col)=>`SUMIFS(Detalle!$${col}$2:$${col}$${last},Detalle!$A$2:$A$${last},$B${R})`;
      const cells=[[3,d.pax,INT,null],[4,{formula:f('M'),result:cons},EUR,null],[5,{formula:`D${R}/C${R}`,result:cons/d.pax},EUR,null],[6,{formula:f('N'),result:mer},EUR,null],[7,{formula:`IF(H${R}=0,0,F${R}/H${R})`,result:br?mer/br:0},PCT,null],[8,{formula:f('O'),result:br},EUR,null],[9,{formula:`H${R}/C${R}`,result:br/d.pax},EUR,null]];
      cells.forEach(([c,v,fm])=>{const cell=ws.getCell(R,c);cell.value=v;bodyCell(cell,{nf:fm});});
      const st=ws.getCell(R,10);st.value=d.estado==='cerrado'?'Cerrado':'Abierto';bodyCell(st,{al:'center',color:d.estado==='cerrado'?C.ok:C.gold,bold:true});
      if(hasB){const b=B.P[i];if(b){const x=ws.getCell(R,11);x.value=dt(b.fecha);bodyCell(x,{nf:'dd/mm/yyyy',al:'center'});const y=ws.getCell(R,12);y.value=b.cpax;bodyCell(y,{nf:EUR});
        const z=ws.getCell(R,13);z.value={formula:`IF(L${R}=0,0,E${R}/L${R}-1)`,result:b.cpax?d.cpax/b.cpax-1:0};bodyCell(z,{nf:'+0.0%;-0.0%;0.0%'});}else for(const c of [11,12,13])bodyCell(ws.getCell(R,c),{});}
      ws.getRow(R).height=21;
    });
    // totales
    const tp=A.totPax,tc=A.totC,tm=A.totM,tb=A.totB;
    const T=DT;ws.getRow(T).height=24;
    const tv=[[2,'TOTAL / PONDERADO',null,'left'],[3,{formula:`SUM(C${D0}:C${DL})`,result:tp},INT],[4,{formula:`SUM(D${D0}:D${DL})`,result:tc},EUR],[5,{formula:`D${T}/C${T}`,result:tc/tp},EUR],[6,{formula:`SUM(F${D0}:F${DL})`,result:tm},EUR],[7,{formula:`IF(H${T}=0,0,F${T}/H${T})`,result:tb?tm/tb:0},PCT],[8,{formula:`SUM(H${D0}:H${DL})`,result:tb},EUR],[9,{formula:`H${T}/C${T}`,result:tb/tp},EUR],[10,'']];
    tv.forEach(([c,v,fm,al])=>{const cell=ws.getCell(T,c);cell.value=v;totalCell(cell,{nf:fm,al});});
    if(hasB){[11,12,13].forEach(c=>totalCell(ws.getCell(T,c),{}));ws.getCell(T,12).value=B.cpax;ws.getCell(T,12).numFmt=EUR;ws.getCell(T,13).value={formula:`IF(L${T}=0,0,E${T}/L${T}-1)`,result:B.cpax?A.cpax/B.cpax-1:0};ws.getCell(T,13).numFmt='+0.0%;-0.0%;0.0%';}
    const T2=T+1;ws.getRow(T2).height=22;
    const simple=days.reduce((a,d)=>a+d.cpax,0)/nd,simpleB=days.reduce((a,d)=>a+d.bruto/d.pax,0)/nd;
    for(let c=2;c<=10;c++){const cell=ws.getCell(T2,c);cell.fill=fill(C.bg);cell.font={name:FONT,size:10.5,italic:true,color:{argb:argb(C.mut)}};cell.border={bottom:thin(C.line)};}
    ws.getCell(T2,2).value='Media simple por día';ws.getCell(T2,2).alignment={indent:1};
    ws.getCell(T2,5).value={formula:`AVERAGE(E${D0}:E${DL})`,result:simple};ws.getCell(T2,5).numFmt=EUR;
    ws.getCell(T2,9).value={formula:`AVERAGE(I${D0}:I${DL})`,result:simpleB};ws.getCell(T2,9).numFmt=EUR;
    bar(ws,`D${D0}:D${DL}`,'9FB4D6');bar(ws,`E${D0}:E${DL}`,'E3CC9B');
    ws.addConditionalFormatting({ref:`G${D0}:G${DL}`,rules:[{type:'colorScale',priority:P(),cfvo:[{type:'min'},{type:'max'}],color:[{argb:'FFFFFFFF'},{argb:'FFF2B8A8'}]}]});
    if(hasB)ws.addConditionalFormatting({ref:`M${D0}:M${DL}`,rules:[{type:'cellIs',priority:P(),operator:'greaterThan',formulae:[0],style:{font:{color:{argb:argb(C.bad)},bold:true}}},{type:'cellIs',priority:P(),operator:'lessThan',formulae:[0],style:{font:{color:{argb:argb(C.ok)},bold:true}}}]});
    // matriz día × familia
    const M0=T2+3;section(ws,M0-1,2,lastCol,'Consumo por día y familia (€)');
    const fk=topF.slice(0,8);const hasOtras=famNames.length>fk.length;
    const mh=['Día',...fk,...(hasOtras?['Otras']:[]),'Total'];
    headRow(ws,M0,2,mh,{left:[0],h:40});
    ws.getRow(M0).height=40;
    // las cabeceras de familia deben coincidir con la categoría para SUMIFS
    days.forEach((d,i)=>{const R=M0+1+i;const dc=ws.getCell(R,2);dc.value=dt(d.fecha);bodyCell(dc,{nf:DTF,al:'left'});
      let rest=d.consumo;const nc=fk.length;
      fk.forEach((cat,j)=>{const v=(A.fam.get(cat)&&A.fam.get(cat).byDay.get(d.fecha))||0;rest-=v;const colL=ws.getColumn(3+j).letter;
        const cell=ws.getCell(R,3+j);cell.value={formula:`SUMIFS(Detalle!$M$2:$M$${last},Detalle!$A$2:$A$${last},$B${R},Detalle!$D$2:$D$${last},${colL}$${M0})`,result:v};bodyCell(cell,{nf:EUR0});});
      const totCol=3+nc+(hasOtras?1:0),tl=ws.getColumn(totCol).letter,fl=ws.getColumn(3+nc-1).letter;
      if(hasOtras){const cell=ws.getCell(R,3+nc);cell.value={formula:`${tl}${R}-SUM(C${R}:${fl}${R})`,result:rest};bodyCell(cell,{nf:EUR0});}
      const tcell=ws.getCell(R,totCol);tcell.value={formula:`SUMIFS(Detalle!$M$2:$M$${last},Detalle!$A$2:$A$${last},$B${R})`,result:d.consumo};bodyCell(tcell,{nf:EUR0,bold:true});
      ws.getRow(R).height=21;});
    const MT=M0+1+nd;ws.getRow(MT).height=24;
    const ncol=fk.length+(hasOtras?1:0)+1;
    totalCell(ws.getCell(MT,2),{al:'left'});ws.getCell(MT,2).value='TOTAL';
    for(let j=0;j<ncol;j++){const colL=ws.getColumn(3+j).letter;const cell=ws.getCell(MT,3+j);let v=0;
      if(j<fk.length)v=(A.fam.get(fk[j])?A.fam.get(fk[j]).eur:0);else if(hasOtras&&j===fk.length)v=A.totC-fk.reduce((a,c)=>a+(A.fam.get(c)?A.fam.get(c).eur:0),0);else v=A.totC;
      cell.value={formula:`SUM(${colL}${M0+1}:${colL}${MT-1})`,result:v};totalCell(cell,{nf:EUR0});}
    ws.addConditionalFormatting({ref:`C${M0+1}:${ws.getColumn(2+ncol-1).letter}${MT-1}`,rules:[{type:'colorScale',priority:P(),cfvo:[{type:'min'},{type:'max'}],color:[{argb:'FFFFFFFF'},{argb:'FFA9BCDD'}]}]});
    // leyenda de familias con color
    ws.views=[{showGridLines:false,state:'frozen',ySplit:5,xSplit:2}];
    ws.getRow(MT+2).getCell(2).value='Consumo = Expuesto + Reposición − Aprovechamiento − Merma. Coste bruto = Expuesto + Reposición − Aprovechamiento (incluye el coste de la merma). Precio congelado al cerrar cada artículo.';
    ws.getRow(MT+2).getCell(2).font={name:FONT,size:9.5,italic:true,color:{argb:argb(C.mut)}};
    // anchos adicionales para matriz (columnas C.. hasta 11) ya definidos
    for(let c=3;c<=Math.max(10,2+ncol);c++)if(!ws.getColumn(c).width||ws.getColumn(c).width<14)ws.getColumn(c).width=15;
  }

  /* ===== FAMILIAS ===== */
  {
    const ws=sFam;[2,3,34,16,12,12,15,14,12,3,16,12].forEach((w,i)=>ws.getColumn(i+1).width=w);ws.getColumn(1).width=2;
    const lastCol=hasB?12:9;banner(ws,1,2,lastCol,'Consumo por familia',SUB,info);ws.getRow(4).height=10;
    const heads=['','Familia','Consumo (€)','% s/ total','€ / pax','Merma (€)','Merma % s/ bruto','Artículos'];
    headRow(ws,5,2,heads,{left:[1],h:36});if(hasB){headRow(ws,5,11,['Consumo comp. (€)','Δ %'],{h:36});}
    const fams=A.fams.filter(f=>f.eur>0||f.merE>0);const bfam=hasB?new Map(B.fams.map(f=>[f.cat,f])):null;
    const F0=6,FL=F0+fams.length-1,FT=FL+1;
    fams.forEach((f,i)=>{const R=F0+i;ws.getRow(R).height=22;
      const sw=ws.getCell(R,2);sw.fill=fill(colorOf(f.cat));sw.border={bottom:thin('FFFFFF'),top:thin('FFFFFF')};
      const name=ws.getCell(R,3);name.value=f.cat;bodyCell(name,{al:'left',bold:true});
      const v=ws.getCell(R,4);v.value={formula:`SUMIFS(Detalle!$M$2:$M$${last},Detalle!$D$2:$D$${last},$C${R})`,result:f.eur};bodyCell(v,{nf:EUR});
      const p=ws.getCell(R,5);p.value={formula:`IF($D$${FT}=0,0,D${R}/$D$${FT})`,result:A.totC?f.eur/A.totC:0};bodyCell(p,{nf:PCT});
      const pp=ws.getCell(R,6);pp.value={formula:`D${R}/${dref('C',DT)}`,result:A.totPax?f.eur/A.totPax:0};bodyCell(pp,{nf:EUR});
      const m=ws.getCell(R,7);m.value={formula:`SUMIFS(Detalle!$N$2:$N$${last},Detalle!$D$2:$D$${last},$C${R})`,result:f.merE};bodyCell(m,{nf:EUR});
      const mp=ws.getCell(R,8);mp.value={formula:`IF((D${R}+G${R})=0,0,G${R}/(D${R}+G${R}))`,result:(f.eur+f.merE)?f.merE/(f.eur+f.merE):0};bodyCell(mp,{nf:PCT});
      const a=ws.getCell(R,9);a.value=f.arts.size;bodyCell(a,{nf:INT,al:'center'});
      if(hasB){const b=bfam.get(f.cat);const x=ws.getCell(R,11);x.value=b?b.eur:0;bodyCell(x,{nf:EUR});const y=ws.getCell(R,12);y.value=b&&b.eur?{formula:`IF(K${R}=0,0,D${R}/K${R}-1)`,result:f.eur/b.eur-1}:null;bodyCell(y,{nf:'+0.0%;-0.0%;0.0%'});}
    });
    ws.getRow(FT).height=24;
    totalCell(ws.getCell(FT,2),{});totalCell(ws.getCell(FT,3),{al:'left'});ws.getCell(FT,3).value='TOTAL';
    const tt=[[4,A.totC,EUR],[5,1,PCT],[6,A.totC/A.totPax,EUR],[7,A.totM,EUR],[8,A.totB?A.totM/A.totB:0,PCT],[9,'',null]];
    ws.getCell(FT,4).value={formula:`SUM(D${F0}:D${FL})`,result:A.totC};ws.getCell(FT,5).value={formula:`SUM(E${F0}:E${FL})`,result:1};ws.getCell(FT,6).value={formula:`D${FT}/${dref('C',DT)}`,result:A.totC/A.totPax};ws.getCell(FT,7).value={formula:`SUM(G${F0}:G${FL})`,result:A.totM};ws.getCell(FT,8).value={formula:`IF((D${FT}+G${FT})=0,0,G${FT}/(D${FT}+G${FT}))`,result:(A.totC+A.totM)?A.totM/(A.totC+A.totM):0};
    tt.forEach(([c,,fm])=>{totalCell(ws.getCell(FT,c),{nf:fm,al:c===9?'center':null});});
    if(hasB){[11,12].forEach(c=>totalCell(ws.getCell(FT,c),{}));ws.getCell(FT,11).value=B.totC;ws.getCell(FT,11).numFmt=EUR;ws.getCell(FT,12).value={formula:`IF(K${FT}=0,0,D${FT}/K${FT}-1)`,result:B.totC?A.totC/B.totC-1:0};ws.getCell(FT,12).numFmt='+0.0%;-0.0%;0.0%';
      ws.addConditionalFormatting({ref:`L${F0}:L${FL}`,rules:[{type:'cellIs',priority:P(),operator:'greaterThan',formulae:[0],style:{font:{color:{argb:argb(C.bad)},bold:true}}},{type:'cellIs',priority:P(),operator:'lessThan',formulae:[0],style:{font:{color:{argb:argb(C.ok)},bold:true}}}]});}
    bar(ws,`D${F0}:D${FL}`,'9FB4D6');bar(ws,`E${F0}:E${FL}`,'E3CC9B');
    ws.addConditionalFormatting({ref:`H${F0}:H${FL}`,rules:[{type:'colorScale',priority:P(),cfvo:[{type:'min'},{type:'max'}],color:[{argb:'FFFFFFFF'},{argb:'FFF2B8A8'}]}]});
    ws.views=[{showGridLines:false,state:'frozen',ySplit:5}];
  }

  /* ===== RANKINGS ===== */
  const rkList=(AA,metric,filt)=>{const key=o=>metric==='eur'?o.eur:metric==='mer'?o.merE:o.cons;return [...AA.pc.values()].filter(filt||(()=>true)).filter(o=>key(o)>0).sort((a,b)=>key(b)-key(a));};
  {
    const ws=sRk;[2,6,44,26,13,12,15,12,12,3,15,12,12,12].forEach((w,i)=>ws.getColumn(i+1).width=w);ws.getColumn(1).width=2;
    const lastCol=hasB?14:9;banner(ws,1,2,lastCol,'Rankings de artículos',SUB,info);ws.getRow(4).height=10;
    let r=5;
    const block=(titleTxt,subtxt,metric,N,filt,qtyMode,uniLabel)=>{
      const list=rkList(A,metric,filt).slice(0,N);if(!list.length)return;
      section(ws,r,2,lastCol,titleTxt);r++;
      if(subtxt){ws.mergeCells(r,2,r,lastCol);const c=ws.getCell(r,2);c.value=subtxt;c.font={name:FONT,size:10,italic:true,color:{argb:argb(C.mut)}};r++;}
      const full=rkList(A,metric,filt);const total=full.reduce((a,o)=>a+(metric==='eur'?o.eur:metric==='mer'?o.merE:o.cons),0);
      const pl=hasB?rkList(B,metric,filt):null;const pr=hasB?new Map(pl.map((o,i)=>[o.codigo,i+1])):null,pv=hasB?new Map(pl.map(o=>[o.codigo,metric==='eur'?o.eur:metric==='mer'?o.merE:o.cons])):null;
      const heads=qtyMode?['#','Artículo','Familia','Cantidad','Unidad','Coste (€)','Cantidad / pax','']:['#','Artículo','Familia',metric==='mer'?'Merma (cant.)':'Cantidad','Unidad',metric==='mer'?'Merma (€)':'Consumo (€)','% s/ total','€ / pax'];
      headRow(ws,r,2,heads.slice(0,8),{left:[1,2],h:32});
      if(hasB&&!qtyMode)headRow(ws,r,11,['Valor comp. (€)','Δ %','Puesto ant.','Movimiento'],{h:32});
      const r0=r+1;
      list.forEach((o,i)=>{const R=r0+i;ws.getRow(R).height=21;const name=o.p?o.p.descripcion:o.codigo;
        const cells=[[2,i+1,INT,'center'],[3,name,null,'left'],[4,o.cat,null,'left'],[5,metric==='mer'?o.mer:o.cons,QTY,null],[6,cap(uname(o.uni,2)),null,'center']];
        cells.forEach(([c,v,fm,al])=>{const cell=ws.getCell(R,c);cell.value=v;bodyCell(cell,{nf:fm,al,bold:c===3});});
        const val=metric==='eur'?o.eur:metric==='mer'?o.merE:o.cons;
        if(qtyMode){const e=ws.getCell(R,7);e.value=o.eur;bodyCell(e,{nf:EUR});const f=ws.getCell(R,8);f.value=A.totPax?o.cons/A.totPax:0;bodyCell(f,{nf:'#,##0.000'});}
        else{const e=ws.getCell(R,7);e.value=val;bodyCell(e,{nf:EUR});const f=ws.getCell(R,8);f.value=total?val/total:0;bodyCell(f,{nf:PCT});const g=ws.getCell(R,9);g.value=A.totPax?val/A.totPax:0;bodyCell(g,{nf:EUR});}
        if(hasB&&!qtyMode){const v0=pv.get(o.codigo),r1=pr.get(o.codigo);
          const x=ws.getCell(R,11);x.value=v0==null?null:v0;bodyCell(x,{nf:EUR});
          const y=ws.getCell(R,12);y.value=v0?{formula:`IF(K${R}=0,0,G${R}/K${R}-1)`,result:val/v0-1}:null;bodyCell(y,{nf:'+0.0%;-0.0%;0.0%'});
          const z=ws.getCell(R,13);z.value=r1==null?'nuevo':r1;bodyCell(z,{nf:INT,al:'center'});
          const mv=r1==null?null:r1-(i+1);const w=ws.getCell(R,14);w.value=mv==null?'●':mv>0?'▲ '+mv:mv<0?'▼ '+(-mv):'=';bodyCell(w,{al:'center',bold:true,color:mv==null?C.gold:mv>0?C.ok:mv<0?C.bad:C.mut});}
      });
      bar(ws,`G${r0}:G${r0+list.length-1}`,qtyMode?'C9D3E6':metric==='mer'?'F2B8A8':'9FB4D6');
      if(hasB&&!qtyMode)ws.addConditionalFormatting({ref:`L${r0}:L${r0+list.length-1}`,rules:[{type:'cellIs',priority:P(),operator:'greaterThan',formulae:[0],style:{font:{color:{argb:argb(C.bad)},bold:true}}},{type:'cellIs',priority:P(),operator:'lessThan',formulae:[0],style:{font:{color:{argb:argb(C.ok)},bold:true}}}]});
      r=r0+list.length+2;
    };
    block('Top 25 artículos por consumo (€)','Ordenados por coste de consumo (expuesto + reposición − aprovechamiento − merma).','eur',25);
    block('Top 25 artículos por merma (€)','Lo que más dinero se tira: merma valorada al coste de cada artículo.','mer',25);
    // por cantidad: las 3 unidades más habituales
    const cu0={};for(const o of A.pc.values())if(o.cons>0)cu0[uk(o.uni)]=(cu0[uk(o.uni)]||0)+1;
    const unis=Object.keys(cu0).sort((a,b)=>cu0[b]-cu0[a]).slice(0,3);
    unis.forEach(u=>block(`Top 15 por cantidad · ${cap(uname(u,2))}`,'Las cantidades solo se comparan entre artículos de la misma unidad.','qty',15,o=>uk(o.uni)===u,true));
    ws.views=[{showGridLines:false}];
  }

  /* ===== MOVIMIENTOS ===== */
  {
    const ws=sMov,H=['Fecha','Hora','Código','Artículo','Tipo','Cantidad','Origen','Usuario','Anulado','Anulado por'];
    headRow(ws,1,1,H,{left:[2,3],h:30});[13,10,17,52,13,11,11,16,10,16].forEach((w,i)=>ws.getColumn(i+1).width=w);
    movs.forEach((m,i)=>{const R=i+2,row=ws.getRow(R);
      [dt(m.fecha),m.hora,m.codigo,m.desc,m.tipo==='expuesto'?'Expuesto':'Reposición',m.cant,m.origen==='arrastre'?'Sobrante de ayer':'Manual',m.usuario,m.anulado?'SÍ':'',m.anuladoPor||''].forEach((v,j)=>{const c=row.getCell(j+1);c.value=v;bodyCell(c,{nf:j===0?'dd/mm/yyyy':j===5?QTY:j===2?'@':null,al:[2,3].includes(j)?'left':j===5?'right':'center'});});
    });
    ws.autoFilter={from:{row:1,column:1},to:{row:Math.max(movs.length+1,2),column:10}};
    ws.addConditionalFormatting({ref:`A2:J${Math.max(movs.length+1,2)}`,rules:[{type:'expression',priority:P(),formulae:['$I2="SÍ"'],style:{font:{color:{argb:argb('9AA3B0')},strike:true}}},{type:'expression',priority:P(),formulae:['MOD(ROW(),2)=0'],style:{fill:{type:'pattern',pattern:'solid',bgColor:{argb:argb('F7F9FC')}}}}]});
    ws.pageSetup.printTitlesRow='1:1';
  }

  /* ===== PANEL ===== */
  {
    const ws=sPanel;
    // rejilla: A margen · B-D · E hueco · F-H · I hueco · J-L · M hueco · N-P · Q margen
    const widths=[2,14,14,14,1.6,14,14,14,1.6,14,14,14,1.6,14,14,14,2];widths.forEach((w,i)=>ws.getColumn(i+1).width=w);
    const colX=c=>{let x=0;for(let i=1;i<c;i++)x+=px(widths[i-1]);return x;}; // px acumulados hasta la col c (1-indexed)
    banner(ws,1,2,16,TITLE,SUB+'  ·  Informe de consumo, coste y merma',info);ws.getRow(4).height=12;
    const cards=(r,list)=>{const starts=[2,6,10,14];
      list.forEach((k,i)=>{const c0=starts[i];
        ws.mergeCells(r,c0,r,c0+2);ws.mergeCells(r+1,c0,r+2,c0+2);ws.mergeCells(r+3,c0,r+3,c0+2);
        for(let rr=r;rr<=r+3;rr++)for(let cc=c0;cc<=c0+2;cc++){const cell=ws.getCell(rr,cc);cell.fill=fill(C.bg);
          const bd={};if(cc===c0)bd.left={style:'thick',color:{argb:argb(k.accent||C.gold)}};if(rr===r+3)bd.bottom=thin(C.line);if(rr===r)bd.top=thin(C.line);if(cc===c0+2)bd.right=thin(C.line);cell.border=bd;}
        const L=ws.getCell(r,c0);L.value=k.label.toUpperCase();L.font={name:FONT,size:9.5,bold:true,color:{argb:argb(C.mut)}};L.alignment={vertical:'bottom',horizontal:'left',indent:1};
        const V=ws.getCell(r+1,c0);V.value=k.value;V.numFmt=k.nf;V.font={name:FONT,size:26,bold:true,color:{argb:argb(k.color||C.navy)}};V.alignment={vertical:'middle',horizontal:'left',indent:1};
        const S=ws.getCell(r+3,c0);S.value=k.sub;S.font={name:FONT,size:10,bold:!!k.subBold,color:{argb:argb(k.subColor||C.mut)}};S.alignment={vertical:'middle',horizontal:'left',indent:1};});
      ws.getRow(r).height=22;ws.getRow(r+1).height=21;ws.getRow(r+2).height=21;ws.getRow(r+3).height=22;};
    const d1=(a,b,g)=>hasB?delta(a,b,g):null;
    const sub=(dd,fallback,suf)=>dd?{sub:`${dd.t}  ${suf}`,subColor:dd.c,subBold:true}:{sub:fallback};
    const nvs=days.length,avgPax=A.totPax/nvs;
    const suf='vs periodo comparado';
    section(ws,5,2,16,'Resumen del periodo');
    cards(6,[
      {label:'Coste por pax',value:{formula:`${dref('D',DT)}/${dref('C',DT)}`,result:A.cpax},nf:EUR,...sub(d1(A.cpax,B&&B.cpax,'down'),'Consumo / PAX del periodo',suf)},
      {label:'Coste de consumo',value:{formula:dref('D',DT),result:A.totC},nf:EUR0,...sub(d1(A.totC,B&&B.totC,'down'),`${nvs} días con datos`,suf),accent:C.navy},
      {label:'PAX totales',value:{formula:dref('C',DT),result:A.totPax},nf:INT,...sub(d1(A.totPax,B&&B.totPax,'none'),`Media ${nf(avgPax,0)} pax / día`,suf),accent:C.teal},
      {label:'Merma',value:{formula:dref('F',DT),result:A.totM},nf:EUR0,...(hasB?sub(d1(A.totM,B.totM,'down'),'',suf):{sub:`${nf(A.totB?A.totM/A.totB*100:0,1)} % del coste bruto`}),accent:C.red,color:C.red}
    ]);
    ws.getRow(10).height=10;
    const nvB=hasB?B.P.length||1:1;
    cards(11,[
      {label:'Merma media por día',value:{formula:`${dref('F',DT)}/${nvs}`,result:A.totM/nvs},nf:EUR,...sub(d1(A.totM/nvs,hasB&&B.totM/nvB,'down'),'Merma total / días',suf),accent:C.red},
      {label:'Merma por pax',value:{formula:`${dref('F',DT)}/${dref('C',DT)}`,result:A.totM/A.totPax},nf:EUR,...sub(d1(A.totM/A.totPax,hasB&&B.totM/B.totPax,'down'),'Merma / PAX',suf),accent:C.red},
      {label:'% de lo sacado que se tira',value:{formula:dref('G',DT),result:A.totB?A.totM/A.totB:0},nf:PCT,...sub(d1(A.totB?A.totM/A.totB:0,hasB&&(B.totB?B.totM/B.totB:0),'down'),'Merma / coste bruto',suf),accent:C.red},
      {label:'Coste bruto (sacado − aprovechado)',value:{formula:dref('H',DT),result:A.totB},nf:EUR0,...sub(d1(A.totB,B&&B.totB,'down'),'Consumo + merma',suf),accent:C.gold}
    ]);
    ws.getRow(15).height=10;
    // gráficos
    const CW=px(14)*3+px(1.6)+px(14)*3-6;                  // ancho de cada gráfico (px)
    const CH=330,rowsFor=Math.ceil(CH/ptpx(20)),RH=20;
    const place=(rowTop,side,png,h)=>{const id=wb.addImage({base64:png,extension:'png'});const x=side===0?colX(2):colX(10);
      // col/row en fracciones (0-indexed)
      let col=0,acc=0;for(let i=1;i<=widths.length;i++){const w=px(widths[i-1]);if(acc+w>x){col=i-1+(x-acc)/w;break;}acc+=w;}
      ws.addImage(id,{tl:{col,row:rowTop-1+0.08},ext:{width:CW,height:h||CH},editAs:'oneCell'});};
    const reserve=(rowStart,n)=>{for(let i=0;i<n;i++)ws.getRow(rowStart+i).height=RH;};
    const labels=days.map(d=>wd(d.fecha)+' '+ddmm(d.fecha));
    const sublab=hasB?days.map((_,i)=>B.P[i]?'vs '+ddmm(B.P[i].fecha):''):null;
    let r=16;section(ws,r,2,16,'Evolución diaria');r++;
    const chartRows=Math.ceil((CH+10)/ptpx(RH));
    reserve(r,chartRows);
    const bars1=[{name:hasB?'Periodo':'Coste de consumo',color:C.navy,vals:days.map(d=>d.consumo)}];
    if(hasB)bars1.push({name:'Comparación',color:C.gold,vals:days.map((_,i)=>B.P[i]?B.P[i].consumo:null)});
    place(r,0,combo(ctx,{w:CW,h:CH,title:'Consumo por día',sub:'Coste de consumo (€) y coste por pax (€)',labels,sublabels:sublab,bars:bars1,line:{name:'€ por pax',color:C.gold,vals:days.map(d=>d.cpax),axis:v=>nf(v,2),fmt:v=>nf(v,2)},fmt:v=>nf(v,0)}));
    place(r,1,combo(ctx,{w:CW,h:CH,title:'Merma por día',sub:'Coste de la merma (€) y % sobre coste bruto',labels,sublabels:sublab,bars:[{name:'Merma (€)',color:C.red,vals:days.map(d=>d.merma)},...(hasB?[{name:'Comparación',color:'E8B4A6',vals:days.map((_,i)=>B.P[i]?B.P[i].merma:null)}]:[])],line:{name:'% s/ coste bruto',color:C.navy,vals:days.map(d=>d.bruto?d.merma/d.bruto*100:0),axis:v=>nf(v,0)+'%',fmt:v=>nf(v,1)+'%'},fmt:v=>nf(v,0)}));
    r+=chartRows;
    ws.getRow(r-1).addPageBreak();
    section(ws,r,2,16,'Familias');r++;reserve(r,chartRows);
    const keys=[...topF,'Otras'];
    const stacks=days.map(d=>{const s={};for(const g of A.fams){const v=g.byDay.get(d.fecha)||0;if(v>0){const k=topF.includes(g.cat)?g.cat:'Otras';s[k]=(s[k]||0)+v;}}return s;});
    place(r,0,stack(ctx,{w:CW,h:CH,title:'Consumo por día y familia',sub:'Coste de consumo (€) apilado por familia',labels,stacks,keys,colorOf:k=>k==='Otras'?FAM[8]:colorOf(k)}));
    const dItems=A.fams.filter(f=>f.eur>0).slice(0,8).map(f=>({name:f.cat,v:f.eur,color:colorOf(f.cat),right:nf(f.eur,0)+' € · '+nf(A.totC?f.eur/A.totC*100:0,1)+' %'}));
    const rest=A.totC-dItems.reduce((a,i)=>a+i.v,0);if(rest>0.5)dItems.push({name:'Otras',v:rest,color:FAM[8],right:nf(rest,0)+' € · '+nf(rest/A.totC*100,1)+' %'});
    place(r,1,donut(ctx,{w:CW,h:CH,title:'Reparto por familia',sub:'Peso de cada familia en el coste de consumo',items:dItems,total:nf(A.totC,0)+' €',centerLabel:'consumo total'}));
    r+=chartRows;
    section(ws,r,2,16,'Artículos que más pesan');r++;reserve(r,chartRows);
    const t10=rkList(A,'eur').slice(0,10),m10=rkList(A,'mer').slice(0,10);
    place(r,0,hbar(ctx,{w:CW,h:CH,title:'Top 10 por consumo (€)',sub:'Artículos con mayor coste de consumo del periodo',items:t10.map(o=>({name:o.p?o.p.descripcion:o.codigo,v:o.eur})),color:C.navy,fmt:v=>nf(v,0)+' €'}));
    place(r,1,hbar(ctx,{w:CW,h:CH,title:'Top 10 por merma (€)',sub:'Artículos donde más dinero se tira',items:m10.map(o=>({name:o.p?o.p.descripcion:o.codigo,v:o.merE})),color:C.red,fmt:v=>nf(v,0)+' €'}));
    r+=chartRows;
    ws.mergeCells(r,2,r+1,16);const nt=ws.getCell(r,2);
    nt.value='Criterio: Consumo = Expuesto + Reposición − Aprovechamiento − Merma. Coste bruto = Expuesto + Reposición − Aprovechamiento (incluye el coste de la merma). El precio de cada artículo queda congelado al cerrar el día. Las hojas «Por día», «Familias» y «Rankings» amplían este panel; «Detalle» y «Movimientos» contienen los datos de origen con fórmulas.';
    nt.font={name:FONT,size:9.5,italic:true,color:{argb:argb(C.mut)}};nt.alignment={wrapText:true,vertical:'top',indent:1};
    ws.getRow(r).height=24;ws.getRow(r+1).height=24;
    ws.pageSetup.printArea=`A1:Q${r+1}`;
  }

  /* ===== EVOLUCIÓN MENSUAL ===== */
  if(sEvo){
    const ws=sEvo,E=meta.evo,n=E.length,MES=['ene','feb','mar','abr','may','jun','jul','ago','sep','oct','nov','dic'],ML=m=>MES[+m.slice(5,7)-1]+' '+m.slice(2,4);
    const widths=[2,28,9,12,15,12,15,16,13,14,12];widths.forEach((w,i)=>ws.getColumn(i+1).width=w);
    banner(ws,1,2,11,'Evolución mensual',SUB,'Coste, PAX y coste por pax por mes  ·  cada mes usa su lista de precios  ·  hasta '+ddmmyyyy(meta.hasta)+'  ·  «Validado» = días validados por el jefe de cocina / días cerrados');ws.getRow(4).height=10;
    headRow(ws,5,2,['Mes','Días','PAX','Coste consumo (€)','€ / pax','Δ € / pax vs mes ant.','€ / pax a precios actuales','Merma (€)','Merma % s/ bruto','Validado'],{left:[0],h:40});
    const R0=6,RL=R0+n-1,RT=RL+1;
    E.forEach((m,i)=>{const r=R0+i,row=ws.getRow(r);row.height=22;
      const set=(c,v,o)=>{const x=row.getCell(c);x.value=v;bodyCell(x,o);};
      set(2,ML(m.mes)+(m.parcial?' (parcial)':''),{al:'left',bold:true});set(3,m.dias,{nf:INT,al:'center'});set(4,m.pax,{nf:INT});set(5,m.cons,{nf:EUR});
      set(6,{formula:`IF(D${r}=0,0,E${r}/D${r})`,result:m.pax?m.cons/m.pax:0},{nf:EUR,bold:true});
      set(7,i?{formula:`IF(F${r-1}=0,0,F${r}/F${r-1}-1)`,result:(m.pax&&E[i-1].pax&&E[i-1].cons)?(m.cons/m.pax)/(E[i-1].cons/E[i-1].pax)-1:0}:null,{nf:'+0.0%;-0.0%;0.0%',al:'center'});
      set(8,m.pax?m.consAct/m.pax:0,{nf:EUR});set(9,m.mer,{nf:EUR});set(10,m.bru?m.mer/m.bru:0,{nf:PCT});set(11,m.val+' / '+m.cerr,{al:'center'});});
    const sum=k=>E.reduce((a,m)=>a+m[k],0),tp=sum('pax'),tc=sum('cons'),ta=sum('consAct'),tm=sum('mer'),tb=sum('bru');
    const trow=ws.getRow(RT);trow.height=24;
    for(let c=2;c<=11;c++)totalCell(trow.getCell(c),{al:c===2?'left':c===3||c===11?'center':'right'});
    trow.getCell(2).value='Total periodo';trow.getCell(3).value={formula:`SUM(C${R0}:C${RL})`,result:sum('dias')};trow.getCell(4).value={formula:`SUM(D${R0}:D${RL})`,result:tp};trow.getCell(5).value={formula:`SUM(E${R0}:E${RL})`,result:tc};
    trow.getCell(6).value={formula:`IF(D${RT}=0,0,E${RT}/D${RT})`,result:tp?tc/tp:0};trow.getCell(8).value=tp?ta/tp:0;trow.getCell(9).value={formula:`SUM(I${R0}:I${RL})`,result:tm};trow.getCell(10).value=tb?tm/tb:0;trow.getCell(11).value=sum('val')+' / '+sum('cerr');
    trow.getCell(3).numFmt=INT;trow.getCell(4).numFmt=INT;trow.getCell(5).numFmt=EUR;trow.getCell(6).numFmt=EUR;trow.getCell(8).numFmt=EUR;trow.getCell(9).numFmt=EUR;trow.getCell(10).numFmt=PCT;
    if(n>1){bar(ws,`E${R0}:E${RL}`,'A9BCDD');ws.addConditionalFormatting({ref:`F${R0}:F${RL}`,rules:[{type:'colorScale',priority:P(),cfvo:[{type:'min'},{type:'max'}],color:[{argb:'FFFFFFFF'},{argb:'FFF2D4A8'}]}]});
      ws.addConditionalFormatting({ref:`G${R0+1}:G${RL}`,rules:[{type:'expression',priority:P(),formulae:[`G${R0+1}>0.05`],style:{font:{color:{argb:argb(C.bad)},bold:true}}},{type:'expression',priority:P(),formulae:[`G${R0+1}<-0.05`],style:{font:{color:{argb:argb(C.ok)},bold:true}}}]});}
    // gráficos
    let r=RT+2;section(ws,r,2,11,'Evolución por mes');r++;
    const pxw=w=>Math.round(w*7+5),totalW=widths.slice(1).reduce((a,w)=>a+pxw(w),0),CW=Math.floor((totalW-40)/2),CH=320,RH=20,rows=Math.ceil(CH/(RH*96/72))+1;
    for(let i=0;i<rows;i++)ws.getRow(r+i).height=RH;
    const labels=E.map(m=>ML(m.mes)),sublab=E.map(m=>m.dias+' días'+(m.parcial?' · parcial':''));
    const colAt=x=>{let acc=0;for(let i=0;i<widths.length;i++){const w=pxw(widths[i]);if(acc+w>x)return i+(x-acc)/w;acc+=w;}return widths.length-1;};
    const put=(png,x)=>{const id=wb.addImage({base64:png,extension:'png'});ws.addImage(id,{tl:{col:colAt(x),row:r-1+0.08},ext:{width:CW,height:CH},editAs:'oneCell'});};
    const x0=pxw(widths[0]);
    put(combo(ctx,{w:CW,h:CH,title:'Coste por pax por mes',sub:'Azul: precios de cada mes · Dorado: mismos consumos a precios actuales',labels,sublabels:sublab,bars:[{name:'€/pax real',color:C.navy,vals:E.map(m=>m.pax?m.cons/m.pax:0)},{name:'€/pax a precios actuales',color:C.gold,vals:E.map(m=>m.pax?m.consAct/m.pax:0)}],fmt:v=>nf(v,2)}),x0);
    put(combo(ctx,{w:CW,h:CH,title:'PAX por mes',sub:'Comensales totales de desayuno por mes',labels,sublabels:sublab,bars:[{name:'PAX',color:C.teal,vals:E.map(m=>m.pax)}],fmt:v=>nf(v,0)}),x0+CW+24);
    r+=rows+1;
    // familias x mes (€/pax)
    const cats=new Map();for(const m of E)for(const[k,v]of m.fam)cats.set(k,(cats.get(k)||0)+v);
    const fams=[...cats.entries()].sort((a,b)=>b[1]-a[1]).map(x=>x[0]);
    if(fams.length){section(ws,r,2,11,'€ por pax por familia y mes');r++;
      const mcols=Math.min(n,10);const Es=E.slice(-mcols);
      headRow(ws,r,2,['Familia',...Es.map(m=>ML(m.mes))],{left:[0],h:28});
      fams.forEach((f,i)=>{const rr=r+1+i,row=ws.getRow(rr);row.height=20;const c0=row.getCell(2);c0.value=f;bodyCell(c0,{al:'left',bold:true});
        Es.forEach((m,j)=>{const c=row.getCell(3+j);c.value=m.pax?(m.fam.get(f)||0)/m.pax:0;bodyCell(c,{nf:EUR});});});
      ws.addConditionalFormatting({ref:`C${r+1}:${ws.getColumn(2+mcols).letter}${r+fams.length}`,rules:[{type:'colorScale',priority:P(),cfvo:[{type:'min'},{type:'max'}],color:[{argb:'FFFFFFFF'},{argb:'FFA9BCDD'}]}]});
      r+=fams.length+2;}
    ws.pageSetup.printArea=`A1:L${r}`;
  }
  const buf=await wb.xlsx.writeBuffer();
  return buf;
}
root.KDExport={build};
})(typeof window!=='undefined'?window:globalThis);

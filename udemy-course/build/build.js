const pptxgen=require('pptxgenjs');const D=require('./data.js');const fs=require('fs');
const INK="1B2A41",AMB="F2A93B",TEAL="1F7A8C",LG="F3F5F7",TX="22303C",MU="5F6B78",RED="C8553D",GRN="3B8E6B",F="Meiryo";
const p=new pptxgen();p.layout='LAYOUT_16x9';p.title="75歳、AIに挑む";
const sh=()=>({type:'outer',color:'000000',opacity:.12,blur:6,offset:2,angle:90});
let md="# 75歳、AIに挑む — スライド原稿（ナレーション台本）\n\n> 各スライドの「ナレーション」は講師が読み上げる原稿です。「AIの回答（要約）」は実際の会話ログのスクリーンショットに差し替えてください。\n\n";
let total=0;D.forEach(s=>total+=s.min);
md+=`想定総時間：${total}分（企画書では約240分）。スライド：`;
let count=0;
function foot(sl,sec,dark){sl.addText(sec,{x:.5,y:5.2,w:6,h:.3,fontSize:10,color:dark?"9FB0C8":MU,fontFace:F,margin:0,isTextBox:true});}
D.forEach((S,si)=>{
 S.slides.forEach((d,i)=>{count++;
  const sl=p.addSlide();sl.addNotes(d.n);
  const num=`${si}`;
  const title=(txt)=>sl.addText(txt,{x:.5,y:.3,w:9,h:.8,fontSize:26,bold:true,color:INK,fontFace:F,margin:0,valign:'middle',isTextBox:true});
  if(d.t==="title"){
   sl.background={color:INK};
   sl.addText("Udemy講座",{x:.7,y:.7,w:4,h:.4,fontSize:14,color:AMB,bold:true,fontFace:F,margin:0,isTextBox:true});
   sl.addText(d.title,{x:.7,y:1.3,w:8.6,h:1.3,fontSize:54,bold:true,color:"FFFFFF",fontFace:F,margin:0,isTextBox:true});
   sl.addText(d.sub,{x:.7,y:2.8,w:8.6,h:.9,fontSize:20,color:AMB,fontFace:F,margin:0,isTextBox:true});
   sl.addText(d.tag,{x:.7,y:4.0,w:8.6,h:.7,fontSize:14,color:"DCE4F0",fontFace:F,margin:0,isTextBox:true});
   sl.addText("講師：梟風　2026年11月版",{x:.7,y:5.0,w:6,h:.3,fontSize:11,color:"9FB0C8",fontFace:F,margin:0,isTextBox:true});
  } else if(d.t==="take"){
   sl.background={color:INK};
   sl.addText(d.label,{x:.7,y:.9,w:8.6,h:.4,fontSize:14,bold:true,color:AMB,fontFace:F,margin:0,isTextBox:true});
   sl.addText(d.text,{x:.7,y:1.6,w:8.6,h:1.8,fontSize:36,bold:true,color:"FFFFFF",fontFace:F,margin:0,valign:'middle',isTextBox:true});
   sl.addText(d.sub,{x:.7,y:3.7,w:8.6,h:.8,fontSize:16,color:"DCE4F0",fontFace:F,margin:0,isTextBox:true});
   foot(sl,S.sec,true);
  } else {
   sl.background={color:"FFFFFF"};title(d.title);foot(sl,S.sec,false);
   sl.addText(`${count}`,{x:9,y:5.2,w:.5,h:.3,fontSize:10,color:MU,align:'right',fontFace:F,margin:0,isTextBox:true});
   if(d.t==="pts"){
    let y0=1.3;if(d.lead){sl.addText(d.lead,{x:.5,y:1.15,w:9,h:.4,fontSize:16,color:MU,fontFace:F,margin:0,isTextBox:true});y0=1.8;}
    const n=d.items.length;const cols=n<=3?n:(n===4?2:3);
    const rows=Math.ceil(n/cols);const gap=.3,W=9;const cw=(W-gap*(cols-1))/cols;const H=5.0-y0;const ch=Math.min(rows===1?2.6:(H-gap*(rows-1))/rows,2.8);
    d.items.forEach((it,k)=>{const r=Math.floor(k/cols),c=k%cols;const x=.5+c*(cw+gap),y=y0+r*(ch+gap);
     sl.addShape(p.ShapeType.roundRect,{x,y,w:cw,h:ch,fill:{color:LG},rectRadius:.12,shadow:sh()});
     const cp=rows>1;const hy=cp?y+.2:y+.8,dy=cp?y+.8:y+1.3,hx=cp?x+.85:x+.2;
     sl.addShape(p.ShapeType.ellipse,{x:x+.2,y:y+.2,w:.5,h:.5,fill:{color:k%2?TEAL:AMB}});
     sl.addText(`${k+1}`,{x:x+.2,y:y+.2,w:.5,h:.5,align:'center',valign:'middle',fontSize:16,bold:true,color:k%2?"FFFFFF":INK,fontFace:F,margin:0,isTextBox:true});
     sl.addText(it[0],{x:hx,y:cp?y+.2:hy,w:cw-(hx-x)-.2,h:cp?.5:.5,fontSize:n>4?16:18,bold:true,color:INK,fontFace:F,margin:0,valign:cp?'middle':'top',isTextBox:true});
     sl.addText(it[1],{x:x+.2,y:dy,w:cw-.4,h:ch-(dy-y)-.1,fontSize:14,color:TX,fontFace:F,margin:0,valign:'top',isTextBox:true});});
   }
   if(d.t==="chat"){
    sl.addShape(p.ShapeType.roundRect,{x:3.3,y:1.3,w:6.2,h:.95,fill:{color:AMB},rectRadius:.15});
    sl.addText("私の依頼",{x:3.5,y:1.35,w:3,h:.3,fontSize:11,bold:true,color:INK,fontFace:F,margin:0,isTextBox:true});
    sl.addText("「"+d.q+"」",{x:3.5,y:1.65,w:5.8,h:.55,fontSize:16,bold:true,color:INK,fontFace:F,margin:0,valign:'middle',isTextBox:true});
    sl.addShape(p.ShapeType.roundRect,{x:.5,y:2.45,w:7.0,h:1.35,fill:{color:LG},rectRadius:.15,shadow:sh()});
    sl.addText("AIの回答（要約）",{x:.7,y:2.5,w:3,h:.3,fontSize:11,bold:true,color:TEAL,fontFace:F,margin:0,isTextBox:true});
    sl.addText(d.a,{x:.7,y:2.8,w:6.6,h:.95,fontSize:14,color:TX,fontFace:F,margin:0,valign:'top',isTextBox:true});
    sl.addShape(p.ShapeType.roundRect,{x:.5,y:4.0,w:9,h:.9,fill:{color:INK},rectRadius:.12});
    sl.addText([{text:"学びのポイント　",options:{bold:true,color:AMB}},{text:d.tip,options:{color:"FFFFFF"}}],{x:.75,y:4.0,w:8.5,h:.9,fontSize:16,fontFace:F,margin:0,valign:'middle',isTextBox:true});
   }
   if(d.t==="steps"){
    const n=d.steps.length;const gap=.2;const cw=(9-gap*(n-1))/n;
    d.steps.forEach((st,k)=>{const x=.5+k*(cw+gap);
     sl.addShape(p.ShapeType.roundRect,{x,y:1.5,w:cw,h:2.4,fill:{color:LG},rectRadius:.12,shadow:sh()});
     sl.addShape(p.ShapeType.ellipse,{x:x+cw/2-.28,y:1.7,w:.56,h:.56,fill:{color:TEAL}});
     sl.addText(`${k+1}`,{x:x+cw/2-.28,y:1.7,w:.56,h:.56,align:'center',valign:'middle',fontSize:18,bold:true,color:"FFFFFF",fontFace:F,margin:0,isTextBox:true});
     sl.addText(st[0],{x:x+.1,y:2.4,w:cw-.2,h:.5,align:'center',fontSize:n>4?16:18,bold:true,color:INK,fontFace:F,margin:0,isTextBox:true});
     sl.addText(st[1],{x:x+.1,y:2.95,w:cw-.2,h:.9,align:'center',valign:'top',fontSize:14,color:TX,fontFace:F,margin:0,isTextBox:true});});
    if(d.note)sl.addText(d.note,{x:.5,y:4.2,w:9,h:.7,fontSize:16,bold:true,color:INK,fontFace:F,margin:0,valign:'middle',isTextBox:true});
   }
   if(d.t==="tr"){
    sl.addShape(p.ShapeType.roundRect,{x:.5,y:1.3,w:4.3,h:2.5,fill:{color:"FBEAE6"},rectRadius:.12});
    sl.addText("つまずき",{x:.7,y:1.4,w:3,h:.35,fontSize:14,bold:true,color:RED,fontFace:F,margin:0,isTextBox:true});
    sl.addText(d.problem,{x:.7,y:1.8,w:3.9,h:1.9,fontSize:16,color:TX,fontFace:F,margin:0,valign:'top',isTextBox:true});
    sl.addText("→",{x:4.8,y:2.2,w:.4,h:.5,align:'center',fontSize:24,bold:true,color:INK,fontFace:F,margin:0,isTextBox:true});
    sl.addShape(p.ShapeType.roundRect,{x:5.2,y:1.3,w:4.3,h:2.5,fill:{color:"E3F2EB"},rectRadius:.12});
    sl.addText("AIに伝えたら",{x:5.4,y:1.4,w:3,h:.35,fontSize:14,bold:true,color:GRN,fontFace:F,margin:0,isTextBox:true});
    sl.addText(d.fix,{x:5.4,y:1.8,w:3.9,h:1.9,fontSize:16,color:TX,fontFace:F,margin:0,valign:'top',isTextBox:true});
    sl.addShape(p.ShapeType.roundRect,{x:.5,y:4.0,w:9,h:.9,fill:{color:INK},rectRadius:.12});
    sl.addText([{text:"学びのポイント　",options:{bold:true,color:AMB}},{text:d.tip,options:{color:"FFFFFF"}}],{x:.75,y:4.0,w:8.5,h:.9,fontSize:16,fontFace:F,margin:0,valign:'middle',isTextBox:true});
   }
   if(d.t==="cmp"){
    [[d.lh,d.l,.5,LG,MU],[d.rh,d.r,5.2,"E6F2F4",TEAL]].forEach(([h,arr,x,bg,hc])=>{
     sl.addShape(p.ShapeType.roundRect,{x,y:1.3,w:4.3,h:3.6,fill:{color:bg},rectRadius:.12});
     sl.addText(h,{x:x+.25,y:1.45,w:3.8,h:.5,fontSize:20,bold:true,color:hc,fontFace:F,margin:0,isTextBox:true});
     sl.addText(arr.map((t,k)=>({text:t,options:{bullet:true,breakLine:k<arr.length-1,paraSpaceAfter:8}})),{x:x+.25,y:2.1,w:3.8,h:2.6,fontSize:16,color:TX,fontFace:F,margin:0,valign:'top',isTextBox:true});});
   }
   if(d.t==="stat"){
    const n=d.stats.length;const gap=.3;const cw=(9-gap*(n-1))/n;
    d.stats.forEach((s,k)=>{const x=.5+k*(cw+gap);
     sl.addShape(p.ShapeType.roundRect,{x,y:1.5,w:cw,h:3,fill:{color:LG},rectRadius:.12,shadow:sh()});
     sl.addText(s[0],{x,y:1.9,w:cw,h:1.2,align:'center',valign:'middle',fontSize:s[0].length>3?36:64,bold:true,color:TEAL,fontFace:F,margin:0,isTextBox:true});
     sl.addText(s[1],{x:x+.15,y:3.3,w:cw-.3,h:.9,align:'center',valign:'top',fontSize:16,color:TX,fontFace:F,margin:0,isTextBox:true});});
   }
  }
 });
});
// md
md=md.replace(/スライド：$/,"");
md=md.replace("想定総時間","想定総時間")+"\n";
let k=0;
D.forEach((S,si)=>{
 md+=`\n## ${si===0?"オリエンテーション":"セクション"+si+"「"+S.sec+(S.goal?"――"+S.goal:"")+"」"}（約${S.min}分）\n`;
 S.slides.forEach(d=>{k++;md+=`\n### スライド${k}：${d.title||d.text}\n`;
  const body=d.t==="chat"?`- 私の依頼：「${d.q}」\n- AIの回答（要約）：${d.a}\n- 学びのポイント：${d.tip}`:d.t==="tr"?`- つまずき：${d.problem}\n- 解決：${d.fix}\n- 学びのポイント：${d.tip}`:d.t==="pts"?d.items.map(i=>`- ${i[0]}：${i[1]}`).join("\n"):d.t==="steps"?d.steps.map((s,j)=>`${j+1}. ${s[0]}：${s[1]}`).join("\n")+(d.note?`\n- 補足：${d.note}`:""):d.t==="cmp"?`- ${d.lh}：${d.l.join("／")}\n- ${d.rh}：${d.r.join("／")}`:d.t==="stat"?d.stats.map(s=>`- ${s[0]} ${s[1]}`).join("\n"):d.t==="take"?`- ${d.text}（${d.sub}）`:`- ${d.sub}`;
  md+=`\n**画面の内容**\n${body}\n\n**ナレーション**\n> ${d.n}\n`;});
});
fs.writeFileSync('../スライド原稿.md',md);
p.writeFile({fileName:'../75歳AIに挑む_スライド.pptx'}).then(()=>console.log('slides',count));

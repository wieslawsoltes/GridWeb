import { safeColor } from './layout.js';
export function chartData(sheet, chart) {
  const r=chart.range,last=Math.min(r.r2,r.r1+200),cols=Math.min(r.c2,r.c1+8),labels=[],series=[];
  for(let c=r.c1+1;c<=cols;c++)series.push({name:sheet.GetCell(r.r1,c).Text,values:[]});
  for(let row=r.r1+1;row<=last;row++){labels.push(sheet.GetCell(row,r.c1).Value);series.forEach((s,i)=>{const v=sheet.GetCell(row,r.c1+1+i).Value;s.values.push(typeof v==='number'&&Number.isFinite(v)?v:null);});}
  return{labels,series,truncated:r.r2>last||r.c2>cols};
}
const palette=['#107c41','#5b9bd5','#e4a343','#8064a2','#49a68e','#cf6679','#526a80'];
export function drawChart(ctx,sheet,chart,rect,{dark=false,selected=false}={}){
  const {x,y,width,height}=rect,{labels,series,truncated}=chartData(sheet,chart),foreground=dark?'#e7eee9':'#263b2c',muted=dark?'#a9b8ae':'#6a7b70',background=dark?'#202a24':'#fff';
  ctx.save();ctx.beginPath();ctx.rect(x,y,width,height);ctx.clip();ctx.fillStyle=background;ctx.fillRect(x,y,width,height);ctx.strokeStyle=selected?'#107c41':dark?'#4a5a4e':'#d7e3da';ctx.lineWidth=selected?2:1;ctx.strokeRect(x+.5,y+.5,width-1,height-1);
  ctx.fillStyle=foreground;ctx.font='600 14px system-ui';ctx.textAlign='left';ctx.fillText(chart.title,x+18,y+26,width-36);
  if(!labels.length||!series.length){ctx.font='12px system-ui';ctx.fillText('Select a label column and numeric series.',x+20,y+60);ctx.restore();return;}
  const p={x:x+55,y:y+48,w:Math.max(20,width-78),h:Math.max(20,height-106)},all=series.flatMap(s=>s.values).filter(v=>v!=null),lo=Math.min(0,...all),hi=Math.max(0,...all),span=hi-lo||1;
  const fy=v=>p.y+p.h-(v-lo)/span*p.h,fx=i=>p.x+(i+.5)*p.w/labels.length;
  ctx.font='10px system-ui';ctx.textBaseline='middle';
  if(chart.type==='pie'){
    const values=series[0].values.map(v=>Math.max(0,v??0)),total=values.reduce((a,b)=>a+b,0);let angle=-Math.PI/2;const radius=Math.min(p.w,p.h)/2;
    values.forEach((v,i)=>{const end=angle+(total?v/total:0)*Math.PI*2;ctx.beginPath();ctx.moveTo(p.x+p.w/2,p.y+p.h/2);ctx.arc(p.x+p.w/2,p.y+p.h/2,radius,angle,end);ctx.closePath();ctx.fillStyle=palette[i%palette.length];ctx.fill();angle=end;});
    if(!total){ctx.fillStyle=muted;ctx.fillText('Pie chart requires positive values.',p.x,p.y+25);}
  }else if(chart.type==='bar'){
    const fw=v=>p.x+(v-lo)/span*p.w,group=p.h/labels.length;
    for(let i=0;i<5;i++){const val=lo+i*span/4,xx=fw(val);ctx.strokeStyle=dark?'#354338':'#e8eee9';ctx.beginPath();ctx.moveTo(xx,p.y);ctx.lineTo(xx,p.y+p.h);ctx.stroke();ctx.fillStyle=muted;ctx.textAlign='center';ctx.fillText(abbreviate(val),xx,p.y+p.h+13);}
    series.forEach((s,si)=>s.values.forEach((v,i)=>{if(v==null)return;ctx.fillStyle=palette[si%palette.length];const h=Math.max(1,group*.7/series.length),yy=p.y+i*group+group*.15+si*h;ctx.fillRect(Math.min(fw(0),fw(v)),yy,Math.abs(fw(v)-fw(0)),h-1);}));
    ctx.fillStyle=muted;ctx.textAlign='right';labels.forEach((l,i)=>{if(i%Math.max(1,Math.ceil(labels.length/12))===0)ctx.fillText(String(l??''),p.x-7,p.y+(i+.5)*group,48);});
  }else{
    for(let i=0;i<5;i++){const val=lo+i*span/4,yy=fy(val);ctx.strokeStyle=dark?'#354338':'#e8eee9';ctx.beginPath();ctx.moveTo(p.x,yy);ctx.lineTo(p.x+p.w,yy);ctx.stroke();ctx.fillStyle=muted;ctx.textAlign='right';ctx.fillText(abbreviate(val),p.x-8,yy);}
    series.forEach((s,si)=>{
      ctx.fillStyle=ctx.strokeStyle=palette[si%palette.length];ctx.lineWidth=2;
      if(chart.type==='column'){const group=p.w/labels.length,bw=group*.7/series.length;s.values.forEach((v,i)=>{if(v==null)return;ctx.fillRect(fx(i)-group*.35+si*bw,Math.min(fy(0),fy(v)),Math.max(1,bw-1),Math.max(1,Math.abs(fy(v)-fy(0))));});}
      else if(chart.type==='scatter'){const xs=labels.map(Number),min=Math.min(...xs.filter(Number.isFinite)),max=Math.max(...xs.filter(Number.isFinite));s.values.forEach((v,i)=>{if(v==null||!Number.isFinite(xs[i]))return;ctx.beginPath();ctx.arc(p.x+(xs[i]-min)/(max-min||1)*p.w,fy(v),3,0,Math.PI*2);ctx.fill();});}
      else{ctx.beginPath();let started=false;s.values.forEach((v,i)=>{if(v==null){started=false;return;}if(!started){ctx.moveTo(fx(i),fy(v));started=true;}else ctx.lineTo(fx(i),fy(v));});ctx.stroke();if(chart.type==='area'&&s.values.every(v=>v!=null)){ctx.lineTo(fx(labels.length-1),fy(0));ctx.lineTo(fx(0),fy(0));ctx.closePath();ctx.globalAlpha=.14;ctx.fill();ctx.globalAlpha=1;}}
    });
    ctx.fillStyle=muted;ctx.textAlign='center';labels.forEach((l,i)=>{if(i%Math.max(1,Math.ceil(labels.length/8))===0)ctx.fillText(String(l??''),fx(i),p.y+p.h+15,p.w/Math.min(labels.length,8)-4);});
  }
  ctx.textAlign='left';ctx.font='10px system-ui';let lx=x+18;
  const legend=chart.type==='pie'?labels.map((name,i)=>({name:String(name??''),i})):series.map((s,i)=>({name:s.name,i}));
  for(const item of legend.slice(0,5)){if(lx>x+width-70)break;ctx.fillStyle=palette[item.i%palette.length];ctx.fillRect(lx,y+height-24,8,8);ctx.fillStyle=muted;ctx.fillText(item.name,lx+13,y+height-20,90);lx+=Math.min(115,ctx.measureText(item.name).width+28);}
  if(truncated){ctx.textAlign='right';ctx.fillText('First 200 rows / 8 series',x+width-10,y+height-7);}
  if(selected){ctx.fillStyle='#107c41';ctx.fillRect(x+width-7,y+height-7,7,7);}
  ctx.restore();
}
function abbreviate(n){return Math.abs(n)>=1e6?(n/1e6).toFixed(1)+'M':Math.abs(n)>=1000?(n/1000).toFixed(0)+'k':Number(n.toPrecision(3)).toString();}

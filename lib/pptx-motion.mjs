// 给 pptxgenjs 生成的幻灯片补上“切换效果”和“进入动画”（pptxgenjs 本身不支持动画）。
// 写入的是 PowerPoint 标准结构（p:transition、p:timing），WPS 演示和 PowerPoint 都能识别；
// 动画在翻到这一页后自动依次播放，不需要点击。
import JSZip from 'jszip';

const SPEED={fast:'fast',med:'med',slow:'slow'};
const DIRS={l:'l',r:'r',u:'u',d:'d'};
export function transitionXml(t){
  if(!t||!t.type)return '';
  const spd=SPEED[t.speed]||'med',dir=DIRS[t.dir]||'l';
  const inner={fade:'<p:fade/>',push:`<p:push dir="${dir}"/>`,wipe:`<p:wipe dir="${dir}"/>`,cover:`<p:cover dir="${dir}"/>`,split:'<p:split orient="horz" dir="out"/>'}[t.type];
  if(!inner)return '';
  return `<p:transition spd="${spd}">${inner}</p:transition>`;
}
function bhvr(id,spid,dur,attr){return `<p:cBhvr${attr?' additive="base"':''}><p:cTn id="${id}" dur="${dur}" fill="hold"/><p:tgtEl><p:spTgt spid="${spid}"/></p:tgtEl>${attr?`<p:attrNameLst><p:attrName>${attr}</p:attrName></p:attrNameLst>`:''}</p:cBhvr>`;}
function anim(id,spid,dur,attr,from,to){
  const val=v=>typeof v==='number'?`<p:fltVal val="${v}"/>`:`<p:strVal val="${v}"/>`;
  return `<p:anim calcmode="lin" valueType="num">${bhvr(id,spid,dur,attr)}<p:tavLst><p:tav tm="0"><p:val>${val(from)}</p:val></p:tav><p:tav tm="100000"><p:val>${val(to)}</p:val></p:tav></p:tavLst></p:anim>`;
}
// 每种效果：PowerPoint 预设编号＋实际行为。行为决定播放效果，预设编号让软件在动画窗格里显示正确的名字。
const EFFECTS={
  fade:{preset:10,sub:0,body:(n,s,d)=>`<p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="${n()}" dur="${d}"/><p:tgtEl><p:spTgt spid="${s}"/></p:tgtEl></p:cBhvr></p:animEffect>`},
  wipe:{preset:22,sub:8,body:(n,s,d)=>`<p:animEffect transition="in" filter="wipe(left)"><p:cBhvr><p:cTn id="${n()}" dur="${d}"/><p:tgtEl><p:spTgt spid="${s}"/></p:tgtEl></p:cBhvr></p:animEffect>`},
  rise:{preset:42,sub:0,body:(n,s,d)=>`<p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="${n()}" dur="${d}"/><p:tgtEl><p:spTgt spid="${s}"/></p:tgtEl></p:cBhvr></p:animEffect>`+anim(n(),s,d,'ppt_x','#ppt_x','#ppt_x')+anim(n(),s,d,'ppt_y','#ppt_y+.1','#ppt_y')},
  zoom:{preset:53,sub:16,body:(n,s,d)=>anim(n(),s,d,'ppt_w',0,'#ppt_w')+anim(n(),s,d,'ppt_h',0,'#ppt_h')+`<p:animEffect transition="in" filter="fade"><p:cBhvr><p:cTn id="${n()}" dur="${d}"/><p:tgtEl><p:spTgt spid="${s}"/></p:tgtEl></p:cBhvr></p:animEffect>`}
};
// items：[{spid, effect, delay(ms), dur(ms)}]，按出现顺序；第一项接在切换之后，其余同时开始、按 delay 错开。
export function timingXml(items){
  const list=items.filter(x=>x&&x.spid&&EFFECTS[x.effect]);
  if(!list.length)return '';
  let id=4;const next=()=>++id;
  const effects=list.map((x,i)=>{
    const e=EFFECTS[x.effect],ctn=next(),dur=Math.max(200,Math.min(2000,Math.round(x.dur||500))),delay=Math.max(0,Math.min(5000,Math.round(x.delay||0)));
    const setId=next();
    const set=`<p:set><p:cBhvr><p:cTn id="${setId}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn><p:tgtEl><p:spTgt spid="${x.spid}"/></p:tgtEl><p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="visible"/></p:to></p:set>`;
    return `<p:par><p:cTn id="${ctn}" presetID="${e.preset}" presetClass="entr" presetSubtype="${e.sub}" fill="hold" grpId="0" nodeType="${i===0?'afterEffect':'withEffect'}"><p:stCondLst><p:cond delay="${delay}"/></p:stCondLst><p:childTnLst>${set}${e.body(next,x.spid,dur)}</p:childTnLst></p:cTn></p:par>`;
  }).join('');
  const bld=[...new Set(list.map(x=>x.spid))].map(s=>`<p:bldP spid="${s}" grpId="0" animBg="1"/>`).join('');
  return `<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst><p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst><p:par><p:cTn id="3" fill="hold"><p:stCondLst><p:cond delay="indefinite"/><p:cond evt="onBegin" delay="0"><p:tn val="2"/></p:cond></p:stCondLst><p:childTnLst><p:par><p:cTn id="4" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>${effects}</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn><p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst><p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq></p:childTnLst></p:cTn></p:par></p:tnLst><p:bldLst>${bld}</p:bldLst></p:timing>`;
}
// 读出幻灯片里每个形状的 id（按 name 对应），把切换和动画插到 </p:sld> 之前。
export function shapeIds(xml){const map=new Map();for(const m of xml.matchAll(/<p:cNvPr id="(\d+)" name="([^"]*)"/g))map.set(m[2],m[1]);return map;}
export function addMotionToSlide(xml,{transition,items}){
  if(/<p:timing\b|<p:transition\b/.test(xml))return xml;
  const ids=shapeIds(xml),resolved=(items||[]).map(x=>({...x,spid:ids.get(x.name)})).filter(x=>x.spid);
  const extra=transitionXml(transition)+timingXml(resolved);
  if(!extra)return xml;
  const at=xml.lastIndexOf('</p:sld>');if(at<0)return xml;
  // p:transition / p:timing 必须排在 p:extLst 前面
  const ext=xml.lastIndexOf('<p:extLst>',at),pos=ext>xml.lastIndexOf('</p:cSld>')?ext:at;
  return xml.slice(0,pos)+extra+xml.slice(pos);
}
export async function addMotion(buffer,slides){
  const zip=await JSZip.loadAsync(buffer);
  for(let i=0;i<slides.length;i++){
    const name=`ppt/slides/slide${i+1}.xml`,file=zip.file(name);if(!file)continue;
    zip.file(name,addMotionToSlide(await file.async('string'),slides[i]));
  }
  return zip.generateAsync({type:'nodebuffer',compression:'DEFLATE'});
}

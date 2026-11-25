(()=>{'use strict';const $=id=>document.getElementById(id),NS='http://www.w3.org/2000/svg';let token='',plan=null,revision=0,readyRevision=-1,sequence=0,timer,exporting=false,pending=null;
const LIMIT={session:8000,preview:20000,export:60000,open:90000};
const TIMEOUT={preview:'整理版式用时过长，已停止等待。请稍后点击“生成预览”重试。',export:'生成 PPTX 用时过长，已停止等待。请稍后重试；内容较多时可先拆分大纲。',open:'等待 WPS 启动的时间较长，已停止等待。请先看看 WPS 中是否已出现新演示；没有的话再点击重试。',session:'本地服务没有响应，请重新打开办公工具箱后再试。'};
// 企业品牌（主色、Logo）和图文页图片：只保存在这台电脑的浏览器存储里，随请求发给本机服务。
const BRAND_KEY='officeToolbox.studio.brand';let brand={on:'',color:'#0B5FFF',logo:''},images=[];
try{const b=JSON.parse(localStorage.getItem(BRAND_KEY)||'null');if(b&&typeof b==='object')brand={...brand,...b};}catch(e){}
function saveBrand(){try{localStorage.setItem(BRAND_KEY,JSON.stringify(brand));}catch(e){status('浏览器存储空间不足，品牌设置只在本次有效。',true);}}
function input(){const d={title:$('deckTitle').value,subtitle:$('subtitle').value,outline:$('outline').value,theme:document.querySelector('input[name=theme]:checked').value,images};if(brand.on)d.brand={color:brand.color,logo:brand.logo||undefined};return d;}
function status(text,error=false,busy=false){const box=$('status');box.textContent=text;box.className=error?'error':busy?'busy':'';box.setAttribute('aria-busy',String(busy));}
function ready(){const ok=!!plan&&readyRevision===revision&&!exporting;$('download').disabled=!ok;$('openWps').disabled=!ok;}
function stale(on){if(!plan)return;$('pages').classList.toggle('stale',on);$('pageCount').textContent=plan.pages.length+' 页'+(on?' · 待更新':'');}
function lock(on){$('previewButton').disabled=on;document.querySelectorAll('input,textarea,#example').forEach(el=>{el.disabled=on;});}
// 带超时的本地请求：超时、断连、被新请求取代都会给出明确结果，不会一直等待。
async function request(url,options,action,signal,read){
  const control=new AbortController();let timedOut=false;const stop=setTimeout(()=>{timedOut=true;control.abort();},LIMIT[action]),relay=()=>control.abort();
  if(signal)signal.addEventListener('abort',relay);
  try{const response=await fetch(url,{...options,signal:control.signal});return {response,body:await read(response)};}
  catch(e){if(timedOut)throw Error(TIMEOUT[action]);if(signal&&signal.aborted){const cancelled=Error('已取消');cancelled.cancelled=true;throw cancelled;}throw Error('无法连接本地服务，请确认办公工具箱仍在运行后重试。');}
  finally{clearTimeout(stop);if(signal)signal.removeEventListener('abort',relay);}
}
async function api(action,data,signal,retry=true){
  if(!token){const {response,body}=await request('/api/session',{},'session',signal,r=>r.json().catch(()=>null));if(!response.ok||!body||!body.token)throw Error('本地服务未连接，请重新打开办公工具箱。');token=body.token;}
  const {response,body}=await request('/api/studio/'+action,{method:'POST',headers:{'Content-Type':'application/json','X-Coach-Token':token},body:JSON.stringify(data)},action,signal,r=>r.ok&&action==='export'?r.blob():r.json().catch(()=>null));
  if(response.status===403&&retry){token='';return api(action,data,signal,false);}
  if(!response.ok)throw Error((body&&body.error)||'处理失败，请重试。');
  return body;
}
function element(tag,attrs){const node=document.createElementNS(NS,tag);Object.entries(attrs).forEach(([key,value])=>node.setAttribute(key,value));return node;}
const FACES={'Microsoft YaHei':"'Microsoft YaHei','PingFang SC','Noto Sans CJK SC',sans-serif",'SimSun':"'SimSun','Songti SC','Noto Serif CJK SC',serif",'Yu Gothic':"'Yu Gothic','Hiragino Sans','Noto Sans CJK JP',sans-serif",'Yu Mincho':"'Yu Mincho','Hiragino Mincho ProN','Noto Serif CJK JP',serif"};
const ROLES={cover:'封面',agenda:'目录',divider:'章节页',content:'内容页',closing:'结束页'},KINDS={grid:'卡片',stat:'数据',timeline:'时间线',quote:'引言',bullets:'要点'};
// 页面文字由 i18n.js 自动翻译，这里保持中文原文，避免二次翻译。
function tr(x){return x;}
function shape(item){const op=String(1-(item.alpha||0)/100);
  if(item.type==='rect'){const r=element('rect',{x:item.x,y:item.y,width:item.w,height:item.h,rx:item.radius||0,fill:'#'+item.fill,'fill-opacity':op});if(item.line){r.setAttribute('stroke','#'+item.line);r.setAttribute('stroke-width',item.lineW||1);r.setAttribute('stroke-opacity',String(1-(item.lineAlpha||0)/100));}return r;}
  if(item.type==='ellipse')return element('ellipse',{cx:item.x+item.w/2,cy:item.y+item.h/2,rx:item.w/2,ry:item.h/2,fill:'#'+item.fill,'fill-opacity':op});
  if(item.type==='image'){const im=element('image',{x:item.x,y:item.y,width:item.w,height:item.h,preserveAspectRatio:'xMidYMid meet'});im.setAttribute('href',item.data);return im;}
  if(item.type==='chart'){const g=element('g',{}),max=Math.max(...item.values,0)||1,n=item.values.length,pad=60,bw=(item.w-pad)/n,base=item.y+item.h-40;
    g.append(element('line',{x1:item.x+pad-10,y1:base,x2:item.x+item.w,y2:base,stroke:'#'+item.muted,'stroke-width':1}));
    item.values.forEach((v,i)=>{const h=Math.max(0,(item.h-90)*v/max),x=item.x+pad+i*bw+bw*.2;g.append(element('rect',{x,y:base-h,width:bw*.6,height:h,fill:'#'+item.color}));
      const t=element('text',{x:x+bw*.3,y:base-h-10,'text-anchor':'middle','font-size':16,fill:'#'+item.text});t.textContent=String(v);g.append(t);
      const l=element('text',{x:x+bw*.3,y:base+26,'text-anchor':'middle','font-size':16,fill:'#'+item.text});l.textContent=item.labels[i];g.append(l);});return g;}
  if(item.type==='line')return element('line',{x1:item.x,y1:item.y,x2:item.x+item.w,y2:item.y+item.h,stroke:'#'+item.color,'stroke-width':item.width||1,'stroke-opacity':op});
  const anchor=item.align==='center'?'middle':item.align==='right'?'end':'start',x=item.align==='center'?item.x+item.w/2:item.align==='right'?item.x+item.w:item.x,lh=item.lh||1.3;
  const text=element('text',{x,'text-anchor':anchor,'font-family':FACES[item.face]||FACES['Microsoft YaHei'],'font-size':item.fontSize,'font-weight':item.bold?700:400,'font-style':item.italic?'italic':'normal',fill:'#'+item.color,'fill-opacity':op});
  item.lines.forEach((line,i)=>{const span=element('tspan',{x,y:item.y+item.fontSize*.92+i*item.fontSize*lh});span.textContent=line;text.append(span);});return text;}
function render(data){$('pages').replaceChildren();$('empty').hidden=true;$('pageCount').textContent=data.pages.length+' 页';data.pages.forEach((page,index)=>{const article=document.createElement('article'),label=document.createElement('div'),number=document.createElement('span'),role=document.createElement('span');article.className='slide-card';label.className='slide-label';number.textContent=String(index+1).padStart(2,'0');role.textContent=tr(ROLES[page.role]||'内容页')+(page.layout?' · '+tr(KINDS[page.layout]||''):'');label.append(number,role);const svg=element('svg',{viewBox:`0 0 ${data.width} ${data.height}`,role:'img','aria-label':`第 ${index+1} 页预览`});svg.append(element('rect',{x:0,y:0,width:data.width,height:data.height,fill:'#'+page.background}));page.elements.forEach(item=>svg.append(shape(item)));article.append(label,svg);$('pages').append(article);});$('pages').classList.remove('stale');$('warnings').hidden=false;$('warnings').textContent=data.warnings.join(' ');}
async function preview(){
  clearTimeout(timer);if(exporting)return;if(pending)pending.abort();
  const control=pending=new AbortController(),current=++sequence,sourceRevision=revision,data=input();
  $('previewButton').disabled=true;status('正在整理页面布局…',false,true);
  const slow=setTimeout(()=>{if(current===sequence)status('本地服务响应较慢，仍在整理页面…最长再等待约 15 秒。',false,true);},5000);
  try{
    const result=await api('preview',data,control.signal);
    if(current!==sequence||sourceRevision!==revision)return;
    try{render(result);}catch(ignore){throw Error('预览显示出错，请重新点击“生成预览”。');}
    plan=result;readyRevision=sourceRevision;
    status('预览已更新，共 '+result.pages.length+' 页。确认内容后可下载或在 WPS 演示中打开。');
  }catch(e){if(current===sequence&&!e.cancelled){readyRevision=-1;stale(true);status(e.message,true);}}
  finally{clearTimeout(slow);if(current===sequence){pending=null;$('previewButton').disabled=false;}ready();}
}
function syncThemes(){document.querySelectorAll('.theme').forEach(el=>el.classList.toggle('selected',el.querySelector('input').checked));}
const requestedTheme=new URLSearchParams(location.search).get('theme');if(requestedTheme&&/^[a-z-]+$/.test(requestedTheme)){const themeRadio=document.querySelector('input[name=theme][value="'+requestedTheme+'"]');if(themeRadio)themeRadio.checked=true;}
syncThemes();
function changed(){syncThemes();revision++;readyRevision=-1;ready();stale(true);if(plan){status('内容已修改，正在更新预览…',false,true);clearTimeout(timer);timer=setTimeout(preview,650);}else status('内容已更新，点击“生成预览”查看版式。');}
$('studioForm').onsubmit=e=>{e.preventDefault();preview();};document.querySelectorAll('input,textarea').forEach(el=>el.addEventListener(el.type==='radio'?'change':'input',changed));
$('example').onclick=()=>{const jp=window.OfficeI18n&&window.OfficeI18n.lang&&window.OfficeI18n.lang()==='ja';$('deckTitle').value=jp?'プロジェクト進捗報告':'项目阶段工作汇报';$('subtitle').value=jp?'成果と次の計画':'阶段成果与下一步计划';$('outline').value=jp?'成果\n初版の機能開発を完了\nユーザーの声を集めて操作を改善\nテストと課題管理の仕組みを整備\n\n主な数値\n35% 処理時間の短縮\n1200 人のアクティブユーザー\n98% 満足度\n\n今後の計画\n2026年1月：要件確定\n2026年3月：試験導入\n2026年6月：全社展開\n\n役割分担\n企画：要件と優先度の整理\n開発：毎週使える版を提供\n品質：品質の最終確認\n運用：現場の声を収集\n\nまとめ\n複雑なことをシンプルに、シンプルなことを極限まで。':'阶段成果\n完成第一版功能开发\n收集用户反馈并优化使用流程\n建立测试与问题跟踪记录\n\n关键数据\n35% 处理效率提升\n1200 位活跃用户\n98% 用户满意度\n\n推进计划\n2026年1月：需求确认\n2026年3月：上线试点\n2026年6月：全面推广\n\n团队分工\n产品：整理需求与优先级\n研发：按周交付可用版本\n测试：守住质量底线\n运营：收集一线反馈\n\n一句话总结\n把复杂的事情做简单，把简单的事情做到极致。';changed();preview();};
async function deliver(action){
  if(exporting||readyRevision!==revision)return;
  clearTimeout(timer);if(pending){pending.abort();pending=null;}sequence++;
  exporting=true;ready();lock(true);status(action==='open'?'正在生成并启动 WPS 演示…':'正在生成可编辑 PPTX…',false,true);
  try{
    const data=input(),result=await api(action,data);
    if(action==='export'){const url=URL.createObjectURL(result),link=document.createElement('a');link.href=url;link.download=(data.title.replace(/[\\/:*?"<>|]/g,'_')||'演示文稿')+'.pptx';document.body.append(link);link.click();link.remove();setTimeout(()=>URL.revokeObjectURL(url),30000);status('PPTX 已生成并发起下载。如果没有看到下载提示，可改用“在 WPS 演示中打开”。');}
    else status('已启动 WPS 演示打开新文件，共 '+result.pages+' 页。\n文件保存在：'+result.filePath);
  }catch(e){status(e.message,true);}
  finally{exporting=false;lock(false);ready();}
}
// 用 AI 写大纲：服务端用客户自己设置的 AI 服务生成，填进左侧表单后自动预览。
async function aiState(){const box=$('aiState');try{const r=await fetch('/api/ai/config');const d=await r.json();const c=d.config||{},p=(d.providers||[]).find(x=>x.id===c.provider);
  box.replaceChildren();if(c.ready){box.textContent=tr('将使用：')+(p?(window.OfficeI18n&&window.OfficeI18n.lang&&window.OfficeI18n.lang()==='ja'&&p.nameJa?p.nameJa:p.name):c.provider)+' · '+c.model+'。'+(p&&p.region==='local'?'内容只在这台电脑上处理。':'主题会发送给该服务商；'+({mask:'个人信息会先自动隐藏。',remind:'含个人信息时会先问你。',off:'隐私保护已关闭。'}[(d.privacy&&d.privacy.mode)||'mask']));$('aiGo').disabled=false;return;}
  $('aiGo').disabled=true;box.append(tr('还没有接入 AI。')+' ');const a=document.createElement('a');a.href='settings.html#ai';a.textContent=tr('去设置');const g=document.createElement('a');g.href='ai-guide.html';g.textContent=tr('不会申请 API Key？看教程');box.append(a,' · ',g);}
  catch(e){box.textContent=tr('无法读取 AI 设置，请确认办公工具箱仍在运行。');}}
$('aiGo').onclick=async()=>{const topic=$('aiTopic').value.trim();if(!topic){status('请先写下演示主题。',true);return;}
  $('aiGo').disabled=true;status('AI 正在写大纲，通常需要 10～60 秒…',false,true);
  try{if(!token){const s=await fetch('/api/session').then(r=>r.json());token=s.token;}
    const lang=(window.OfficeI18n&&window.OfficeI18n.lang&&window.OfficeI18n.lang()==='ja')||/[぀-ヿ]/.test(topic)?'ja':'zh';
    const body={topic,audience:$('aiAudience').value,pages:$('aiPages').value,lang};
    const outline=b=>fetch('/api/ai/ppt-outline',{method:'POST',headers:{'Content-Type':'application/json','X-Coach-Token':token},body:JSON.stringify(b)});
    let r=await outline(body);if(r.status===403){token=(await fetch('/api/session').then(x=>x.json())).token;r=await outline(body);}let d=await r.json().catch(()=>({}));
    // 隐私设置为“发送前提醒”时，服务会先返回 428，用户确认后再发
    if(r.status===428&&d.needConfirm){if(!window.confirm(d.error))throw Error('已取消，没有发送给 AI。');r=await outline({...body,privacyConfirmed:true});d=await r.json().catch(()=>({}));}
    if(!r.ok)throw Error(d.error||'AI 没有生成大纲，请重试。');
    $('deckTitle').value=d.title||'';$('subtitle').value=d.subtitle||'';$('outline').value=d.outline||'';changed();await preview();status('AI 大纲已填好（'+d.model+'）。请检查内容后再下载或打开。'+(d.privacy&&d.privacy.masked?' 发送前已隐藏 '+d.privacy.masked+' 处个人信息。':''));}
  catch(e){status(e.message,true);}finally{aiState();}};
aiState();
function readFile(file){return new Promise((resolve,reject)=>{if(!/^image\/(png|jpeg)$/.test(file.type))return reject(Error('只支持 PNG 或 JPG 图片。'));if(file.size>2*1024*1024)return reject(Error(file.name+' 超过 2MB，请压缩后再用。'));const r=new FileReader();r.onload=()=>resolve(r.result);r.onerror=()=>reject(Error('读取图片失败。'));r.readAsDataURL(file);});}
function brandUi(){$('brandColor').value=brand.color||'#0B5FFF';$('brandOn').value=brand.on?'1':'';$('brandLogoState').textContent=brand.logo?'已设置 Logo':'未设置 Logo';$('brandLogoClear').hidden=!brand.logo;$('deckImagesState').textContent=images.length?'已添加 '+images.length+' 张图片':'未添加图片';}
$('brandColor').oninput=()=>{brand.color=$('brandColor').value;brand.on='1';saveBrand();brandUi();changed();};
$('brandOn').onchange=()=>{brand.on=$('brandOn').value;saveBrand();changed();};
$('brandLogo').onchange=async()=>{const f=$('brandLogo').files[0];if(!f)return;try{brand.logo=await readFile(f);brand.on='1';saveBrand();brandUi();changed();}catch(e){status(e.message,true);}$('brandLogo').value='';};
$('brandLogoClear').onclick=()=>{brand.logo='';saveBrand();brandUi();changed();};
$('deckImages').onchange=async()=>{try{const list=[...$('deckImages').files].slice(0,8);images=[];for(const f of list)images.push(await readFile(f));brandUi();changed();}catch(e){images=[];brandUi();status(e.message,true);}};
brandUi();
$('download').onclick=()=>deliver('export');$('openWps').onclick=()=>deliver('open');
window.addEventListener('unhandledrejection',()=>status('页面出现意外问题，请重试。如果仍无反应，请关闭此窗口，再从 WPS 功能区重新打开。',true));
$('loadHint').remove();
// 从“Word 转 PPT”带过来的大纲（2 小时内、且当前大纲为空时填入一次）。
try{const raw=localStorage.getItem('officeToolbox.studio.draft'),d=raw&&JSON.parse(raw);
  if(d&&d.outline&&Date.now()-Number(d.at||0)<7200000&&!$('outline').value.trim()){
    $('deckTitle').value=String(d.title||'');$('subtitle').value=String(d.subtitle||'');$('outline').value=String(d.outline);
    const radio=document.querySelector('input[name=theme][value="'+String(d.theme||'').replace(/[^a-z-]/g,'')+'"]');if(radio)radio.checked=true;
    localStorage.removeItem('officeToolbox.studio.draft');changed();preview().then(()=>status('已填入从 Word 文档整理的大纲，可以修改内容或换一种样式后再制作。'),()=>{});}
}catch(e){}
})();

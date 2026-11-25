// 演示 · 第二批新功能（2026-11-24）：每页导出图片 / 拼长图、讲义打印（导出讲义 PDF）、PPT 瘦身（删掉没用到的版式和母版）。
(function(root){'use strict';
var T=root.OfficeTools,L=T.lib;
function n(v){return Number(v)||0;}
function text(v){return String(v==null?'':v);}
function ja(){try{return root.OfficeI18n&&root.OfficeI18n.lang()==='ja';}catch(e){return false;}}
function M(zh,jp){return ja()?jp:zh;}
function tt(s){return root.OfficeI18n?root.OfficeI18n.t(s):s;}
async function tool(name,body){var s=await L.localFetch('/api/session',{}).then(function(r){return r.json();});var r=await L.localFetch('/api/tools/'+name,{method:'POST',headers:{'Content-Type':'application/json','X-Coach-Token':s.token},body:JSON.stringify(body)});var d=await r.json().catch(function(){return {};});if(!r.ok)throw Error(d.error||M('本机服务没有完成，请重试。','ローカルサービスで処理できませんでした。もう一度お試しください。'));return d;}
function deckName(p){return text(p.Name).replace(/\.[^.]+$/,'').replace(/[\\\/:*?"<>|]/g,'_')||M('演示','プレゼン');}
function sep(folder){return /\//.test(folder)&&!/\\/.test(folder)?'/':'\\';}
function pickFolder(a,title,start){var dlg=a.FileDialog(4);dlg.Title=tt(title);try{if(start)dlg.InitialFileName=start+sep(start);}catch(ignore){}if(dlg.Show()!==-1)return '';return text(dlg.SelectedItems.Item(1)).replace(/[\\\/]+$/,'');}
// 文件名被占用时加 -2、-3……（问本机服务哪些已经存在）
async function freeNames(list){var x=await tool('exists',{paths:list}).catch(function(){return {exists:[]};});return (x.exists||[]).some(Boolean);}
function slideSize(p){var w=1920,h=1080;try{var sw=n(p.PageSetup.SlideWidth),sh=n(p.PageSetup.SlideHeight);if(sw&&sh)h=Math.round(w*sh/sw);}catch(ignore){}return {w:w,h:h};}

// ---- 每页导出图片 / 长图 ----
async function slidesToImages(c){var a=c.a,p=a.ActivePresentation,o=c.options||{},mode=['each','long'].indexOf(o.mode)>=0?o.mode:'each',count=n(p.Slides.Count);if(!count)throw Error(M('演示文稿里还没有幻灯片。','スライドがまだありません。'));if(count>200)throw Error(M('超过 200 页，请分几次导出。','200 枚を超えています。何回かに分けて書き出してください。'));
  var folder=pickFolder(a,'选择图片保存位置',text(p.Path));if(!folder)return M('已取消。','キャンセルしました。');var s=sep(folder),base=deckName(p),size=slideSize(p),k=1,names;
  for(;;){var tag=k>1?'('+k+')':'';names=[];for(var i=1;i<=count;i++)names.push(folder+s+base+tag+'-'+String(i).padStart(2,'0')+'.png');if(!await freeNames(names)||k>20)break;k++;}
  var done=[],skipped=0;for(var j=1;j<=count;j++){var sl=p.Slides.Item(j);try{if(sl.SlideShowTransition&&sl.SlideShowTransition.Hidden&&n(sl.SlideShowTransition.Hidden)!==0){skipped++;continue;}}catch(ignore){}sl.Export(names[j-1],'PNG',size.w,size.h);done.push(names[j-1]);}
  if(!done.length)throw Error(M('没有导出图片（可能全部是隐藏页）。','画像を書き出せませんでした（すべて非表示スライドの可能性があります）。'));
  if(mode==='long'){if(done.length<2)return M('只有 1 页，已导出：','1 枚だけなので、そのまま書き出しました：')+done[0];var r=await tool('stitch',{paths:done,name:base+M('-长图','-縦長'),open:true});return M('已导出 '+done.length+' 页，并拼成一张长图：',done.length+' 枚を書き出し、1 枚の縦長画像にまとめました：')+r.path+(skipped?M('（跳过了 '+skipped+' 个隐藏页）','（非表示スライド '+skipped+' 枚は除外）'):'')+M('。每页的图片也在同一个文件夹里。','\n各スライドの画像も同じフォルダーにあります。');}
  try{await tool('reveal',{path:done[0]});}catch(ignore){}
  return M('已把 '+done.length+' 页导出成图片（'+size.w+' × '+size.h+'），放在：',done.length+' 枚を画像（'+size.w+' × '+size.h+'）に書き出しました。保存先：')+folder+(skipped?M('。跳过了 '+skipped+' 个隐藏页','\n非表示スライド '+skipped+' 枚は除外しました'):'')+M('。原文件没有改动。','\n元のファイルは変更していません。');}
T.register({id:'slidesToImages',label:'每页导出图片 / 长图',desc:'每页存成一张清晰的 PNG，也可以拼成一张长图发微信、LINE',icon:'picture',group:'slideout',undo:false,ownConfirm:['wpp'],keywords:'导出图片 每页 长图 图片 PNG 发朋友圈 画像 書き出し 長い画像 スライド 画像化',fields:[{key:'mode',label:'导出成',options:[{value:'each',label:'每页一张图片'},{value:'long',label:'每页一张，再拼成一张长图'}],value:'each'}],hosts:{wpp:slidesToImages}});

// ---- 讲义打印 ----
var HANDOUT={1:10,2:2,3:3,4:8,6:4,9:9};
async function handoutPdf(c){var a=c.a,p=a.ActivePresentation,per=String((c.options||{}).per||'6');if(!HANDOUT[per])per='6';var type=HANDOUT[per];
  var folder=text(p.Path);if(!folder){folder=pickFolder(a,'演示还没有保存，选择讲义 PDF 的保存位置','');if(!folder)return M('已取消。','キャンセルしました。');}
  var s=sep(folder),base=folder+s+deckName(p)+M('-讲义'+per+'页','-配布資料'+per+'枚'),file=base+'.pdf',k=2;while(await freeNames([file])&&k<50)file=base+'-'+(k++)+'.pdf';
  try{p.ExportAsFixedFormat(file,2,2,-1,1,type,0,null,1);}catch(e){
    try{p.PrintOptions.OutputType=type;try{p.PrintOptions.FrameSlides=-1;}catch(ignore){}}catch(e2){throw Error(M('这台 WPS 不支持自动导出讲义。请按 Ctrl+P，在“打印内容”里选“讲义”，每页 '+per+' 张。','この WPS では配布資料を自動で書き出せません。Ctrl+P を押し、印刷対象で「配布資料」（1 ページ '+per+' 枚）を選んでください。'));}
    return M('已把打印设置改成“讲义，每页 '+per+' 张”。现在按 Ctrl+P 就能打印。（这台 WPS 不支持直接导出讲义 PDF。）','印刷設定を「配布資料（1 ページ '+per+' 枚）」にしました。Ctrl+P で印刷できます。（この WPS では配布資料の PDF を直接書き出せません。）');}
  try{await tool('reveal',{path:file});}catch(ignore){}
  return M('已导出讲义 PDF（每页 '+per+' 张，带边框'+(per==='3'?'，右边留有记笔记的横线':'')+'）：','配布資料の PDF を書き出しました（1 ページ '+per+' 枚、枠付き'+(per==='3'?'、右側にメモ欄':'')+'）：')+file+M('\n打开这个 PDF 打印即可。原演示没有改动。','\nこの PDF を開いて印刷してください。元のプレゼンは変更していません。');}
T.register({id:'handoutPdf',label:'讲义打印',desc:'一页放 2～9 张幻灯片，导出讲义 PDF 方便打印发给大家',icon:'pdf',group:'slideout',undo:false,ownConfirm:['wpp'],keywords:'讲义 打印 一页多张 handout 配布資料 印刷 ハンドアウト 6スライド',fields:[{key:'per',label:'每页几张',options:[{value:'2',label:'2 张'},{value:'3',label:'3 张（右边可记笔记）'},{value:'4',label:'4 张'},{value:'6',label:'6 张'},{value:'9',label:'9 张'},{value:'1',label:'1 张（带边框）'}],value:'6'}],hosts:{wpp:handoutPdf}});

// ---- PPT 瘦身 ----
function slimDeck(c){var a=c.a,p=a.ActivePresentation,used={},usedDesign={};
  for(var i=1;i<=n(p.Slides.Count);i++){var s=p.Slides.Item(i),dn='',ln='';try{dn=text(s.Design.Name);}catch(ignore){}try{ln=text(s.CustomLayout.Name);}catch(ignore){}used[dn+'|'+ln]=1;usedDesign[dn]=1;}
  var layouts=[],designs=[];
  for(var d=1;d<=n(p.Designs.Count);d++){var de=p.Designs.Item(d),name=text(de.Name);if(!usedDesign[name]){designs.push(de);continue;}var cl=de.SlideMaster.CustomLayouts;for(var k=n(cl.Count);k>=1;k--){var lay=cl.Item(k);if(!used[name+'|'+text(lay.Name)])layouts.push(lay);}}
  if(designs.length>=n(p.Designs.Count))designs.shift();
  if(!layouts.length&&!designs.length)return M('没有用不到的版式和母版。文件还是太大的话，多半是图片太大：请用“文件与图片 → 压缩文档里的图片”。','使っていないレイアウトやマスターはありません。それでもファイルが大きい場合は画像が原因のことが多いので、「ファイルと画像 → 文書内の画像を圧縮」を使ってください。');
  if(!L.ask(M('将删除 '+layouts.length+' 个没有用到的版式'+(designs.length?'、'+designs.length+' 套没有用到的母版':'')+'，幻灯片内容不变。\n部分演示操作不能撤销，建议先保存一份副本。是否继续？','使っていないレイアウト '+layouts.length+' 個'+(designs.length?'とマスター '+designs.length+' 組':'')+'を削除します。スライドの内容は変わりません。\n元に戻せない操作があるので、先にコピーを保存することをおすすめします。続行しますか？')))return M('已取消。','キャンセルしました。');
  var gone=0,goneD=0;layouts.forEach(function(l){try{l.Delete();gone++;}catch(ignore){}});designs.forEach(function(de){try{de.Delete();goneD++;}catch(ignore){}});
  return M('已删除 '+gone+' 个没用到的版式'+(goneD?'、'+goneD+' 套母版':'')+'。保存后文件会变小一些。\n图片多的演示，再用“文件与图片 → 压缩文档里的图片”另存一份，效果最明显。','使っていないレイアウト '+gone+' 個'+(goneD?'とマスター '+goneD+' 組':'')+'を削除しました。保存すると少し小さくなります。\n画像が多いプレゼンは、「ファイルと画像 → 文書内の画像を圧縮」で別名保存すると効果が大きいです。');}
T.register({id:'slimDeck',label:'PPT 瘦身',desc:'删掉没用到的版式和母版；再配合“压缩图片”让文件变小',icon:'trim',group:'slideout',undo:false,ownConfirm:['wpp'],keywords:'瘦身 变小 压缩 文件太大 发不出去 母版 版式 軽量化 容量 ファイルサイズ 削減',hosts:{wpp:slimDeck}});
})(window);

// 办公工具箱 · 批量与拆分（WPS）：批量转 PDF、批量替换文字、批量加文字水印、按标题拆分文档、把选中内容发送到另一个打开的文档。
// 原文件一律只读打开、不保存；结果写到用户选的文件夹，同名文件已存在时自动加序号，不覆盖。
(function(root){
  'use strict';
  var T=root.OfficeTools;if(!T)return;var L=T.lib;
  var MAX=100;
  var KINDS={wps:{desc:'文字文档',pattern:'*.docx;*.doc;*.wps;*.rtf;*.docm',re:/\.(docx?|wps|rtf|docm)$/i},et:{desc:'表格',pattern:'*.xlsx;*.xls;*.et;*.xlsm;*.csv',re:/\.(xlsx?|et|xlsm|csv)$/i},wpp:{desc:'演示文稿',pattern:'*.pptx;*.ppt;*.dps;*.pptm',re:/\.(pptx?|dps|pptm)$/i}};
  function sepOf(p){return /\//.test(p)&&!/\\/.test(p)?'/':'\\';}
  function join(folder,name){folder=String(folder).replace(/[\\\/]+$/,'');return folder+sepOf(folder)+name;}
  function baseName(p){return String(p).split(/[\\\/]/).pop();}
  function stem(p){return baseName(p).replace(/\.[^.]+$/,'');}
  function dirOf(p){var s=String(p),i=Math.max(s.lastIndexOf('\\'),s.lastIndexOf('/'));return i>0?s.slice(0,i):'';}
  function clean(name){return String(name).replace(/[\\\/:*?"<>|\r\n\t\x07\x0b\x0c]/g,'_').replace(/\s+/g,' ').trim().slice(0,60)||'未命名';}
  function pickFiles(a,h,title){var k=KINDS[h],dlg=a.FileDialog(3);dlg.Title=title;dlg.AllowMultiSelect=true;try{dlg.Filters.Clear();dlg.Filters.Add(k.desc,k.pattern);}catch(ignore){}if(dlg.Show()!==-1)return null;var list=[];for(var i=1;i<=dlg.SelectedItems.Count;i++)list.push(String(dlg.SelectedItems.Item(i)));
    list=list.filter(function(p){return k.re.test(p);});if(!list.length)throw Error('没有选中'+k.desc+'。');if(list.length>MAX)throw Error('一次最多处理 '+MAX+' 个文件，请分批。');list.sort(function(x,y){return baseName(x).localeCompare(baseName(y),'zh-Hans-CN',{numeric:true});});return list;}
  function pickFolder(a,title,start){var dlg=a.FileDialog(4);dlg.Title=title;try{if(start)dlg.InitialFileName=start+sepOf(start);}catch(ignore){}if(dlg.Show()!==-1)return null;return String(dlg.SelectedItems.Item(1)).replace(/[\\\/]+$/,'');}
  // 核对磁盘和本批次名称；服务不可用时停止，避免覆盖。
  async function freeNames(paths){
    if(!paths.length)return [];
    var message='无法确认输出文件是否重名。请重新启动办公工具箱后再试。';
    try{
      var session=await L.localFetch('/api/session',{});if(!session.ok)throw Error(message);
      var s=await session.json();if(!s.token)throw Error(message);
      var used=Object.create(null),result=[],pending=paths.map(function(p,i){return {path:p,index:i,n:1};});
      for(var round=0;pending.length&&round<1000;round++){
        var candidates=pending.map(function(x){return x.n===1?x.path:x.path.replace(/(\.[^.\\\/]+)$/,'-'+x.n+'$1');});
        var response=await L.localFetch('/api/tools/exists',{method:'POST',headers:{'Content-Type':'application/json','X-Coach-Token':s.token},body:JSON.stringify({paths:candidates})});
        if(!response.ok)throw Error(message);
        var data=await response.json();if(!Array.isArray(data.exists)||data.exists.length!==candidates.length||data.exists.some(function(v){return typeof v!=='boolean';}))throw Error(message);
        var next=[];pending.forEach(function(x,i){var p=candidates[i],key=p.replace(/\\/g,'/').toLowerCase();if(data.exists[i]||used[key]){x.n++;next.push(x);}else{used[key]=true;result[x.index]=p;}});pending=next;
      }
      if(pending.length)throw Error(message);return result;
    }catch(e){throw Error(message);}
  }
  function openDoc(a,h,p){if(h==='wps')return a.Documents.Open(p,false,true,false);if(h==='et')return a.Workbooks.Open(p,0,true);return a.Presentations.Open(p,true,false,false);}
  function closeDoc(h,d){try{if(h==='wps')d.Close(0);else if(h==='et')d.Close(false);else d.Close();}catch(ignore){}}
  function report(done,failed,what,folder){return '已'+what+' '+done.length+' 个文件，保存在：'+folder+(failed.length?'\n'+failed.length+' 个没能处理：'+failed.slice(0,5).map(function(f){return baseName(f.file)+'（'+f.error+'）';}).join('、')+(failed.length>5?'…':''):'')+'\n原文件都没有改动。';}
  function errText(e){return String(e&&e.message||e||'未知错误').split('\n')[0].slice(0,60);}
  // ---- 批量转 PDF（文字、表格、演示都可以） ----
  async function batchPdf(c){var a=c.a,h=c.h,files=pickFiles(a,h,'选择要转成 PDF 的'+KINDS[h].desc+'（可多选）');if(!files)return '已取消。';
    var folder=pickFolder(a,'选择 PDF 保存到哪个文件夹',dirOf(files[0]));if(!folder)return '已取消。';
    var outs=await freeNames(files.map(function(f){return join(folder,clean(stem(f))+'.pdf');}));
    if(!(await L.ask('将把 '+files.length+' 个'+KINDS[h].desc+'转成 PDF，保存到：\n'+folder+'\n处理时会依次打开每个文件（只读），完成后关闭。是否开始？')))return '已取消。';
    var done=[],failed=[],keep=T.lib.active(a,h);
    for(var i=0;i<files.length;i++){var d=null;try{d=openDoc(a,h,files[i]);if(h==='wps')d.ExportAsFixedFormat(outs[i],17);else if(h==='et')d.ExportAsFixedFormat(0,outs[i]);else d.SaveAs(outs[i],32);done.push(outs[i]);}catch(e){failed.push({file:files[i],error:errText(e)});}finally{if(d)closeDoc(h,d);}}
    try{if(keep&&keep.Activate)keep.Activate();}catch(ignore){}
    return report(done,failed,'转成 PDF',folder);}
  // ---- 批量替换文字（文字文档） ----
  function countIn(text,find){var n=0,i=0;while((i=text.indexOf(find,i))>=0){n++;i+=find.length;}return n;}
  async function batchReplace(c){var a=c.a,o=c.options||{},find=String(o.find||''),rep=String(o.replace==null?'':o.replace);if(!find)throw Error('请填写要查找的文字。');if(find.length>250||rep.length>250)throw Error('查找和替换的文字各不超过 250 个字。');
    var files=pickFiles(a,'wps','选择要批量替换的文字文档（可多选）');if(!files)return '已取消。';var folder=pickFolder(a,'选择改好的文档保存到哪个文件夹',dirOf(files[0]));if(!folder)return '已取消。';
    var same=files.some(function(f){return dirOf(f).toLowerCase()===folder.toLowerCase();});
    var outs=await freeNames(files.map(function(f){return join(folder,clean(stem(f))+(same?'-替换':'')+'.docx');}));
    if(!(await L.ask('将在 '+files.length+' 个文档里把“'+find+'”换成'+(rep?'“'+rep+'”':'空（删除）')+'，改好的文档另存到：\n'+folder+'\n原文件不改。是否开始？')))return '已取消。';
    var done=[],failed=[],total=0,none=[];
    for(var i=0;i<files.length;i++){var d=null;try{d=openDoc(a,'wps',files[i]);var n=countIn(String(d.Content.Text||''),find);var stories=d.StoryRanges;
      for(var s=1;s<=stories.Count;s++){var st=stories.Item(s);while(st){var f=st.Find;f.ClearFormatting();f.Replacement.ClearFormatting();f.Execute(find,true,false,false,false,false,true,0,false,rep,2);try{st=st.NextStoryRange;}catch(e){st=null;}}}
      total+=n;if(!n)none.push(baseName(files[i]));d.SaveAs2(outs[i],16);done.push(outs[i]);}catch(e){failed.push({file:files[i],error:errText(e)});}finally{if(d)closeDoc('wps',d);}}
    return report(done,failed,'替换并另存',folder).replace('\n原文件','\n正文里一共替换了 '+total+' 处（页眉、页脚、文本框里的也一起换了）。'+(none.length?'\n'+none.length+' 个文档正文里没有找到：'+none.slice(0,5).join('、')+(none.length>5?'…':''):'')+'\n原文件');}
  // ---- 批量加文字水印（文字文档） ----
  function addWatermark(d,text){var n=0;for(var i=1;i<=d.Sections.Count;i++){var hd=d.Sections.Item(i).Headers.Item(1);if(i>1&&hd.LinkToPrevious)continue;var sh=hd.Shapes.AddTextEffect(0,text,T.lib.localFont('宋体'),1,false,false,0,0);try{sh.Name='办公工具箱水印-'+i;}catch(ignore){}
      sh.TextEffect.NormalizedHeight=false;sh.Line.Visible=false;sh.Fill.Visible=true;sh.Fill.Solid();sh.Fill.ForeColor.RGB=12632256;try{sh.Fill.Transparency=0.5;}catch(ignore){}sh.Rotation=315;sh.LockAspectRatio=true;sh.Height=Math.min(120,Math.max(40,400/Math.max(2,text.length)));sh.Width=Math.min(500,sh.Height*text.length);
      sh.WrapFormat.AllowOverlap=true;sh.WrapFormat.Side=3;sh.WrapFormat.Type=3;sh.RelativeHorizontalPosition=0;sh.RelativeVerticalPosition=0;sh.Left=-999995;sh.Top=-999995;n++;}return n;}
  async function batchWatermark(c){var a=c.a,text=String((c.options||{}).text||'').trim();if(!text)throw Error('请填写水印文字，例如“内部资料”。');if(text.length>20)throw Error('水印文字不超过 20 个字。');
    var files=pickFiles(a,'wps','选择要加水印的文字文档（可多选）');if(!files)return '已取消。';var folder=pickFolder(a,'选择加好水印的文档保存到哪个文件夹',dirOf(files[0]));if(!folder)return '已取消。';
    var same=files.some(function(f){return dirOf(f).toLowerCase()===folder.toLowerCase();});var outs=await freeNames(files.map(function(f){return join(folder,clean(stem(f))+(same?'-水印':'')+'.docx');}));
    if(!(await L.ask('将给 '+files.length+' 个文档加上斜着的浅灰色水印“'+text+'”，另存到：\n'+folder+'\n原文件不改。是否开始？')))return '已取消。';
    var done=[],failed=[];for(var i=0;i<files.length;i++){var d=null;try{d=openDoc(a,'wps',files[i]);if(!addWatermark(d,text))throw Error('没有可以放水印的页眉');d.SaveAs2(outs[i],16);done.push(outs[i]);}catch(e){failed.push({file:files[i],error:errText(e)});}finally{if(d)closeDoc('wps',d);}}
    return report(done,failed,'加水印并另存',folder);}
  // ---- 按标题拆分文档 ----
  function headingParts(d,level){var ps=d.Paragraphs,n=ps.Count,heads=[];if(n>20000)throw Error('文档太长（超过 2 万段），请先手动分成几部分。');
    for(var i=1;i<=n;i++){var p=ps.Item(i),lv=Number(p.OutlineLevel);if(lv>=1&&lv<=level){var r=p.Range,t=String(r.Text||'').replace(/[\r\x07\x0b\x0c]/g,'').trim();if(t)heads.push({start:Number(r.Start),title:t});}}
    var end=Number(d.Content.End),parts=[];if(!heads.length)return parts;
    if(heads[0].start>0){var pre=String(d.Range(0,heads[0].start).Text||'').replace(/\s/g,'');if(pre)parts.push({start:0,end:heads[0].start,title:'前言'});}
    heads.forEach(function(h,k){parts.push({start:h.start,end:k+1<heads.length?heads[k+1].start:end,title:h.title});});return parts;}
  async function splitByHeading(c){var a=c.a,d=c.d,level=Math.min(3,Math.max(1,Number((c.options||{}).level)||1));var parts=headingParts(d,level);
    if(!parts.length)throw Error('没有找到'+(level===1?'“标题 1”':'1～'+level+' 级标题')+'。请先给每一部分的标题设置标题样式（或用“设为一级标题”）。');if(parts.length<2)throw Error('只有一个部分，不需要拆分。');if(parts.length>MAX)throw Error('超过 '+MAX+' 个部分，请改用更高一级的标题拆分。');
    var folder=pickFolder(a,'选择拆分出的文档保存到哪个文件夹',d.Path?String(d.Path):'');if(!folder)return '已取消。';
    var w=String(parts.length).length<2?2:String(parts.length).length,pad=function(n){n=String(n);while(n.length<w)n='0'+n;return n;};
    var outs=await freeNames(parts.map(function(p,i){return join(folder,pad(i+1)+'-'+clean(p.title)+'.docx');}));
    if(!(await L.ask('将按标题拆成 '+parts.length+' 个文档：\n'+parts.slice(0,8).map(function(p,i){return baseName(outs[i]);}).join('\n')+(parts.length>8?'\n…':'')+'\n保存到：'+folder+'\n原文档不改。是否开始？')))return '已取消。';
    var ps=d.PageSetup,setup={o:ps.Orientation,t:ps.TopMargin,b:ps.BottomMargin,l:ps.LeftMargin,r:ps.RightMargin,w:ps.PageWidth,h:ps.PageHeight},done=[],failed=[];
    for(var i=0;i<parts.length;i++){var nd=null;try{nd=a.Documents.Add();try{var q=nd.PageSetup;q.Orientation=setup.o;q.PageWidth=setup.w;q.PageHeight=setup.h;q.TopMargin=setup.t;q.BottomMargin=setup.b;q.LeftMargin=setup.l;q.RightMargin=setup.r;}catch(ignore){}
      nd.Content.FormattedText=d.Range(parts[i].start,parts[i].end).FormattedText;var last=nd.Paragraphs.Item(nd.Paragraphs.Count).Range;if(nd.Paragraphs.Count>1&&!String(last.Text||'').replace(/[\r\s]/g,''))try{last.Delete();}catch(ignore){}
      nd.SaveAs2(outs[i],16);done.push(outs[i]);}catch(e){failed.push({file:parts[i].title,error:errText(e)});}finally{if(nd)closeDoc('wps',nd);}}
    try{d.Activate();}catch(ignore){}
    return report(done,failed,'拆分出',folder).replace('\n原文件都没有改动。','\n页眉页脚和页码设置不会带过去，需要时在新文档里再加。原文档没有改动。');}
  // ---- 把选中内容发送到另一个打开的文档（连图片、表格和格式） ----
  function others(a,h,cur){var list=[],col=h==='wps'?a.Documents:h==='et'?a.Workbooks:a.Presentations;for(var i=1;i<=col.Count;i++){var x=col.Item(i);if(x!==cur&&String(x.FullName)!==String(cur.FullName))list.push(x);}return list;}
  async function sendToDoc(c){var a=c.a,h=c.h,d=c.d,o=c.options||{},list=others(a,h,d);
    if(!list.length)throw Error('请先打开要接收内容的另一个'+KINDS[h].desc+'。');
    var pick=Number(o.target)||0;if(list.length>1&&!(pick>=1&&pick<=list.length))return '打开的'+KINDS[h].desc+'有 '+list.length+' 个，请在“发送到第几个”里填序号：\n'+list.map(function(x,i){return (i+1)+'. '+x.Name;}).join('\n');
    var target=list[(pick||1)-1];
    if(h==='wps'){var sel=a.Selection;if(!sel||Number(sel.End)<=Number(sel.Start))throw Error('请先选中要发送的内容（文字、图片、表格都可以）。');var src=sel.Range;
      if(!(await L.ask('将把选中的内容（连图片、表格和格式）插到《'+target.Name+'》的末尾。是否继续？')))return '已取消。';
      var dest,e=Number(target.Content.End)-1;dest=target.Range(e,e);if(String(target.Range(Math.max(0,e-1),e).Text||'')!=='\r'&&e>0){dest.InsertParagraphAfter();e=Number(target.Content.End)-1;dest=target.Range(e,e);}
      dest.FormattedText=src.FormattedText;return '已发送到《'+target.Name+'》（在那个文档里可以按 Ctrl+Z 撤销）。原文没有改动。';}
    if(h==='et'){var r=a.Selection;if(!r||!r.Cells)throw Error('请先选中要发送的单元格。');if(r.Areas&&r.Areas.Count>1)throw Error('请只选一块连续的区域。');
      var ws=target.ActiveSheet,cell=null;try{target.Activate();cell=a.ActiveCell;}catch(e){cell=null;}try{d.Activate();}catch(ignore){}if(!cell||String(cell.Worksheet.Parent.Name)!==String(target.Name))cell=ws.Range('A1');
      var dest2=cell.Resize(r.Rows.Count,r.Columns.Count),used=0;for(var i=1;i<=Math.min(r.Rows.Count,200);i++)for(var j=1;j<=Math.min(r.Columns.Count,50);j++){var v=dest2.Cells.Item(i,j).Value2;if(v!==undefined&&v!==null&&v!=='')used++;}
      if(!(await L.ask('将把选中的 '+r.Address(false,false)+' 复制到《'+target.Name+'》的“'+ws.Name+'”'+dest2.Address(false,false)+(used?'，其中 '+used+' 个单元格已有内容，会被覆盖':'')+'。WPS 表格的撤销不能还原这一步。是否继续？')))return '已取消。';
      r.Copy(dest2);return '已复制到《'+target.Name+'》的 '+ws.Name+'!'+dest2.Address(false,false)+'（格式和公式一起带过去）。';}
    var win=a.ActiveWindow,slides=[];try{var sr=win.Selection.SlideRange;for(var k=1;k<=sr.Count;k++)slides.push(sr.Item(k));}catch(e){}if(!slides.length)try{slides.push(win.View.Slide);}catch(e){}if(!slides.length)throw Error('请先点一下要发送的那一页（可以在左边缩略图里多选）。');
    if(!(await L.ask('将把 '+slides.length+' 页复制到《'+target.Name+'》的最后。是否继续？')))return '已取消。';
    slides.sort(function(x,y){return x.SlideIndex-y.SlideIndex;});var before=target.Slides.Count;slides.forEach(function(s,i){s.Copy();target.Slides.Paste(before+i+1);});
    return '已把 '+slides.length+' 页复制到《'+target.Name+'》第 '+(before+1)+' 页起（版式会跟着那个演示的母版）。';}
  T.register({id:'batchPdf',label:'批量转 PDF',desc:'一次把多个文件转成 PDF，原文件不变',icon:'pdf',group:'file',undo:false,hosts:{wps:batchPdf,et:batchPdf,wpp:batchPdf}});
  T.register({id:'batchReplace',label:'批量替换文字',desc:'多个文档里同一处文字一起改，另存新文件',icon:'replace',group:'file',undo:false,fields:[{key:'find',label:'查找',placeholder:'例如：2025 年'},{key:'replace',label:'换成',placeholder:'例如：2026 年（留空表示删除）',required:false}],hosts:{wps:batchReplace}});
  T.register({id:'batchWatermark',label:'批量加水印',desc:'多个文档加上同样的文字水印，另存新文件',icon:'watermark',group:'file',undo:false,fields:[{key:'text',label:'水印文字',placeholder:'例如：内部资料'}],hosts:{wps:batchWatermark}});
  T.register({id:'splitByHeading',label:'按标题拆分文档',desc:'每个“标题 1”一份，拆成多个文档',icon:'columns',group:'file',undo:false,fields:[{key:'level',label:'按几级标题拆（1～3）',placeholder:'不填按 1 级',required:false}],hosts:{wps:splitByHeading}});
  T.register({id:'sendToDoc',label:'发送到另一个文档',desc:'选中的内容连图片、格式复制到另一个打开的文档',icon:'insert',group:'file',undo:false,fields:[{key:'target',label:'发送到第几个（只有一个时不用填）',placeholder:'例如：2',required:false}],hosts:{wps:sendToDoc,et:sendToDoc,wpp:sendToDoc}});
  root.OfficeBatch={headingParts:headingParts,freeNames:freeNames,join:join,addWatermark:addWatermark};
})(typeof window!=='undefined'?window:globalThis);

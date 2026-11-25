// 办公工具箱 · 数据脱敏：各程序怎么取文字、怎么改回去（2026-11-25）。
// 侧栏 mask-panel.js 只管界面和流程；这里按程序分开写：
//   wps / et / wpp：WPS 的 JS 接口（同步调用）；office-word / office-excel：Microsoft Office 的 Office.js（异步）。
// 每个程序都提供：
//   info()              当前文件的名字、位置、有没有保存过
//   segments(scope)     要检查的文字片段：[{key, text, ...}]。scope = all（整份）/ selection（选中的）
//   write(seg, edits)   把片段里的几处换掉，edits = [{start,end,text,expect}]，返回每处成功与否
//   exact               表格的单元格是“整格记原值”（公式、数字格式都能原样还原），其他程序按出现顺序还原
//   copy(path)          另存副本（WPS）；Office 做不到时没有这个函数
//   prepare()           改之前：关掉修订、接受已有修订（副本里）
//   clean()             清掉文件属性里的作者、单位
//   mark(id)/marked()   在文件里记下对照表编号，打开脱敏后的文件就能直接还原
//   save()              保存
(function(root){'use strict';
function n(x){return Number(x);}
function txt(s){return s==null?'':String(s);}
function ja(){try{return root.OfficeI18n&&root.OfficeI18n.lang()==='ja';}catch(e){return false;}}
function M(zh,jp){return ja()?jp:zh;}
var MARK='NiumaMaskId',PROPS=['Author','Last Author','Company','Manager'];
// 按位置从后往前换，换之前核对原文，对不上就跳过（文字里有域、隐藏内容时位置可能对不上）。
function eachDesc(edits,fn){var ok=edits.map(function(){return false;});edits.map(function(e,i){return {e:e,i:i};}).sort(function(x,y){return y.e.start-x.e.start;}).forEach(function(x){try{ok[x.i]=!!fn(x.e);}catch(e){ok[x.i]=false;}});return ok;}
function builtin(d,name,v){try{var p=(d.BuiltInDocumentProperties||d.BuiltinDocumentProperties).Item(name);if(v===undefined)return txt(p.Value);p.Value=v;return true;}catch(e){return v===undefined?'':false;}}
function cleanProps(d){var done=0;PROPS.forEach(function(p){if(builtin(d,p)&&builtin(d,p,''))done++;});return done;}
function customGet(d){try{return txt(d.CustomDocumentProperties.Item(MARK).Value);}catch(e){return '';}}
function customSet(d,id){try{try{d.CustomDocumentProperties.Item(MARK).Delete();}catch(ignore){}d.CustomDocumentProperties.Add(MARK,false,4,id);return true;}catch(e){return false;}}

// ---------------- WPS 文字 ----------------
// 所有“文字流”：正文、脚注、尾注、批注、文本框、页眉页脚……每种可能有好几段（各节的页眉），用 NextStoryRange 串起来。
function wordStories(d){var out=[];for(var t=1;t<=17;t++){var r=null;try{r=d.StoryRanges.Item(t);}catch(e){r=null;}var k=0;while(r&&k<300){out.push({type:t,idx:k,range:r});k++;var nx=null;try{nx=r.NextStoryRange;}catch(e){nx=null;}r=nx;}}return out;}
function wordParas(range,key,lo,hi,out,limit){var ps=range.Paragraphs,c=n(ps.Count);for(var i=1;i<=c;i++){if(out.length>=limit)throw Error(M('文档太长（超过 '+limit+' 段），请选中一部分再处理。','文書が長すぎます（'+limit+' 段落超）。一部を選択してから処理してください。'));var r=ps.Item(i).Range,t=txt(r.Text).replace(/[\r\u0007]+$/,'');if(!t)continue;out.push({key:key,start:n(r.Start),text:t,range:r,lo:lo,hi:hi});}}
function wpsWord(a){var d=a.ActiveDocument;if(!d)throw Error(M('请先打开一份文字文档。','先に文書を開いてください。'));
 return {host:'wps',exact:false,
  info:function(){return {name:txt(d.Name),path:txt(d.FullName),saved:!(d.Saved===false||d.Saved===0),hasPath:/[\\\/]/.test(txt(d.FullName))};},
  segments:function(scope){var out=[],LIMIT=30000;
   if(scope==='selection'){var s=a.Selection&&a.Selection.Range;if(!s||!txt(s.Text).replace(/[\r\n\u0007\s]/g,''))throw Error(M('请先选中要脱敏的文字。','マスキングする文字を選択してください。'));var type=n(s.StoryType)||1;wordParas(s,'st'+type+'#*',n(s.Start),n(s.End),out,LIMIT);return out;}
   wordStories(d).forEach(function(st){wordParas(st.range,'st'+st.type+'#'+st.idx,null,null,out,LIMIT);});return out;},
  // 选中范围以外的不改（段落比选区大时）
  inside:function(seg,h){return seg.lo==null||(seg.start+h.start>=seg.lo&&seg.start+h.end<=seg.hi);},
  write:function(seg,edits){return eachDesc(edits,function(e){var r=seg.range.Duplicate;r.SetRange(seg.start+e.start,seg.start+e.end);if(txt(r.Text)!==e.expect)return false;r.Text=e.text;return true;});},
  copy:function(path){var fmt;try{fmt=d.SaveFormat;}catch(e){fmt=undefined;}if(fmt!==undefined&&fmt!==null)d.SaveAs2(path,fmt);else d.SaveAs2(path);},
  prepare:function(copied){var note='';try{d.TrackRevisions=false;}catch(ignore){}var rc=0;try{rc=n(d.Revisions.Count);}catch(ignore){}if(rc){if(!copied)throw Error(M('文档里还有 '+rc+' 处没接受的修订（删掉的原文还在修订里）。请先接受或拒绝全部修订，或者用“另存副本”方式脱敏。','未処理の変更履歴が '+rc+' 件あります（削除した元の文字が履歴に残っています）。先にすべて承諾・元に戻すか、「コピーを保存」でマスキングしてください。'));try{d.Revisions.AcceptAll();note=M('副本里接受了全部修订（'+rc+' 处），删掉的原文不会留在修订里。','コピーでは変更履歴（'+rc+' 件）をすべて承諾し、削除された元の文字が残らないようにしました。');}catch(e){throw Error(M('没能接受文档里的修订，请先手动“接受所有修订”再脱敏。','変更履歴を承諾できませんでした。先に手動で「すべての変更を反映」してからマスキングしてください。'));}}return note;},
  clean:function(){return cleanProps(d);},
  mark:function(id){var ok=false;try{try{d.Variables.Item(MARK).Delete();}catch(ignore){}d.Variables.Add(MARK,id);ok=true;}catch(ignore){}return customSet(d,id)||ok;},
  marked:function(){try{var v=txt(d.Variables.Item(MARK).Value);if(v)return v;}catch(ignore){}return customGet(d);},
  save:function(){d.Save();try{d.UndoClear();}catch(ignore){}},
  undoClear:function(){try{d.UndoClear();}catch(ignore){}},
  names:function(){return [];}};}

// ---------------- WPS 表格 ----------------
function grid(r,prop){var rows=n(r.Rows.Count),cols=n(r.Columns.Count),v;try{v=r[prop];}catch(e){return null;}if(rows===1&&cols===1)return [[v]];if(Array.isArray(v)&&v.length===rows&&Array.isArray(v[0]))return v;var out=[];for(var i=1;i<=rows;i++){var line=[];for(var j=1;j<=cols;j++){var x;try{x=r.Cells.Item(i,j)[prop];}catch(e){x='';}line.push(x);}out.push(line);}return out;}
function cellText(v){if(typeof v==='number'){if(!Number.isFinite(v)||Math.abs(v)>=1e15||Math.floor(v)!==v)return '';return String(v);}return typeof v==='string'?v:'';}
function wpsSheet(a){var wb=a.ActiveWorkbook;if(!wb)throw Error(M('请先打开一个表格文件。','先にブックを開いてください。'));var LIMIT=300000;
 function rangeSegs(sh,r,out){var name=txt(sh.Name),r0=n(r.Row),c0=n(r.Column),rows=n(r.Rows.Count),cols=n(r.Columns.Count);if(out.cells+rows*cols>LIMIT)throw Error(M('要检查的格子超过 30 万个，请选中一部分再处理。','対象のセルが 30 万を超えています。一部を選択してから処理してください。'));out.cells+=rows*cols;
  var v=grid(r,'Value2')||[],f=grid(r,'Formula')||[];for(var i=0;i<rows;i++)for(var j=0;j<cols;j++){var t=cellText(v[i]&&v[i][j]);if(!t)continue;var fx=txt(f[i]&&f[i][j]);out.list.push({key:name+'!R'+(r0+i)+'C'+(c0+j),kind:'cell',sheet:sh,row:r0+i,col:c0+j,text:t,raw:v[i][j],formula:/^=/.test(fx)?fx:null});}}
 return {host:'et',exact:true,
  info:function(){return {name:txt(wb.Name),path:txt(wb.FullName),saved:!(wb.Saved===false||wb.Saved===0),hasPath:/[\\\/]/.test(txt(wb.FullName))};},
  segments:function(scope){var out={list:[],cells:0};
   if(scope==='selection'){var s=a.Selection;if(!s||!s.Rows)throw Error(M('请先选中要脱敏的单元格。','マスキングするセルを選択してください。'));var areas=[];try{for(var k=1;k<=n(s.Areas.Count);k++)areas.push(s.Areas.Item(k));}catch(e){areas=[s];}var sh=a.ActiveSheet;if(sh.ProtectContents)throw Error(M('工作表受保护，请先取消保护。','シートが保護されています。保護を解除してください。'));areas.forEach(function(r){rangeSegs(sh,r,out);});return out.list;}
   for(var i=1;i<=n(wb.Worksheets.Count);i++){var ws=wb.Worksheets.Item(i);if(ws.ProtectContents)throw Error(M('工作表“'+ws.Name+'”受保护，请先取消保护。','シート「'+ws.Name+'」が保護されています。保護を解除してください。'));
    var ur=ws.UsedRange;if(ur&&!(n(ur.Rows.Count)===1&&n(ur.Columns.Count)===1&&!cellText(ur.Value2)))rangeSegs(ws,ur,out);
    // 批注（备注）、页眉页脚
    try{var cs=ws.Comments;for(var c=1;c<=n(cs.Count);c++){var cm=cs.Item(c),ct=txt(cm.Text());if(ct)out.list.push({key:txt(ws.Name)+'!note!'+txt(cm.Parent.Address(false,false)),kind:'note',obj:cm,text:ct});}}catch(ignore){}
    try{var ps=ws.PageSetup;['LeftHeader','CenterHeader','RightHeader','LeftFooter','CenterFooter','RightFooter'].forEach(function(p){var pt=txt(ps[p]);if(pt)out.list.push({key:txt(ws.Name)+'!ps!'+p,kind:'ps',obj:ps,prop:p,text:pt});});}catch(ignore){}}
   return out.list;},
  inside:function(){return true;},
  read:function(seg){if(seg.kind==='cell')return cellText(seg.sheet.Cells.Item(seg.row,seg.col).Value2);if(seg.kind==='note')return txt(seg.obj.Text());return txt(seg.obj[seg.prop]);},
  // 整格写：记下原来的值或公式、数字格式，还原时原样放回
  put:function(seg,value){if(seg.kind==='cell'){var c=seg.sheet.Cells.Item(seg.row,seg.col),nf;try{nf=txt(c.NumberFormat);}catch(e){nf='';}c.NumberFormat='@';c.Value2=value;return {b:seg.formula||seg.raw,nf:nf,a:value};}
   if(seg.kind==='note'){var before=seg.text;try{seg.obj.Shape.TextFrame.Characters().Text=value;}catch(e){seg.obj.Text(value);}return {b:before,a:value};}
   var old=seg.text;seg.obj[seg.prop]=value;return {b:old,a:value};},
  back:function(seg,entry){if(seg.kind==='cell'){var c=seg.sheet.Cells.Item(seg.row,seg.col);if(cellText(c.Value2)!==entry.a)return false;if(entry.nf!=null)try{c.NumberFormat=entry.nf;}catch(ignore){}if(typeof entry.b==='string'&&/^=/.test(entry.b))c.Formula=entry.b;else c.Value2=entry.b;return true;}
   if(seg.kind==='note'){if(txt(seg.obj.Text())!==entry.a)return false;try{seg.obj.Shape.TextFrame.Characters().Text=entry.b;}catch(e){seg.obj.Text(entry.b);}return true;}
   if(txt(seg.obj[seg.prop])!==entry.a)return false;seg.obj[seg.prop]=entry.b;return true;},
  copy:function(path){var fmt;try{fmt=wb.FileFormat;}catch(e){fmt=undefined;}if(fmt)wb.SaveAs(path,fmt);else wb.SaveAs(path);},
  prepare:function(){return '';},
  clean:function(){return cleanProps(wb);},
  mark:function(id){return customSet(wb,id);},
  marked:function(){return customGet(wb);},
  save:function(){wb.Save();},
  // 工作表名称里有词表的词：不自动改（会牵动公式和链接），提醒客户自己改
  names:function(words){var hit=[];for(var i=1;i<=n(wb.Worksheets.Count);i++){var nm=txt(wb.Worksheets.Item(i).Name);if(words.some(function(w){return nm.indexOf(w.word)>=0;}))hit.push(nm);}return hit;}};}

// ---------------- WPS 演示 ----------------
function slideShapes(shapes,list){for(var i=1;i<=n(shapes.Count);i++){var s=shapes.Item(i);list.push(s);}return list;}
function shapeSegs(s,base,out){var t=0;try{t=n(s.Type);}catch(e){}
 if(t===6){try{for(var g=1;g<=n(s.GroupItems.Count);g++)shapeSegs(s.GroupItems.Item(g),base,out);}catch(ignore){}return;}
 var id;try{id=n(s.Id);}catch(e){id=out.length;}
 var hasTable=false;try{hasTable=!!s.HasTable;}catch(e){}
 if(hasTable){var tb=s.Table;for(var r=1;r<=n(tb.Rows.Count);r++)for(var c=1;c<=n(tb.Columns.Count);c++){try{var tr=tb.Cell(r,c).Shape.TextFrame.TextRange,tt=txt(tr.Text);if(tt)out.push({key:base+'#'+id+'/r'+r+'c'+c,tr:tr,text:tt});}catch(ignore){}}return;}
 try{if(s.HasTextFrame&&s.TextFrame.HasText){var x=s.TextFrame.TextRange,v=txt(x.Text);if(v)out.push({key:base+'#'+id,tr:x,text:v});}}catch(ignore){}}
function wpsSlides(a){var d=a.ActivePresentation;if(!d)throw Error(M('请先打开一份演示文稿。','先にプレゼンテーションを開いてください。'));
 return {host:'wpp',exact:false,
  info:function(){return {name:txt(d.Name),path:txt(d.FullName),saved:!(d.Saved===false||d.Saved===0),hasPath:/[\\\/]/.test(txt(d.FullName))};},
  segments:function(scope){var out=[];
   if(scope==='selection'){var w=a.ActiveWindow,sel=w&&w.Selection,type=0;try{type=n(sel.Type);}catch(e){}if(type<2)throw Error(M('请先选中要脱敏的文本框、表格或图形。','マスキングするテキストボックス・表・図形を選択してください。'));var sid=n(w.View.Slide.SlideID),sr=sel.ShapeRange;for(var k=1;k<=n(sr.Count);k++)shapeSegs(sr.Item(k),'s'+sid,out);return out;}
   for(var i=1;i<=n(d.Slides.Count);i++){var sl=d.Slides.Item(i),base='s'+n(sl.SlideID);slideShapes(sl.Shapes,[]).forEach(function(s){shapeSegs(s,base,out);});
    // 演讲者备注
    try{slideShapes(sl.NotesPage.Shapes,[]).forEach(function(s){var pt=0;try{pt=n(s.PlaceholderFormat.Type);}catch(e){}if(pt===2)shapeSegs(s,base+'n',out);});}catch(ignore){}}
   return out;},
  inside:function(){return true;},
  write:function(seg,edits){return eachDesc(edits,function(e){var r=seg.tr.Characters(e.start+1,e.end-e.start);if(txt(r.Text)!==e.expect)return false;r.Text=e.text;return true;});},
  copy:function(path){d.SaveAs(path);},
  prepare:function(){return '';},
  clean:function(){return cleanProps(d);},
  mark:function(id){var ok=false;try{try{d.Tags.Delete(MARK);}catch(ignore){}d.Tags.Add(MARK,id);ok=true;}catch(ignore){}return customSet(d,id)||ok;},
  marked:function(){try{var v=txt(d.Tags.Item(MARK));if(v)return v;}catch(ignore){}return customGet(d);},
  save:function(){d.Save();},
  names:function(){return [];}};}

// ---------------- Microsoft Office（Office.js）----------------
// Office.js 不能“另存为”，所以 Office 版直接改当前文件：侧栏会先提醒客户自己另存一份副本。
// Word：正文的每个段落（含表格里的段落）；页眉页脚、文本框、批注改不到，会说明。
function officeWord(api){var W=api.Word;
 return {host:'office-word',exact:false,office:true,
  info:async function(){var name='',path='';try{var u=txt(api.Office.context.document.url);path=u;name=u.split(/[\\\/]/).pop();}catch(ignore){}return {name:name||M('当前文档','現在の文書'),path:path,saved:true,hasPath:!!path};},
  segments:async function(scope){return W.run(async function(c){var sel=c.document.getSelection();sel.load('text');await c.sync();if(scope==='selection'&&!txt(sel.text).replace(/\s/g,''))throw Error(M('请先选中要脱敏的文字。','マスキングする文字を選択してください。'));
   var ps=(scope==='selection'?sel:c.document.body).paragraphs;ps.load('items/text');await c.sync();if(ps.items.length>20000)throw Error(M('文档太长，请选中一部分再处理。','文書が長すぎます。一部を選択してから処理してください。'));
   return ps.items.map(function(p,i){return {key:'body',index:i,scope:scope,text:txt(p.text)};}).filter(function(s){return s.text;});});},
  inside:function(){return true;},
  // 一次把很多段落一起改：按段落找到要换的文字（同一段里第几次出现），整批同步
  writeMany:async function(jobs){return W.run(async function(c){var sel=c.document.getSelection();var ps=(jobs.length&&jobs[0].seg.scope==='selection'?sel:c.document.body).paragraphs;ps.load('items');await c.sync();var qs=[];
    jobs.forEach(function(j){var p=ps.items[j.seg.index],t=j.seg.text;j.edits.forEach(function(e,k){if(!p||!e.expect||e.expect.length>255){qs.push(null);return;}var nth=0,at=-1;while((at=t.indexOf(e.expect,at+1))>=0&&at<e.start)nth++;var q=p.search(e.expect,{matchCase:true,matchWildcards:false});q.load('items');qs.push({q:q,nth:nth,e:e,j:j,k:k});});});
    await c.sync();var ok=jobs.map(function(j){return j.edits.map(function(){return false;});});
    qs.slice().reverse().forEach(function(x){if(!x)return;var r=x.q.items[x.nth];if(!r)return;r.insertText(x.e.text,'Replace');ok[jobs.indexOf(x.j)][x.k]=true;});await c.sync();return ok;});},
  prepare:async function(){return W.run(async function(c){try{c.document.changeTrackingMode='Off';await c.sync();}catch(ignore){}return '';});},
  clean:async function(){return W.run(async function(c){try{var p=c.document.properties;p.author='';p.company='';p.manager='';p.lastAuthor='';await c.sync();return 1;}catch(e){return 0;}});},
  mark:async function(id){return W.run(async function(c){try{var cp=c.document.properties.customProperties;cp.add(MARK,id);await c.sync();return true;}catch(e){return false;}});},
  marked:async function(){return W.run(async function(c){try{var cp=c.document.properties.customProperties.getItemOrNullObject(MARK);cp.load('value');await c.sync();return cp.isNullObject?'':txt(cp.value);}catch(e){return '';}});},
  names:function(){return [];}};}
// Excel：单元格整格记原值（公式、数字格式都还原）。批注、页眉页脚 Office.js 改不到，会说明。
function officeExcel(api){var X=api.Excel;
 async function collect(c,scope){var ranges=[];if(scope==='selection'){var r=c.workbook.getSelectedRange();r.load('address');ranges.push(r);}else{var ws=c.workbook.worksheets;ws.load('items/name');await c.sync();ws.items.forEach(function(w){var u=w.getUsedRangeOrNullObject(true);u.load('address');ranges.push(u);});}
  await c.sync();ranges=ranges.filter(function(r){return !r.isNullObject;});ranges.forEach(function(r){r.load('values,formulas,numberFormat,rowIndex,columnIndex,worksheet/name');});await c.sync();return ranges;}
 return {host:'office-excel',exact:true,office:true,
  info:async function(){var name='',path='';try{var u=txt(api.Office.context.document.url);path=u;name=u.split(/[\\\/]/).pop();}catch(ignore){}return {name:name||M('当前工作簿','現在のブック'),path:path,saved:true,hasPath:!!path};},
  segments:async function(scope){return X.run(async function(c){var out=[],total=0;(await collect(c,scope)).forEach(function(r){var nm=r.worksheet.name;total+=r.values.length*(r.values[0]||[]).length;if(total>300000)throw Error(M('要检查的格子超过 30 万个，请选中一部分再处理。','対象のセルが 30 万を超えています。一部を選択してから処理してください。'));
    r.values.forEach(function(row,i){row.forEach(function(v,j){var t=cellText(v);if(!t)return;var f=txt(r.formulas[i][j]);out.push({key:nm+'!R'+(r.rowIndex+i+1)+'C'+(r.columnIndex+j+1),kind:'cell',sheetName:nm,row:r.rowIndex+i,col:r.columnIndex+j,text:t,raw:v,formula:/^=/.test(f)?f:null,nf:txt(r.numberFormat[i][j])});});});});return out;});},
  inside:function(){return true;},
  putMany:async function(items){return X.run(async function(c){items.forEach(function(x){var cell=c.workbook.worksheets.getItem(x.seg.sheetName).getCell(x.seg.row,x.seg.col);cell.numberFormat=[['@']];cell.values=[[x.value]];});await c.sync();return items.map(function(x){return {b:x.seg.formula||x.seg.raw,nf:x.seg.nf,a:x.value};});});},
  backMany:async function(items){return X.run(async function(c){var cells=items.map(function(x){var p=x.key.match(/^(.*)!R(\d+)C(\d+)$/);if(!p)return null;var cell=c.workbook.worksheets.getItem(p[1]).getCell(n(p[2])-1,n(p[3])-1);cell.load('values');return cell;});await c.sync();var ok=items.map(function(x,i){var cell=cells[i];if(!cell||cellText(cell.values[0][0])!==x.entry.a)return false;if(x.entry.nf!=null)cell.numberFormat=[[x.entry.nf]];if(typeof x.entry.b==='string'&&/^=/.test(x.entry.b))cell.formulas=[[x.entry.b]];else cell.values=[[x.entry.b]];return true;});await c.sync();return ok;});},
  prepare:async function(){return '';},
  clean:async function(){return X.run(async function(c){try{var p=c.workbook.properties;p.author='';p.company='';p.manager='';await c.sync();return 1;}catch(e){return 0;}});},
  mark:async function(id){return X.run(async function(c){try{c.workbook.properties.custom.add(MARK,id);await c.sync();return true;}catch(e){return false;}});},
  marked:async function(){return X.run(async function(c){try{var p=c.workbook.properties.custom.getItemOrNullObject(MARK);p.load('value');await c.sync();return p.isNullObject?'':txt(p.value);}catch(e){return '';}});},
  names:function(){return [];}};}

// 当前是哪个程序
function adapter(){var T=root.OfficeTools;
 if(T&&T.app&&T.app()){var a=T.app(),h=T.host();return h==='et'?wpsSheet(a):h==='wpp'?wpsSlides(a):wpsWord(a);}
 var api={Word:root.Word,Excel:root.Excel,PowerPoint:root.PowerPoint,Office:root.Office},host='';try{host=String(root.Office.context.host||'');}catch(e){}
 if(/Word/i.test(host))return officeWord(api);if(/Excel/i.test(host))return officeExcel(api);
 if(/PowerPoint/i.test(host))throw Error(M('PowerPoint 版暂时做不到自动脱敏（Office 接口改不了全部文字）。请用 WPS 演示打开这份文件来脱敏，或者用“替换”手动改。','PowerPoint 版では自動マスキングができません（Office の仕組みで全文を変更できないため）。WPS プレゼンテーションで開いてマスキングするか、「置換」で手動で直してください。'));
 throw Error(M('请在 WPS 或 Office 里打开工具箱再使用。','WPS または Office でツールボックスを開いてからお使いください。'));}
root.MaskHost={adapter:adapter,wpsWord:wpsWord,wpsSheet:wpsSheet,wpsSlides:wpsSlides,officeWord:officeWord,officeExcel:officeExcel,cellText:cellText,MARK:MARK};
})(typeof window!=='undefined'?window:globalThis);

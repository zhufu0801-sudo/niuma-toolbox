// 办公工具箱 · 内容搬运与文件模块（三个宿主通用）。
(function(root){
  'use strict';
  var T=root.OfficeTools,L=T.lib,labels=L.labels,active=L.active,wordSelection=L.wordSelection,cells=L.cells,api=L.api,selectedEnd=L.selectedEnd,position=L.position;
  function attempt(fn){try{return fn();}catch(ignore){return undefined;}}
  function hex(value){value=Number(value);return Number.isInteger(value)&&value>=0&&value<=16777215&&value!==9999999?'#'+[value&255,(value>>8)&255,(value>>16)&255].map(function(n){return ('0'+n.toString(16)).slice(-2);}).join(''):undefined;}
  function richStyle(range,ppt){var result={},font=attempt(function(){return range.Font;});if(!font)return result;var name=attempt(function(){return String(font.NameFarEast||font.Name||'');}),size=Number(attempt(function(){return font.Size;}));if(name&&name.length<=80&&!/[<>;{}\x00-\x1f]/.test(name))result.font=name;if(size>0&&size<=144)result.size=size;['Bold','Italic'].forEach(function(key){var value=attempt(function(){return font[key];});if(value===true||value===-1||value===1||value===false||value===0)result[key.toLowerCase()]=!!value;});var color=hex(attempt(function(){return ppt?font.Color.RGB:font.Color;}));if(color)result.color=color;var alignment=Number(attempt(function(){return range.ParagraphFormat.Alignment;}))-(ppt?1:0);if(alignment>=0&&alignment<=3)result.alignment=alignment;var level=Number(attempt(function(){return range.Paragraphs.Item(1).OutlineLevel;}));if(level>=1&&level<=9)result.level=level;else if(level===10)result.level=0;return result;}
  function captureFormat(a,h,block){var result={version:1,base:{}},range,table;
    if(h==='wps'){range=a.Selection.Range;result.base=richStyle(range);if(block.type==='table')table=a.Selection.Tables.Item(1);else if(block.text.split('\n').length<=200){result.paragraphs=block.text.split('\n').map(function(_,i){return richStyle(attempt(function(){return range.Paragraphs.Item(i+1).Range;}));});delete result.base.level;}}
    else if(h==='et'){range=a.Selection;result.base=richStyle(range);}
    else{var selection=a.ActiveWindow.Selection;range=selection.Type===3?selection.TextRange:attempt(function(){return selection.ShapeRange.Item(1).TextFrame.TextRange;});if(block.type==='table')table=selection.ShapeRange.Item(1).Table;result.base=richStyle(range,true);if(block.type==='text'&&block.text.split('\n').length<=200)result.paragraphs=block.text.split('\n').map(function(_,i){var style=richStyle(attempt(function(){return range.Paragraphs(i+1,1);})||range,true);if(i===0&&attempt(function(){var type=selection.ShapeRange.Item(1).PlaceholderFormat.Type;return type===1||type===3;}))style.level=1;return style;});}
    if(block.type==='table'){result.cells=block.rows.map(function(row,r){return row.map(function(_,c){var cell=h==='et'?attempt(function(){return range.Cells.Item(r+1,c+1);}):attempt(function(){return table.Cell(r+1,c+1);});var text=h==='wps'?cell&&cell.Range:h==='wpp'?attempt(function(){return cell.Shape.TextFrame.TextRange;}):cell,style=richStyle(text,h==='wpp');var background=hex(attempt(function(){return h==='et'?cell.Interior.Color:h==='wps'?cell.Shading.BackgroundPatternColor:cell.Shape.Fill.ForeColor.RGB;}));if(background)style.background=background;return style;});});var widths=block.rows[0].map(function(_,c){return Number(attempt(function(){return h==='et'?range.Columns.Item(c+1).Width:table.Columns.Item(c+1).Width;}));});if(widths.every(function(v){return v>0&&v<=2000;}))result.columnWidths=widths;}
    return result;
  }
  function officeColor(value){return parseInt(value.slice(1,3),16)+parseInt(value.slice(3,5),16)*256+parseInt(value.slice(5,7),16)*65536;}
  function applyStyle(range,style,ppt,keepDeck,spreadsheet){if(!style)return;var failed=false;function set(fn){try{fn();}catch(ignore){failed=true;}}
    if(!ppt&&style.level)set(function(){range.Style=-1-style.level;});if(style.font&&!keepDeck)set(function(){range.Font.Name=style.font;if(!ppt&&!spreadsheet)range.Font.NameFarEast=style.font;});if(style.size&&!keepDeck)set(function(){range.Font.Size=style.size;});['bold','italic'].forEach(function(key){if(style[key]!==undefined)set(function(){range.Font[key==='bold'?'Bold':'Italic']=style[key]?-1:0;});});if(style.color)set(function(){if(ppt)range.Font.Color.RGB=officeColor(style.color);else range.Font.Color=officeColor(style.color);});if(style.alignment!==undefined)set(function(){range.ParagraphFormat.Alignment=style.alignment+(ppt?1:0);});return failed;
  }
  function formatWord(d,range,block,table){var format=block.format;if(!format)return false;var missed=false;
    if(table){for(var r=0;r<block.rows.length;r++)for(var c=0;c<block.rows[r].length;c++){var cell=table.Cell(r+1,c+1),style=Object.assign({},format.base,format.cells&&format.cells[r]&&format.cells[r][c]);missed=applyStyle(cell.Range,style,false)||missed;if(style.background)try{cell.Shading.BackgroundPatternColor=officeColor(style.background);}catch(ignore){missed=true;}}if(format.columnWidths)try{var widths=format.columnWidths,total=widths.reduce(function(x,y){return x+y;},0),available=Number(d.PageSetup.PageWidth)-Number(d.PageSetup.LeftMargin)-Number(d.PageSetup.RightMargin);if(!(available>0))available=450;for(var k=0;k<widths.length;k++)table.Columns.Item(k+1).Width=widths[k]/total*Math.min(total,available);}catch(ignore){missed=true;}return missed;}
    var offset=Number(range.Start);block.text.split('\n').forEach(function(text,i){var target=d.Range(offset,offset+text.length),style=Object.assign({},format.base,format.paragraphs&&format.paragraphs[i]);missed=applyStyle(target,style,false)||missed;offset+=text.length+1;});return missed;
  }
  function formatSheet(destination,block,skipFit){var f=block.format,missed=false;if(f){for(var r=0;r<(block.type==='table'?block.rows.length:block.text.split('\n').length);r++)for(var c=0;c<(block.type==='table'?block.rows[0].length:1);c++){var style=Object.assign({},f.base,block.type==='table'?f.cells&&f.cells[r]&&f.cells[r][c]:f.paragraphs&&f.paragraphs[r]);try{var cell=destination.Cells.Item(r+1,c+1);var basic=Object.assign({},style);delete basic.level;delete basic.alignment;missed=applyStyle(cell,basic,false,false,true)||missed;if(style.alignment!==undefined)cell.HorizontalAlignment=[-4131,-4108,-4152,-4130][style.alignment];if(style.background)cell.Interior.Color=officeColor(style.background);if(style.level)cell.Font.Bold=true;}catch(ignore){missed=true;}}}if(block.type==='table')try{destination.Rows.Item(1).Font.Bold=true;}catch(ignore){if(f)missed=true;}try{if(!skipFit)destination.Columns.AutoFit();destination.WrapText=true;}catch(ignore){}return missed;}
  L.applyModuleStyle=applyStyle;L.officeColor=officeColor;
  function capture(a,h){
    var block;
    if(h==='wps'){
      var selection=wordSelection(a),tables=selection.Tables;
      if(tables&&tables.Count){var t=tables.Item(1);if(tables.Count!==1||t.Range.Start<selection.Range.Start||t.Range.End>selection.Range.End)throw Error('请完整选中一张表格，或仅选择表格外的文字。');if(t.Rows.Count>20||t.Columns.Count>10)throw Error('跨文档表格最多 20 行、10 列。');var rows=[];for(var i=1;i<=t.Rows.Count;i++){var row=[];for(var j=1;j<=t.Columns.Count;j++)row.push(String(t.Cell(i,j).Range.Text).replace(/[\r\u0007]+$/,''));rows.push(row);}block={type:'table',rows:rows};}
      else {if(String(selection.Text).length>10000)throw Error('跨文档文字最多 1 万字，请缩小选区。');block={type:'text',text:String(selection.Text).replace(/\r/g,'\n')};}
    }else if(h==='et'){
      var range=cells(a);if(range.Rows.Count>20||range.Columns.Count>10)throw Error('跨文档表格最多 20 行、10 列。');var data=[];for(var r=1;r<=range.Rows.Count;r++){var line=[];for(var c=1;c<=range.Columns.Count;c++)line.push(String(range.Cells.Item(r,c).Text||''));data.push(line);}block={type:'table',rows:data};
    }else{var text='',selected=a.ActiveWindow.Selection,shape=attempt(function(){if(selected.ShapeRange.Count!==1)return null;return selected.ShapeRange.Item(1);});if(shape&&shape.HasTable){var pt=shape.Table,nr=Number(pt.Rows.Count),nc=Number(pt.Columns.Count);if(nr>20||nc>10)throw Error('模块表格最多 20 行、10 列。');var pr=[];for(var ri=1;ri<=nr;ri++){var line=[];for(var ci=1;ci<=nc;ci++)line.push(String(pt.Cell(ri,ci).Shape.TextFrame.TextRange.Text||''));pr.push(line);}block={type:'table',rows:pr};}else{try{if(selected.Type===3)text=String(selected.TextRange.Text||'');else if(shape&&shape.HasTextFrame)text=String(shape.TextFrame.TextRange.Text||'');}catch(ignore){}if(!text.trim())throw Error('请先选中幻灯片里的文字或一张表格。');if(text.length>10000)throw Error('跨文档文字最多 1 万字。');block={type:'text',text:text.replace(/\r\n?/g,'\n')};}}
    return block;
  }
  // 把一段内容（文字或表格）放进当前文档，三种宿主各自适配：
  // 文字：插在选区末尾，文字自动沿用插入点的格式，表格适应页面宽度；
  // 演示：在当前页之后新建页面，字体跟随本演示；表格：写入选中的空白区域，绝不覆盖已有内容。
  function insertBlock(a,h,d,block,what){
    if(h==='wpp')return '已在当前页之后新建 '+slides(a,d,block)+' 页，字体跟随本演示。';
    if(h==='wps'){
      var record;try{try{record=a.UndoRecord;record.StartCustomRecord('插入'+what);}catch(ignore){record=null;}var range=selectedEnd(a),t;if(block.type==='table'){t=addWordTable(d,range,block.rows);try{t.AutoFitBehavior(2);}catch(ignore){}}else range.Text=block.text.replace(/\n/g,'\r');var missed;try{missed=formatWord(d,range,block,t);}catch(ignore){missed=true;}return '已在选区末尾插入'+what+'。'+(missed?'部分原格式未能应用，请检查；可再用“匹配周围格式”。':'');}finally{if(record)try{record.EndCustomRecord();}catch(ignore){}}
    }
    var targetRange=cells(a),rows=block.type==='table'?block.rows:block.text.split('\n').map(function(t){return [t];});
    if(rows.length>200||rows[0].length>30)throw Error('内容超出快速插入范围。');
    var destination=targetRange.Cells.Item(1,1).Resize(rows.length,rows[0].length);
    if(destination.MergeCells!==false&&destination.MergeCells!==0)throw Error('目标区域存在合并单元格或无法判断，请选择未合并的空白区域。');
    if(Number(a.WorksheetFunction.CountA(destination))>0)throw Error('目标区域包含内容，请先选择足够大的空白区域，避免覆盖。');
    try{destination.Value2=rows.map(function(row){return row.map(function(cell){return /^\s*[=+@-]/.test(cell)?String.fromCharCode(39)+cell:cell;});});}catch(e){var restored=true;try{destination.ClearContents();}catch(ignore){restored=false;}throw Error('插入失败：'+e.message+(restored?' 已清理本次空白目标区域。':' 请使用 WPS 撤销并检查目标区域。'));}
    var missed=formatSheet(destination,block);return '已写入空白目标区域，表格按静态显示值插入。'+(missed?' 部分格式未能应用，请检查。':'');
  }
  async function insertTransfer(a,h,d){
    var target=position(a,h,d),transfer=await api('GET');
    if(position(a,h,active(a,h))!==target)throw Error('当前文档、工作表或选区已切换，请重新点击插入。');
    if(!L.ask('将中转区的'+(transfer.block.type==='table'?'表格':'文字')+'插入当前'+labels[h]+'文档？原内容不会被删除。'))return '已取消。';
    d=active(a,h);if(position(a,h,d)!==target)throw Error('当前文档、工作表或选区已切换，请重新点击插入。');
    return insertBlock(a,h,d,transfer.block,'中转内容');
  }
  // ---- 内容卡片 ----
  function checkBlock(b){
    if(!b||(b.type!=='text'&&b.type!=='table'))throw Error('卡片内容无法识别，请重新拖入或存成卡片。');
    if(b.type==='text'){if(!String(b.text||'').trim())throw Error('卡片内容是空的。');if(String(b.text).length>10000)throw Error('一张卡片最多 1 万字。');return {type:'text',text:String(b.text)};}
    if(!Array.isArray(b.rows)||!b.rows.length||b.rows.length>20)throw Error('卡片里的表格最多 20 行。');
    var width=Math.max.apply(null,b.rows.map(function(r){return Array.isArray(r)?r.length:0;}));if(!width||width>10)throw Error('卡片里的表格最多 10 列。');
    return {type:'table',rows:b.rows.map(function(r){var row=r.map(function(x){return String(x==null?'':x);});while(row.length<width)row.push('');return row;})};
  }
  async function cardApi(method,query,body){var session=await L.localFetch('/api/session',{}).then(function(r){if(!r.ok)throw Error('本地服务未连接。');return r.json();});var r=await L.localFetch('/api/cards'+(query||''),{method:method,headers:{'Content-Type':'application/json','X-Coach-Token':session.token},body:body?JSON.stringify(body):undefined});var data=await r.json().catch(function(){return {};});if(!r.ok)throw Error(data.error||'卡片操作失败，请重试。');return data;}
  T.cards={
    list:function(){return cardApi('GET').then(function(d){return d.cards||[];});},
    add:function(block,source,anchor){return cardApi('POST','',{block:checkBlock(block),format:block.format,sourceHost:source,source:anchor}).then(function(d){return d.card;});},
    remove:function(id){return cardApi('DELETE','?id='+encodeURIComponent(id));},
    clear:function(){return cardApi('DELETE');},
    saveDesktop:async function(id){var session=await L.localFetch('/api/session',{}).then(function(r){return r.json();}),response=await L.localFetch('/api/modules/'+encodeURIComponent(id)+'/desktop',{method:'POST',headers:{'X-Coach-Token':session.token}}),result=await response.json();if(!response.ok)throw Error(result.error||'保存未完成，请重试。');return result;},
    check:checkBlock
  };
  T.register({id:'saveCard',label:'设为内容模块',desc:'把选中的文字或表格存起来反复用',icon:'cards',group:'move',panel:false,undo:false,hosts:{wps:run_save,et:run_save,wpp:run_save}});
  function run_save(c){var block=capture(c.a,c.h);block.format=captureFormat(c.a,c.h,block);return T.cards.add(block,c.h).then(function(card){return '已设为内容模块：'+L.mark(card.title)+'。可在模块栏插入其他文档。';});}
  T.register({id:'insertCard',label:'插入内容卡片',desc:'放进当前文档并适配格式',icon:'insert',group:'move',panel:false,undo:false,hosts:{wps:run_card,et:run_card,wpp:run_card}});
  function run_card(c){var a=c.a,h=c.h,d=c.d,block=checkBlock(c.options.block),target=position(a,h,d);if(c.options.format)block.format=c.options.format;d=active(a,h);if(position(a,h,d)!==target)throw Error('当前文档或选区已切换，请重新点击插入。');return insertBlock(a,h,d,block,'卡片内容');}
  // 演示与文字的表格插入函数由各自模块提供，按需取用，避免加载顺序问题。
  function slides(a,d,b){return L.slides(a,d,b);}
  function addWordTable(d,r,rows){return L.addWordTable(d,r,rows);}
  L.capture=capture;L.captureFormat=captureFormat;L.checkBlock=checkBlock;L.formatWord=formatWord;L.formatSheet=formatSheet;L.insertBlock=insertBlock;
  T.register({id:'capture',icon:'capture',group:'move',panel:false,undo:false,hosts:{wps:run_capture,et:run_capture,wpp:run_capture}});
  function run_capture(c){var content=capture(c.a,c.h);return api('POST',{block:content,sourceHost:c.h}).then(function(){return '已放入本机中转区。切换到目标文档后点击“插入中转内容”。原文档不变。';});}
  T.register({id:'insertTransfer',icon:'insert',group:'move',panel:false,undo:false,hosts:{wps:run_insert,et:run_insert,wpp:run_insert}});
  function run_insert(c){return insertTransfer(c.a,c.h,c.d);}
  T.register({id:'insertLatest',label:'插入最新卡片',desc:'把最近存的一张卡片放进当前文档',icon:'insert',group:'move',panel:false,undo:false,hosts:{wps:run_latest,et:run_latest,wpp:run_latest}});
  function run_latest(c){var a=c.a,h=c.h,target=position(a,h,c.d);return T.cards.list().then(function(list){if(!list.length)throw Error('还没有卡片。先选中内容，点侧栏里的“把当前选区设为模块”，或把内容拖进卡片栏。');var card=list[0];if(position(a,h,active(a,h))!==target)throw Error('当前文档或选区已切换，请重新点击插入。');if(!L.ask('把卡片“'+L.mark(card.title)+'”插入当前'+labels[h]+'文档？原内容不会被删除。'))return '已取消。';var d=active(a,h);if(position(a,h,d)!==target)throw Error('当前文档或选区已切换，请重新点击插入。');var block=checkBlock(card.block);block.format=card.format;return insertBlock(a,h,d,block,'卡片内容');});}
  // 一键备份：把磁盘上的当前文件复制一份到“文档\办公工具箱备份”，文件名带时间，不改动原文件。
  // ---- 粘贴栏：复制过、收集过的内容排成一列，点一下放进当前文档（2026-11-21） ----
  async function pasteApi(method,suffix,body){var session=await L.localFetch('/api/session',{}).then(function(r){if(!r.ok)throw Error('本地服务未连接。');return r.json();});var r=await L.localFetch('/api/pastebar'+(suffix||''),{method:method,headers:{'Content-Type':'application/json','X-Coach-Token':session.token},body:body?JSON.stringify(body):undefined});var data=await r.json().catch(function(){return {};});if(!r.ok)throw Error(data.error||'粘贴栏操作失败，请重试。');return data;}
  T.pasteBar={
    list:function(){return pasteApi('GET');},
    pin:function(id,pinned){return pasteApi('POST','/pin',{id:id,pinned:!!pinned});},
    remove:function(id){return pasteApi('DELETE','?id='+encodeURIComponent(id));},
    clear:function(){return pasteApi('DELETE');},
    fromClipboard:function(){return pasteApi('POST','/clipboard');}
  };
  T.register({id:'pasteBarCollect',label:'放进粘贴栏',desc:'把选中的文字或表格收进粘贴栏，之后点一下就能放进别的文档',icon:'capture',group:'move',panel:false,undo:false,keywords:'粘贴栏 剪贴板 复制历史 收集 貼り付け クリップボード 履歴 集める',hosts:{wps:run_collect,et:run_collect,wpp:run_collect}});
  function run_collect(c){var block=capture(c.a,c.h);return pasteApi('POST','',{block:block,sourceHost:c.h}).then(function(d){return '已放进粘贴栏（共 '+d.items.length+' 条）。';});}
  T.register({id:'pasteBarInsert',label:'从粘贴栏插入',desc:'把粘贴栏里的一条内容放进当前位置',icon:'insert',group:'move',panel:false,undo:false,ownConfirm:['wps','et','wpp'],hosts:{wps:run_pasteInsert,et:run_pasteInsert,wpp:run_pasteInsert}});
  // 文字：插在光标处（选中的文字之后），沿用周围格式；表格：写到选中的空白格子，不会覆盖已有内容；
  // 演示：光标在文本框里时直接插进文字，否则在当前页之后新建一页。
  function run_pasteInsert(c){var a=c.a,h=c.h,d=c.d,block=checkBlock(c.options.block);
    if(h==='wpp'&&block.type==='text'){var sel=attempt(function(){return a.ActiveWindow.Selection;});if(sel&&Number(sel.Type)===3){var tr=sel.TextRange;tr.Text=block.text.replace(/\n/g,'\r');return '已插入到当前文本框。';}}
    if(h==='wps'&&block.type==='text'){var record;try{try{record=a.UndoRecord;record.StartCustomRecord('从粘贴栏插入');}catch(ignore){record=null;}var r=a.Selection.Range;if(Number(r.Start)!==Number(r.End)){r=selectedEnd(a);}r.Text=block.text.replace(/\n/g,'\r');try{r.Select();a.Selection.Collapse(0);}catch(ignore){}return '已插入，格式跟随光标处的文字。';}finally{if(record)try{record.EndCustomRecord();}catch(ignore){}}}
    return insertBlock(a,h,d,block,'粘贴栏内容');}
  T.register({id:'backupFile',label:'备份当前文件',desc:'复制一份带时间的备份',icon:'backup',group:'file',undo:false,hosts:{wps:run_backup,et:run_backup,wpp:run_backup}});
  function run_backup(c){var d=c.d;if(!d.Path)throw Error('这个文档还没保存过，请先保存，再备份。');
    if(d.Saved===false&&!L.ask('文档有还没保存的修改。备份的是上次保存的版本，不包含这些修改。\n要连同修改一起备份，请先点“取消”并保存文档。\n继续备份上次保存的版本吗？'))return '已取消。';
    return L.localFetch('/api/session',{}).then(function(r){if(!r.ok)throw Error('本地服务未连接。');return r.json();}).then(function(session){return L.localFetch('/api/backup',{method:'POST',headers:{'Content-Type':'application/json','X-Coach-Token':session.token},body:JSON.stringify({path:String(d.FullName)})});}).then(function(r){return r.json().catch(function(){return {};}).then(function(data){if(!r.ok)throw Error(data.error||'备份失败，请重试。');return '已备份到：'+data.backupPath;});});}
  T.register({id:'exportPdf',label:'导出 PDF',desc:'使用当前版式生成新文件',icon:'pdf',group:'file',panel:false,undo:false,hosts:{wps:run_pdf,et:run_pdf,wpp:run_pdf}});
  // 导出 PDF：用原文件名；同名已存在时加 -2、-3（以前在名字后面加一串时间数字，客户看不懂）。
  async function freePdfName(base){var list=[base+'.pdf'];for(var n=2;n<=30;n++)list.push(base+'-'+n+'.pdf');try{var s=await (await L.localFetch('/api/session',{})).json(),r=await L.localFetch('/api/tools/exists',{method:'POST',headers:{'Content-Type':'application/json','X-Coach-Token':s.token},body:JSON.stringify({paths:list})}),x=await r.json();if(x&&x.exists)for(var i=0;i<list.length;i++)if(!x.exists[i])return list[i];}catch(e){}return base+'-'+Date.now()+'.pdf';}
  async function run_pdf(c){var a=c.a,d=c.d,h=c.h,dialog=a.FileDialog(4);dialog.Title='选择 PDF 保存文件夹';if(d.Path)dialog.InitialFileName=d.Path+'\\';if(dialog.Show()!==-1)return '已取消导出。';var folder=String(dialog.SelectedItems.Item(1)),name=String(d.Name).replace(/\.[^.]+$/,'').replace(/[\\/:*?"<>|]/g,'_'),output=await freePdfName(folder.replace(/[\\/]$/,'')+'\\'+name);if(h==='wps')d.ExportAsFixedFormat(output,17);else if(h==='et')d.ExportAsFixedFormat(0,output);else d.ExportAsFixedFormat(output,2);return 'PDF 已导出：'+output;}
})(window);

// 办公工具箱 · 让 WPS 版的“技巧导航”和“中文公式助手”在 Microsoft Office 侧栏里运行。
// 面板代码与 WPS 版共用；这里提供它们需要的宿主接口，底层换成 Office.js（Excel）。规则（检查、副本、只写空白格）与 WPS 版一致。
(function(root){
  'use strict';
  var VERSION=(root.OfficeToolboxVersion||'office');
  var Q=root.QuickHost||{},F=root.OfficeFormula,E=root.QuickEngine;
  function Excel(){if(!root.Excel||!root.Excel.run)throw Error('请在 Excel 中打开工具箱侧栏。');return root.Excel;}
  // ---- 侧栏内的确认框（Office 侧栏里不能用 window.confirm）----
  var asker=null;
  function ask(message){return asker?asker(message):Promise.resolve(false);}
  // ---- 设置：收藏和快捷键保存在本侧栏（每台电脑、每个 Office 账号各自一份）----
  var KEY='officeToolbox.office.settings',subs=[];
  function load(){try{return JSON.parse(root.localStorage.getItem(KEY))||{};}catch(e){return {};}}
  var state={settings:Object.assign({toolkit:{favorites:[],bindings:[]}},load())};
  root.WorkbenchSettings={snapshot:function(){return state;},ready:function(){return Promise.resolve(state);},subscribe:function(fn){subs.push(fn);},
    update:function(patch){Object.keys(patch||{}).forEach(function(k){state.settings[k]=Object.assign({},state.settings[k],patch[k]);});try{root.localStorage.setItem(KEY,JSON.stringify(state.settings));}catch(e){}subs.forEach(function(fn){try{fn(state);}catch(e){}});return Promise.resolve(state);}};
  // ---- WPS 面板用到的 OfficeTools 接口 ----
  var hostKey='excel',activeSheet='';
  root.OfficeTools={version:VERSION,hostReady:false,host:function(){return hostKey==='excel'?'et':hostKey==='word'?'wps':'wpp';},app:function(){return {ActiveWorkbook:{},ActiveSheet:{Name:activeSheet}};},
    isHidden:function(){return false;},
    list:function(){return (root.OfficeAddinTools?root.OfficeAddinTools.list(hostKey):[]).map(function(x){return {id:x.id,label:x.label,desc:x.desc,panel:true,group:x.group||'',fields:x.fields||null};});},
    lib:{ask:ask,localFetch:function(){return Promise.resolve({ok:true,json:function(){return Promise.resolve({version:VERSION});}});}}};
  root.OfficeToolsUi={run:function(id,extra){if(root.OfficeToolboxPane)root.OfficeToolboxPane.go(id,extra||{});},showView:function(){}};
  // ---- 选区快照（和 WPS 版 QuickHost.capture 返回同样的结构）----
  function local(address){return String(address||'').replace(/^.*!/,'');}
  async function snapshot(context){
    var r=context.workbook.getSelectedRange();r.load('address,rowCount,columnCount,rowIndex,columnIndex,values,formulas');var ws=r.worksheet;ws.load('name');var wb=context.workbook;try{wb.load('name');}catch(e){}
    var merged=null;try{merged=r.getMergedAreasOrNullObject();merged.load('areas/items/address,isNullObject');}catch(e){merged=null;}
    await context.sync();
    if(r.rowCount*r.columnCount>10000)throw Error('一次最多处理 10000 个单元格。');activeSheet=String(ws.name);
    var p={book:String(wb.name||'workbook'),sheet:String(ws.name),address:local(r.address),row:r.rowIndex+1,column:r.columnIndex+1,
      rows:r.values.map(function(row){return row.map(function(v){return v===''?undefined:v;});}),formulas:r.formulas.map(function(row){return row.map(function(f){return typeof f==='string'&&f.charAt(0)==='='?f:'';});}),merges:[]};
    if(merged&&!merged.isNullObject&&merged.areas){var areas=merged.areas.items.map(function(a){var x=ws.getRange(local(a.address));x.load('address,rowIndex,columnIndex,rowCount,columnCount,values,formulas');return x;});await context.sync();
      areas.forEach(function(m){if(m.rowIndex<r.rowIndex||m.columnIndex<r.columnIndex||m.rowIndex+m.rowCount>r.rowIndex+r.rowCount||m.columnIndex+m.columnCount>r.columnIndex+r.columnCount)throw Error('请把合并单元格完整选中。');var f=m.formulas[0][0];if(typeof f==='string'&&f.charAt(0)==='=')throw Error('合并区域含公式，请先单独处理。');p.merges.push({address:local(m.address),value:m.values[0][0],rows:m.rowCount,cols:m.columnCount});});}
    return p;}
  function same(a,b){return JSON.stringify([a.rows,a.formulas,a.merges.map(function(m){return m.address;}).sort()])===JSON.stringify([b.rows,b.formulas,b.merges.map(function(m){return m.address;}).sort()]);}
  async function recheck(context,p){var r=context.workbook.getSelectedRange();var ws=context.workbook.worksheets.getItem(p.sheet);var src=ws.getRange(p.address);src.select();await context.sync();var now=await snapshot(context);
    if(now.sheet!==p.sheet||now.address!==p.address||!same(now,p))throw Error('来源数据已改变，请重新读取和预览。');return ws;}
  function lit(v){return E&&E.literal?E.literal(v):v;}
  async function apply(planned){var id=planned.id,s=planned.source,d=planned.data;
    return Excel().run(async function(context){var src=await recheck(context,s);
      if(id==='uniqueCount'||id==='groupSummary'){var dst=context.workbook.worksheets.add();var rows=[[s.rows[0][d.key]||'分类','次数'].concat(id==='groupSummary'?['合计']:[])].concat(d.result.map(function(r){return id==='groupSummary'?r:r.slice(0,2);}));
        var out=dst.getRangeByIndexes(0,0,rows.length,rows[0].length);out.values=rows.map(function(r){return r.map(lit);});dst.getRangeByIndexes(0,0,1,rows[0].length).format.font.bold=true;out.format.autofitColumns();dst.load('name');await context.sync();
        return '已生成 '+d.result.length+' 类汇总，工作表「'+dst.name+'」。原数据保留，汇总为静态结果。';}
      var copy=src.copy('End');copy.load('name');await context.sync();var r=copy.getRange(s.address);
      try{
        if(id==='unmergeFill')s.merges.forEach(function(m){var x=copy.getRange(m.address),fill=[];x.unmerge();for(var a=0;a<m.rows;a++){var line=[];for(var b=0;b<m.cols;b++)line.push(lit(m.value));fill.push(line);}x.values=fill;});
        if(id==='mergeSame')d.runs.forEach(function(run){var x=r.getCell(run.start,0).getResizedRange(run.length-1,0);if(run.length>1)r.getCell(run.start+1,0).getResizedRange(run.length-2,0).clear('Contents');x.merge(false);x.format.verticalAlignment='Center';});
        if(id==='cleanInvisible'||id==='clearConstants'||id==='textIds')s.rows.forEach(function(row,i){row.forEach(function(v,j){if(s.formulas[i][j])return;var c=r.getCell(i,j);
          if(id==='cleanInvisible'&&typeof v==='string'&&E.clean(v)!==v)c.values=[[lit(E.clean(v))]];if(id==='clearConstants'&&v!==null&&v!==undefined&&v!=='')c.clear('Contents');if(id==='textIds')c.numberFormat=[['@']];});});
        if(id==='dropdown'||id==='preventDuplicates'){r.dataValidation.clear();var col=F.col(s.column),first=col+s.row,span='$'+col+'$'+s.row+':$'+col+'$'+(s.row+s.rows.length-1);
          r.dataValidation.rule=id==='dropdown'?{list:{inCellDropDown:true,source:d.choices.join(',')}}:{custom:{formula:'=OR('+first+'="",SUMPRODUCT(--('+span+'='+first+'))<=1)'}};
          r.dataValidation.ignoreBlanks=true;r.dataValidation.errorAlert={showAlert:true,style:'Stop',title:'请检查输入',message:id==='dropdown'?'请从下拉选项中选择。':'这一项已经存在，请勿重复录入。'};}
        if(id==='filterValues')copy.autoFilter.apply(r,d.column-1,{filterOn:'Values',values:d.values.map(String)});
        await context.sync();
      }catch(e){throw Error(e.message+' 原表保留，副本「'+copy.name+'」可能只完成部分处理，请检查。');}
      copy.activate();await context.sync();return '已生成处理副本「'+copy.name+'」。原表「'+s.sheet+'」保留。检查满意后再保存文件。';});}
  if(root.QuickHost){root.QuickHost.capture=function(){return Excel().run(function(context){return snapshot(context);});};root.QuickHost.apply=apply;}
  // ---- 中文公式助手（和 WPS 版 OfficeFormulaHost 同样的检查）----
  function canonical(f){return String(f).replace(/\s+/g,'').toUpperCase();}
  async function targetCells(context,p){var ws=context.workbook.worksheets.getItemOrNullObject(p.target.sheet);ws.load('isNullObject');await context.sync();if(ws.isNullObject)throw Error('找不到工作表：'+p.target.sheet);
    var prot=ws.protection;prot.load('protected');var cells=[];for(var i=0;i<p.formulas.length;i++){var c=ws.getRange(F.address({sheet:p.target.sheet,r1:p.target.r1+i,r2:p.target.r1+i,c1:p.target.c1,c2:p.target.c1},false));c.load('values,formulas');var m=null;try{m=c.getMergedAreasOrNullObject();m.load('isNullObject');}catch(e){}cells.push({c:c,m:m});}
    await context.sync();if(prot.protected)throw Error('结果工作表已保护，请先在 WPS 中解除保护。'.replace('WPS','Excel'));
    cells.forEach(function(x){var f=x.c.formulas[0][0];if(typeof f==='string'&&f.charAt(0)==='=')throw Error('结果区域存在公式或无法确认状态。');if(x.c.values[0][0]!==''&&x.c.values[0][0]!==null)throw Error('结果区域已有内容，请选择空白位置。');if(x.m&&!x.m.isNullObject)throw Error('结果区域含合并单元格。');});return cells.map(function(x){return x.c;});}
  async function bookName(context){var wb=context.workbook;try{wb.load('name');await context.sync();}catch(e){}return String(wb.name||'workbook');}
  root.OfficeFormulaHost={
    capture:function(){return Excel().run(async function(context){var r=context.workbook.getSelectedRange();r.load('rowIndex,columnIndex,rowCount,columnCount');var ws=r.worksheet;ws.load('name');var book=await bookName(context);await context.sync();
      if(r.rowCount*r.columnCount>100000)throw Error('区域超过 10 万个单元格。');activeSheet=String(ws.name);return {book:book,reference:F.full({sheet:String(ws.name),r1:r.rowIndex+1,r2:r.rowIndex+r.rowCount,c1:r.columnIndex+1,c2:r.columnIndex+r.columnCount})};});},
    preview:function(app,input){return Excel().run(async function(context){var book=await bookName(context);if(input.book&&input.book!==book)throw Error('已切换工作簿，请重新选择区域。');var p=F.plan(input);
      for(var i=0;i<p.dependencies.length;i++){var dep=p.dependencies[i],ws=context.workbook.worksheets.getItemOrNullObject(dep.sheet);ws.load('isNullObject');await context.sync();if(ws.isNullObject)throw Error('找不到工作表：'+dep.sheet);}
      await targetCells(context,p);return {book:book,plan:p,at:Date.now()};});},
    apply:function(app,prepared){return Excel().run(async function(context){var book=await bookName(context);if(book!==prepared.book)throw Error('当前工作簿已切换或另存，请重新预览。');var p=prepared.plan,cells=await targetCells(context,p);
      cells.forEach(function(c,i){c.formulas=[[p.formulas[i]]];});await context.sync();cells.forEach(function(c){c.load('text');});await context.sync();
      var errors=[];cells.forEach(function(c,i){var t=String(c.text&&c.text[0]&&c.text[0][0]||'');if(/^#(?:N\/A|NAME\?|VALUE!|REF!|DIV\/0!|NUM!|SPILL!|CALC!)/.test(t))errors.push(i+1);});return {book:prepared.book,plan:p,at:Date.now(),errors:errors};});},
    undo:function(app,record){return Excel().run(async function(context){var book=await bookName(context);if(book!==record.book)throw Error('当前工作簿已切换或另存，请重新预览。');var p=record.plan,ws=context.workbook.worksheets.getItem(p.target.sheet),cells=[];
      for(var i=0;i<p.formulas.length;i++){var c=ws.getRange(F.address({sheet:p.target.sheet,r1:p.target.r1+i,r2:p.target.r1+i,c1:p.target.c1,c2:p.target.c1},false));c.load('formulas');cells.push(c);}await context.sync();
      cells.forEach(function(c,i){if(canonical(c.formulas[0][0])!==canonical(p.formulas[i]))throw Error('结果公式已被修改，已停止撤回，保留你的修改。');});cells.forEach(function(c){c.clear('Contents');});await context.sync();return cells.length;});}
  };
  // 公式助手需要知道当前工作表名称（WPS 里是同步读取，这里在打开侧栏和读取选区时更新）。
  function refreshSheet(){try{return Excel().run(async function(context){var ws=context.workbook.worksheets.getActiveWorksheet();ws.load('name');await context.sync();activeSheet=String(ws.name);});}catch(e){return Promise.resolve();}}
  root.OfficeShim={setHost:function(h){hostKey=h;root.OfficeTools.hostReady=!!h;if(h==='excel'){refreshSheet();try{root.Office.context.document.addHandlerAsync&&root.Excel.run(async function(context){context.workbook.worksheets.onActivated.add(function(){return refreshSheet();});await context.sync();}).catch(function(){});}catch(e){}}},setAsker:function(fn){asker=fn;}};
})(window);
// Microsoft Office 版：把只适用于 WPS 的操作指引换成 Office 里的对应入口（中文＋日文），其余卡片内容不变。
(function(root){
 'use strict';
 var G='Office 内置技巧',O={
  readingGuide:{steps:'Microsoft 365：「视图」→「焦点单元格」→ 用旁边的小箭头选颜色。旧版 Excel 没有这个开关，可以选中整行或整列查看。',note:'只改变显示效果，不修改单元格底色。此入口提供操作指引。'},
  renameGuide:{desc:'点窗口顶部的文件名就能改',steps:'先保存文件 → 点窗口顶部的文件名 → 在弹出的「文件名」框里输入新名字 → 按 Enter。旧版 Office 请关闭文件后在文件资源管理器里改名。',note:'此入口提供操作指引；重命名可能影响已有跨文件链接，请先检查。'},
  pivotGuide:{note:'需要已有透视表；字段搜索框在「数据透视表字段」窗格顶部。简单汇总可用「按分类汇总」。'},
  mcE18:{desc:'放在单元格中，或用“图片适应单元格”',steps:'Microsoft 365：「插入 → 图片 → 放在单元格中」，或选中已插入的图片 → 右键「放在单元格中」，图片随单元格大小显示。保持浮动图片时：一次插入多张 → 每张拖到目标格附近 → 选中这些单元格，用工具箱的「图片适应单元格」统一缩放居中。',note:'来源：作者题目＋官方资料补全。旧版 Excel 没有「放在单元格中」，请用浮动图片加「图片适应单元格」。'},
  mcW28:{steps:'Windows 11：用「截图工具」截下图片 →「文本操作」→「复制所有文本」→ 粘贴到 Word 核对并修改 → 重新排版或替换原图。',note:'来源：作者题目＋官方资料补全。识别会有错字，金额、编号逐个核对；不要修改他人证件、票据或合同图片里的内容冒充原件。'},
  mcE28:{desc:'生成图片后插入，扫码确认',steps:'Excel 没有内置二维码：用单位认可的二维码生成工具把网址或文字做成图片 →「插入 → 图片」→ 调整大小，四周留白 → 插入后用手机扫一次确认。'},
  mcO04:{desc:'从自动恢复记录里找',steps:'重新打开 Excel / Word → 左侧「文档恢复」里查看；或「文件 → 信息 → 管理文档（管理工作簿）→ 恢复未保存的文档」→ 打开候选文件核对时间和内容 → 另存为新文件。',note:'来源：作者题目＋官方资料补全。能否找回取决于是否开启了自动恢复，并且已经保存过记录。'},
  mcO08:{label:'把自动恢复位置改到其他盘',desc:'文件 → 选项 → 保存',steps:'「文件 → 选项 → 保存」→ 查看「自动恢复文件位置」→ 目标盘存在、可写、空间够时再改路径 → 保留「保存自动恢复信息」→ 编辑测试文档几分钟，确认新位置有恢复文件。',note:'来源：视频画面（原视频讲 WPS，这里换成 Office 的对应入口）。改路径不会迁移旧文件，也不会马上释放 C 盘。'}
 },JA={'Office 内置技巧':'Office の標準機能','生成图片后插入，扫码确认':'画像を作って挿入し、読み取って確認','Excel 没有内置二维码：用单位认可的二维码生成工具把网址或文字做成图片 →「插入 → 图片」→ 调整大小，四周留白 → 插入后用手机扫一次确认。':'Excel には QR コード機能がありません：社内で認められた QR コード作成ツールで URL や文字を画像にする →「挿入 → 画像」→ 大きさを調整し、周囲に余白を残す → 挿入後にスマートフォンで一度読み取って確認します。',
  'Microsoft 365：「视图」→「焦点单元格」→ 用旁边的小箭头选颜色。旧版 Excel 没有这个开关，可以选中整行或整列查看。':'Microsoft 365：「表示」→「フォーカス セル」→ 横の小さな矢印で色を選びます。古い Excel にはこの機能がないため、行全体または列全体を選択して確認してください。',
  '只改变显示效果，不修改单元格底色。此入口提供操作指引。':'表示だけを変え、セルの塗りつぶしは変更しません。ここでは操作手順を案内します。',
  '点窗口顶部的文件名就能改':'ウィンドウ上部のファイル名をクリックして変更',
  '先保存文件 → 点窗口顶部的文件名 → 在弹出的「文件名」框里输入新名字 → 按 Enter。旧版 Office 请关闭文件后在文件资源管理器里改名。':'ファイルを保存 → ウィンドウ上部のファイル名をクリック →「ファイル名」欄に新しい名前を入力 → Enter。古い Office ではファイルを閉じてからエクスプローラーで名前を変更してください。',
  '需要已有透视表；字段搜索框在「数据透视表字段」窗格顶部。简单汇总可用「按分类汇总」。':'ピボットテーブルが必要です。フィールドの検索欄は「ピボットテーブルのフィールド」ウィンドウの上部にあります。簡単な集計は「分類ごとに集計」を使えます。',
  '放在单元格中，或用“图片适应单元格”':'「セル内に配置」または「画像をセルに合わせる」',
  'Microsoft 365：「插入 → 图片 → 放在单元格中」，或选中已插入的图片 → 右键「放在单元格中」，图片随单元格大小显示。保持浮动图片时：一次插入多张 → 每张拖到目标格附近 → 选中这些单元格，用工具箱的「图片适应单元格」统一缩放居中。':'Microsoft 365：「挿入 → 画像 → セル内に配置」、または挿入済みの画像を右クリック →「セル内に配置」で、画像がセルの大きさに合わせて表示されます。浮動画像のままにする場合：まとめて挿入 → それぞれ目的のセルの近くへドラッグ → そのセルを選択し、ツールボックスの「画像をセルに合わせる」でまとめて縮小・中央配置します。',
  '来源：作者题目＋官方资料补全。旧版 Excel 没有「放在单元格中」，请用浮动图片加「图片适应单元格」。':'出典：作者の題目＋公式資料で補足。古い Excel には「セル内に配置」がないため、浮動画像と「画像をセルに合わせる」を使ってください。',
  'Windows 11：用「截图工具」截下图片 →「文本操作」→「复制所有文本」→ 粘贴到 Word 核对并修改 → 重新排版或替换原图。':'Windows 11：「Snipping Tool」で画像を切り取る →「テキスト アクション」→「すべてのテキストをコピー」→ Word に貼り付けて確認・修正 → レイアウトし直すか元の画像を置き換えます。',
  '从自动恢复记录里找':'自動回復の記録から探す',
  '重新打开 Excel / Word → 左侧「文档恢复」里查看；或「文件 → 信息 → 管理文档（管理工作簿）→ 恢复未保存的文档」→ 打开候选文件核对时间和内容 → 另存为新文件。':'Excel / Word を開き直す → 左側の「ドキュメントの回復」を確認。または「ファイル → 情報 → ドキュメントの管理（ブックの管理）→ 保存されていないドキュメントの回復」→ 候補のファイルを開いて日時と内容を確認 → 新しいファイルとして保存します。',
  '来源：作者题目＋官方资料补全。能否找回取决于是否开启了自动恢复，并且已经保存过记录。':'出典：作者の題目＋公式資料で補足。復元できるかどうかは、自動回復が有効で記録が保存されているかによります。',
  '把自动恢复位置改到其他盘':'自動回復の保存先を別のドライブに',
  '文件 → 选项 → 保存':'ファイル → オプション → 保存',
  '「文件 → 选项 → 保存」→ 查看「自动恢复文件位置」→ 目标盘存在、可写、空间够时再改路径 → 保留「保存自动恢复信息」→ 编辑测试文档几分钟，确认新位置有恢复文件。':'「ファイル → オプション → 保存」→「自動回復用ファイルの場所」を確認 → 移動先のドライブが存在し、書き込みでき、空きが十分なときだけパスを変更 →「自動回復用データを保存する」はオンのまま → テスト文書を数分編集し、新しい場所に回復用ファイルができることを確認します。',
  '来源：视频画面（原视频讲 WPS，这里换成 Office 的对应入口）。改路径不会迁移旧文件，也不会马上释放 C 盘。':'出典：動画の画面（元の動画は WPS 向け。ここでは Office の該当メニューに置き換え）。パスを変えても古いファイルは移動されず、C ドライブもすぐには空きません。'};
 function apply(list){(list||[]).forEach(function(c){var o=O[c.id];if(c.group==='原生技巧')c.group=G;if(o)Object.keys(o).forEach(function(k){c[k]=o[k];});});}
 apply(root.QuickCatalog&&root.QuickCatalog.items);apply(root.MethodCards&&root.MethodCards.cards);
 var J=root.OfficeI18nJa;if(J)Object.keys(JA).forEach(function(k){J[k]=JA[k];});
 root.OfficeShimGuides={overrides:O};
})(typeof window!=='undefined'?window:globalThis);
// Microsoft Office 版：AI 翻译与润色的读取 / 写回（Excel 逐格、Word 选中文字、PowerPoint 选中文字）。
(function(root){
 'use strict';
 function host(){return root.OfficeTools&&root.OfficeTools.host&&root.OfficeTools.host();}
 root.AiTextHost={
  read:async function(){var h=host();
   if(h==='et')return Excel.run(async function(ctx){var r=ctx.workbook.getSelectedRange();r.load('values,formulas,rowCount,columnCount,address');await ctx.sync();
     if(r.rowCount*r.columnCount>2000)throw Error('选区太大，请少选一些单元格（最多 200 个有文字的单元格）。');var items=[];
     r.values.forEach(function(row,i){row.forEach(function(v,j){var f=r.formulas[i][j];if(typeof f==='string'&&f.charAt(0)==='=')return;if(typeof v==='string'&&v.trim())items.push({row:i,col:j,text:v});});});
     if(!items.length)throw Error('选区里没有文字单元格（公式和数字不会处理）。');if(items.length>200)throw Error('一次最多处理 200 个文字单元格，请分批选。');
     return {kind:'cells',items:items,rows:r.rowCount,cols:r.columnCount,address:r.address};});
   if(h==='wpp')return PowerPoint.run(async function(ctx){var tr=ctx.presentation.getSelectedTextRangeOrNullObject();tr.load('text');await ctx.sync();
     if(tr.isNullObject||!String(tr.text||'').trim())throw Error('请先在文本框里选中文字。');return {kind:'text',text:String(tr.text).replace(/\r/g,'\n')};});
   return Word.run(async function(ctx){var s=ctx.document.getSelection();s.load('text');await ctx.sync();var text=String(s.text||'');if(!text.replace(/[\r\n\s]/g,''))throw Error('请先选中要处理的文字。');return {kind:'text',text:text.replace(/\r$/,'').replace(/\r/g,'\n')};});},
  write:async function(src,result,mode,confirmed){var h=host();
   if(src.kind==='cells')return Excel.run(async function(ctx){var r=ctx.workbook.getSelectedRange();r.load('address');await ctx.sync();if(r.address!==src.address)throw Error('选区已改变，请重新读取。');
     var target=mode==='replace'?r:r.getOffsetRange(0,src.cols);target.load('values');await ctx.sync();
     if(mode!=='replace'&&!confirmed){var used=0;src.items.forEach(function(it){var v=target.values[it.row][it.col];if(v!==''&&v!==null&&v!==undefined)used++;});if(used)return {confirm:'右侧区域有 '+used+' 个单元格已有内容，写入会覆盖。是否继续？'};}
     src.items.forEach(function(it,k){var v=result[k];target.getCell(it.row,it.col).values=[[/^[=+\-@]/.test(v)?"'"+v:v]];});await ctx.sync();
     return mode==='replace'?'已替换 '+src.items.length+' 个单元格（可按 Ctrl+Z 撤销）。':'已把结果写到右侧一列（'+src.items.length+' 个单元格），原文保留。';});
   var out=String(result);
   if(h==='wpp')return PowerPoint.run(async function(ctx){var tr=ctx.presentation.getSelectedTextRangeOrNullObject();tr.load('text');await ctx.sync();if(tr.isNullObject)throw Error('请重新选中文字。');
     tr.text=mode==='replace'?out:String(tr.text)+'\n'+out;await ctx.sync();return mode==='replace'?'已替换选中的文字（可按 Ctrl+Z 撤销）。':'已把结果加在原文后面。';});
   return Word.run(async function(ctx){var s=ctx.document.getSelection();if(mode==='replace')s.insertText(out.replace(/\r?\n/g,'\r'),'Replace');else out.split(/\r\n?|\n/).forEach(function(line){s=s.insertParagraph(line,'After');});await ctx.sync();return mode==='replace'?'已替换选中的文字（可按 Ctrl+Z 撤销）。':'已把结果插在原文后面。';});},
  // AI 对话的“插入文档”。
  insert:async function(text,confirmed){var h=host(),out=String(text).replace(/\s+$/,'');if(!out)throw Error('没有可插入的内容。');
   if(h==='et')return Excel.run(async function(ctx){var cell=ctx.workbook.getSelectedRange().getCell(0,0);cell.load('address');var lines=out.split('\n');if(lines.length>500)throw Error('内容超过 500 行，请复制后手动粘贴。');var target=cell.getResizedRange(lines.length-1,0);target.load('values');await ctx.sync();
     var formula=lines.length===1&&/^=/.test(lines[0].trim()),at=String(cell.address).replace(/^.*!/,'');
     if(!confirmed){var used=target.values.filter(function(r){return r[0]!==''&&r[0]!==null;}).length;return {confirm:(formula?'将把公式写入 '+at:'将从 '+at+' 开始向下写入 '+lines.length+' 行')+(used?'，其中 '+used+' 个单元格已有内容，会被覆盖':'')+'。是否继续？'};}
     if(formula)cell.formulas=[[lines[0].trim()]];else target.values=lines.map(function(l){return [/^[=+\-@]/.test(l)?"'"+l:l];});await ctx.sync();return formula?'已把公式写入 '+at+'。':'已从 '+at+' 开始写入 '+lines.length+' 行。';});
   if(h==='wpp')return PowerPoint.run(async function(ctx){var tr=ctx.presentation.getSelectedTextRangeOrNullObject();tr.load('text');await ctx.sync();
     if(!tr.isNullObject){tr.text=String(tr.text)+'\n'+out;await ctx.sync();return '已插在选中的文字后面。';}
     var slides=ctx.presentation.getSelectedSlides();slides.load('items');await ctx.sync();if(!slides.items.length)throw Error('请先点一下要插入的那一页。');slides.items[0].shapes.addTextBox(out,{left:60,top:60,width:600,height:300});await ctx.sync();return '已在当前页加了一个文本框（可以拖动位置）。';});
   return Word.run(async function(ctx){var s=ctx.document.getSelection();s.load('text');await ctx.sync();if(String(s.text||'').replace(/[\r\n\s]/g,''))out.split(/\r\n?|\n/).forEach(function(line){s=s.insertParagraph(line,'After');});else s.insertText(out.replace(/\r?\n/g,'\r'),'End');await ctx.sync();return '已插入到文档（可按 Ctrl+Z 撤销）。';});}
 };
})(typeof window!=='undefined'?window:globalThis);
// Microsoft Office 版：AI 帮写公式的读取（表头 + 前 5 行示例）和写入（只写空白单元格）。
(function(root){
 'use strict';
 function plain(a){return String(a||'').replace(/^.*!/,'');}
 root.AiFormulaHost={
  context:function(){return Excel.run(async function(ctx){var ws=ctx.workbook.worksheets.getActiveWorksheet(),used=ws.getUsedRange(true),cell=ctx.workbook.getActiveCell();
    ws.load('name');used.load('address,columnCount,rowCount,columnIndex');cell.load('address');await ctx.sync();
    var cols=Math.min(40,used.columnCount),rows=Math.min(6,used.rowCount),head=used.getCell(0,0).getResizedRange(rows-1,cols-1);head.load('values');await ctx.sync();
    var headers=[];for(var j=0;j<cols;j++){var n=used.columnIndex+j+1,s='';while(n>0){var m=(n-1)%26;s=String.fromCharCode(65+m)+s;n=Math.floor((n-1)/26);}var v=head.values[0][j];headers.push({col:s,name:v==null?'':String(v)});}
    return {sheet:ws.name,range:plain(used.address),target:plain(cell.address),headers:headers,rows:head.values.slice(1)};});},
  write:function(target,formula,sheet){return Excel.run(async function(ctx){var ws=ctx.workbook.worksheets.getActiveWorksheet();ws.load('name');await ctx.sync();
    if(sheet&&ws.name!==sheet)throw Error('当前工作表已切换，请回到“'+sheet+'”再写入，或重新让 AI 生成。');
    var c=ws.getRange(target);c.load('cellCount,formulas,values');await ctx.sync();if(c.cellCount!==1)throw Error('结果位置只能是一个单元格。');
    var f=c.formulas[0][0];if((typeof f==='string'&&f!=='')||(f!==''&&f!==null&&f!==undefined))throw Error('结果单元格 '+target+' 不是空的，请换一个空白单元格。');
    c.formulas=[[formula]];await ctx.sync();c.load('values');await ctx.sync();var v=c.values[0][0];
    return typeof v==='string'&&/^#/.test(v)?'公式已写入 '+target+'，但结果显示 '+v+'，请检查引用或数据类型。':'公式已写入 '+target+'，结果：'+(v===''||v==null?'（空）':v)+'。';});}
 };
})(typeof window!=='undefined'?window:globalThis);

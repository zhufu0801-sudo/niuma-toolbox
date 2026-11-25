// 办公工具箱 · 表格整理（WPS 表格）：常用但不容易做到的操作，一键完成。
// 原则：会改原数据的先确认；结果优先写到右边空白列或新工作表，不覆盖已有内容；WPS 表格的撤销不能还原工具写入的内容，所以提示先备份。
// Microsoft Excel 版在 public/office/office-sheetplus.js，规则一致。
(function(root){
  'use strict';
  var T=root.OfficeTools,S=root.OfficeSheetLogic,L=T.lib,cells=L.cells,rgb=L.rgb,active=L.active,position=L.position;
  var NOUNDO='WPS 表格的撤销不能还原这一步，重要文件请先“备份当前文件”。';
  function reg(id,label,desc,icon,fn,extra){var d={id:id,label:label,desc:desc,icon:icon,group:'sheetplus',undo:false,hosts:{et:fn}};if(extra)Object.keys(extra).forEach(function(k){d[k]=extra[k];});T.register(d);}
  function n(x){return Number(x);}
  function empty(v){return v===undefined||v===null||v==='';}
  function grid(range){var rows=n(range.Rows.Count),cols=n(range.Columns.Count),v=range.Value2;if(rows===1&&cols===1)return [[v]];if(Array.isArray(v)&&v.length===rows&&Array.isArray(v[0]))return v.map(function(r){return r.slice();});var out=[];for(var r=1;r<=rows;r++){var line=[];for(var c=1;c<=cols;c++)line.push(range.Cells.Item(r,c).Value2);out.push(line);}return out;}
  function put(range,values){if(values.length===1&&values[0].length===1)range.Value2=values[0][0];else range.Value2=values;}
  function safe(v){return typeof v==='string'&&/^\s*[=+@-]/.test(v)?"'"+v:(v==null?'':v);}
  function hasFormula(range){var f=range.HasFormula;return f!==false&&f!==0;}
  function merged(range){var m=range.MergeCells;return m!==false&&m!==0;}
  function sheet(a){return a.ActiveSheet;}
  function book(a){return a.ActiveWorkbook;}
  function addr(r){return String(typeof r.Address==='function'?r.Address(false,false):r.Address);}
  function color(hex){return rgb(parseInt(hex.slice(1,3),16),parseInt(hex.slice(3,5),16),parseInt(hex.slice(5,7),16));}
  // 结果放在选区右边紧挨着的一列：必须是空的，避免覆盖。
  function rightColumn(a,range,width){var out=range.Offset(0,n(range.Columns.Count)).Resize(n(range.Rows.Count),width||1),v=grid(out);for(var i=0;i<v.length;i++)for(var j=0;j<v[i].length;j++)if(!empty(v[i][j]))throw Error('结果要写在选区右边的 '+addr(out)+'，但那里已有内容。请先在右边插入空白列再执行。');if(hasFormula(out))throw Error('选区右边有公式，请先插入空白列。');return out;}
  function sheetNames(wb){var out=[];for(var i=1;i<=n(wb.Worksheets.Count);i++)out.push(String(wb.Worksheets.Item(i).Name));return out;}
  var cleanName=S.cleanName;
  function uniqueName(wb,base){var names=sheetNames(wb).map(function(x){return x.toLowerCase();}),name=cleanName(base),k=2;while(names.indexOf(name.toLowerCase())>=0){var tail='('+k+++')';name=cleanName(base).slice(0,31-tail.length)+tail;}return name;}
  function newSheet(a,name){var wb=book(a),last=wb.Worksheets.Item(n(wb.Worksheets.Count)),ws;try{ws=wb.Worksheets.Add(null,last);}catch(e){ws=wb.Worksheets.Add();}ws.Name=uniqueName(wb,S.nm(name));return ws;}
  function writeAt(ws,rows){if(!rows.length)return;var w=0;rows.forEach(function(r){if(r.length>w)w=r.length;});var data=rows.map(function(r){var x=r.map(safe);while(x.length<w)x.push('');return x;});put(ws.Range('A1').Resize(data.length,w),data);}
  // 新表只写了值：把原表每列的数字格式（日期、百分比、金额）带过去，否则日期会变成 46296 这样的序号。
  function colFormats(ws,row,c0,c1){var out=[];for(var j=c0;j<=c1;j++){var f=null;try{f=String(ws.Cells.Item(row,j).NumberFormat);}catch(e){}out.push(f&&!/^(General|G\/通用格式|@|常规)$/.test(f)?f:null);}return out;}
  function applyFormats(ws,fmts,startRow,count){if(count<1)return;fmts.forEach(function(f,j){if(!f)return;try{ws.Range(ws.Cells.Item(startRow,j+1),ws.Cells.Item(startRow+count-1,j+1)).NumberFormat=f;}catch(ignore){}});try{ws.Columns.AutoFit();}catch(ignore){}}
  function limit(range,max){if(n(range.Rows.Count)*n(range.Columns.Count)>(max||50000))throw Error('选区太大（超过 '+(max||50000)+' 个单元格），请缩小范围。');}
  function settle(a,target){if(position(a,'et',active(a,'et'))!==target)throw Error('当前工作表或选区已切换，请重新执行。');}

  // ---- 1. 聚光灯：高亮当前单元格所在的行和列，跟着选区移动；再点一次关闭 ----
  var spot={on:false,sheet:null,listener:null};
  function spotFormula(r,c){return '=OR(ROW()='+r+',COLUMN()='+c+')';}
  var isSpot=S.isSpot;
  function clearSpot(ws){if(!ws)return 0;var fc=ws.Cells.FormatConditions,k=0;for(var i=n(fc.Count);i>=1;i--){var f;try{f=fc.Item(i).Formula1;}catch(e){continue;}if(isSpot(f)){fc.Item(i).Delete();k++;}}return k;}
  function drawSpot(a){var ws=sheet(a),cell=a.ActiveCell;if(!ws||!cell)return;if(spot.sheet&&spot.sheet!==ws)try{clearSpot(spot.sheet);}catch(ignore){}clearSpot(ws);
    var used=ws.UsedRange,lastR=Math.max(n(used.Row)+n(used.Rows.Count)-1,n(cell.Row)+60,200),lastC=Math.max(n(used.Column)+n(used.Columns.Count)-1,n(cell.Column)+20,30),area=ws.Range(ws.Cells.Item(1,1),ws.Cells.Item(Math.min(lastR,1048576),Math.min(lastC,16384)));
    var f=area.FormatConditions.Add(2,undefined,spotFormula(n(cell.Row),n(cell.Column)));f.Interior.Color=color('#FFE699');try{f.StopIfTrue=false;}catch(ignore){}spot.sheet=ws;}
  function events(a){try{return a.ApiEvent||root.wps&&root.wps.ApiEvent;}catch(e){return null;}}
  reg('spotlight','聚光灯','高亮当前行和列，看长表不串行','spotlight',function(c){var a=c.a,ev=events(a);
    if(spot.on){try{if(ev&&spot.listener)ev.RemoveApiEventListener('SheetSelectionChange',spot.listener);}catch(ignore){}var k=0;try{k=clearSpot(spot.sheet||sheet(a));}catch(ignore){}spot.on=false;spot.listener=null;spot.sheet=null;return '聚光灯已关闭，高亮已清除。';}
    drawSpot(a);
    if(ev&&typeof ev.AddApiEventListener==='function'){spot.listener=function(){try{drawSpot(T.app());}catch(ignore){}};try{ev.AddApiEventListener('SheetSelectionChange',spot.listener);spot.on=true;return '聚光灯已打开：当前行和列会跟着选中的单元格移动。再点一次“聚光灯”关闭并清除高亮（打印前请先关闭，否则高亮会一起打印）。';}catch(e){spot.listener=null;}}
    spot.on=true;return '已高亮当前行和列。这个 WPS 版本不能跟着选区自动移动：换位置后再点一次“聚光灯”关闭，再点一次重新高亮。';},{panel:true});
  T.spotlight={clear:clearSpot,isSpot:isSpot,formula:spotFormula};

  // ---- 2. 隔行填色（条件格式，排序、筛选后仍是隔行）----
  var isBand=S.isBand;
  reg('bandRows','隔行填色','浅色隔行，长表看得清；再点一次取消','list',function(c){var a=c.a,r=cells(a),fc=r.FormatConditions,removed=0;
    for(var i=n(fc.Count);i>=1;i--){var f;try{f=fc.Item(i).Formula1;}catch(e){continue;}if(isBand(f)){fc.Item(i).Delete();removed++;}}
    if(removed)return '已取消选区里的隔行填色。';
    if(n(r.Rows.Count)<3)throw Error('请选中整张表（至少 3 行，可连同表头）。');
    var body=r.Offset(1,0).Resize(n(r.Rows.Count)-1,n(r.Columns.Count)),f2=body.FormatConditions.Add(2,undefined,'=MOD(ROW()-'+n(body.Row)+',2)=1');f2.Interior.Color=color('#DDEBF7');
    return '已给 '+addr(body)+' 隔行填色（第一行当作表头不填）。用的是条件格式，排序、筛选、插入行后仍是隔行。再点一次取消。';});

  // ---- 3. 文本数字 / 文本日期转成真正的数字和日期 ----
  var parseText=S.parseText;
  reg('textToNumber','文本转数字','左上角绿三角、不能求和的数字和日期转成真值','type',function(c){var a=c.a,r=cells(a);limit(r);if(merged(r))throw Error('选区里有合并单元格，请先取消合并。');
    var target=position(a,'et',active(a,'et')),todo=[],kept=0;
    for(var i=1;i<=n(r.Rows.Count);i++)for(var j=1;j<=n(r.Columns.Count);j++){var cell=r.Cells.Item(i,j),v=cell.Value2;if(typeof v!=='string'||hasFormula(cell))continue;var p=parseText(v);if(p)todo.push({cell:cell,p:p,before:v});else if(/^\s*[0-9０-９]/.test(v))kept++;}
    if(!todo.length)return '选区里没有“看起来是数字、实际是文字”的单元格。'+(kept?'有 '+kept+' 个以 0 开头或超过 15 位的编号保留为文字。':'');
    if(!L.ask('将把 '+todo.length+' 个文字格式的数字或日期转成真正的数值'+(kept?'（'+kept+' 个以 0 开头或超过 15 位的编号保留不动）':'')+'。\n'+NOUNDO+'\n是否执行？'))return '已取消。';
    settle(a,target);var dates=0;todo.forEach(function(x){if(x.cell.Value2!==x.before)return;try{if(x.p.date){x.cell.NumberFormat='yyyy/m/d';dates++;}else if(x.p.pct)x.cell.NumberFormat='0%';else if(String(x.cell.NumberFormat)==='@')x.cell.NumberFormat='General';}catch(ignore){}x.cell.Value2=x.p.v;});
    return '已转换 '+todo.length+' 个单元格（其中日期 '+dates+' 个），现在可以求和、排序和计算日期了。';});

  // ---- 4. 提取数字 / 中文 / 英文、去掉括号内容（结果写在右边空白列）----
  function extractor(id,label,desc,fn,word){reg(id,label,desc,'filter',function(c){var a=c.a,r=cells(a);if(n(r.Columns.Count)!==1)throw Error('请只选一列要处理的单元格。');limit(r,20000);var out=rightColumn(a,r),v=grid(r),res=v.map(function(row){return [fn(row[0]==null?'':String(row[0]))];}),hit=res.filter(function(x){return x[0]!=='';}).length;
    if(!hit)return '选区里没有可以提取的'+word+'。';put(out,res.map(function(x){return [safe(x[0])];}));return '已把'+word+'写到右边的 '+addr(out)+'（'+hit+' 行有结果），原列不变。';});}
  extractor('extractNumbers','提取数字','从混合文字里取出数字',function(x){return S.extract('numbers',x);},'数字');
  extractor('extractChinese','提取中文','只保留汉字（含日文汉字、假名）',function(x){return S.extract('chinese',x);},'中文');
  extractor('extractLetters','提取英文','只保留英文字母和空格',function(x){return S.extract('letters',x);},'英文');
  extractor('removeBrackets','去掉括号内容','删除（）()【】[] 及里面的文字',function(x){return S.extract('brackets',x);},'内容');

  // ---- 5. 两表核对：选两块区域（按住 Ctrl），标出不同 ----
  reg('compareRanges','两表核对','按住 Ctrl 选两块区域，标出不一样的地方','check',function(c){var a=c.a,sel=a.Selection;if(!sel||!sel.Areas||n(sel.Areas.Count)!==2)throw Error('请先选中第一块区域，按住 Ctrl 再选第二块区域（可以在不同位置，同一工作表），然后执行。');
    var A=sel.Areas.Item(1),B=sel.Areas.Item(2);limit(A,20000);limit(B,20000);var plan=S.compare(grid(A),grid(B)),yellow=color('#FFD966'),orange=color('#F8CBAD');
    if(plan.same){if(!plan.cells.length)return '两块区域逐格比对，完全一致。';
      if(!L.ask('逐格比对：有 '+plan.cells.length+' 个单元格不一样，将在两块区域里都填黄色标出。\n'+NOUNDO+'\n是否标出？'))return '已取消。';
      plan.cells.forEach(function(m){A.Cells.Item(m[0]+1,m[1]+1).Interior.Color=yellow;B.Cells.Item(m[0]+1,m[1]+1).Interior.Color=yellow;});return '已标出 '+plan.cells.length+' 处不同（两边都填了黄色）。';}
    if(!plan.onlyA.length&&!plan.onlyB.length&&!plan.diff.length)return '按第一列对照：两边的项目和内容都一致。';
    if(!L.ask('两块区域大小不同，按第一列的值对照：\n只在第一块里有 '+plan.onlyA.length+' 行，只在第二块里有 '+plan.onlyB.length+' 行，同一项目内容不同 '+plan.diff.length+' 行。\n将把这些行填色标出（只有一边有：橙色；内容不同：黄色）。\n'+NOUNDO+'\n是否标出？'))return '已取消。';
    plan.onlyA.forEach(function(i){A.Rows.Item(i+1).Interior.Color=orange;});plan.onlyB.forEach(function(i){B.Rows.Item(i+1).Interior.Color=orange;});plan.diff.forEach(function(p){A.Rows.Item(p[0]+1).Interior.Color=yellow;B.Rows.Item(p[1]+1).Interior.Color=yellow;});
    return '已标出：只在第一块 '+plan.onlyA.length+' 行、只在第二块 '+plan.onlyB.length+' 行（橙色），内容不同 '+plan.diff.length+' 行（黄色）。';},{panel:true});

  // ---- 6. 工资条：每行前加表头，行间空一行，生成在新工作表 ----
  reg('payrollSlips','工资条','每人一行带表头，生成新表，可直接打印裁开','copy',function(c){var a=c.a,r=cells(a);if(n(r.Rows.Count)<2)throw Error('请连同表头一起选中工资表（第一行是表头）。');if(merged(r))throw Error('选区里有合并单元格。请选中只有一行表头、没有合并单元格的区域。');limit(r,30000);
    var v=grid(r),head=v[0],rows=v.slice(1).filter(function(x){return x.some(function(y){return !empty(y);});});if(!rows.length)throw Error('表头下面没有数据。');
    var src=sheet(a),out=[];rows.forEach(function(x,i){out.push(head);out.push(x);if(i<rows.length-1)out.push([]);});
    var ws=newSheet(a,'工资条');writeAt(ws,out);
    try{var cols=n(r.Columns.Count);for(var i=0;i<rows.length;i++){var top=i*3+1;var hr=ws.Range(ws.Cells.Item(top,1),ws.Cells.Item(top,cols));hr.Font.Bold=true;hr.Interior.Color=color('#DDEBF7');var box=ws.Range(ws.Cells.Item(top,1),ws.Cells.Item(top+1,cols));box.Borders.LineStyle=1;
      for(var j=1;j<=cols;j++){var nf=r.Cells.Item(i+2,j).NumberFormat;if(nf&&String(nf)!=='General')ws.Cells.Item(top+1,j).NumberFormat=nf;}}ws.Columns.AutoFit();}catch(ignore){}
    try{src.Activate();}catch(ignore){}
    return '已生成工作表“'+ws.Name+'”：'+rows.length+' 人，每人一行带表头，中间空一行方便裁开。原表不变。';});

  // ---- 7. 合并单元格连续编号 ----
  reg('mergedNumbering','合并单元格编号','大小不一的合并单元格填 1、2、3…','list',function(c){var a=c.a,r=cells(a);if(n(r.Columns.Count)!==1)throw Error('请只选要编号的那一列（可以包含大小不一的合并单元格）。');
    var target=position(a,'et',active(a,'et')),tops=[],row=1,total=n(r.Rows.Count);
    while(row<=total){var cell=r.Cells.Item(row,1),area=cell.MergeArea,h=area?n(area.Rows.Count):1;if(n(area.Row)===n(cell.Row))tops.push(cell);row+=Math.max(1,h-(n(cell.Row)-n(area.Row)));}
    var filled=tops.filter(function(x){return !empty(x.Value2);}).length;
    if(!L.ask('将给 '+tops.length+' 个单元格（合并的按一个算）依次填 1～'+tops.length+(filled?'，其中 '+filled+' 个已有内容会被覆盖':'')+'。\n'+NOUNDO+'\n是否执行？'))return '已取消。';
    settle(a,target);tops.forEach(function(x,i){x.Value2=i+1;});return '已编号 1～'+tops.length+'。增删行后再执行一次即可重新编号。';});

  // ---- 8. 不打乱顺序的排名（右边空白列写 RANK 公式）----
  reg('rankKeepOrder','排名不打乱顺序','右边一列显示名次，原顺序不变','up',function(c){var a=c.a,r=cells(a);if(n(r.Columns.Count)!==1)throw Error('请只选要排名的那一列数字（不含表头）。');var v=grid(r),nums=v.filter(function(x){return typeof x[0]==='number';}).length;if(nums<2)throw Error('选区里的数字少于 2 个。文字格式的数字请先用“文本转数字”。');
    var out=rightColumn(a,r),ref=String(r.Address(true,true)),first=r.Cells.Item(1,1).Address(false,false);out.Cells.Item(1,1).Formula='=IF(ISNUMBER('+first+'),RANK('+first+','+ref+',0),"")';if(n(r.Rows.Count)>1)out.Cells.Item(1,1).Copy(out);
    return '已在 '+addr(out)+' 写入排名公式：最大的为第 1 名，并列名次相同；改数字后名次自动更新，原顺序不变。';});

  // ---- 9. 到期提醒（条件格式）----
  var isDue=S.isDue;
  reg('dueReminder','到期提醒','已过期标红、30 天内到期标黄，每天自动更新','check',function(c){var a=c.a,r=cells(a),fc=r.FormatConditions,removed=0;for(var i=n(fc.Count);i>=1;i--){var f;try{f=fc.Item(i).Formula1;}catch(e){continue;}if(isDue(f)){fc.Item(i).Delete();removed++;}}if(removed)return '已取消选区里的到期提醒。';
    if(n(r.Columns.Count)!==1)throw Error('请只选到期日期那一列（不含表头）。');var v=grid(r),dates=v.filter(function(x){return typeof x[0]==='number';}).length;if(!dates)throw Error('选区里没有日期。文字格式的日期请先用“文本转数字”。');
    // WPS 和 Excel 一样，条件格式公式里的相对引用以“活动单元格”为准：先把活动单元格放到选区左上角。
    var tl=r.Cells.Item(1,1);try{tl.Activate();}catch(ignore){}var c1=String(tl.Address(false,false)),red=r.FormatConditions.Add(2,undefined,'=AND(ISNUMBER('+c1+'),'+c1+'<TODAY())');red.Interior.Color=color('#F8CBAD');red.Font.Color=color('#9C0006');
    var yel=r.FormatConditions.Add(2,undefined,'=AND(ISNUMBER('+c1+'),'+c1+'>=TODAY(),'+c1+'<=TODAY()+30)');yel.Interior.Color=color('#FFEB9C');
    if(String(red.Formula1).replace(/\$/g,'').indexOf(c1)<0){try{red.Delete();yel.Delete();}catch(ignore){}throw Error('没能设置到期提醒：请从日期列第一格向下选中，再执行一次。');}
    return '已设置到期提醒：已过期标红，30 天内到期标黄，每天打开自动更新。再点一次取消。';});

  // ---- 10. 进度条（数据条）----
  reg('progressBar','进度条','数字或百分比显示成单元格里的进度条','columns',function(c){var a=c.a,r=cells(a),v=grid(r),nums=[];v.forEach(function(row){row.forEach(function(x){if(typeof x==='number')nums.push(x);});});if(!nums.length)throw Error('选区里没有数字。');
    var bar=r.FormatConditions.AddDatabar(),pct=nums.every(function(x){return x>=0&&x<=1;});try{bar.BarColor.Color=color('#5B9BD5');}catch(ignore){}
    try{if(pct){bar.MinPoint.Modify(0,0);bar.MaxPoint.Modify(0,1);}else if(nums.every(function(x){return x>=0&&x<=100;})){bar.MinPoint.Modify(0,0);bar.MaxPoint.Modify(0,100);}}catch(ignore){}
    return '已添加进度条（'+(pct?'按 0～100%':'按数值大小')+'）。删除请用“开始 → 条件格式 → 清除规则”。';});

  // ---- 11. 多列合并成一格 / 一格拆成多行 / 多列堆成一列 ----
  reg('joinColumns','多列合一格','每行几列的内容合成一格（右边空白列）','columns',function(c){var a=c.a,r=cells(a);if(n(r.Columns.Count)<2)throw Error('请选中要合并的多列（至少两列）。');limit(r,30000);var sep=c.options&&c.options.separator!=null?String(c.options.separator):'';var out=rightColumn(a,r),v=grid(r),res=v.map(function(row){return [safe(row.filter(function(x){return !empty(x);}).map(String).join(sep))];});put(out,res);
    return '已把每行的内容合成一格，写在 '+addr(out)+(sep?'（用“'+sep+'”分隔）':'')+'。原来的列不变，确认后可以删掉。';},{fields:[{key:'separator',label:'分隔符（可留空）',placeholder:'例如：、 或 -',required:false}]});
  reg('splitToRows','一格拆成多行','“张三、李四”拆成两行，其他列跟着复制，生成新表','down',function(c){var a=c.a,r=cells(a);if(n(r.Columns.Count)!==1)throw Error('请只选要拆分的那一列（不含表头）。同一行的其他列会一起复制。');var ws=sheet(a),used=ws.UsedRange,c0=n(used.Column),c1=c0+n(used.Columns.Count)-1,col=n(r.Column);
    if(col<c0||col>c1)throw Error('选区不在表格数据范围内。');var head=n(r.Row)>n(used.Row)?grid(ws.Range(ws.Cells.Item(n(used.Row),c0),ws.Cells.Item(n(used.Row),c1)))[0]:null,rows=grid(ws.Range(ws.Cells.Item(n(r.Row),c0),ws.Cells.Item(n(r.Row)+n(r.Rows.Count)-1,c1))),k=col-c0,res=S.splitRows(head,rows,k),out=res.rows,split=res.split;
    if(!split)return '选区里没有用顿号、逗号、分号、斜杠或换行隔开的多个内容。';if(out.length>100000)throw Error('拆分后超过 10 万行，请分批处理。');
    var nws=newSheet(a,'拆分结果');writeAt(nws,out);applyFormats(nws,colFormats(ws,n(r.Row),c0,c1),head?2:1,out.length-(head?1:0));try{ws.Activate();}catch(ignore){}return '已生成工作表“'+nws.Name+'”：'+split+' 行被拆开，共 '+(out.length-(head?1:0))+' 行。原表不变。';});
  reg('stackColumns','多列堆成一列','几列的内容依次接成一列，生成新表','down',function(c){var a=c.a,r=cells(a);if(n(r.Columns.Count)<2)throw Error('请选中要合成一列的多列。');limit(r);var v=grid(r),out=[];for(var j=0;j<v[0].length;j++)for(var i=0;i<v.length;i++)if(!empty(v[i][j]))out.push([v[i][j]]);
    if(!out.length)return '选区是空的。';var ws=newSheet(a,'合成一列'),src=sheet(a);writeAt(ws,out);try{src.Activate();}catch(ignore){}return '已把 '+v[0].length+' 列按列顺序接成一列，共 '+out.length+' 个（跳过空白），写在新工作表“'+ws.Name+'”。';});

  // ---- 12. 打印：每页打印表头、一页宽 ----
  reg('printTitles','每页打印表头','打印多页时每页都带表头','header',function(c){var a=c.a,r=cells(a),ws=sheet(a),top=n(r.Row),rows=n(r.Rows.Count);if(rows>5)throw Error('请只选表头所在的行（1～5 行），例如点一下第 1 行的任意单元格。');
    var spec='$'+top+':$'+(top+rows-1);ws.PageSetup.PrintTitleRows=spec;return '已设置：打印时每页顶端重复第 '+top+(rows>1?'～'+(top+rows-1):'')+' 行。可在“页面 → 打印标题”里修改。';});
  reg('fitOnePageWide','打印一页宽','所有列缩到一页宽，行数不限','columns',function(c){var ps=sheet(c.a).PageSetup;ps.Zoom=false;ps.FitToPagesWide=1;ps.FitToPagesTall=false;return '已设置打印缩放：所有列放进一页宽，行多就自动分页。打印预览里确认字是否太小，太小可改为横向纸张。';});

  // ---- 13. 删除空白行 ----
  reg('deleteBlankRows','删除空白行','删除选区里整行都空的行','trash',function(c){var a=c.a,r=cells(a),ws=sheet(a);if(n(r.Rows.Count)<2)throw Error('请选中要检查的整张表。');var used=ws.UsedRange,c0=n(used.Column),c1=c0+n(used.Columns.Count)-1,blank=[];
    var v=grid(ws.Range(ws.Cells.Item(n(r.Row),c0),ws.Cells.Item(n(r.Row)+n(r.Rows.Count)-1,c1)));v.forEach(function(row,i){if(row.every(empty))blank.push(n(r.Row)+i);});
    if(!blank.length)return '选区里没有整行都空的行（这一行在整张表里任何一列有内容都不算空）。';
    if(!L.ask('将删除 '+blank.length+' 个整行都空的行（第 '+blank.slice(0,10).join('、')+(blank.length>10?' 等':'')+' 行），下面的行上移。\n'+NOUNDO+'\n是否执行？'))return '已取消。';
    for(var i=blank.length-1;i>=0;i--)ws.Rows.Item(blank[i]).Delete();return '已删除 '+blank.length+' 个空白行。';});

  // ---- 14. 工作表目录、批量改名 ----
  reg('sheetIndex','工作表目录','生成带链接的目录页，点名字就跳过去','list',function(c){var a=c.a,wb=book(a),names=sheetNames(wb),title=S.nm('目录'),oldName=names.filter(S.isIndex)[0],old=oldName?wb.Worksheets.Item(oldName):null;
    if(old){var top=grid(old.Range('A1:B1'))[0];if(!S.isIndex(String(top[0])))throw Error('已有一个名叫“目录”的工作表且不是工具箱生成的，请先改名后再生成。');}
    var ws=old||wb.Worksheets.Add(wb.Worksheets.Item(1));if(!old)ws.Name=title;ws.Cells.Clear();ws.Cells.Item(1,1).Value2=String(ws.Name);ws.Cells.Item(1,2).Value2=S.nm('（工具箱生成，再次执行会更新）');ws.Cells.Item(1,1).Font.Bold=true;
    var k=2;sheetNames(wb).forEach(function(s2){if(S.isIndex(s2))return;var q="'"+s2.replace(/'/g,"''")+"'";ws.Cells.Item(k,1).Formula='=HYPERLINK("#'+q.replace(/"/g,'""')+'!A1","'+s2.replace(/"/g,'""')+'")';k++;});
    try{ws.Columns.Item(1).ColumnWidth=30;}catch(ignore){}ws.Activate();return '已生成“目录”工作表：'+(k-2)+' 个工作表，点名字就能跳过去。增删工作表后再执行一次即可更新。';},{panel:true});
  reg('renameSheets','批量改工作表名','把工作表名里的文字统一替换','replace',function(c){var o=c.options||{};var wb=book(c.a),plan=S.renamePlan(sheetNames(wb),o);if(!plan.length)return '没有需要改名的工作表。';
    if(!L.ask('将改名 '+plan.length+' 个工作表，例如：'+plan.slice(0,3).map(function(p){return p[0]+' → '+p[1];}).join('；')+'。\n公式里引用到这些表的地方会自动跟着改。是否执行？'))return '已取消。';
    plan.forEach(function(p,i){wb.Worksheets.Item(p[0]).Name='~tbx'+i;});plan.forEach(function(p,i){wb.Worksheets.Item('~tbx'+i).Name=p[1];});return '已改名 '+plan.length+' 个工作表。';},{fields:[{key:'find',label:'把这段文字',placeholder:'例如：2025'},{key:'replace',label:'换成',placeholder:'例如：2026（可留空表示删除）'},{key:'prefix',label:'前面加',placeholder:'可留空'},{key:'suffix',label:'后面加',placeholder:'可留空'}]});

  // ---- 15. 按列拆分成多个表、多表合并、每个工作表另存为文件 ----
  reg('splitByColumn','按列拆分成多表','按当前列的值，每类一张新工作表','copy',function(c){var a=c.a,r=cells(a),cell=a.ActiveCell;if(n(r.Rows.Count)<3)throw Error('请连同表头选中整张表，并让当前单元格（白色那格）落在要按它拆分的列上。');if(merged(r))throw Error('选区里有合并单元格，请先取消合并。');limit(r);
    var k=n(cell.Column)-n(r.Column);if(k<0||k>=n(r.Columns.Count))throw Error('当前单元格不在选区里。请先选表，再按 Tab 或 Enter 把白色当前格移到要拆分的列。');
    var v=grid(r),head=v[0],g=S.groupBy(v,k),groups=g.groups,order=g.order;
    if(order.length>60)throw Error('“'+head[k]+'”列有 '+order.length+' 种不同的值，超过 60 张表。请换一列拆分。');
    if(!L.ask('将按“'+(head[k]==null?'第 '+(k+1)+' 列':head[k])+'”拆成 '+order.length+' 张新工作表（每张带表头）：'+order.slice(0,6).map(function(x){return '“'+x+'”';}).join('、')+(order.length>6?' 等':'')+'。原表不变。是否执行？'))return '已取消。';
    var src=sheet(a),fm=colFormats(src,n(r.Row)+1,n(r.Column),n(r.Column)+n(r.Columns.Count)-1);order.forEach(function(key){var ws=newSheet(a,key);writeAt(ws,[head].concat(groups[key]));try{ws.Rows.Item(1).Font.Bold=true;}catch(ignore){}applyFormats(ws,fm,2,groups[key].length);});try{src.Activate();}catch(ignore){}
    return '已拆成 '+order.length+' 张工作表，原表不变。需要单独的文件可以再用“工作表另存为文件”。';});
  reg('mergeSheets','多表合并','表头相同的工作表上下合并到一张新表','insert',function(c){var a=c.a,wb=book(a),list=[];
    for(var i=1;i<=n(wb.Worksheets.Count);i++){var ws=wb.Worksheets.Item(i),name=String(ws.Name);if(S.isIndex(name)||S.isMerged(name))continue;var ur=ws.UsedRange;list.push({name:name,values:grid(ur),ws:ws,row:n(ur.Row),col:n(ur.Column),cols:n(ur.Columns.Count)});}
    var plan=S.mergePlan(list);if(plan.parts<2)throw Error('表头和第一张表相同的工作表不到 2 张，没法合并。'+(plan.mismatch.length?'表头不同的有：'+plan.mismatch.join('、'):''));
    if(!L.ask('将把 '+plan.parts+' 张表头相同的工作表（共 '+(plan.rows.length-1)+' 行）合并到一张新表，并加一列“来源工作表”。'+(plan.mismatch.length?'\n表头不同、不合并：'+plan.mismatch.join('、'):'')+'\n原表不变。是否执行？'))return '已取消。';
    var ws2=newSheet(a,'多表合并');writeAt(ws2,plan.rows);try{ws2.Rows.Item(1).Font.Bold=true;}catch(ignore){}var first=list.filter(function(x){return x.values.length>1;})[0];if(first)applyFormats(ws2,colFormats(first.ws,first.row+1,first.col,first.col+first.cols-1),2,plan.rows.length-1);ws2.Activate();
    return '已合并 '+plan.parts+' 张工作表，共 '+(plan.rows.length-1)+' 行，写在“'+ws2.Name+'”。'+(plan.mismatch.length?'跳过表头不同的：'+plan.mismatch.join('、')+'。':'');});
  reg('exportSheets','工作表另存为文件','每个工作表单独存成一个 xlsx','backup',function(c){var a=c.a,wb=book(a),names=sheetNames(wb).filter(function(x){return !S.isIndex(x);});if(names.length<1)throw Error('没有可导出的工作表。');
    var dialog=a.FileDialog(4);dialog.Title='选择保存位置（每个工作表一个文件）';try{if(wb.Path)dialog.InitialFileName=String(wb.Path)+'\\';}catch(ignore){}if(dialog.Show()!==-1)return '已取消。';var folder=String(dialog.SelectedItems.Item(1)).replace(/[\\\/]$/,''),stamp=new Date(),tag=String(stamp.getHours()).padStart(2,'0')+String(stamp.getMinutes()).padStart(2,'0')+String(stamp.getSeconds()).padStart(2,'0'),base=String(wb.Name).replace(/\.[^.]+$/,''),done=[];
    if(!L.ask('将把 '+names.length+' 个工作表分别另存为文件，放到：\n'+folder+'\n文件名形如“'+base+'-工作表名-'+tag+'.xlsx”，不会覆盖已有文件。原文件不变。是否执行？'))return '已取消。';
    names.forEach(function(name){wb.Worksheets.Item(name).Copy();var nb=a.ActiveWorkbook,file=folder+'\\'+cleanName(base+'-'+name+'-'+tag)+'.xlsx';try{nb.SaveAs(file,51);done.push(file);}finally{try{nb.Close(false);}catch(ignore){}}});try{wb.Activate();}catch(ignore){}
    return '已另存 '+done.length+' 个文件到 '+folder+'。';});

  // ---- 16. 一键图表 ----
  reg('quickChart','一键图表','按选区自动选柱形图或折线图（可按 Ctrl 选不相邻的列）','presentation',function(c){var a=c.a,sel=a.Selection,ws=sheet(a),r,v,last;
    if(sel&&sel.Areas&&n(sel.Areas.Count)>1){var parts=[];for(var i=1;i<=n(sel.Areas.Count);i++)parts.push(sel.Areas.Item(i));var h=n(parts[0].Rows.Count),top0=n(parts[0].Row);parts.forEach(function(p){if(n(p.Rows.Count)!==h||n(p.Row)!==top0)throw Error('按 Ctrl 选的几列需要从同一行开始、行数相同（例如 A1:A6 和 C1:C6）。');limit(p);});
      v=parts.map(grid).reduce(function(acc,g){return acc.map(function(row,i){return row.concat(g[i]);});},Array.apply(null,Array(h)).map(function(){return [];}));r=sel;last=parts[parts.length-1];}
    else{r=cells(a);v=grid(r);last=r;}
    if(v.length<2||v[0].length<2)throw Error('请连同表头选中数据（第一列是类别或日期，后面是数字）。不相邻的列可以按住 Ctrl 选。');var dates=v.slice(1).every(function(x){return typeof x[0]==='number'&&x[0]>20000&&x[0]<80000;})&&/日期|时间|月|年|date/i.test(String(v[0][0]||'日期'));
    var type=dates?4:51,left=n(last.Left)+n(last.Width)+20,top=n(last.Top),chart;try{chart=ws.Shapes.AddChart(type,left,top,480,280).Chart;}catch(e){chart=ws.ChartObjects().Add(left,top,480,280).Chart;chart.ChartType=type;}chart.SetSourceData(r);try{chart.HasTitle=true;chart.ChartTitle.Text=String(v[0][1]==null?'图表':v[0][1]);}catch(ignore){}
    return '已在选区右边插入'+(dates?'折线图（第一列是日期）':'柱形图')+'。可以拖动位置，双击修改样式。';},{panel:true});

  // ---- 17. 两列互换、转置 ----
  reg('swapColumns','两列互换','选两列（相邻或按 Ctrl 选），内容对调','replace',function(c){var a=c.a,sel=a.Selection,A,B;if(sel.Areas&&n(sel.Areas.Count)===2){A=sel.Areas.Item(1);B=sel.Areas.Item(2);}else{var r=cells(a);if(n(r.Columns.Count)!==2)throw Error('请选中相邻的两列，或按住 Ctrl 选两列同样行数的区域。');A=r.Columns.Item(1);B=r.Columns.Item(2);}
    if(n(A.Rows.Count)!==n(B.Rows.Count)||n(A.Columns.Count)!==1||n(B.Columns.Count)!==1)throw Error('两边需要都是一列、行数相同。');if(hasFormula(A)||hasFormula(B))throw Error('选区里有公式，对调会打乱引用。请先“粘贴为数值”。');if(merged(A)||merged(B))throw Error('选区里有合并单元格。');limit(A,20000);
    var va=grid(A),vb=grid(B);if(!L.ask('将对调 '+addr(A)+' 和 '+addr(B)+' 的内容（只换值，不换格式）。\n'+NOUNDO+'\n是否执行？'))return '已取消。';put(A,vb.map(function(x){return [safe(x[0])];}));put(B,va.map(function(x){return [safe(x[0])];}));return '已对调两列内容。';});
  reg('transposeRange','行列转置','行变列、列变行，生成在新工作表','replace',function(c){var a=c.a,r=cells(a);limit(r);var v=grid(r),t=v[0].map(function(_,j){return v.map(function(row){return row[j];});});var ws=newSheet(a,'转置'),src=sheet(a);writeAt(ws,t);try{src.Activate();}catch(ignore){}return '已把 '+v.length+' 行 × '+v[0].length+' 列转成 '+t.length+' 行 × '+t[0].length+' 列，写在新工作表“'+ws.Name+'”（只复制值）。';});

  // ---- 18. 姓名对齐（两字、三字姓名左右对齐，不改内容）----
  reg('alignNames','姓名对齐','两字、三字姓名两端对齐，内容不加空格','heading2',function(c){var a=c.a,r=cells(a);r.HorizontalAlignment=-4117;try{r.AddIndent=false;r.IndentLevel=1;}catch(ignore){}return '已设为“分散对齐”：两字、三字姓名两端对齐，单元格内容没有加空格，查找和匹配不受影响。';});

  // ---- 19. 按颜色求和计数 ----
  reg('colorSumCount','按颜色求和计数','按单元格底色分别统计个数和合计','filter',function(c){var a=c.a,r=cells(a);limit(r,20000);var stat={},order=[];
    for(var i=1;i<=n(r.Rows.Count);i++)for(var j=1;j<=n(r.Columns.Count);j++){var cell=r.Cells.Item(i,j),it=cell.Interior,key=n(it.ColorIndex)===-4142?'无填充':'#'+('000000'+n(it.Color).toString(16)).slice(-6).replace(/(..)(..)(..)/,'$3$2$1').toUpperCase(),v=cell.Value2;if(!stat[key]){stat[key]={count:0,sum:0,nums:0};order.push(key);}if(!empty(v))stat[key].count++;if(typeof v==='number'){stat[key].sum+=v;stat[key].nums++;}}
    var colored=order.filter(function(k){return k!=='无填充';});if(!colored.length)return '选区里的单元格都没有填充颜色。';
    return '按底色统计（条件格式显示的颜色不算）：\n'+order.map(function(k){var s=stat[k];return k+'：'+s.count+' 个，数字合计 '+(Math.round(s.sum*1e6)/1e6)+'（'+s.nums+' 个数字）';}).join('\n');},{panel:true});

  // ---- 20. 批量加前缀 / 后缀 ----
  reg('addPrefixSuffix','批量加前后缀','选区文字前后统一加内容','plus',function(c){var a=c.a,r=cells(a),o=c.options||{},pre=String(o.prefix||''),suf=String(o.suffix||'');if(!pre&&!suf)throw Error('请填写要加的前缀或后缀。');limit(r,20000);if(merged(r))throw Error('选区里有合并单元格。');
    var target=position(a,'et',active(a,'et')),todo=[],skipped=0;for(var i=1;i<=n(r.Rows.Count);i++)for(var j=1;j<=n(r.Columns.Count);j++){var cell=r.Cells.Item(i,j),v=cell.Value2;if(hasFormula(cell)){skipped++;continue;}if(empty(v))continue;todo.push({cell:cell,v:v});}
    if(!todo.length)return '选区里没有可以加前后缀的内容（空白和公式会跳过）。';
    if(!L.ask('将给 '+todo.length+' 个单元格'+(pre?'前面加“'+pre+'”':'')+(pre&&suf?'、':'')+(suf?'后面加“'+suf+'”':'')+'，结果变成文字'+(skipped?'（跳过 '+skipped+' 个公式）':'')+'。\n'+NOUNDO+'\n是否执行？'))return '已取消。';
    settle(a,target);todo.forEach(function(x){x.cell.Value2="'"+pre+(typeof x.v==='number'?String(x.v):x.v)+suf;});return '已处理 '+todo.length+' 个单元格。';},{fields:[{key:'prefix',label:'前缀',placeholder:'例如：BJ-'},{key:'suffix',label:'后缀',placeholder:'例如：号'}]});

  // ---- 21. 给公式包上 IFERROR ----
  reg('wrapIferror','公式错误显示为空','选中的公式外面包上 IFERROR，#N/A 等显示为空','clean',function(c){var a=c.a,r=cells(a);limit(r,20000);var target=position(a,'et',active(a,'et')),todo=[];
    for(var i=1;i<=n(r.Rows.Count);i++)for(var j=1;j<=n(r.Columns.Count);j++){var cell=r.Cells.Item(i,j);if(!hasFormula(cell))continue;var f=String(cell.Formula);if(/^=\s*IFERROR\(/i.test(f)||cell.HasArray)continue;todo.push({cell:cell,f:f});}
    if(!todo.length)return '选区里没有需要处理的公式（已包了 IFERROR 的和数组公式会跳过）。';
    if(!L.ask('将把 '+todo.length+' 个公式改成 =IFERROR(原公式,"")，出错时显示为空。\n注意：真正的错误也会被隐藏，核对数据时请先取消。\n'+NOUNDO+'\n是否执行？'))return '已取消。';
    settle(a,target);todo.forEach(function(x){x.cell.Formula='=IFERROR('+x.f.slice(1)+',"")';});return '已处理 '+todo.length+' 个公式。';});

  // ---- 0.14：抖音常见技巧的一键版 ----
  reg('maskData','敏感信息打码','身份证、手机号、银行卡、邮箱中间换成 *','clean',function(c){var a=c.a,r=cells(a);limit(r,20000);if(n(r.Columns.Count)!==1)throw Error('请只选一列（例如身份证号那一列）。');
    var v=grid(r),out=v.map(function(row){var m=S.mask(row[0]);return [m==null?'':m];}),hit=out.filter(function(x){return x[0]!=='';}).length;if(!hit)throw Error('选区里没有认得出的身份证号、手机号、银行卡号或邮箱。');
    var target=rightColumn(a,r,1);put(target,out);try{target.NumberFormat='@';put(target,out);}catch(ignore){}
    return '已在右边一列写好打码后的内容（'+hit+' 个），原数据不动。发给别人时复制右边这列，或删掉原来那列。';});
  reg('unitSum','带单位求和','“12元”“3.5kg”这样的文字也能求和','filter',function(c){var a=c.a,r=cells(a);limit(r,50000);var vals=[];grid(r).forEach(function(row){row.forEach(function(x){vals.push(x);});});var u=S.unitSum(vals);
    if(!u.count)throw Error('选区里没有能取出数字的单元格。');var unit=u.units.length===1?u.units[0]:'';
    return '合计：'+u.sum+unit+'（'+u.count+' 个单元格'+(u.units.length>1?'，单位不统一：'+u.units.slice(0,5).join('、'):'')+(u.skipped?'，'+u.skipped+' 个没有数字已跳过':'')+'）。结果没有写进表格，需要的话请手动填写。';});
  reg('columnsForPrint','细长表格分栏打印','两三列的长名单并排成几栏，省纸','columns',function(c){var a=c.a,r=cells(a);limit(r,50000);var cols=n(r.Columns.Count),rows=n(r.Rows.Count);if(rows<6)throw Error('请连同表头选中整张长表（至少 6 行）。');if(cols>4)throw Error('这个工具适合 1～4 列的细长表格，列太多并排后会太宽。');
    var groups=Number(c.options.groups)||(cols<=2?3:2);if(!(groups>=2&&groups<=6))throw Error('分几栏请填 2～6。');var v=grid(r),plan=S.columnsFor(v,groups),ws=newSheet(a,'分栏打印');writeAt(ws,plan.rows);
    try{ws.Rows.Item(1).Font.Bold=true;ws.Columns.AutoFit();ws.PageSetup.PrintTitleRows='$1:$1';ws.PageSetup.Zoom=false;ws.PageSetup.FitToPagesWide=1;ws.PageSetup.FitToPagesTall=false;}catch(ignore){}
    return '已在新工作表“'+String(ws.Name)+'”里分成 '+groups+' 栏（每栏 '+plan.per+' 行），每页都打印表头、整页宽度缩到一页。原表不变。';},{fields:[{key:'groups',label:'分几栏（不填自动）',placeholder:'2～6',required:false}]});
  reg('listToTable','一列整理成表格','“姓名、电话、地址”一条接一条的长列，整理成一行一条','table',function(c){var a=c.a,r=cells(a);limit(r,50000);if(n(r.Columns.Count)!==1)throw Error('请只选一列。');var k=Number(c.options.perRow);if(!(k>=2&&k<=20))throw Error('请填写每条记录有几行（2～20），例如“姓名、电话、地址”就是 3。');
    var list=grid(r).map(function(x){return x[0];}),plan=S.listToTable(list,k);if(!plan.rows.length)throw Error('选区是空的。');var ws=newSheet(a,'整理成表');writeAt(ws,plan.rows);try{ws.Columns.AutoFit();}catch(ignore){}
    return '已在新工作表“'+String(ws.Name)+'”里整理成 '+plan.rows.length+' 行、每行 '+k+' 列'+(plan.rest?'（最后一条只有 '+plan.rest+' 项，请检查是不是中间少了内容）':'')+'。原表不变。';},{fields:[{key:'perRow',label:'每条记录几行',placeholder:'例如 3',required:true}]});
  reg('rangeToImage','选区存为图片','表格区域另存为 PNG，代替长截图','pdf',function(c){var a=c.a,r=cells(a),ws=sheet(a);limit(r,200000);var dlg=a.FileDialog(4);dlg.Title='选择图片保存位置';if(dlg.Show()!==-1)return '已取消。';
    var folder=String(dlg.SelectedItems.Item(1)).replace(/[\\\/]$/,''),d=new Date(),tag=d.getFullYear()+String(d.getMonth()+1).padStart(2,'0')+String(d.getDate()).padStart(2,'0')+'-'+String(d.getHours()).padStart(2,'0')+String(d.getMinutes()).padStart(2,'0')+String(d.getSeconds()).padStart(2,'0'),file=folder+'\\'+S.cleanName(String(ws.Name))+'-'+tag+'.png';
    r.CopyPicture(1,2);var co=ws.ChartObjects().Add(n(r.Left),n(r.Top),n(r.Width),n(r.Height));try{co.Activate();co.Chart.Paste();co.Chart.Export(file);}finally{try{co.Delete();}catch(ignore){}}
    return '已保存图片：'+file;});
  reg('picturesByName','按名称批量插入图片','选中姓名列，按文件名把照片插到右边格子里','insert',function(c){var a=c.a,r=cells(a),ws=sheet(a);limit(r,5000);if(n(r.Columns.Count)!==1)throw Error('请只选一列名字（例如姓名或产品编号）。');
    var names=grid(r).map(function(x){return String(x[0]==null?'':x[0]).trim();}),target=r.Offset(0,1),dlg=a.FileDialog(3);dlg.Title='选择图片（可多选，文件名要和表格里的名字一样）';dlg.AllowMultiSelect=true;if(dlg.Show()!==-1)return '已取消。';
    var files={};for(var i=1;i<=n(dlg.SelectedItems.Count);i++){var f=String(dlg.SelectedItems.Item(i)),base=f.replace(/^.*[\\\/]/,'').replace(/\.[^.]+$/,'').trim();if(/\.(jpe?g|png|gif|bmp|webp|tiff?)$/i.test(f))files[base.toLowerCase()]=f;}
    var todo=[],miss=[];names.forEach(function(nm,k){if(!nm)return;var f=files[nm.toLowerCase()];if(f)todo.push({k:k,f:f});else miss.push(nm);});if(!todo.length)throw Error('没有文件名和表格里的名字对得上（例如“张三.jpg”对应“张三”）。');
    if(!L.ask('将把 '+todo.length+' 张图片插到选区右边一列的格子里（按格子大小缩放，跟着格子移动）。'+(miss.length?'\n没找到图片：'+miss.slice(0,8).join('、')+(miss.length>8?'…':''):'')+'\n是否继续？'))return '已取消。';
    var done=0;todo.forEach(function(t){var cell=target.Cells.Item(t.k+1,1),L0=n(cell.Left)+2,T0=n(cell.Top)+2,W=n(cell.Width)-4,H=n(cell.Height)-4;if(W<8||H<8)return;var p=ws.Shapes.AddPicture(t.f,0,-1,L0,T0,-1,-1);try{p.LockAspectRatio=-1;}catch(ignore){}var k2=Math.min(W/n(p.Width),H/n(p.Height));p.Width=n(p.Width)*k2;p.Left=L0+(W-n(p.Width))/2;p.Top=T0+(H-n(p.Height))/2;try{p.Placement=1;}catch(ignore){}done++;});
    return '已插入 '+done+' 张图片'+(miss.length?'；'+miss.length+' 个名字没找到图片':'')+(done<todo.length?'；'+(todo.length-done)+' 个格子太小，先把行高、列宽拉大再试':'')+'。';});
  T.register({id:'diagonalHeader',hosts:{et:function(c){var a=c.a,cell=a.ActiveCell;if(!cell)throw Error('请先点表头左上角的单元格。');var b=cell.Borders.Item(5),on=n(b.LineStyle)!==-4142&&n(b.LineStyle)!==0,t=String(cell.Value2==null?'':cell.Value2);if(on&&!/\n/.test(t)){b.LineStyle=-4142;return '已去掉 '+addr(cell)+' 的斜线。';}b.LineStyle=1;
    if(/\n/.test(t)){cell.WrapText=true;cell.HorizontalAlignment=-4131;cell.VerticalAlignment=-4108;return '已给 '+addr(cell)+' 加斜线并设为自动换行。第一行用空格推到右边即可。';}
    return '已给 '+addr(cell)+' 加了斜线。格子里先写右上角的标题，按 Alt+Enter 换行，再写左下角的标题，然后再点一次本按钮排好。';}}});

  // 第二批表格工具（sheet-data.js）共用这些读写函数。
  T.SheetPlus={uniqueName:uniqueName,grid:grid,put:put,safe:safe,empty:empty,rightColumn:rightColumn,newSheet:newSheet,writeAt:writeAt,limit:limit,addr:addr,hasFormula:hasFormula};
})(window);

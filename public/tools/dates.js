// 日期工具（2026-11-21）：插入今天日期、统一日期格式、算工作日、按年 / 季度 / 月汇总。
(function(root){'use strict';
var T=root.OfficeTools,L=T.lib,D=root.OfficeDates,cells=L.cells,active=L.active;
function n(v){return Number(v)||0;}
function ja(){try{return root.OfficeI18n&&root.OfficeI18n.lang()==='ja';}catch(e){return false;}}
var FORMAT_FIELD={key:'format',label:'日期格式',options:D.FORMATS.map(function(f){return {value:f.id,label:f.label};}),value:'slash'};
function fmtOf(o){var id=String((o||{}).format||'slash');return D.FORMATS.find(function(f){return f.id===id;})||D.FORMATS[0];}
function today(){var d=new Date();return new Date(d.getFullYear(),d.getMonth(),d.getDate());}
// 表格里写日期：能用数字格式表示的写真正的日期（可以排序、计算），和暦、带星期的写成文字。
function writeDate(cell,d,f){if(f.excel){cell.Value2=D.toSerial(d);cell.NumberFormat=f.excel;}else{cell.NumberFormat='@';cell.Value2=D.format(d,f.id);}}
function looksDateFormat(nf){nf=String(nf||'');return /[yY]|年|e|g/.test(nf)&&/[dD]|日/.test(nf)||/^m{1,2}[\/.-]d{1,2}/i.test(nf);}

// ---- 插入今天日期 ----
function wordInsert(a,text){var r=a.Selection.Range;if(n(r.Start)!==n(r.End)){r=a.Selection.Range.Duplicate;r.Collapse(0);}r.Text=text;try{r.Select();a.Selection.Collapse(0);}catch(ignore){}}
T.register({id:'insertToday',label:'插入今天日期',desc:'中文、日文、和暦、带星期等格式，放在光标处',icon:'calendar',group:'date',undo:false,ownConfirm:['wps','et','wpp'],keywords:'日期 今天 今日 日付 和暦 令和 年月日 星期 曜日',fields:[FORMAT_FIELD],hosts:{
  wps:function(c){var f=fmtOf(c.options),text=D.format(today(),f.id);wordInsert(c.a,text);return '已插入“'+text+'”。';},
  et:function(c){var f=fmtOf(c.options),cell=cells(c.a).Cells.Item(1,1),d=today();if(String(cell.Value2==null?'':cell.Value2)!==''&&!L.ask('选中的格子里已经有内容，要换成今天的日期吗？'))return '已取消。';writeDate(cell,d,f);return '已在格子里写入今天：'+D.format(d,f.id)+'。';},
  wpp:function(c){var f=fmtOf(c.options),text=D.format(today(),f.id),sel=c.a.ActiveWindow.Selection;if(n(sel.Type)!==3)throw Error('请先把光标放在幻灯片的文本框里。');sel.TextRange.InsertAfter?sel.TextRange.InsertAfter(text):(sel.TextRange.Text=String(sel.TextRange.Text||'')+text);return '已插入“'+text+'”。';}}});

// ---- 统一日期格式 ----
var IN_TEXT=/(?:令和|平成|昭和|[RH](?=\d))(?:元|\d{1,2})(?:年\d{1,2}月\d{1,2}日|\.\d{1,2}\.\d{1,2})|\d{4}(?:年\d{1,2}月\d{1,2}日|\/\d{1,2}\/\d{1,2}|-\d{1,2}-\d{1,2}|\.\d{1,2}\.\d{1,2})(?:[（(][日月火水木金土][）)]| ?星期[一二三四五六日天])?/g;
function sheetUnify(c){var a=c.a,f=fmtOf(c.options),r=cells(a),rows=n(r.Rows.Count),cols=n(r.Columns.Count);if(rows*cols>20000)throw Error('选区太大（超过 2 万个格子），请缩小范围。');
  var year=today().getFullYear(),todo=[],skipped=0;
  for(var i=1;i<=rows;i++)for(var j=1;j<=cols;j++){var cell=r.Cells.Item(i,j),v=cell.Value2,d=null;if(v==null||v==='')continue;
    if(typeof v==='number'){if(looksDateFormat(cell.NumberFormat))d=D.fromSerial(v);else{var s=String(v);if(/^\d{8}$/.test(s))d=D.parse(s);}}else d=D.parse(v,{defaultYear:year});
    if(d)todo.push([cell,d]);else skipped++;}
  if(!todo.length)throw Error('选区里没有认得出的日期。支持 2026/11/21、2026-11-21、2026年11月21日、20261121、令和8年11月21日、R8.11.21 等写法。');
  if(!L.ask('选区里认出 '+todo.length+' 个日期，统一成“'+f.label+'”'+(skipped?'；另有 '+skipped+' 个格子不是日期，保持不变':'')+'。\nWPS 表格里这一步无法撤销，执行前会自动备份已保存的文件。继续吗？'))return '已取消。';
  todo.forEach(function(x){writeDate(x[0],x[1],f);});
  return '已统一 '+todo.length+' 个日期为“'+f.label+'”'+(skipped?'，'+skipped+' 个不是日期的格子没动':'')+'。';}
function wordUnify(c){var a=c.a,d=c.d,f=fmtOf(c.options),s=L.wordSelection(a),start=n(s.Start),text=String(s.Text),hits=[],m;IN_TEXT.lastIndex=0;
  while((m=IN_TEXT.exec(text))){var p=D.parse(m[0]);if(p){var out=D.format(p,f.id);if(out!==m[0])hits.push({at:start+m.index,len:m[0].length,out:out});}}
  if(!hits.length)throw Error('选中的文字里没有需要改的日期（已经是“'+f.label+'”的不会再改）。');
  var record;try{try{record=a.UndoRecord;record.StartCustomRecord('统一日期格式');}catch(ignore){record=null;}
    for(var i=hits.length-1;i>=0;i--){var h=hits[i];d.Range(h.at,h.at+h.len).Text=h.out;}}finally{if(record)try{record.EndCustomRecord();}catch(ignore){}}
  return '已把 '+hits.length+' 处日期改成“'+f.label+'”。可以用 WPS 的撤销恢复。';}
T.register({id:'unifyDates',label:'统一日期格式',desc:'2026/11/21、令和8年…、20261121 等写法统一成一种',icon:'calendar',group:'date',undo:false,ownConfirm:['et','wps'],keywords:'日期 格式 统一 日付 統一 和暦 西暦 年月日 乱',fields:[FORMAT_FIELD],hosts:{et:sheetUnify,wps:wordUnify}});

// ---- 算工作日 ----
var REGION_FIELD={key:'region',label:'按哪里的节假日',options:[{value:'cn',label:'中国（含调休）'},{value:'jp',label:'日本（祝日）'},{value:'jpYearEnd',label:'日本（祝日＋年末年始 12/29～1/3）'}],value:'cn'};
function region(o){var r=String((o||{}).region||'cn');return r==='jpYearEnd'?{r:'jp',opts:{yearEnd:true}}:{r:r==='jp'?'jp':'cn',opts:{}};}
function workdayCalc(c){var o=c.options||{},g=region(o),y=today().getFullYear(),start=o.start?D.parse(o.start,{defaultYear:y}):today();
  if(!start)throw Error('开始日期看不懂，请写成 2026/11/21 这样的格式。');
  var lines=[],days=String(o.days||'').trim(),endText=String(o.end||'').trim(),fmt=function(d){return D.format(d,g.r==='jp'||ja()?'ja':'cnWeek');};
  if(endText){var end=D.parse(endText,{defaultYear:y});if(!end)throw Error('结束日期看不懂，请写成 2026/12/31 这样的格式。');var k=D.countWorkdays(start,end,g.r,g.opts);
    lines.push(D.format(start<end?start:end,'cn')+' 到 '+D.format(start<end?end:start,'cn')+'（含首尾）共 '+k.workdays+' 个工作日。');
    lines.push('去掉周末 '+k.weekend+' 天、节假日 '+k.holidays+' 天'+(k.makeup?'，含调休上班 '+k.makeup+' 天':'')+'。');
    if(k.missingYears.length)lines.push('注意：没有 '+k.missingYears.join('、')+' 年的放假安排，这些年份只按周六周日算。');}
  if(days){var count=parseInt(days,10);if(!count||Math.abs(count)>2000)throw Error('天数请填 1～2000 之间的数字（往前数可以填负数）。');var at=D.addWorkdays(start,count,g.r,g.opts);
    lines.push('从 '+D.format(start,'cn')+(count>0?' 往后数 ':' 往前数 ')+Math.abs(count)+' 个工作日是：'+fmt(at)+'。');
    if(g.r==='cn'&&D.cnYears().indexOf(at.getFullYear())<0)lines.push('注意：没有 '+at.getFullYear()+' 年的放假安排，只按周六周日算。');}
  if(!lines.length){var info=D.dayInfo(start,g.r,g.opts);lines.push(fmt(start)+(info.work?' 是工作日':' 不上班')+(info.name?'（'+info.name+'）':'')+'。');lines.push('填“结束日期”可以算两天之间有几个工作日；填“天数”可以算几个工作日以后是哪天。');}
  return lines.join('\n');}
T.register({id:'workdayCalc',label:'算工作日',desc:'两天之间有几个工作日、几个工作日后是哪天（中国调休、日本祝日）',icon:'calendar',group:'date',undo:false,ownConfirm:['wps','et','wpp'],keywords:'工作日 天数 交期 期限 调休 节假日 祝日 営業日 稼働日 何日後 納期',fields:[{key:'start',label:'开始日期（不填 = 今天）',placeholder:'例如：2026/11/21',required:false},{key:'end',label:'结束日期（可不填）',placeholder:'例如：2026/12/31',required:false},{key:'days',label:'或者：几个工作日以后（可不填）',placeholder:'例如：10',required:false},REGION_FIELD],hosts:{wps:workdayCalc,et:workdayCalc,wpp:workdayCalc}});

// ---- 按年 / 季度 / 月汇总 ----
function dateSummary(c){var a=c.a,o=c.options||{},kind=['year','quarter','month'].indexOf(o.period)>=0?o.period:'month',r=cells(a),rows=n(r.Rows.Count),cols=n(r.Columns.Count);
  if(rows<2||cols<2)throw Error('请选中至少两列：一列日期、一列数字（可以带表头）。');if(rows*cols>50000)throw Error('选区太大，请缩小范围。');
  var y=today().getFullYear(),grid=[],colFmt=[];for(var i=1;i<=rows;i++){var line=[];for(var j=1;j<=cols;j++)line.push(r.Cells.Item(i,j).Value2);grid.push(line);}
  // 数字格式按列看第 2 行（第 1 行可能是表头）。
  for(var jf=1;jf<=cols;jf++)colFmt.push(r.Cells.Item(Math.min(2,rows),jf).NumberFormat);
  function asDate(v,i,j){if(typeof v==='number')return looksDateFormat(colFmt[j])?D.fromSerial(v):null;return v==null?null:D.parse(v,{defaultYear:y});}
  var best=-1,bestHits=0;for(var j2=0;j2<cols;j2++){var hits=0;for(var i2=0;i2<rows;i2++)if(asDate(grid[i2][j2],i2,j2))hits++;if(hits>bestHits){bestHits=hits;best=j2;}}
  if(best<0||bestHits<1)throw Error('选区里找不到日期列。日期要写在一列里，例如 2026/11/21。');
  var val=-1;for(var j3=cols-1;j3>=0;j3--){if(j3===best)continue;var nums=0;for(var i3=0;i3<rows;i3++)if(typeof grid[i3][j3]==='number')nums++;if(nums>0){val=j3;break;}}
  if(val<0)throw Error('选区里找不到数字列（金额、数量等）。');
  var header=!asDate(grid[0][best],0,best),groups={},order=[],skipped=0;
  for(var i4=header?1:0;i4<rows;i4++){var d=asDate(grid[i4][best],i4,best),v=grid[i4][val];if(!d){if(grid[i4][best]!=null&&grid[i4][best]!=='')skipped++;continue;}var p=D.period(d,kind,ja()?'ja':'zh');if(!groups[p.key]){groups[p.key]={label:p.label,sum:0,count:0};order.push(p.key);}groups[p.key].sum+=typeof v==='number'?v:0;groups[p.key].count++;}
  order.sort();// 新表是用户的文档内容：日文界面写日文表头。
  var J=ja(),title=(J?{year:'年別集計',quarter:'四半期別集計',month:'月別集計'}:{year:'按年汇总',quarter:'按季度汇总',month:'按月汇总'})[kind],valueName=header&&grid[0][val]?String(grid[0][val]):J?'合計':'合计';
  var wb=a.ActiveWorkbook,last=wb.Worksheets.Item(n(wb.Worksheets.Count)),ws;try{ws=wb.Worksheets.Add(null,last);}catch(e){ws=wb.Worksheets.Add();}
  var name=title,k=2;while(true){try{ws.Name=name;break;}catch(e){name=title+'('+(k++)+')';if(k>50)break;}}
  var out=[J?['期間',valueName+'（合計）','件数']:['期间',valueName+'（合计）','笔数']].concat(order.map(function(key){var g=groups[key];return [g.label,g.sum,g.count];}));
  var total=order.reduce(function(s,key){return s+groups[key].sum;},0),cnt=order.reduce(function(s,key){return s+groups[key].count;},0);out.push([J?'総計':'总计',total,cnt]);
  ws.Range(ws.Cells.Item(1,1),ws.Cells.Item(out.length,3)).Value2=out;try{ws.Range(ws.Cells.Item(1,1),ws.Cells.Item(1,3)).Font.Bold=true;ws.Range(ws.Cells.Item(out.length,1),ws.Cells.Item(out.length,3)).Font.Bold=true;ws.Columns.Item(1).ColumnWidth=16;ws.Columns.Item(2).ColumnWidth=16;}catch(ignore){}
  return '已在新工作表“'+String(ws.Name)+'”里按'+{year:'年',quarter:'季度',month:'月'}[kind]+'汇总 '+cnt+' 笔，共 '+order.length+' 个期间'+(skipped?'；'+skipped+' 行日期看不懂，没算进去':'')+'。原表不变。';}
T.register({id:'dateSummary',label:'按年 / 季度 / 月汇总',desc:'选中日期列和金额列，在新工作表里按期间合计',icon:'calendar',group:'date',undo:false,keywords:'汇总 按月 按季度 按年 月度 季度 年度 集計 月別 四半期 年別 合計',fields:[{key:'period',label:'按什么汇总',options:[{value:'month',label:'按月'},{value:'quarter',label:'按季度'},{value:'year',label:'按年'}],value:'month'}],hosts:{et:dateSummary}});
})(window);

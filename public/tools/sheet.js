// 办公工具箱 · 表格模块（WPS 表格）。
(function(root){
  'use strict';
  var T=root.OfficeTools,L=T.lib,active=L.active,cells=L.cells,rgb=L.rgb,replaceOptions=L.replaceOptions,position=L.position;
  function editSheetText(a,range,id,options){
    if(range.Rows.Count*range.Columns.Count>2000)throw Error('文本整理一次最多处理 2000 个单元格，请缩小选区。');
    if(range.MergeCells!==false&&range.MergeCells!==0)throw Error('请先选择未合并的单元格区域。');
    var spec=id==='replaceText'?replaceOptions(options):null,changes=[],skipped=0,target=position(a,'et',active(a,'et'));
    for(var r=1;r<=range.Rows.Count;r++)for(var c=1;c<=range.Columns.Count;c++){
      var cell=range.Cells.Item(r,c);if(cell.HasFormula!==false&&cell.HasFormula!==0){skipped++;continue;}
      var value=cell.Value2;if(typeof value!=='string')continue;
      var next=id==='trimCells'?value.replace(/^[\s\u3000]+|[\s\u3000]+$/g,''):value.split(spec.find).join(spec.replacement);
      if(next!==value)changes.push({cell:cell,before:value,after:next});
    }
    if(!changes.length)return '没有需要修改的文本单元格。公式和数值不会被修改。';
    if(!L.ask('将修改选区中 '+changes.length+' 个文本单元格，跳过公式与数值。是否执行？'))return '已取消。';
    if(position(a,'et',active(a,'et'))!==target)throw Error('当前文档、工作表或选区已切换，请重新执行。');
    changes.forEach(function(change){if(change.cell.Value2!==change.before||change.cell.HasFormula)throw Error('单元格内容已变化，请重新检查后执行。');});
    var touched=[];try{changes.forEach(function(change){touched.push(change);change.cell.Value2="'"+change.after;});}
    catch(e){var failed=false;touched.reverse().forEach(function(change){try{change.cell.Value2="'"+change.before;}catch(ignore){failed=true;}});throw Error('文本整理失败：'+e.message+(failed?'。部分内容未恢复，请检查文档。':'。已恢复本次涉及的原文本。'));}
    return '已修改 '+changes.length+' 个文本单元格，跳过 '+skipped+' 个公式或无法判断的单元格；数值不变。';
  }
  T.register({id:'formatHeader',label:'美化表头',desc:'突出选区第一行',icon:'header',group:'table',hosts:{et:function(c){var range=cells(c.a),header=range.Rows.Item(1);header.Font.Bold=true;var prefs=T.settings().table;function color(hex){return rgb(parseInt(hex.slice(1,3),16),parseInt(hex.slice(3,5),16),parseInt(hex.slice(5,7),16));}header.Font.Color=color(prefs.headerColor);header.Interior.Color=color(prefs.headerBackground);try{header.Borders.Item(9).LineStyle=1;}catch(ignore){}return '已美化选区第一行。';}}});
  T.register({id:'fitColumns',label:'自动列宽',desc:'按内容调整宽度',icon:'columns',group:'table',hosts:{et:function(c){cells(c.a).Columns.AutoFit();return '已按选区内容调整列宽。';}}});
  T.register({id:'wrapCells',label:'自动换行',desc:'让单元格内容完整显示',icon:'wrap',group:'table',hosts:{et:function(c){var range=cells(c.a);range.WrapText=true;range.Rows.AutoFit();return '已开启自动换行并调整行高。';}}});
  T.register({id:'filterCells',label:'筛选开关',desc:'快速筛出所需数据',icon:'filter',group:'table',hosts:{et:function(c){cells(c.a).AutoFilter();return '已切换选区筛选按钮。';}}});
  T.register({id:'trimCells',label:'清理空格',desc:'移除文本首尾空白',icon:'trim',group:'table',hosts:{et:function(c){return editSheetText(c.a,cells(c.a),'trimCells',c.options);}}});
  // 删除重复行：整行内容完全相同的只保留第一次出现的那行；不重复的行按原顺序上移，选区以外的内容不动。
  function readValues(range,rows,cols){var v=range.Value2;if(Array.isArray(v)&&v.length===rows&&Array.isArray(v[0]))return v.map(function(r){return r.slice();});var out=[];for(var r=1;r<=rows;r++){var line=[];for(var c=1;c<=cols;c++)line.push(range.Cells.Item(r,c).Value2);out.push(line);}return out;}
  // 写回文字时，像数字、日期、时间、百分比、以 0 开头的编号这类文字，WPS 会自动改成数字或日期（13800000001 变成 1.38E+10，0123 丢掉 0）。
  // 前面加 ' 告诉 WPS 这是文字，单元格里不会显示这个 '。
  function looksConvertible(s){var t=String(s).trim();return /^[=+@-]/.test(t)||/^(true|false)$/i.test(t)||(/\d/.test(t)&&/^[\d\s.,，%％¥￥$€+\-−()（）\/:：年月日eE]+$/.test(t));}
  function literal(v){return typeof v==='string'&&looksConvertible(v)?"'"+v:(v==null?'':v);}
  // 只点了一个单元格时，自动选中它所在的整块数据（像 Ctrl+A 那样）。
  function dataRange(a){var r=cells(a);try{if(Number(r.Rows.Count)===1&&Number(r.Columns.Count)===1){var region=r.CurrentRegion;if(region&&Number(region.Rows.Count)>1){region.Select();return a.Selection&&a.Selection.Rows?a.Selection:region;}}}catch(ignore){}return r;}
  T.register({id:'removeDuplicates',label:'删除重复行',desc:'重复的行只保留第一条',icon:'dedupe',group:'table',ownConfirm:['et'],keywords:'重复 去重 一样的 相同 删掉重复 重複 ダブり',hosts:{et:function(c){
    var a=c.a,range=dataRange(a),rows=Number(range.Rows.Count),cols=Number(range.Columns.Count);
    if(rows<2)throw Error('请选中至少两行数据（可以连同表头一起选）。');
    if(range.MergeCells!==false&&range.MergeCells!==0)throw Error('选区里有合并单元格，请先取消合并再删除重复行。');
    if(range.HasFormula!==false&&range.HasFormula!==0)throw Error('选区里有公式，删除重复行会打乱公式引用。请先复制选区并“粘贴为数值”，再操作。');
    var target=position(a,'et',active(a,'et')),values=readValues(range,rows,cols),seen={},unique=[];
    // 只差前后空格的行不删，但要告诉客户（例如「 李娜」和「李娜」）。
    var loose={},nearly=0;
    values.forEach(function(row){var key=JSON.stringify(row.map(function(x){return x==null?'':x;}));var soft=JSON.stringify(row.map(function(x){return x==null?'':String(x).replace(/^[\s\u3000]+|[\s\u3000]+$/g,'');}));if(!seen[key]){seen[key]=true;unique.push(row);if(loose[soft])nearly++;}loose[soft]=true;});
    var removed=rows-unique.length,hint=nearly?'\n另有 '+nearly+' 行只是多了空格，看起来一样但不完全相同，这次不删；可以先用“清理空格”再删重复。':'';
    if(!removed)return '选区内没有重复的行。'+hint.replace(/^\n/,'');
    if(!L.ask('选区共 '+rows+' 行（已自动包含整块数据），其中 '+removed+' 行与上面的某一行完全相同。\n将删除这些重复行，其余行按原顺序上移；选区以外的内容不动。'+hint+'\n是否执行？'))return '已取消。';
    if(position(a,'et',active(a,'et'))!==target)throw Error('当前工作表或选区已切换，请重新执行。');
    try{
      range.Cells.Item(1,1).Resize(unique.length,cols).Value2=unique.map(function(r){return r.map(literal);});
      range.Cells.Item(unique.length+1,1).Resize(removed,cols).ClearContents();
    }catch(e){var restored=true;try{range.Value2=values.map(function(r){return r.map(literal);});}catch(ignore){restored=false;}throw Error('删除重复行失败：'+e.message+(restored?'。已恢复原来的内容。':'。部分内容未恢复，请检查选区内容。'));}
    return '已删除 '+removed+' 行重复数据，保留 '+unique.length+' 行。'+(nearly?'另有 '+nearly+' 行只差空格，没有删除。':'')+'WPS 撤销不能恢复这一步；如需还原，请关闭文件时选择不保存，或打开备份。';
  }}});
  T.register({id:'replaceText',hosts:{et:function(c){return editSheetText(c.a,cells(c.a),'replaceText',c.options);}}});
  // ---- 以下三项整理自 houge-office 技能 0.3.0 的方法卡（E16、E01、E02），按其“失败边界”加了检查 ----
  function merged(range){return range.MergeCells!==false&&range.MergeCells!==0;}
  function blank(cell){var v=cell.Value2;return (v===undefined||v===null||v==='')&&(cell.HasFormula===false||cell.HasFormula===0||cell.HasFormula===undefined);}
  T.register({id:'centerAcross',label:'跨列居中',desc:'标题居中但不合并单元格',icon:'header',group:'table',hosts:{et:function(c){
    var range=cells(c.a);if(range.Columns.Count<2)throw Error('请横向选择标题所在的一行及右侧要跨过的空白单元格，例如 A1:F1。');
    if(merged(range))throw Error('选区里有合并单元格。请先在副本中取消合并并核对内容，再使用跨列居中。');
    for(var r=1;r<=range.Rows.Count;r++){var filled=0;for(var k=1;k<=range.Columns.Count;k++)if(!blank(range.Cells.Item(r,k)))filled++;if(filled>1)throw Error('第 '+r+' 行选区里有多个非空单元格。跨列居中只适合“左边一格标题、右边都空”，不会清空其他内容。');}
    range.HorizontalAlignment=7;return '已设置跨列居中：标题看起来居中，单元格仍各自独立，可单独选取、排序和筛选。';}}});
  T.register({id:'shrinkToFit',label:'缩小字体填充',desc:'不改列宽也能显示完整',icon:'type',group:'table',hosts:{et:function(c){
    var range=cells(c.a),wrapped=range.WrapText===true||range.WrapText===-1;if(wrapped)range.WrapText=false;range.ShrinkToFit=true;
    return '已开启缩小字体填充，单元格的值不变。'+(wrapped?'原来的自动换行已关闭（两者不能同时生效）。':'')+'文字很长时字会很小，请检查可读性或改用自动换行。';}}});
  T.register({id:'fillBlanks',label:'空白填上一格',desc:'合并拆开后、分类留空时补齐',icon:'down',group:'table',hosts:{et:function(c){
    var range=cells(c.a);if(range.Rows.Count*range.Columns.Count>2000)throw Error('一次最多处理 2000 个单元格，请缩小选区。');
    if(merged(range))throw Error('选区里有合并单元格。请先用“拆开合并并补齐”或取消合并。');
    var target=position(c.a,'et',active(c.a,'et')),fills=[],edge=0;
    for(var k=1;k<=range.Columns.Count;k++){var last;for(var r=1;r<=range.Rows.Count;r++){var cell=range.Cells.Item(r,k);if(blank(cell)){if(last===undefined){edge++;continue;}fills.push({cell:cell,value:last});}else last=cell.Value2;}last=undefined;}
    if(!fills.length)return '选区里没有需要填充的真正空白单元格。只含空格的单元格和公式不算空白。';
    if(!L.ask('将把 '+fills.length+' 个空白单元格填成同列上方最近的内容'+(edge?'（每列顶端的 '+edge+' 个空白没有上一格，会跳过）':'')+'。\n请先确认空白确实表示“同上”，而不是缺失或 0。是否执行？'))return '已取消。';
    if(position(c.a,'et',active(c.a,'et'))!==target)throw Error('当前文档、工作表或选区已切换，请重新执行。');
    var done=[];try{fills.forEach(function(f){if(!blank(f.cell))throw Error('单元格内容已变化，请重新执行。');f.cell.Value2=f.value;done.push(f);});}
    catch(e){var failed=false;done.reverse().forEach(function(f){try{f.cell.Value2='';}catch(ignore){failed=true;}});throw Error('填充失败：'+e.message+(failed?'。部分单元格未恢复，请检查。':'。已恢复本次填入的单元格。'));}
    return '已填充 '+fills.length+' 个空白单元格（填入的是值，不是公式）。WPS 撤销不能恢复这一步，重要文件请先备份。';}}});
  T.register({id:'fitPicturesToCells',label:'图片适应单元格',desc:'选区内图片缩放居中、随单元格移动',icon:'columns',group:'table',hosts:{et:function(c){
    var range=cells(c.a),sheet=c.a.ActiveSheet,shapes=sheet&&sheet.Shapes,n=Number(shapes&&shapes.Count)||0,r0=Number(range.Row),c0=Number(range.Column),r1=r0+range.Rows.Count-1,c1=c0+range.Columns.Count-1,done=0,pad=2;
    if(!n)throw Error('当前工作表没有浮动图片。嵌入单元格的图片（DISPIMG）由 WPS 自动适应，不需要处理。');
    for(var i=1;i<=n;i++){var s=shapes.Item(i);if(Number(s.Type)!==13)continue;var cell=s.TopLeftCell;if(!cell)continue;var row=Number(cell.Row),col=Number(cell.Column);if(row<r0||row>r1||col<c0||col>c1)continue;
      var box=cell.MergeArea||cell,bw=Number(box.Width)-pad*2,bh=Number(box.Height)-pad*2;if(!(bw>0&&bh>0))continue;
      var k=Math.min(bw/Number(s.Width),bh/Number(s.Height));try{s.LockAspectRatio=-1;}catch(ignore){}var w=Number(s.Width)*k,h=Number(s.Height)*k;s.Width=w;s.Height=h;
      s.Left=Number(box.Left)+(Number(box.Width)-w)/2;s.Top=Number(box.Top)+(Number(box.Height)-h)/2;try{s.Placement=1;}catch(ignore){}done++;}
    if(!done)return '选区里没有左上角落在其中的图片。请连同图片所在的单元格一起选中。';
    return '已调整 '+done+' 张图片：按比例缩放到所在单元格内并居中，之后会随单元格移动和缩放。图片显示太小时，先把行高列宽调大再执行一次。';}}});
  T.register({id:'matchFormat',panel:false,hosts:{et:function(){throw Error('表格中请使用 WPS 自带的“格式刷”：先点参照单元格，再刷新粘贴的区域。');}}});
})(window);

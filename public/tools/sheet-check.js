// 办公工具箱 · 表格体检（WPS 表格）：找出“公式查不到、求和不对、数字变样”的常见原因。只检查、不修改；能一键修的给出对应工具。
(function(root){
  'use strict';
  var T=root.OfficeTools,L=T.lib;
  function n(x){return Number(x);}
  function hasFormula(c){var f=c.HasFormula;return f!==false&&f!==0;}
  function addr(c){return String(typeof c.Address==='function'?c.Address(false,false):c.Address);}
  var S=root.OfficeSheetLogic,classify=S.classify,report=S.report;
  T.register({id:'sheetCheck',label:'表格体检',desc:'找出查不到、求和不对、数字变样的原因',icon:'check',group:'sheetplus',undo:false,hosts:{et:function(c){
    var a=c.a,ws=a.ActiveSheet,sel=a.Selection,range=sel&&n(sel.Rows.Count)*n(sel.Columns.Count)>1?sel:ws.UsedRange,total=n(range.Rows.Count)*n(range.Columns.Count);
    if(total>200000)throw Error('检查范围超过 20 万个单元格，请先选中要检查的部分。');
    var found={},extra=[],scope=range===sel?'选区 '+addr(range):'整张工作表（'+addr(range)+'）';
    for(var i=1;i<=n(range.Rows.Count);i++)for(var j=1;j<=n(range.Columns.Count);j++){var cell=range.Cells.Item(i,j),v=cell.Value2,f=hasFormula(cell)?String(cell.Formula):'';if((v===undefined||v===null||v==='')&&!f)continue;
      var kinds=classify({value:v,formula:f,text:cell.Text});kinds.forEach(function(k){var x=found[k]||(found[k]={count:0,where:[]});x.count++;if(x.where.length<5)x.where.push(addr(cell));});}
    try{if(n(a.Calculation)===-4135)extra.push('计算方式是“手动”：改了数字公式不更新。到“公式 → 计算选项”改为“自动”。');}catch(ignore){}
    try{var circ=ws.CircularReference;if(circ)extra.push('有循环引用（'+addr(circ)+'）：公式直接或间接引用了自己，结果不可靠。');}catch(ignore){}
    if(range!==sel){try{var used=ws.UsedRange,lastR=n(used.Row)+n(used.Rows.Count)-1,lastC=n(used.Column)+n(used.Columns.Count)-1,end=ws.Cells.Find('*',undefined,-4163,2,1,2),endC=ws.Cells.Find('*',undefined,-4163,2,2,2);if(end&&endC){var realR=n(end.Row),realC=n(endC.Column);if(lastR-realR>500||lastC-realC>50)extra.push('“已使用区域”到 '+addr(ws.Cells.Item(lastR,lastC))+'，但数据只到第 '+realR+' 行、第 '+realC+' 列：多出来的空行空列会让文件变大、滚动条变短。选中多余的行列整行删除后保存即可。');}}catch(ignore){}}
    var lines=report(found,extra);
    if(!lines.length)return '体检完成（'+scope+'）：没有发现常见问题。';
    return '体检完成（'+scope+'），发现 '+lines.length+' 类问题：\n'+lines.join('\n')+'\n体检只查看，没有修改任何内容。';}}});
  T.SheetCheck={classify:classify,split:S.split,LABEL:S.LABEL};
})(window);

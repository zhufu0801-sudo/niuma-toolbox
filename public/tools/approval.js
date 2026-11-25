// 押印欄 / 审批栏（2026-11-21）：在文档或表格里插入“担当・課長・部長”这类盖章格子。
// 文字文档：插在光标处，靠右，上面一行写职位，下面是约 2 厘米见方的空格子；表格：从选中的格子开始写，不覆盖已有内容。
(function(root){'use strict';
var T=root.OfficeTools,L=T.lib,cells=L.cells;
function n(v){return Number(v)||0;}
function ja(){try{return root.OfficeI18n&&root.OfficeI18n.lang()==='ja';}catch(e){return false;}}
function labels(o){var raw=String((o||{}).labels||'').trim()||(ja()?'担当,課長,部長':'经办,审核,批准');var list=raw.split(/[,，、\s\/／|]+/).map(function(s){return s.trim();}).filter(Boolean);
  if(!list.length)throw Error('请填写至少一个职位，例如：担当,課長,部長。');if(list.length>8)throw Error('最多 8 个格子。');if(list.some(function(s){return s.length>10;}))throw Error('每个职位名称最多 10 个字。');return list;}
var BOX=57,HEAD=18;   // 磅：盖章格约 2 厘米，职位行约 0.6 厘米
function word(c){var a=c.a,d=c.d,list=labels(c.options),r=a.Selection.Range;if(n(r.Start)!==n(r.End)){r=a.Selection.Range.Duplicate;r.Collapse(0);}
  var record;try{try{record=a.UndoRecord;record.StartCustomRecord('插入押印欄');}catch(ignore){record=null;}
    var t=L.addWordTable(d,r,[list,list.map(function(){return '';})]);
    try{t.Rows.Alignment=2;}catch(ignore){}
    try{for(var j=1;j<=list.length;j++){t.Columns.Item(j).Width=BOX;var h=t.Cell(1,j).Range;h.ParagraphFormat.Alignment=1;h.Font.Size=9;}}catch(ignore){}
    try{var r1=t.Rows.Item(1),r2=t.Rows.Item(2);r1.HeightRule=2;r1.Height=HEAD;r2.HeightRule=2;r2.Height=BOX;}catch(ignore){}
  }finally{if(record)try{record.EndCustomRecord();}catch(ignore){}}
  return '已在光标处插入押印欄（'+list.join(' / ')+'），靠右放置。可以用 WPS 的撤销恢复。';}
function sheet(c){var a=c.a,list=labels(c.options),start=cells(a).Cells.Item(1,1),ws=a.ActiveSheet,row=n(start.Row),col=n(start.Column);
  var area=ws.Range(ws.Cells.Item(row,col),ws.Cells.Item(row+1,col+list.length-1));
  if(n(a.WorksheetFunction.CountA(area))>0)throw Error('从选中格子开始的 2 行 × '+list.length+' 列里已经有内容，请先选一块空白的位置。');
  area.Value2=[list,list.map(function(){return '';})];
  try{area.Borders.LineStyle=1;area.HorizontalAlignment=-4108;ws.Rows.Item(row).RowHeight=HEAD;ws.Rows.Item(row+1).RowHeight=BOX;for(var j=0;j<list.length;j++)ws.Columns.Item(col+j).ColumnWidth=8.5;}catch(ignore){}
  return '已从选中的格子开始插入押印欄（'+list.join(' / ')+'）。';}
T.register({id:'approvalBox',label:'插入押印欄（审批栏）',desc:'担当・課長・部長 / 经办・审核・批准 的盖章格子',icon:'stamp',group:'stamp',undo:false,ownConfirm:['wps','et'],keywords:'押印欄 承認欄 印鑑 ハンコ 判子 稟議 決裁 审批栏 盖章 签字栏 签批 担当 課長 部長',fields:[{key:'labels',label:'职位（用逗号隔开，可不填）',placeholder:'例如：担当,課長,部長 或 经办,审核,批准',required:false}],hosts:{wps:word,et:sheet}});
})(window);

// 中日本地化小工具（WPS 表格）。计算在 locale-engine.js；这里负责读取选区和写回。
// 派生结果（大写、生日性别、和历等）写到选区右侧，原数据不动；“统一格式”类在原处修改，执行前确认。
(function(root){
 'use strict';
 var T=root.OfficeTools,E=root.LocaleEngine;if(!T||!E)return;var L=T.lib;
 // 写进单元格的提示（不是界面文字，不会被自动翻译）按界面语言给出
 function tx(s){var i=root.OfficeI18n;return i&&i.t?i.t(s):s;}
 function grab(c,oneCol){var r=L.cells(c.a);if(r.Rows.Count*r.Columns.Count>2000)throw Error('一次最多处理 2000 个单元格，请缩小选区。');if(oneCol&&r.Columns.Count!==1)throw Error('请只选一列。');return r;}
 function values(r){var out=[];for(var i=1;i<=r.Rows.Count;i++)for(var j=1;j<=r.Columns.Count;j++){var cell=r.Cells.Item(i,j);if(cell.HasFormula)continue;var v=cell.Value2;if(v===undefined||v===null||v==='')continue;out.push({cell:cell,i:i,j:j,v:v});}return out;}
 function blank(v){return v===undefined||v===null||v==='';}
 // 写到右侧：width 是每个单元格输出几列。目标格已有内容时先确认。
 async function toRight(r,rows,width,put){
  if(!rows.length)return 0;var base=r.Offset(0,r.Columns.Count),used=0;
  rows.forEach(function(x){for(var k=0;k<width;k++)if(!blank(base.Cells.Item(x.i,x.j+k).Value2))used++;});
  if(used&&!(await L.ask('右侧有 '+used+' 个单元格已有内容，写入会覆盖。是否继续？')))throw Error('已取消。');
  rows.forEach(function(x){put(x,function(k){return base.Cells.Item(x.i,x.j+k);});});return rows.length;
 }
 function note(n,skip,what){return '已在选区右侧写入 '+n+' 个'+what+(skip?'，'+skip+' 个无法识别的单元格已跳过':'')+'。原数据未改动。';}
 T.register({id:'rmbUpper',label:'金额转人民币大写',desc:'1234.5 → 壹仟贰佰叁拾肆元伍角',icon:'type',group:'locale',hosts:{et:async function(c){
  var r=grab(c),rows=[],skip=0;values(r).forEach(function(x){var s=E.rmbUpper(x.v);if(s)rows.push({i:x.i,j:x.j,s:s});else skip++;});
  if(!rows.length)throw Error('选区里没有可以识别的金额。');var n=await toRight(r,rows,1,function(x,at){at(0).Value2=x.s;});return note(n,skip,'大写金额');}}});
 T.register({id:'idCardInfo',label:'身份证提取生日性别',desc:'校验号码，写出生日期和性别',icon:'columns',group:'locale',hosts:{et:async function(c){
  var r=grab(c,true),rows=[],bad=0;values(r).forEach(function(x){var info=E.idInfo(x.v);if(info.error)bad++;rows.push({i:x.i,j:x.j,info:info});});
  if(!rows.length)throw Error('选区里没有身份证号。');
  var n=await toRight(r,rows,2,function(x,at){if(x.info.error){at(0).Value2=x.info.error;at(1).Value2='';}else{at(0).Value2=x.info.birth;at(1).Value2=x.info.gender;}});
  return '已在右侧两列写入出生日期和性别（'+n+' 行）'+(bad?'，其中 '+bad+' 个号码有问题，已在右侧标出':'')+'。原号码未改动。身份证号属于个人信息，请妥善保管文件。';}}});
 T.register({id:'phoneCheck',label:'检查手机号',desc:'位数、号段是否正确',icon:'check',group:'locale',hosts:{et:async function(c){
  var r=grab(c),rows=[],bad=0;values(r).forEach(function(x){var s=E.phoneCheck(x.v);if(!s)return;if(!/^有效/.test(s))bad++;rows.push({i:x.i,j:x.j,s:s});});
  if(!rows.length)throw Error('选区里没有手机号。');var n=await toRight(r,rows,1,function(x,at){at(0).Value2=x.s;});
  return '已检查 '+n+' 个号码：'+(n-bad)+' 个有效，'+bad+' 个有问题（结果写在右侧一列）。';}}});
 T.register({id:'zenHan',label:'全角半角统一',desc:'英数字半角、片假名全角',icon:'type',group:'locale',hosts:{et:async function(c){
  var r=grab(c),changes=[];values(r).forEach(function(x){var s=E.zenHan(x.v);if(s!==null)changes.push({cell:x.cell,s:s,old:x.v});});
  if(!changes.length)return '选区里没有需要统一的全角英数字或半角片假名。';
  if(!(await L.ask('将修改 '+changes.length+' 个单元格：英文字母和数字改成半角，半角片假名改成全角（中日文标点和全角空格不动）。\n例：'+changes[0].old+' → '+changes[0].s+'\n是否执行？')))return '已取消。';
  changes.forEach(function(x){x.cell.Value2=x.s;});return '已统一 '+changes.length+' 个单元格。WPS 撤销可能不能恢复这一步，重要文件请先备份。';}}});
 T.register({id:'toWareki',label:'西历转和历',desc:'2026/11/12 → 令和8年11月12日',icon:'list',group:'locale',hosts:{et:async function(c){
  var r=grab(c),rows=[],skip=0;values(r).forEach(function(x){var s=E.toWareki(x.v);if(s)rows.push({i:x.i,j:x.j,s:s});else skip++;});
  if(!rows.length)throw Error('选区里没有可以识别的日期（1868 年 10 月 23 日以后）。');var n=await toRight(r,rows,1,function(x,at){at(0).Value2=x.s;});return note(n,skip,'和历日期');}}});
 T.register({id:'toSeireki',label:'和历转西历',desc:'令和8年11月12日 / R8.11.12 → 2026/11/12',icon:'list',group:'locale',hosts:{et:async function(c){
  var r=grab(c),rows=[],skip=0,bad=0;values(r).forEach(function(x){var s=E.toSeireki(x.v);if(!s){skip++;return;}if(s.error)bad++;rows.push({i:x.i,j:x.j,s:s});});
  if(!rows.length)throw Error('选区里没有可以识别的和历日期，例如“令和8年11月12日”或“R8.11.12”。');
  var n=await toRight(r,rows,1,function(x,at){var cell=at(0);if(x.s.error){cell.Value2=tx(x.s.error);return;}cell.NumberFormat='yyyy/m/d';cell.Value2=x.s.serial;});
  return note(n,skip,'西历日期（真正的日期，可以计算）')+(bad?' 其中 '+bad+' 个日期不在对应年号范围内，已标出。':'');}}});
 T.register({id:'postalFormat',label:'邮编格式统一',desc:'日本 1234567 → 123-4567',icon:'type',group:'locale',hosts:{et:async function(c){
  var r=grab(c),changes=[];values(r).forEach(function(x){var s=E.postal(x.v);if(s!==null)changes.push({cell:x.cell,s:s,old:x.v});});
  if(!changes.length)return '选区里的邮编格式已经统一，或没有 6～7 位的邮编。';
  if(!(await L.ask('将统一 '+changes.length+' 个邮编：日本 7 位写成 123-4567，中国 6 位去掉空格；单元格会设为文本格式，避免开头的 0 丢失。\n例：'+changes[0].old+' → '+changes[0].s+'\n是否执行？')))return '已取消。';
  changes.forEach(function(x){x.cell.NumberFormat='@';x.cell.Value2=x.s;});return '已统一 '+changes.length+' 个邮编。';}}});
 T.register({id:'furigana',label:'加注假名（读音）',desc:'用 PHONETIC 显示输入时的读音',icon:'type',group:'locale',hosts:{et:async function(c){
  var r=grab(c,true),rows=[];values(r).forEach(function(x){if(typeof x.v==='string')rows.push({i:x.i,j:x.j,addr:x.cell.Address(false,false)});});
  if(!rows.length)throw Error('选区里没有文字。');var n=await toRight(r,rows,1,function(x,at){at(0).Formula='=PHONETIC('+x.addr+')';});
  return '已在右侧写入 '+n+' 个读音公式。只有在本表格里用日文输入法打出的文字才有读音信息；从别处粘贴来的文字会原样显示。';}}});
})(window);

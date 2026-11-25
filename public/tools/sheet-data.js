// 表格 · 第二批新功能（2026-11-24）：万元显示、号码校验、合并多个工作簿、批量二维码 / 条形码、考勤统计。
// 规则和“表格整理”一样：结果写在右边空白列或新工作表，不覆盖已有内容；会改原表的先确认。
(function(root){'use strict';
var T=root.OfficeTools,L=T.lib,S=T.SheetPlus,ID=root.OfficeIdCheck,cells=L.cells;
function n(v){return Number(v)||0;}
function text(v){return String(v==null?'':v);}
function ja(){try{return root.OfficeI18n&&root.OfficeI18n.lang()==='ja';}catch(e){return false;}}
function M(zh,jp){return ja()?jp:zh;}
var CANCEL=function(){return M('已取消。','キャンセルしました。');};
function reg(id,def){def.id=id;def.group=def.group||'sheetdata';def.undo=false;T.register(def);}
async function tool(name,body){var s=await L.localFetch('/api/session',{}).then(function(r){return r.json();});var r=await L.localFetch('/api/tools/'+name,{method:'POST',headers:{'Content-Type':'application/json','X-Coach-Token':s.token},body:JSON.stringify(body)});var d=await r.json().catch(function(){return {};});if(!r.ok)throw Error(d.error||M('本机服务没有完成，请重试。','ローカルサービスで処理できませんでした。もう一度お試しください。'));return d;}

// ---- 万元显示 ----
var UNITS={wan:{f:'0!.0,"万"',label:'万（1 位小数）'},yi:{f:'0!.00,,"亿"',label:'亿（2 位小数）'},qian:{f:'#,##0,',label:'千（不带单位字）'},baiwan:{f:'#,##0,,',label:'百万（不带单位字）'},senyen:{f:'#,##0,"千円"',label:'千円'},general:{f:'General',label:'恢复成普通数字'}};
function wanUnit(c){var a=c.a,r=cells(a),u=UNITS[(c.options||{}).unit]?c.options.unit:'wan',x=UNITS[u],nums=0,rows=n(r.Rows.Count),cols=n(r.Columns.Count);S.limit(r,200000);
  var v=S.grid(r);for(var i=0;i<v.length;i++)for(var j=0;j<v[i].length;j++)if(typeof v[i][j]==='number')nums++;if(!nums)throw Error(M('选区里没有数字。','選択範囲に数値がありません。'));
  r.NumberFormat=x.f;
  var lab=root.OfficeI18n?root.OfficeI18n.t(x.label):x.label;return u==='general'?M('已把选区恢复成普通数字格式。','選択範囲を標準の数値表示に戻しました。'):M('已把选区里的 '+nums+' 个数字显示成“'+lab+'”。只改显示方式，格子里的数值没变，求和、公式照常算；想改回来选“恢复成普通数字”。','選択範囲の '+nums+' 個の数値を「'+lab+'」で表示しました。表示だけの変更で、セルの値はそのままなので合計や数式も今までどおりです。元に戻すには「標準の数値表示に戻す」を選んでください。');}
reg('wanUnit',{label:'万元显示',desc:'金额按万、亿、千、千円显示，数值不变',icon:'type',ownConfirm:['et'],keywords:'万元 万 亿 千元 单位 金额 显示 千円 百万円 単位 表示',fields:[{key:'unit',label:'显示成',options:Object.keys(UNITS).map(function(k){return {value:k,label:UNITS[k].label};}),value:'wan'}],hosts:{et:wanUnit}});

// ---- 号码校验 ----
var ID_TYPES=[{value:'auto',label:'自动判断'},{value:'uscc',label:'统一社会信用代码'},{value:'id',label:'身份证号'},{value:'invoice',label:'发票号码'},{value:'jpcorp',label:'法人番号 / インボイス登録番号'},{value:'bank',label:'银行卡号'}];
function idCheckCells(c){var a=c.a,r=cells(a),type=(c.options||{}).type;if(n(r.Columns.Count)!==1)throw Error(M('请只选一列号码。','番号の列を 1 列だけ選んでください。'));S.limit(r,20000);var out=S.rightColumn(a,r,1),v=S.grid(r),res=[],good=0,bad=0;
  v.forEach(function(row,i){var x=row[0],cell=r.Cells.Item(i+1,1),msg='';if(x==null||x===''){res.push(['']);return;}
    if(typeof x==='number'&&Math.abs(x)>=1e15)msg=M('✗ 这个号码按数字保存，超过 15 位的部分已经丢失。请把这一列设成文本后重新录入。','✗ 数値として保存されているため、15 桁を超える部分が失われています。この列を文字列に設定して入力し直してください。');
    else{var s=typeof x==='number'?(text(cell.Text)||String(x)):text(x),k=ID.check(s,type==='auto'?'':type,{lang:ja()?'ja':'zh'});if(!k){res.push(['']);return;}msg=k.msg;if(typeof x==='number'&&k.ok===false&&/^\d{1,7}$/.test(String(x))&&(type==='invoice'||!type||type==='auto'))msg+=M('（开头的 0 可能被表格去掉了）','（先頭の 0 が消えている可能性があります）');}
    if(/^✓/.test(msg))good++;else bad++;res.push([msg]);});
  if(!good&&!bad)throw Error(M('选区里没有号码。','選択範囲に番号がありません。'));S.put(out,res);try{out.Columns.Item(1).ColumnWidth=40;}catch(ignore){}
  return M('已在右边一列写出校验结果：通过 '+good+' 个，有问题 '+bad+' 个。只按编码规则检查校验位，不能证明号码真实存在；发票真伪请到税务局查验平台查。','右隣の列にチェック結果を書きました：OK '+good+' 件、問題あり '+bad+' 件。番号の規則とチェックデジットだけを確認しており、実在するかどうかは分かりません。');}
reg('idCheck',{label:'号码校验',desc:'统一社会信用代码、身份证、发票号、法人番号是否写错',icon:'check',ownConfirm:['et'],keywords:'校验 统一社会信用代码 信用代码 身份证 发票号码 银行卡 法人番号 インボイス 登録番号 T番号 チェックデジット',fields:[{key:'type',label:'号码种类',options:ID_TYPES,value:'auto'}],hosts:{et:idCheckCells}});

// ---- 合并多个工作簿 ----
function baseName(p){return text(p).replace(/^.*[\\\/]/,'').replace(/\.[^.]+$/,'');}
function mergeWorkbooks(c){var a=c.a,wb0=a.ActiveWorkbook,mode=(c.options||{}).mode==='sheets'?'sheets':'stack',dlg=a.FileDialog(3);dlg.Title=M('选择要合并的工作簿（可多选）','結合するブックを選択（複数選択可）');dlg.AllowMultiSelect=true;try{dlg.Filters.Clear();dlg.Filters.Add(M('表格','表計算'),'*.xlsx;*.xls;*.xlsm;*.et;*.csv');}catch(ignore){}
  if(dlg.Show()!==-1)return CANCEL();var files=[],self=text(wb0.FullName).toLowerCase();for(var i=1;i<=n(dlg.SelectedItems.Count);i++){var f=text(dlg.SelectedItems.Item(i));if(f.toLowerCase()!==self)files.push(f);}
  if(!files.length)throw Error(M('请选择其他工作簿（不能选当前这个）。','今開いているブック以外を選んでください。'));if(files.length>100)throw Error(M('一次最多合并 100 个文件。','一度に結合できるのは 100 ファイルまでです。'));files.sort();
  if(!L.ask(mode==='stack'?M('将把 '+files.length+' 个文件第一张工作表的数据上下接在一起，按表头对齐列，放到当前工作簿的新工作表里，并加一列“来源文件”。\n原文件只读打开，不会改动。是否继续？',files.length+' 個のファイルの先頭シートのデータを縦につなげ、見出しで列をそろえて、このブックの新しいシートに入れます。「ファイル名」の列も付けます。\n元のファイルは読み取り専用で開くので変更されません。続行しますか？'):M('将把 '+files.length+' 个文件里的每张工作表复制到当前工作簿（表名前加文件名）。\n原文件只读打开，不会改动。是否继续？',files.length+' 個のファイルのシートをすべてこのブックにコピーします（シート名の前にファイル名を付けます）。\n元のファイルは読み取り専用で開くので変更されません。続行しますか？')))return CANCEL();
  var header=[],index={},rows=[],sheets=0,failed=[],total=0,src=ja()?'ファイル名':'来源文件';try{a.ScreenUpdating=false;}catch(ignore){}
  try{files.forEach(function(f){var wb=null;try{wb=a.Workbooks.Open(f,0,true);
      if(mode==='sheets'){for(var k=1;k<=n(wb.Worksheets.Count);k++){var ws=wb.Worksheets.Item(k),last=wb0.Worksheets.Item(n(wb0.Worksheets.Count));ws.Copy(null,last);var copy=wb0.Worksheets.Item(n(wb0.Worksheets.Count));try{copy.Name=S.uniqueName(wb0,baseName(f)+'-'+text(ws.Name));}catch(ignore){}sheets++;}return;}
      var ur=wb.Worksheets.Item(1).UsedRange,v=S.grid(ur);if(!v.length)return;var head=v[0].map(function(h,j){var t=text(h).trim();return t||(M('列','列')+(j+1));}),map=head.map(function(h){if(!(h in index)){index[h]=header.length;header.push(h);}return index[h];});
      for(var r=1;r<v.length;r++){if(v[r].every(function(x){return x==null||x==='';}))continue;if(++total>100000)throw Error(M('合并后超过 10 万行，请分几次。','結合すると 10 万行を超えます。何回かに分けてください。'));var line=[];map.forEach(function(col,j){line[col]=v[r][j];});rows.push({f:baseName(f),line:line});}
    }catch(e){if(/10 万/.test(e.message))throw e;failed.push(baseName(f)+'（'+e.message+'）');}finally{if(wb)try{wb.Close(false);}catch(ignore){}}});
  }finally{try{a.ScreenUpdating=true;}catch(ignore){}try{wb0.Activate();}catch(ignore){}}
  var fail=failed.length?M('\n没打开的文件：','\n開けなかったファイル：')+failed.slice(0,5).join('、')+(failed.length>5?'…':''):'';
  if(mode==='sheets')return M('已复制 '+sheets+' 张工作表到当前工作簿。原文件没有改动。','シート '+sheets+' 枚をこのブックにコピーしました。元のファイルは変更していません。')+fail;
  if(!rows.length)throw Error(M('选的文件里没有读到数据。','選んだファイルからデータを読み取れませんでした。')+fail);
  var out=[[src].concat(header)].concat(rows.map(function(x){var line=[x.f];for(var j=0;j<header.length;j++)line.push(x.line[j]==null?'':x.line[j]);return line;}));
  var ws2=S.newSheet(a,ja()?'統合結果':'合并结果');S.writeAt(ws2,out);try{ws2.Rows.Item(1).Font.Bold=true;ws2.Columns.AutoFit();}catch(ignore){}
  return M('已把 '+(files.length-failed.length)+' 个文件的 '+rows.length+' 行数据合并到新工作表“'+text(ws2.Name)+'”，共 '+header.length+' 列（表头不同的列各自单独一列）。原文件没有改动。',(files.length-failed.length)+' 個のファイルの '+rows.length+' 行を新しいシート「'+text(ws2.Name)+'」にまとめました。列は '+header.length+' 列です（見出しが違う列はそれぞれ別の列）。元のファイルは変更していません。')+fail;}
reg('mergeWorkbooks',{label:'合并多个工作簿',desc:'选几个 Excel 文件，数据上下接在一起，或把工作表都复制过来',icon:'insert',keywords:'合并 工作簿 多个文件 汇总 文件夹 多个表格 ブック 結合 統合 複数ファイル まとめる',ownConfirm:['et'],fields:[{key:'mode',label:'怎么合并',options:[{value:'stack',label:'数据上下接在一起（按表头对齐）'},{value:'sheets',label:'每张工作表原样复制过来'}],value:'stack'}],hosts:{et:mergeWorkbooks}});

// ---- 批量二维码 / 条形码 ----
async function codesBatch(c){var a=c.a,r=cells(a),kind=(c.options||{}).kind==='barcode'?'barcode':'qr',ws=a.ActiveSheet;if(n(r.Columns.Count)!==1)throw Error(M('请只选一列内容（网址、编号等）。','内容（URL や番号など）の列を 1 列だけ選んでください。'));if(n(r.Rows.Count)>500)throw Error(M('一次最多 500 个。','一度に作れるのは 500 個までです。'));
  var out=S.rightColumn(a,r,1),items=S.grid(r).map(function(x,i){var v=x[0];return typeof v==='number'?text(r.Cells.Item(i+1,1).Text)||String(v):text(v);});var count=items.filter(function(x){return x.trim();}).length;if(!count)throw Error(M('选区是空的。','選択範囲が空です。'));
  if(!L.ask(M('将为 '+count+' 个格子生成'+(kind==='qr'?'二维码':'条形码')+'，插到右边一列，并把这些行调高、这一列调宽。\n图片同时保存在“生成文件”文件夹里。是否继续？',count+' 個のセルの'+(kind==='qr'?'QR コード':'バーコード')+'を作って右隣の列に貼り付けます。行の高さと列の幅も広げます。\n画像は作成ファイルのフォルダーにも保存します。続行しますか？')))return CANCEL();
  var res=await tool('codes',{items:items,kind:kind}),done=0;
  try{out.ColumnWidth=kind==='qr'?12:26;}catch(ignore){}
  res.files.forEach(function(file,i){if(!file)return;var cell=out.Cells.Item(i+1,1);try{cell.RowHeight=kind==='qr'?70:46;}catch(ignore){}
    var L0=n(cell.Left)+3,T0=n(cell.Top)+3,W=n(cell.Width)-6,H=n(cell.Height)-6;if(W<10||H<10)return;var p=ws.Shapes.AddPicture(file,0,-1,L0,T0,-1,-1);try{p.LockAspectRatio=-1;}catch(ignore){}
    var k=Math.min(W/n(p.Width),H/n(p.Height));p.Width=n(p.Width)*k;p.Left=L0+(W-n(p.Width))/2;p.Top=T0+(H-n(p.Height))/2;try{p.Placement=1;}catch(ignore){}done++;});
  var err=res.errors&&res.errors.length?M('\n没生成的：','\n作れなかったもの：')+res.errors.slice(0,5).map(function(e){return M('第 '+(e.index+1)+' 行','行 '+(e.index+1))+'（'+(root.OfficeI18n?root.OfficeI18n.t(e.error):e.error)+'）';}).join(M('；','、')):'';
  return M('已插入 '+done+' 个'+(kind==='qr'?'二维码':'条形码')+'，图片也存在：',(kind==='qr'?'QR コード':'バーコード')+'を '+done+' 個貼り付けました。画像の保存先：')+res.folder+err;}
reg('codesBatch',{label:'批量二维码 / 条形码',desc:'选中一列网址或编号，每格生成一个二维码或条形码，插在右边',icon:'picture',ownConfirm:['et'],keywords:'二维码 条形码 条码 批量 生成 QR コード バーコード 一括',fields:[{key:'kind',label:'生成什么',options:[{value:'qr',label:'二维码（网址、中文都可以）'},{value:'barcode',label:'条形码（Code 128，只能英文和数字）'}],value:'qr'}],hosts:{et:codesBatch}});

// ---- 考勤统计 ----
function normMark(v){var t=text(v).trim();if(/^[√✓✔☑]$/.test(t))return '√';if(/^[○〇◯]$/.test(t))return '○';if(/^[×✗✘xX]$/.test(t))return '×';return t;}
function attendanceStats(c){var a=c.a,r=cells(a);S.limit(r,100000);var v=S.grid(r);if(v.length<2||v[0].length<3)throw Error(M('请连同表头选中整张考勤表：第一行是日期，左边一列是姓名，中间是 √、休、假、迟 等记号。','見出しを含めて勤怠表全体を選んでください：1 行目が日付、左の列が氏名、間に √・休・有・遅 などの記号がある表です。'));
  var cols=v[0].length,nameCol=-1;for(var j=0;j<cols&&nameCol<0;j++){var texts=0;for(var i=1;i<v.length;i++){var x=v[i][j];if(typeof x==='string'&&x.trim().length>=2&&!/^\d+$/.test(x.trim()))texts++;}if(texts>=Math.max(1,(v.length-1)/2))nameCol=j;}
  if(nameCol<0)throw Error(M('找不到姓名列。请让姓名在记号的左边，并且连同表头一起选中。','氏名の列が見つかりません。氏名を記号の左に置き、見出しも含めて選択してください。'));
  var dayCols=[];for(var j2=nameCol+1;j2<cols;j2++){var seen={},ok=true,any=false;for(var i2=1;i2<v.length;i2++){var y=v[i2][j2];if(y==null||y==='')continue;any=true;var t=normMark(y);if(t.length>4){ok=false;break;}seen[t]=1;}if(ok&&any&&Object.keys(seen).length<=12)dayCols.push(j2);}
  if(!dayCols.length)throw Error(M('姓名右边没有找到记考勤的列（每格应是 √、休、假、迟 这样的短记号）。','氏名の右に勤怠の列が見つかりません（各セルは √・休・有・遅 のような短い記号にしてください）。'));
  var people=[],byName={},totals={};
  for(var i3=1;i3<v.length;i3++){var name=text(v[i3][nameCol]).trim();if(!name)continue;var p=byName[name];if(!p){p=byName[name]={name:name,c:{}};people.push(p);}dayCols.forEach(function(j){var m=normMark(v[i3][j]);if(!m)return;p.c[m]=(p.c[m]||0)+1;totals[m]=(totals[m]||0)+1;});}
  var marks=Object.keys(totals).sort(function(x,y){return totals[y]-totals[x];});if(!marks.length)throw Error(M('考勤列里都是空的。','勤怠の列がすべて空です。'));
  var J=ja(),out=[[J?'氏名':'姓名'].concat(marks).concat([J?'記録日数':'记录天数'])];
  people.forEach(function(p){var sum=0,line=[p.name];marks.forEach(function(m){var k=p.c[m]||0;sum+=k;line.push(k);});line.push(sum);out.push(line);});
  var tot=[J?'合計':'合计'];marks.forEach(function(m){tot.push(totals[m]);});tot.push(marks.reduce(function(s,m){return s+totals[m];},0));out.push(tot);
  var ws=S.newSheet(a,J?'勤怠集計':'考勤统计');S.writeAt(ws,out);try{ws.Rows.Item(1).Font.Bold=true;ws.Rows.Item(out.length).Font.Bold=true;ws.Columns.AutoFit();}catch(ignore){}
  var top=marks.slice(0,8).map(function(m){return M('“'+m+'” '+totals[m]+' 次','「'+m+'」'+totals[m]+' 回');}).join('、')+(marks.length>8?'…':'');
  return M('已在新工作表“'+text(ws.Name)+'”里统计 '+people.length+' 人、'+dayCols.length+' 天：'+top+'。原表不变。','新しいシート「'+text(ws.Name)+'」に '+people.length+' 人・'+dayCols.length+' 日分を集計しました：'+top+'。元の表は変更していません。');}
reg('attendanceStats',{label:'考勤统计',desc:'按 √、休、假、迟等记号，统计每人各出现几次',icon:'table',ownConfirm:['et'],keywords:'考勤 出勤 排班 打卡 请假 迟到 统计 天数 勤怠 出勤簿 シフト 集計 有給',hosts:{et:attendanceStats}});
})(window);

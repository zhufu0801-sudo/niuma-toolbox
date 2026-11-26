// 办公工具箱 · 宏代码检查（2026-11-25）：不运行、不用 AI，只看 VBA 代码的写法。
// 查三类：写错了（块没写完、引号括号不配对、跳转标签不存在…）、容易出 bug 的写法、有安全风险的写法（宏病毒常见）。
// 能确定怎么改的给出“修改后的代码”：Integer → Long、加 Option Explicit、Declare 加 PtrSafe。
// 侧栏（vba-panel.js）和测试共用；返回的说明有中文和日文两份。
(function(root){'use strict';
var SEV={error:0,risk:1,warn:2,tip:3};
// ---- 预处理：去注释、记住字符串位置、合并续行（" _"） ----
function stripLine(s){var out='',inStr=false;for(var i=0;i<s.length;i++){var c=s[i];
  if(inStr){if(c==='"'){if(s[i+1]==='"'){out+='xx';i++;continue;}inStr=false;out+='"';continue;}out+='x';continue;}
  if(c==='"'){inStr=true;out+='"';continue;}if(c==="'")break;if((c==='R'||c==='r')&&/^rem(\s|$)/i.test(s.slice(i))&&/(^|[\s:])$/.test(out))break;out+=c;}
 return {code:out,keep:s.slice(0,out.length),openStr:inStr};}
// 返回逻辑行：{text（去注释，字符串内容换成 x）, raw（原文）, line（起始行号，从 1 开始）, openStr}
function logical(src){var lines=String(src||'').replace(/\r\n?/g,'\n').split('\n'),out=[],cur=null;
 lines.forEach(function(raw,i){var st=stripLine(raw),code=st.code;var cont=/\s_\s*$/.test(code)&&!st.openStr;var piece=cont?code.replace(/\s_\s*$/,' '):code;
  var keep=cont?st.keep.replace(/\s_\s*$/,' '):st.keep;
  if(cur){cur.text+=piece;cur.keep+=keep;cur.raw+='\n'+raw;cur.openStr=cur.openStr||st.openStr;}else cur={text:piece,keep:keep,raw:raw,line:i+1,openStr:st.openStr,count:0};cur.count++;
  if(!cont){out.push(cur);cur=null;}});
 if(cur)out.push(cur);return out;}
// 一行里用冒号分开的几句（字符串已经替换掉，不会误切）；行首的标签（Label: / 100）单独拿出来
function statements(t){var parts=[],depth=0,buf='';for(var i=0;i<t.length;i++){var c=t[i];if(c==='(')depth++;else if(c===')')depth--;if(c===':'&&depth===0&&!/:=$/.test(buf+c)&&t[i+1]!=='='){parts.push(buf);buf='';continue;}buf+=c;}parts.push(buf);return parts.map(function(s){return s.trim();}).filter(Boolean);}
function words(s){return s.toLowerCase().replace(/\s+/g,' ').trim();}
function F(sev,code,line,zh,ja,fixZh,fixJa,extra){var f={sev:sev,code:code,line:line,msg:zh,msgJa:ja,fix:fixZh||'',fixJa:fixJa||''};if(extra)for(var k in extra)f[k]=extra[k];return f;}

// ---- 块结构 ----
var OPEN=[
 {re:/^(public\s+|private\s+|friend\s+)?(static\s+)?sub\s+([A-Za-z_\u0080-\uffff][\w\u0080-\uffff]*)/i,kind:'sub'},
 {re:/^(public\s+|private\s+|friend\s+)?(static\s+)?function\s+([A-Za-z_\u0080-\uffff][\w\u0080-\uffff]*)/i,kind:'function'},
 {re:/^(public\s+|private\s+|friend\s+)?(static\s+)?property\s+(get|let|set)\s+([A-Za-z_\u0080-\uffff][\w\u0080-\uffff]*)/i,kind:'property'},
 {re:/^(public\s+|private\s+)?type\s+[A-Za-z_\u0080-\uffff][\w\u0080-\uffff]*/i,kind:'type'},{re:/^(public\s+|private\s+)?enum\s+[A-Za-z_\u0080-\uffff][\w\u0080-\uffff]*/i,kind:'enum'},
 {re:/^with\b/i,kind:'with'},{re:/^select\s+case\b/i,kind:'select'},{re:/^for\s+each\s+([A-Za-z_\u0080-\uffff][\w\u0080-\uffff]*)/i,kind:'for',each:true},{re:/^for\s+([A-Za-z_\u0080-\uffff][\w\u0080-\uffff]*)\s*=/i,kind:'for'},
 {re:/^do\b/i,kind:'do'},{re:/^while\b/i,kind:'while'}];
var CLOSE={'end sub':'sub','end function':'function','end property':'property','end type':'type','end enum':'enum','end with':'with','end select':'select','loop':'do','wend':'while','end if':'if'};
var NAME={sub:['Sub','Sub'],function:['Function','Function'],property:['Property','Property'],type:['Type','Type'],enum:['Enum','Enum'],with:['With','With'],select:['Select Case','Select Case'],for:['For','For'],do:['Do','Do'],while:['While','While'],if:['If','If']};
var ENDER={sub:'End Sub',function:'End Function',property:'End Property',type:'End Type',enum:'End Enum',with:'End With',select:'End Select',for:'Next',do:'Loop',while:'Wend',if:'End If'};
function isBlockIf(s){// If ... Then 后面没有语句才是块 If
 var m=/^(else)?if\b(.*)\bthen\b(.*)$/i.exec(s);if(!m)return null;return !m[3].trim();}

function lint(src,opts){opts=opts||{};var L=logical(src),out=[],stack=[],procs={},proc=null,hasExplicit=false,firstCodeLine=0,declares=[];
 function push(f){out.push(f);}
 function closeProc(){if(!proc)return;var p=proc;
  // 恢复设置
  [['screenupdating','Application.ScreenUpdating','= True'],['displayalerts','Application.DisplayAlerts','= True'],['enableevents','Application.EnableEvents','= True']].forEach(function(x){if(p.off[x[0]]&&!p.on[x[0]])push(F('warn','RESTORE',p.off[x[0]],x[1]+' 关掉后，这个过程里没有再打开。中途出错或结束后会一直关着。',x[1]+' を無効にした後、このプロシージャ内で元に戻していません。エラーや終了後も無効のままになります。','在结尾（和出错处理里）加上 '+x[1]+' '+x[2]+'。','最後（とエラー処理）に '+x[1]+' '+x[2]+' を追加してください。'));});
  if(p.calcManual&&!p.calcRestore)push(F('warn','RESTORE',p.calcManual,'计算方式改成“手动”后没有改回来，之后表格不会自动重算。','計算方法を「手動」にした後、元に戻していません。以後シートが自動で再計算されません。','结尾加上 Application.Calculation = xlCalculationAutomatic。','最後に Application.Calculation = xlCalculationAutomatic を追加してください。'));
  if(p.resumeNext&&!p.errReset)push(F('warn','ERR_RESUME',p.resumeNext,'On Error Resume Next 之后一直没有恢复，后面的错误都会被悄悄跳过，问题很难发现。','On Error Resume Next の後、元に戻していません。以降のエラーがすべて黙って無視され、問題に気付けません。','需要忽略错误的语句后面加 On Error GoTo 0。','エラーを無視したい文の直後に On Error GoTo 0 を入れてください。'));
  Object.keys(p.gotos).forEach(function(lab){if(!p.labels[lab.toLowerCase()])push(F('error','LABEL',p.gotos[lab],'跳转到的标签“'+lab+'”在这个过程里不存在，运行时会报错。','ジャンプ先のラベル「'+lab+'」がこのプロシージャ内にありません。実行時にエラーになります。','检查标签名是否写错，标签要写成“'+lab+':”放在行首。','ラベル名の誤りを確認してください。ラベルは行頭に「'+lab+':」と書きます。'));});
  if(p.selects>=5)push(F('tip','SELECT',p.line,'这个过程里用了 '+p.selects+' 次 Select / Activate，会慢，也容易因为选错位置出错。','このプロシージャで Select / Activate を '+p.selects+' 回使っています。遅くなり、選択位置の違いで誤動作しやすくなります。','直接写 Worksheets("表名").Range("A1").Value = …，不用先选中。','Worksheets("シート名").Range("A1").Value = … のように、選択せず直接指定してください。'));
  proc=null;}
 L.forEach(function(lt){var t=lt.text,line=lt.line;
  if(lt.openStr)push(F('error','QUOTE',line,'这一行的双引号没有配对。','この行のダブルクォーテーションが対になっていません。','检查字符串是否少了一个 "（字符串里的 " 要写成 ""）。','文字列の " が足りないか確認してください（文字列内の " は "" と書きます）。'));
  var depth=0,minDepth=0;for(var i=0;i<t.length;i++){if(t[i]==='(')depth++;else if(t[i]===')'){depth--;if(depth<minDepth)minDepth=depth;}}
  if(depth!==0||minDepth<0)push(F('error','PAREN',line,'这一行的括号没有配对。','この行の括弧が対になっていません。','数一下左右括号的个数。','開き括弧と閉じ括弧の数を確認してください。'));
  var stm=statements(t),rawStm=statements(lt.keep||'');
  function handle(s,si,keepS,inline){
   var lab=/^([A-Za-z_\u0080-￿][\w\u0080-￿]*|\d+)\s*$/.exec(s);
   if(si===0&&lab&&proc&&/^\s*[\w\u0080-￿]+\s*:/.test(t)&&!/^(else|end|loop|next|wend|exit|stop|beep|resume|return|doevents)$/i.test(lab[1])){proc.labels[lab[1].toLowerCase()]=line;return;}
   if(si===0&&/^\d+\s/.test(s)&&proc){var num=s.match(/^\d+/)[0];proc.labels[num]=line;s=s.slice(num.length).trim();if(!s)return;}
   var w=words(s);if(!w)return;
   if(!firstCodeLine&&!/^option\b/.test(w)&&!/^attribute\b/.test(w))firstCodeLine=line;
   if(/^option\s+explicit$/.test(w)){hasExplicit=true;return;}
   if(/^attribute\b/.test(w))return;
   // Declare（64 位要 PtrSafe）
   if(/^(public\s+|private\s+)?declare\s+/.test(w)){declares.push(line);if(!/\bptrsafe\b/.test(w))push(F('error','PTRSAFE',line,'Declare 没有写 PtrSafe，在 64 位的 Office / WPS 里会编译不过。','Declare に PtrSafe がありません。64 ビット版の Office / WPS ではコンパイルエラーになります。','改成 Declare PtrSafe …；涉及指针、句柄的参数改成 LongPtr。','Declare PtrSafe … に変更し、ポインターやハンドルの引数は LongPtr にしてください。',{autofix:'PTRSAFE'}));
    if(/\b(virtualalloc|virtualprotect|createthread|rtlmovememory|writeprocessmemory|loadlibrary|getprocaddress|shellexecute|winexec|urldownloadtofile)\b/.test(w))push(F('risk','API',line,'调用了 Windows 底层接口（'+(w.match(/\b(virtualalloc|virtualprotect|createthread|rtlmovememory|writeprocessmemory|loadlibrary|getprocaddress|shellexecute|winexec|urldownloadtofile)\b/)[0])+'），宏病毒常用这种写法。','Windows の低レベル API（'+(w.match(/\b(virtualalloc|virtualprotect|createthread|rtlmovememory|writeprocessmemory|loadlibrary|getprocaddress|shellexecute|winexec|urldownloadtofile)\b/)[0])+'）を呼び出しています。マクロウイルスによく見られる書き方です。','来历不明的文件请不要启用宏。','出所不明のファイルではマクロを有効にしないでください。'));return;}
   // 过程开始 / 结束
   for(var k=0;k<OPEN.length;k++){var m=OPEN[k].re.exec(s);if(!m)continue;var kind=OPEN[k].kind;
    if(kind==='sub'||kind==='function'||kind==='property'){if(proc)push(F('error','BLOCK',line,'上一个过程还没有 '+ENDER[proc.kind]+' 就开始了新的过程。','前のプロシージャに '+ENDER[proc.kind]+' がないまま、次のプロシージャが始まっています。','在新过程前补上 '+ENDER[proc.kind]+'。','次のプロシージャの前に '+ENDER[proc.kind]+' を追加してください。'));
     while(stack.length){var b=stack.pop();if(b.kind!=='sub'&&b.kind!=='function'&&b.kind!=='property')push(F('error','BLOCK',b.line,NAME[b.kind][0]+' 没有对应的 '+ENDER[b.kind]+'。',NAME[b.kind][1]+' に対応する '+ENDER[b.kind]+' がありません。','在过程结束前补上 '+ENDER[b.kind]+'。','プロシージャの終わりまでに '+ENDER[b.kind]+' を追加してください。'));}
     closeProc();var name=m[m.length-1];var key=(kind==='property'?'p:'+m[4]+':'+name:name).toLowerCase();if(procs[key])push(F('error','DUP',line,'过程名“'+name+'”重复了（第 '+procs[key]+' 行也有）。','プロシージャ名「'+name+'」が重複しています（'+procs[key]+' 行目にもあります）。','改成不同的名字。','別の名前にしてください。'));procs[key]=line;
     proc={kind:kind,name:name,line:line,labels:{},gotos:{},off:{},on:{},selects:0,resumeNext:0,errReset:0,calcManual:0,calcRestore:0};
     if(/^(auto_open|autoopen|autoexec|autoclose|auto_close|document_open|document_new|document_close|workbook_open|workbook_activate|workbook_beforeclose|presentation_open)$/i.test(name))push(F('risk','AUTORUN',line,'“'+name+'” 会在打开（或关闭）文件时自动运行。自己写的没问题；来历不明的文件要小心。','「'+name+'」はファイルを開く（閉じる）ときに自動実行されます。自作なら問題ありませんが、出所不明のファイルは注意してください。','确认这段代码是你自己或可信的人写的。','自分または信頼できる人が書いたコードか確認してください。'));
     stack.push({kind:kind,line:line});break;}
    if(kind==='for'){stack.push({kind:'for',line:line,v:m[1].toLowerCase(),each:!!OPEN[k].each,asc:!OPEN[k].each&&!/\bstep\s+-/.test(w),deletes:0});break;}
    if(kind==='do'){stack.push({kind:'do',line:line,cond:/^do\s+(while|until)\b/.test(w),exit:false});break;}
    stack.push({kind:kind,line:line});break;}
   var bi=isBlockIf(s);
   if(bi===false&&!inline){var mm=/^((?:else)?if\b.*?\bthen\b)\s*(.+)$/i.exec(s),mk=/^((?:else)?if\b.*?\bthen\b)\s*(.+)$/i.exec(keepS||'');if(mm){var tails=mm[2].split(/\belse\b/i),ktails=mk?mk[2].split(/\belse\b/i):tails;tails.forEach(function(p,pi){if(p.trim())handle(p.trim(),1,(ktails[pi]||p).trim(),true);});}}
   if(/^if\b/i.test(s)&&bi===true)stack.push({kind:'if',line:line});
   if(/^(else|elseif\b)/i.test(s)&&!stack.some(function(b){return b.kind==='if';})&&bi!==false)push(F('error','BLOCK',line,'Else / ElseIf 前面没有对应的 If。','Else / ElseIf に対応する If がありません。','检查 If 是否写成了单行（Then 后面直接写了语句）。','If が 1 行形式（Then の後に文がある）になっていないか確認してください。'));
   // 结束语句
   var close=null;if(/^next\b/.test(w))close='for';else for(var ck in CLOSE)if(w===ck||w.indexOf(ck+' ')===0){close=CLOSE[ck];break;}
   if(/^loop\b/.test(w))close='do';
   if(close){var top=stack[stack.length-1];
    if(close==='for'){var vars=w.replace(/^next\s*/,'').split(',').map(function(x){return x.trim();}).filter(Boolean);var n=Math.max(1,vars.length);
     for(var j=0;j<n;j++){top=stack[stack.length-1];if(!top||top.kind!=='for'){push(F('error','BLOCK',line,'Next 前面没有对应的 For'+(top?'（还有 '+NAME[top.kind][0]+' 没结束）':'')+'。','Next に対応する For がありません'+(top?'（'+NAME[top.kind][1]+' がまだ閉じていません）':'')+'。','检查 For 和 Next 是否成对，嵌套顺序是否对。','For と Next が対になっているか、入れ子の順序を確認してください。'));break;}
      if(vars[j]&&vars[j]!==top.v)push(F('error','NEXTVAR',line,'Next '+vars[j]+' 和 For '+top.v+'（第 '+top.line+' 行）的变量对不上。','Next '+vars[j]+' と For '+top.v+'（'+top.line+' 行目）の変数が一致しません。','改成 Next '+top.v+'。','Next '+top.v+' にしてください。'));
      if(top.asc&&top.deletes)push(F('warn','DELETE_ASC',top.line,'在正着数的 For 循环里删除行（第 '+top.deletes+' 行），删掉一行后下一行会往上补，会被跳过。','昇順の For ループの中で行を削除しています（'+top.deletes+' 行目）。削除すると次の行が詰まり、1 行飛ばされます。','改成倒着循环：For i = 最后一行 To 1 Step -1。','逆順にしてください：For i = 最終行 To 1 Step -1。'));
      stack.pop();}return;}
    if(!top||top.kind!==close){var known=stack.some(function(b){return b.kind===close;});push(F('error','BLOCK',line,(known?ENDER[close]+' 之前还有 '+NAME[top.kind][0]+'（第 '+top.line+' 行）没结束。':ENDER[close]+' 前面没有对应的 '+NAME[close][0]+'。'),(known?ENDER[close]+' の前に、'+NAME[top.kind][1]+'（'+top.line+' 行目）が閉じていません。':ENDER[close]+' に対応する '+NAME[close][1]+' がありません。'),(known?'在这之前补上 '+ENDER[top.kind]+'。':'检查是否多写了，或者开头那句写错了。'),(known?'その前に '+ENDER[top.kind]+' を追加してください。':'余分に書いていないか、開始側の文が正しいか確認してください。')));
     if(known){while(stack.length&&stack[stack.length-1].kind!==close)stack.pop();top=stack[stack.length-1];}else return;}
    if(close==='do'&&!top.cond&&!/^loop\s+(while|until)\b/.test(w)&&!top.exit)push(F('warn','INFINITE',top.line,'Do … Loop 没有结束条件，里面也没有 Exit Do，会一直循环下去。','Do … Loop に終了条件がなく、Exit Do もありません。無限ループになります。','加上 Do While 条件，或者在合适的地方 Exit Do。','Do While 条件を付けるか、適切な場所で Exit Do してください。'));
    stack.pop();if(close==='sub'||close==='function'||close==='property'){closeProc();}return;}
   if(/^exit\s+do$/.test(w))for(var q=stack.length-1;q>=0;q--)if(stack[q].kind==='do'){stack[q].exit=true;break;}
   if(/^exit\s+(sub|function|property)$/.test(w)&&proc){var ek=w.split(' ')[1];if(ek!==proc.kind)push(F('error','EXIT',line,'在 '+NAME[proc.kind][0]+' 里写了 Exit '+NAME[ek][0]+'。',NAME[proc.kind][1]+' の中で Exit '+NAME[ek][1]+' を使っています。','改成 Exit '+NAME[proc.kind][0]+'。','Exit '+NAME[proc.kind][1]+' にしてください。'));}
   if(w==='end')push(F('warn','END',line,'单独的 End 会立刻结束所有宏，并清空所有变量，打开的文件也不会关。','単独の End はすべてのマクロを即座に終了し、変数もすべて消えます。開いたファイルも閉じられません。','改成 Exit Sub / Exit Function。','Exit Sub / Exit Function にしてください。'));
   // 过程内的检查
   if(proc){var gm=/\b(?:goto|gosub)\s+([\w\u0080-￿]+)/i.exec(s),rm=/^resume\s+([\w\u0080-￿]+)$/i.exec(s);
    [gm,rm].forEach(function(x){if(x&&!/^(0|-1|next)$/i.test(x[1])&&!/^on\s+error\s+goto\s+(0|-1)$/i.test(w))proc.gotos[x[1]]=proc.gotos[x[1]]||line;});
    if(/^on\s+error\s+resume\s+next$/.test(w)&&!proc.resumeNext)proc.resumeNext=line;
    if(/^on\s+error\s+goto\b/.test(w))proc.errReset=line;
    [['screenupdating'],['displayalerts'],['enableevents']].forEach(function(x){var re=new RegExp('application\\.'+x[0]+'\\s*=\\s*(true|false)');var mm=re.exec(w);if(mm){if(mm[1]==='false'&&!proc.off[x[0]])proc.off[x[0]]=line;else if(mm[1]==='true')proc.on[x[0]]=line;}});
    if(/application\.calculation\s*=\s*(xlcalculationmanual|-4135)/.test(w))proc.calcManual=proc.calcManual||line;if(/application\.calculation\s*=\s*(xlcalculationautomatic|-4105)/.test(w))proc.calcRestore=line;
    proc.selects+=(w.match(/\.(select|activate)\b/g)||[]).length;
    if(/(\.|\b)(entirerow\.)?delete\b/.test(w)&&/(rows\s*\(|entirerow|\.delete\s*(shift|$))/.test(w)){for(var z=stack.length-1;z>=0;z--)if(stack[z].kind==='for'){if(stack[z].asc&&new RegExp('\\b'+stack[z].v+'\\b').test(w))stack[z].deletes=stack[z].deletes||line;break;}}}
   // 不分过程的检查
   if(/(=|<>)\s*null\b/.test(w)&&!/\bis\s+nothing\b/.test(w))push(F('error','NULL',line,'和 Null 比较（= Null / <> Null）的结果永远是 Null，条件不会成立。','Null との比較（= Null / <> Null）は常に Null になり、条件が成立しません。','改成 IsNull(变量)。','IsNull(変数) にしてください。'));
   if(/^(let\s+)?[\w.()]+\s*=\s*(new\s+\w|createobject\s*\(|getobject\s*\()/.test(w)&&!/^set\s/.test(w))push(F('error','SET',line,'给对象变量赋值要用 Set，现在这样运行时会报错。','オブジェクト変数への代入には Set が必要です。このままでは実行時エラーになります。','在这一句前面加上 Set。','文の先頭に Set を付けてください。'));
   if(/\b(65536|1048576)\b/.test(w))push(F('tip','HARDROW',line,'写死了最大行号（65536 / 1048576），换个版本的表格就不对，也会很慢。','最大行番号（65536 / 1048576）を直接書いています。形式が違うブックでは正しくなく、処理も遅くなります。','改成 Cells(Rows.Count, 1).End(xlUp).Row 找到最后一行。','Cells(Rows.Count, 1).End(xlUp).Row で最終行を求めてください。'));
   if(/\bas\s+integer\b/.test(w))push(F('tip','INTEGER',line,'Integer 最大只到 32767，用来存行号或大数会溢出。','Integer は 32767 までです。行番号や大きな数を入れるとオーバーフローします。','改成 Long。','Long にしてください。',{autofix:'INTEGER'}));
   // 安全风险
   var R=[[/\bshell\s*\(|\bshell\s+"|wscript\.shell|\bcmd(\.exe)?\b|powershell|\.run\s*\(?\s*"/,'SHELL','运行外部程序或命令','外部プログラム・コマンドの実行'],[/\bkill\s|\.deletefile\b|\.deletefolder\b|\brmdir\s/,'DELETE','删除电脑上的文件','パソコン上のファイル削除'],[/urldownloadtofile|msxml2\.(server)?xmlhttp|winhttp|xmlhttp|internetexplorer\.application|\.responsebody\b|adodb\.stream/,'NET','从网上下载或往外发送数据','インターネットからのダウンロード・外部への送信'],[/\.regwrite\b|\.regdelete\b/,'REG','修改注册表','レジストリの変更'],[/vbproject|\.vbcomponents|codemodule|\.addfromstring|\.insertlines/,'SELFMOD','修改宏代码本身（宏病毒自我复制的常见写法）','マクロのコード自体の書き換え（マクロウイルスの自己複製によくある書き方）'],[/\bsendkeys\b/,'SENDKEYS','模拟键盘输入（SendKeys，容易把按键发到别的窗口）','キー入力の送信（SendKeys。別のウィンドウに送られる恐れ）'],[/executeexcel4macro|\bcallbyname\b/,'DYNAMIC','用字符串动态执行代码','文字列からの動的なコード実行']];
   var wk=words(keepS||s);R.forEach(function(r){if(r[0].test(wk))push(F(r[1]==='SENDKEYS'||r[1]==='DYNAMIC'?'warn':'risk',r[1],line,'这里在'+r[2]+'。自己写的没问题；来历不明的文件要小心。','ここで'+r[3]+'を行っています。自作なら問題ありませんが、出所不明のファイルは注意してください。','确认代码来源可信；不需要的话删掉这一句。','コードの出所が信頼できるか確認し、不要ならこの文を削除してください。'));});
   if((s.match(/\bchrw?\s*\(/gi)||[]).length>=8||/\bstrreverse\s*\(/i.test(s))push(F('risk','OBFUSCATE',line,'用大量 Chr() 拼字符串或把字符串倒过来，常用来隐藏真实内容（宏病毒常见）。','Chr() の大量連結や文字列の反転で、本当の内容を隠しています（マクロウイルスによくある手口）。','来历不明的文件请不要启用宏。','出所不明のファイルではマクロを有効にしないでください。'));
  }
  stm.forEach(function(s,si){handle(s,si,rawStm[si]||s,false);});
 });
 while(stack.length){var b=stack.pop();push(F('error','BLOCK',b.line,NAME[b.kind][0]+' 没有对应的 '+ENDER[b.kind]+'。',NAME[b.kind][1]+' に対応する '+ENDER[b.kind]+' がありません。','补上 '+ENDER[b.kind]+'。',ENDER[b.kind]+' を追加してください。'));if(b.kind==='sub'||b.kind==='function'||b.kind==='property')closeProc();}
 closeProc();
 if(!hasExplicit&&L.some(function(l){return /\b(sub|function)\b/i.test(l.text);}))push(F('tip','EXPLICIT',1,'没有写 Option Explicit。变量名打错时不会提示，会悄悄当成新变量。','Option Explicit がありません。変数名を打ち間違えても警告されず、新しい変数として扱われます。','在模块第一行加上 Option Explicit，并给所有变量写 Dim。','モジュールの先頭に Option Explicit を入れ、すべての変数を Dim で宣言してください。',{autofix:'EXPLICIT'}));
 // 同一行同一种问题只报一次
 var seen={};out=out.filter(function(f){var k=f.code+'@'+f.line;if(seen[k])return false;seen[k]=1;return true;});
 out.sort(function(a,b){return SEV[a.sev]-SEV[b.sev]||a.line-b.line;});
 var count={error:0,risk:0,warn:0,tip:0};out.forEach(function(f){count[f.sev]++;});
 return {findings:out,count:count,lines:String(src||'').split(/\r\n?|\n/).length};}

// 自动修改：只改能确定的。返回 {code, changes:[说明]}
function autofix(src,which){which=which||['INTEGER','EXPLICIT','PTRSAFE'];var lines=String(src||'').replace(/\r\n?/g,'\n').split('\n'),changes=[];
 function codePart(s){var st=stripLine(s);return st.code.length;}
 lines=lines.map(function(s,i){var n=codePart(s),code=s.slice(0,n),rest=s.slice(n),orig=code,masked=stripLine(s).code;
  // 只改字符串外面的：在“字符串已换成 x”的那份里找位置，再改原文
  if(which.indexOf('INTEGER')>=0){var re=/\b([Aa][Ss]\s+)Integer\b/g,m,pos=[];while((m=re.exec(masked)))pos.push(m.index+m[1].length);for(var p=pos.length-1;p>=0;p--)code=code.slice(0,pos[p])+'Long'+code.slice(pos[p]+7);}
  if(which.indexOf('PTRSAFE')>=0&&/^\s*((public|private)\s+)?declare\s+(?!ptrsafe)/i.test(code))code=code.replace(/\b(declare)\s+/i,'$1 PtrSafe ');
  if(code!==orig)changes.push({line:i+1,from:s,to:code+rest});return code+rest;});
 var out=lines.join('\r\n');
 if(which.indexOf('EXPLICIT')>=0&&!/^\s*option\s+explicit\b/im.test(src)){var at=0;while(at<lines.length&&/^\s*(attribute\b|option\b)/i.test(lines[at]))at++;lines.splice(at,0,'Option Explicit');out=lines.join('\r\n');changes.unshift({line:at+1,from:'',to:'Option Explicit'});}
 return {code:out,changes:changes};}
var LABEL={error:['错误','エラー'],risk:['安全风险','セキュリティ'],warn:['容易出错','注意'],tip:['建议','ヒント']};
var api={lint:lint,autofix:autofix,logical:logical,LABEL:LABEL};
root.VbaLint=api;if(typeof module!=='undefined'&&module.exports)module.exports=api;
})(typeof window!=='undefined'?window:globalThis);

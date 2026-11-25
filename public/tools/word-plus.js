// 办公工具箱 · 方法卡的一键版（WPS 文字 / 表格）：化学式和单位上下标、落款靠右、斜线表头、合并多个文档、输入后自动加边框。
// 写格式一律用 d.Range(起,止)（WPS 里 Paragraphs.Item(i).Range 能读不能写）。
(function(root){
  'use strict';
  var T=root.OfficeTools,L=T.lib;
  function bare(t){return String(t||'').replace(/[\r\n\x07\x0b\x0c\s　 ​]/g,'');}
  function scope(a,d,what){var s=a.Selection,has=s&&bare(s.Text).length>0;if(has)return s.Range;if(!L.ask('没有选中文字。要对整篇文档执行“'+what+'”吗？'))return null;return d.Content;}
  // ---- 化学式和单位上下标 ----
  // 单位：m2、cm3、km2 等的数字改上标；化学式：只由元素符号和数字组成、含数字的词（H2O、CO2、H2SO4），数字改下标。
  var ELEMENTS='H He Li Be B C N O F Ne Na Mg Al Si P S Cl Ar K Ca Ti Cr Mn Fe Co Ni Cu Zn Br Ag Sn I Ba Au Hg Pb'.split(' ');
  var NOT_CHEM={B2B:1,B2C:1,C2C:1,P2P:1,O2O:1,C2B:1,B2G:1,F2F:1};
  function chemDigits(word){if(NOT_CHEM[word])return null;var m,re=/([A-Z][a-z]?)(\d*)/g,pos=0,syms=0,digits=[];
    while(pos<word.length){re.lastIndex=pos;m=re.exec(word);if(!m||m.index!==pos||!m[0])return null;if(ELEMENTS.indexOf(m[1])<0)return null;syms++;if(m[2])for(var k=0;k<m[2].length;k++)digits.push(pos+m[1].length+k);pos+=m[0].length;}
    if(!digits.length)return null;if(syms<2&&!/^(H|O|N|Cl|F)[23]$/.test(word))return null;return digits;}
  function marks(text){var out=[],m;
    var unit=/(^|[^A-Za-z])((?:k|c|d|m)?m)([23])(?![0-9A-Za-z])/g;while((m=unit.exec(text))){out.push({at:m.index+m[1].length+m[2].length,len:1,sup:true,word:m[2]+m[3]});}
    var word=/(^|[^A-Za-z0-9])([A-Z][A-Za-z0-9]{1,15})(?![A-Za-z0-9])/g;while((m=word.exec(text))){var start=m.index+m[1].length,ds=chemDigits(m[2]);if(ds)ds.forEach(function(i){out.push({at:start+i,len:1,sup:false,word:m[2]});});word.lastIndex=start+m[2].length;}
    return out;}
  function superSubscript(c){var a=c.a,d=c.d,r=scope(a,d,'化学式和单位上下标');if(!r)return '已取消。';var text=String(r.Text),base=Number(r.Start),list=marks(text),done=0,words=[];
    list.forEach(function(x){var cr=d.Range(base+x.at,base+x.at+x.len);if(String(cr.Text)!==text.charAt(x.at))return;if(x.sup)cr.Font.Superscript=true;else cr.Font.Subscript=true;done++;if(words.indexOf(x.word)<0)words.push(x.word);});
    if(!done)return '没有找到要改的内容（如 m2、cm3、H2O、CO2）。';
    return '已设置 '+done+' 处上标 / 下标：'+words.slice(0,8).join('、')+(words.length>8?'…':'')+'（可按 Ctrl+Z 撤销）。';}
  // ---- 落款靠右 ----
  function alignSignature(c){var a=c.a,d=c.d,s=a.Selection;if(!s||!bare(s.Text))throw Error('请先选中落款的几行（单位名称、日期等）。');var ps=s.Range.Paragraphs,n=Number(ps.Count)||0;if(n>12)throw Error('选中的段落太多了：落款通常只有几行，请只选落款部分。');var fixed=0;
    for(var i=n;i>=1;i--){var pr=ps.Item(i).Range,st=Number(pr.Start),t=String(pr.Text),lead=(t.match(/^[ \t　 ]+/)||[''])[0].length;if(lead){d.Range(st,st+lead).Delete();fixed++;}var r=d.Range(Number(ps.Item(i).Range.Start),Number(ps.Item(i).Range.End));r.ParagraphFormat.Alignment=2;r.ParagraphFormat.FirstLineIndent=0;r.ParagraphFormat.CharacterUnitFirstLineIndent=0;}
    return '已把 '+n+' 行设为靠右对齐'+(fixed?'，去掉了 '+fixed+' 行前面用来对齐的空格':'')+'（可按 Ctrl+Z 撤销）。';}
  // ---- 斜线表头 ----
  function diagonalHeader(c){var a=c.a,d=c.d,tb=null;try{tb=a.Selection.Range.Tables.Item(1);}catch(e){}if(!tb)throw Error('请先把光标放到表格里。');
    var cell=tb.Cell(1,1),b=cell.Borders.Item(-7),on=Number(b.LineStyle)!==0,cr=cell.Range,ps=cr.Paragraphs,n=Number(ps.Count)||0,note='';if(on&&n<2){b.LineStyle=0;return '已去掉左上角的斜线。';}b.LineStyle=1;
    if(n>=2){var p1=ps.Item(1).Range,p2=ps.Item(n).Range;d.Range(Number(p1.Start),Number(p1.End)).ParagraphFormat.Alignment=2;d.Range(Number(p2.Start),Number(p2.End)).ParagraphFormat.Alignment=0;note='第一行靠右、最后一行靠左。';}
    else note='在这个格子里先写右上角的标题（如“月份”），回车，再写左下角的标题（如“部门”），然后再点一次本按钮排好左右。';
    return '已在左上角单元格加了斜线。'+note;}
  // ---- 合并多个文档 ----
  function insertFiles(c){var a=c.a,d=c.d,dlg=a.FileDialog(3);dlg.Title='选择要合并进来的文档（可多选）';dlg.AllowMultiSelect=true;if(dlg.Show()!==-1)return '已取消。';
    var files=[];for(var i=1;i<=dlg.SelectedItems.Count;i++)files.push(String(dlg.SelectedItems.Item(i)));files=files.filter(function(f){return /\.(docx?|wps|wpt|rtf|txt)$/i.test(f)&&f.toLowerCase()!==String(d.FullName||'').toLowerCase();});
    if(!files.length)throw Error('没有可合并的文档（支持 doc、docx、wps、rtf、txt，不能选当前文档本身）。');files.sort(function(x,y){return x.localeCompare(y,'zh');});
    if(!L.ask('将按文件名顺序，把下面 '+files.length+' 个文档接在当前文档末尾（每个从新的一页开始），原文件不变：\n'+files.map(function(f){return f.replace(/^.*[\\\/]/,'');}).slice(0,12).join('\n')+(files.length>12?'\n…':'')+'\n是否继续？'))return '已取消。';
    var ok=0,fail=[];files.forEach(function(f){try{var end=Number(d.Content.End)-1,r=d.Range(end,end);r.InsertBreak(7);end=Number(d.Content.End)-1;d.Range(end,end).InsertFile(f);ok++;}catch(e){fail.push(f.replace(/^.*[\\\/]/,''));}});
    if(!ok)throw Error('没能合并：'+fail.join('、')+'。请确认文件没有被其他程序打开。');
    return '已合并 '+ok+' 个文档'+(fail.length?'；没能合并：'+fail.join('、'):'')+'。检查无误后请另存为新文件。';}
  // ---- WPS 表格：输入内容后自动加边框 ----
  function autoBorder(c){var a=c.a,r=a.Selection;if(!r||Number(r.Areas&&r.Areas.Count||1)>1)throw Error('请选中一块连续区域（例如 A2:F500）。');if(Number(r.Rows.Count)<2&&Number(r.Columns.Count)<2)throw Error('请选中要自动加边框的整块区域（例如 A2:F500）。');
    if(Number(r.Rows.Count)>100000)throw Error('选区太大，请少选一些行。');var tl=r.Cells.Item(1,1);try{tl.Activate();}catch(ignore){}
    var first=String(tl.Address(false,false)).replace(/[0-9]+$/,''),row=String(tl.Row),last=String(r.Cells.Item(1,Number(r.Columns.Count)).Address(false,false)).replace(/[0-9]+$/,''),formula='=COUNTA($'+first+row+':$'+last+row+')>0';
    var f=r.FormatConditions.Add(2,undefined,formula);[-4131,-4152,-4160,-4107].forEach(function(k){try{f.Borders.Item(k).LineStyle=1;}catch(ignore){try{f.Borders(k).LineStyle=1;}catch(e){}}});
    if(String(f.Formula1).replace(/\$/g,'').indexOf(first+row)<0){try{f.Delete();}catch(ignore){}throw Error('没能设置：请从区域左上角开始选中，再执行一次。');}
    return '已设置：在 '+String(r.Address(false,false))+' 里，一行只要有内容就自动加边框（开始 → 条件格式 → 管理规则 里可以删除）。';}
  // ---- 0.14：抖音常见技巧的一键版（WPS 文字）----
  function paras(range){var list=[],ps=range.Paragraphs,n=Number(ps.Count)||0;if(n>20000)throw Error('一次最多处理 2 万个段落，请分段选择。');for(var i=1;i<=n;i++)list.push(ps.Item(i));return list;}
  function inTable(r){try{return !!r.Information(12);}catch(e){return false;}}
  function listed(r){try{return Number(r.ListFormat.ListType)>0;}catch(e){return false;}}
  // 段落规范化：只改正文段落（标题、居中/靠右、表格、编号列表不动）：开头空格 → 首行缩进 2 字符；Shift+Enter 换行符 → 分段。
  function paraNormalize(c){var a=c.a,d=c.d,r=scope(a,d,'段落规范化');if(!r)return '已取消。';var body=[],lb=0,sp=0;
    paras(r).forEach(function(p){var pr=p.Range;if(Number(p.OutlineLevel)!==10||inTable(pr)||listed(pr))return;var al=Number(pr.ParagraphFormat.Alignment);if(al===1||al===2)return;var t=String(pr.Text);if(!bare(t))return;body.push({start:Number(pr.Start),end:Number(pr.End),text:t});});
    if(!body.length)return '没有找到可以整理的正文段落（标题、居中或靠右的行、表格、编号列表不会动）。';
    var rs=Number(r.Start),re=Number(r.End),cut=0;
    for(var i=body.length-1;i>=0;i--){var b=body[i],t=b.text,k;for(k=t.length-1;k>=0;k--)if(t.charAt(k)==='\x0b'){var x=d.Range(b.start+k,b.start+k+1);if(String(x.Text)==='\x0b'){x.Text='\r';lb++;}}
      var lead=(t.match(/^[ \t　 ]+/)||[''])[0].length;if(lead){d.Range(b.start,b.start+lead).Delete();sp++;cut+=lead;}}
    var fixed=0;
    // 删掉空格后范围变短了，按新的结束位置重新取段落（不依赖程序自动调整范围）。
    paras(d.Range(rs,Math.max(rs,re-cut))).forEach(function(p){var pr=p.Range;if(Number(p.OutlineLevel)!==10||inTable(pr)||listed(pr))return;var al=Number(pr.ParagraphFormat.Alignment);if(al===1||al===2||!bare(String(pr.Text)))return;var x=d.Range(Number(pr.Start),Number(pr.End)).ParagraphFormat;try{x.CharacterUnitFirstLineIndent=2;}catch(e){x.FirstLineIndent=Number(d.Range(Number(pr.Start),Number(pr.Start)+1).Font.Size||12)*2;}fixed++;});
    return '已整理 '+fixed+' 个正文段落：首行缩进 2 字符'+(sp?'，删掉了 '+sp+' 段开头用来缩进的空格':'')+(lb?'，'+lb+' 处手动换行符（Shift+Enter）改成了分段':'')+'。标题、居中或靠右的行、表格和编号列表没有动（可按 Ctrl+Z 撤销）。';}
  function pageOf(d,pos){try{d.Repaginate();}catch(e){}return Number(d.Range(pos,pos).Information(3))||0;}
  // 表格压到一页：先清除造成跨页的设置，再逐步压缩（单元格上下边距、段前段后、行高、字号），每一步都检查是否已经回到一页。
  function tableOnePage(c){var a=c.a,d=c.d,tb=null;try{tb=a.Selection.Range.Tables.Item(1);}catch(e){}if(!tb)throw Error('请先把光标放到要处理的表格里。');
    var tr=tb.Range,st=Number(tr.Start),en=Number(tr.End);function pages(){return pageOf(d,Math.max(st,en-1))-pageOf(d,st)+1;}
    if(pages()<=1)return '这张表格已经在一页里了。如果后面多出一页空白页，请用“删除空白页”。';var done=[];
    var all=d.Range(st,en);paras(all).forEach(function(p){var f=d.Range(Number(p.Range.Start),Number(p.Range.End)).ParagraphFormat;if(f.PageBreakBefore){f.PageBreakBefore=0;}if(f.KeepWithNext)f.KeepWithNext=0;if(f.KeepTogether)f.KeepTogether=0;});try{tb.Rows.AllowBreakAcrossPages=-1;}catch(ignore){}done.push('清除段前分页、与下段同页');
    if(pages()>1){try{tb.TopPadding=0;tb.BottomPadding=0;}catch(ignore){}var pf=all.ParagraphFormat;pf.SpaceBefore=0;pf.SpaceAfter=0;try{pf.LineSpacingRule=0;}catch(ignore){}done.push('单元格上下边距和段前段后改为 0、单倍行距');}
    if(pages()>1){try{tb.Rows.HeightRule=0;}catch(ignore){}done.push('行高改为自动');}
    var orig=Number(all.Font.Size)||0,size=orig,min=Math.max(8,(orig||10.5)-2);
    while(pages()>1&&size&&size-0.5>=min){size-=0.5;all.Font.Size=size;}
    if(orig&&size!==orig)done.push('字号从 '+orig+' 磅缩到 '+size+' 磅');
    var ok=pages()<=1;
    return (ok?'表格已回到一页。':'表格还是超过一页：它本身就比一页长，请删掉一些行，或在“页面”里把上下页边距改小。')+'\n已做：'+done.join('；')+'（可按 Ctrl+Z 撤销）。'+(ok?'\n如果后面多出一页空白页，请再点“删除空白页”。':'');}
  // 文字挤成一行：多出一两个字换到下一行时，把字稍微压窄（字符缩放），最多压到 70%。
  function lineOf(r){return Number(r.Information(10))||0;}
  function squeeze(d,r){var s=Number(r.Start),e=Number(r.End);while(e>s&&/[\r\x07\n]/.test(String(d.Range(e-1,e).Text)))e--;if(e<=s)return 0;var x=d.Range(s,e);try{d.Repaginate();}catch(ig){}if(lineOf(d.Range(s,s))===lineOf(d.Range(e-1,e-1)))return 0;var sc=Number(x.Font.Scaling)||100;
    while(sc>70){sc-=5;x.Font.Scaling=sc;try{d.Repaginate();}catch(ig){}if(lineOf(d.Range(s,s))===lineOf(d.Range(e-1,e-1)))return sc;}x.Font.Scaling=100;return -1;}
  function fitOneLine(c){var a=c.a,d=c.d,sel=a.Selection;if(!sel||!bare(sel.Text))throw Error('请先选中要挤到一行的文字，或选中表格里的几个格子。');var units=[];
    try{if(Number(sel.Cells.Count)>0&&inTable(sel.Range)){for(var i=1;i<=Number(sel.Cells.Count);i++)units.push(sel.Cells.Item(i).Range);}}catch(e){}if(!units.length)paras(sel.Range).forEach(function(p){units.push(p.Range);});
    if(units.length>200)throw Error('一次最多处理 200 段或 200 个格子。');var ok=0,fail=0;units.forEach(function(u){var r=squeeze(d,u);if(r>0)ok++;else if(r<0)fail++;});
    if(!ok&&!fail)return '选中的内容本来就都在一行里。';
    return '已把 '+ok+' 处文字压到一行（字符缩放）'+(fail?'；'+fail+' 处多出的字太多，压到 70% 也放不下，请把格子拉宽或删减文字':'')+'（可按 Ctrl+Z 撤销）。';}
  // 打印超大字：选中的每个字放在新文档的一页上（横向 A4），适合做门牌、桌牌、标语。
  function printBigText(c){var a=c.a,s=a.Selection,t=String(s&&s.Text||'').replace(/[\s\r\n\x07\x0b\x0c]/g,'');if(!t)throw Error('请先选中要放大打印的文字（例如“会议室”）。');var chars=Array.from(t);if(chars.length>60)throw Error('一次最多 60 个字。');
    var nd=a.Documents.Add(),ps=nd.PageSetup;try{ps.Orientation=1;}catch(ignore){}ps.TopMargin=28;ps.BottomMargin=28;ps.LeftMargin=28;ps.RightMargin=28;
    var r=nd.Content;r.Text=chars.join('\r');var f=nd.Content;f.Font.Name=L.localFont('黑体');try{f.Font.NameFarEast=L.localFont('黑体');}catch(ignore){}f.Font.Size=430;f.ParagraphFormat.Alignment=1;f.ParagraphFormat.SpaceBefore=0;f.ParagraphFormat.SpaceAfter=0;try{f.ParagraphFormat.LineSpacingRule=0;}catch(ignore){}
    for(var i=2;i<=Number(nd.Paragraphs.Count);i++){var pr=nd.Paragraphs.Item(i).Range;nd.Range(Number(pr.Start),Number(pr.End)).ParagraphFormat.PageBreakBefore=-1;}
    return '已新建文档，共 '+chars.length+' 页，每页一个字（横向 A4、'+L.localFont('黑体')+' 430 磅）。原文档没有改动。打印前可以在新文档里换字体。';}
  // 图片：拍照的文件转打印稿（灰度、提亮、加对比，去掉灰底）；手写签名去白底。
  function pics(a){var sel=a.Selection,list=[];try{for(var i=1;i<=Number(sel.InlineShapes.Count);i++)list.push(sel.InlineShapes.Item(i));}catch(e){}try{if(!list.length&&Number(sel.ShapeRange.Count))for(var k=1;k<=Number(sel.ShapeRange.Count);k++)list.push(sel.ShapeRange.Item(k));}catch(e){}if(!list.length)throw Error('请先点一下要处理的图片（可以选中一段包含多张图片的内容）。');return list;}
  function photoToPrint(c){var list=pics(c.a),ok=0;list.forEach(function(p){try{var f=p.PictureFormat;f.ColorType=2;f.Brightness=0.6;f.Contrast=0.85;ok++;}catch(ignore){}});if(!ok)throw Error('没能修改图片颜色：请在“图片工具 → 颜色”里选“灰度”，再调高亮度和对比度。');
    return '已把 '+ok+' 张图片改成适合打印的灰度（提亮、加对比，去掉拍照的灰底）。效果不满意可以在“图片工具 → 重设图片”里还原。';}
  function signatureClean(c){var list=pics(c.a),ok=0;list.forEach(function(p){try{var f=p.PictureFormat;f.ColorType=3;try{f.TransparentBackground=-1;f.TransparencyColor=16777215;}catch(ignore){}ok++;}catch(ignore){}try{if(p.ConvertToShape){var sh=p.ConvertToShape();sh.WrapFormat.Type=3;}}catch(ignore){}});
    if(!ok)throw Error('没能处理图片：请在“图片工具”里用“抠除背景”或“设置透明色”点一下白底。');
    return '已把 '+ok+' 张签名图片改成黑白、白底透明，并浮于文字上方，可以直接拖到签名处。白纸拍得不够白时，先用“照片转打印稿”。';}
  // 删除页眉横线：WPS 默认页眉段落带下边框。
  function removeHeaderLine(c){var d=c.d,n=0;for(var i=1;i<=Number(d.Sections.Count);i++){var sec=d.Sections.Item(i);for(var k=1;k<=3;k++){try{var h=sec.Headers.Item(k);if(!h.Exists&&k>1)continue;var r=h.Range,b=r.ParagraphFormat.Borders;var had=false;try{had=Number(b.Item(-3).LineStyle)!==0;}catch(e){}b.Enable=0;try{b.Item(-3).LineStyle=0;}catch(e){}if(had)n++;}catch(ignore){}}}
    try{var st=d.Styles.Item(-32).ParagraphFormat.Borders;st.Enable=0;}catch(ignore){}
    return n?'已删除 '+n+' 处页眉横线（页眉文字保留）。':'已检查全部页眉，没有找到横线；如果还看得到线，可能是页眉里插入的线条形状或表格边框，请双击页眉后手动删除。';}
  // 试卷：删掉括号里的答案（只删 A～H、√×、对错这类答案，其他括号内容不动），括号留空方便学生作答。
  var ANSWER=/([（(])\s*([A-Ha-h]{1,6}|[√×✓✗✔✘]|对|错|正确|错误|T|F)\s*([）)])/g;
  function deleteAnswers(c){var a=c.a,d=c.d,r=scope(a,d,'删除括号里的答案');if(!r)return '已取消。';var text=String(r.Text),base=Number(r.Start),hits=[],m;ANSWER.lastIndex=0;while((m=ANSWER.exec(text)))hits.push({at:m.index+m[1].length,len:m[0].length-m[1].length-m[3].length,full:m[0]});
    if(!hits.length)return '没有找到括号里的答案（例如（ B ）、（√））。';if(!L.ask('将清空 '+hits.length+' 处括号里的答案（例如 '+hits.slice(0,3).map(function(h){return h.full;}).join('、')+'），括号保留并留出空位。是否执行？'))return '已取消。';
    var ok=0;for(var i=hits.length-1;i>=0;i--){var h=hits[i],x=d.Range(base+h.at,base+h.at+h.len);if(String(x.Text).replace(/\s/g,'')!==h.full.slice(1,-1).replace(/\s/g,''))continue;x.Text='　　';ok++;}
    return '已清空 '+ok+' 处答案（可按 Ctrl+Z 撤销）。';}
  // 合同落款：甲方、乙方写在同一行，用制表位左右对齐，不再敲空格。
  function contractSign(c){var a=c.a,d=c.d,s=a.Selection;if(!s||!bare(s.Text))throw Error('请先选中“甲方：……  乙方：……”这几行落款。');var list=paras(s.Range);if(list.length>15)throw Error('选中的段落太多了，请只选落款部分。');
    var ps=d.PageSetup,width=Number(ps.PageWidth)-Number(ps.LeftMargin)-Number(ps.RightMargin),pos=Math.round(width/2),fixed=0;
    for(var i=list.length-1;i>=0;i--){var pr=list[i].Range,st=Number(pr.Start),t=String(pr.Text).replace(/[\r\x07]+$/,''),m=t.match(/^(\S.*?\S)([ 　\t]{2,}|\t)(\S.*)$/);if(!m)continue;var gs=st+m[1].length,ge=gs+m[2].length;var g=d.Range(gs,ge);if(String(g.Text)!==m[2])continue;g.Text='\t';fixed++;}
    list=paras(s.Range);list.forEach(function(p){var f=d.Range(Number(p.Range.Start),Number(p.Range.End)).ParagraphFormat;try{f.TabStops.ClearAll();}catch(ignore){}f.TabStops.Add(pos,0);f.Alignment=0;f.FirstLineIndent=0;try{f.CharacterUnitFirstLineIndent=0;}catch(ignore){}});
    return '已把 '+list.length+' 行落款设成左右两栏对齐'+(fixed?'（'+fixed+' 行中间的空格换成了制表符）':'（行中间请按 Tab 键分开甲方和乙方）')+'，乙方从页面正中开始（可按 Ctrl+Z 撤销）。';}
  // 表格里插图片不变形：固定列宽，已有图片按格子宽度缩小。
  function tablePictureFit(c){var a=c.a,d=c.d,tb=null;try{tb=a.Selection.Range.Tables.Item(1);}catch(e){}if(!tb)throw Error('请先把光标放到表格里。');try{tb.AllowAutoFit=false;}catch(ignore){}try{tb.AutoFitBehavior(0);}catch(ignore){}
    var shapes=tb.Range.InlineShapes,n=Number(shapes.Count)||0,ok=0;for(var i=1;i<=n;i++){var p=shapes.Item(i);try{var cell=p.Range.Cells.Item(1),w=Number(cell.Width)-Number(tb.LeftPadding||5.4)-Number(tb.RightPadding||5.4)-2;if(w>10&&Number(p.Width)>w){var h=Number(p.Height)*w/Number(p.Width);try{p.LockAspectRatio=-1;}catch(ig){}p.Width=w;if(Math.abs(Number(p.Height)-h)>1)p.Height=h;ok++;}}catch(ignore){}}
    return '已固定这张表格的列宽，以后插入图片不会把表格撑开'+(n?'；'+ok+' 张图片按格子宽度缩小，比例不变':'')+'。插入图片后如果还太大，再点一次本按钮。';}
  T.register({id:'paraNormalize',label:'段落规范化',desc:'空格缩进改首行缩进、手动换行改分段，标题不动',icon:'paragraph',group:'format',hosts:{wps:paraNormalize}});
  T.register({id:'tableOnePage',label:'表格压到一页',desc:'跨到下一页的表格拉回一页',icon:'table',group:'format',hosts:{wps:tableOnePage}});
  T.register({id:'fitOneLine',label:'文字挤成一行',desc:'多出一两个字换行时，压窄字距放回一行',icon:'type',group:'format',hosts:{wps:fitOneLine}});
  T.register({id:'printBigText',label:'打印超大字',desc:'每页一个字，做门牌、桌牌、标语',icon:'pdf',group:'create',undo:false,hosts:{wps:printBigText}});
  T.register({id:'photoToPrint',label:'照片转打印稿',desc:'拍照的文件去灰底，打印更清楚',icon:'picture',group:'format',hosts:{wps:photoToPrint}});
  T.register({id:'signatureClean',label:'签名去白底',desc:'手写签名照片变成透明底电子签名',icon:'picture',group:'format',hosts:{wps:signatureClean}});
  T.register({id:'removeHeaderLine',label:'删除页眉横线',desc:'去掉页眉下面那条线，页眉文字保留',icon:'clean',group:'format',hosts:{wps:removeHeaderLine}});
  T.register({id:'deleteAnswers',label:'删除括号里的答案',desc:'试卷（ B ）（√）里的答案清空',icon:'clean',group:'format',hosts:{wps:deleteAnswers}});
  T.register({id:'contractSign',label:'甲乙双方落款对齐',desc:'甲方在左、乙方在右，用制表位对齐',icon:'columns',group:'format',hosts:{wps:contractSign}});
  T.register({id:'tablePictureFit',label:'表格插图不变形',desc:'固定列宽，图片按格子缩小',icon:'picture',group:'format',hosts:{wps:tablePictureFit}});
  // 金额转人民币大写：选区没把数字选全（WPS 双击只选到小数点前）时自动补全整个数字；
  // 后面有「大写：　　」空位就直接填进去，没有才在数字后面加括号。
  function rmbUpperWps(c){
    var a=c.a,d=c.d||a.ActiveDocument,s=a.Selection,r=s&&s.Range;if(!r)throw Error('请先把光标放在金额数字上，或选中金额。');
    var st=Number(r.Start),en=Number(r.End),text=function(x,y){try{return String(d.Range(Math.max(0,x),Math.max(0,y)).Text||'');}catch(e){return '';}};
    var before=text(st-30,st),after=text(en,en+30),sel=String(r.Text||'');
    var head=(before.match(/[\d,，.]*$/)||[''])[0],tail=(after.match(/^[\d,，.]*/)||[''])[0];
    if(/[^\d,，.\s¥￥元]/.test(sel.trim().replace(/^[¥￥]/,'').replace(/元$/,'')))head=tail='';
    tail=tail.replace(/[.,，]+$/,'');head=head.replace(/^[.,，]+/,'');
    var t=(head+sel+tail).trim().replace(/[,，\s]/g,'').replace(/^[¥￥]/,'').replace(/元$/,'');
    if(!/^-?\d+(\.\d{1,2})?$/.test(t))throw Error(/\.\d{3,}$/.test(t)?'金额最多两位小数（角、分），请检查数字。':'请先把光标放在金额数字上，或选中一个金额（例如 1234.50）。');
    var up=root.LocaleEngine.rmbUpper(Number(t)),numEnd=en+tail.length,rest=text(numEnd,numEnd+40).split(/[\r\n]/)[0],m=rest.match(/大写\s*[:：]?([\s\u3000_＿]+)(?=[)）]|$)/);
    if(m&&m[1].length){var bs=numEnd+m.index+m[0].length-m[1].length;d.Range(bs,bs+m[1].length).Text=up;return '已填入大写：'+up+'。';}
    if(/^\s*[（(][^）)]*[元整角分][）)]/.test(rest))return '数字后面已经有大写了：'+rest.match(/[（(][^）)]*[）)]/)[0]+'。如需更新请先删掉再执行。';
    d.Range(numEnd,numEnd).InsertAfter('（'+up+'）');return '已在数字后面加上大写：'+up+'。';
  }
  T.register({id:'rmbUpper',keywords:'大写 人民币 金额 中文数字 壹贰叁 合同金额',hosts:{wps:rmbUpperWps}});
  T.register({id:'superSubscript',label:'化学式和单位上下标',desc:'m2→m²、H2O→H₂O，一次改好',icon:'type',group:'format',hosts:{wps:superSubscript}});
  T.register({id:'alignSignature',label:'落款靠右',desc:'选中的落款几行靠右对齐，去掉对齐用的空格',icon:'paragraph',group:'format',hosts:{wps:alignSignature}});
  T.register({id:'diagonalHeader',label:'斜线表头',desc:'表格左上角加斜线，两侧文字排好',icon:'table',group:'format',hosts:{wps:diagonalHeader}});
  T.register({id:'insertFiles',label:'合并多个文档',desc:'选几个文档，按文件名顺序接在末尾',icon:'insert',group:'file',undo:false,hosts:{wps:insertFiles}});
  T.register({id:'autoBorder',label:'输入后自动加边框',desc:'一行有内容就出现边框',icon:'table',group:'sheetplus',undo:false,hosts:{et:autoBorder}});
  root.OfficeWordPlus={marks:marks,chemDigits:chemDigits,ANSWER:ANSWER};
})(window);

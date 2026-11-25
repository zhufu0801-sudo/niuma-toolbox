// 办公工具箱 · 文字模块（WPS 文字）。
(function(root){
  'use strict';
  var T=root.OfficeTools,L=T.lib,labels=L.labels,active=L.active,wordSelection=L.wordSelection,tableInput=L.tableInput,selectedEnd=L.selectedEnd,replaceOptions=L.replaceOptions,position=L.position;
  function addWordTable(d,range,rows){var t=d.Tables.Add(range,rows.length,rows[0].length);try{for(var r=0;r<rows.length;r++)for(var c=0;c<rows[r].length;c++)t.Cell(r+1,c+1).Range.Text=rows[r][c];t.Borders.Enable=true;return t;}catch(e){try{t.Delete();}catch(ignore){}throw e;}}
  // 拖拽或粘贴后，让新内容跟随周围的格式。
  function matchWord(a,d){
    var s=wordSelection(a),range=s.Range,start=Number(range.Start),end=Number(range.End),total=Number(d.Content.End);
    // 先找选区前后的正文段落做参照；紧挨着标题时不能把新内容也变成标题。
    var candidates=[];if(start>0)candidates.push(d.Range(start-1,start));if(end<total-1)candidates.push(d.Range(end,end+1));
    if(!candidates.length)throw Error('文档里没有可以参照的周围文字。');
    var ref=candidates[0];for(var i=0;i<candidates.length;i++){try{if(Number(candidates[i].Paragraphs.Item(1).OutlineLevel)===10){ref=candidates[i];break;}}catch(ignore){}}
    var refPara=ref.Paragraphs.Item(1).Range,mixed=function(v){return v===undefined||v===null||v===''||Number(v)===9999999;};
    try{range.Style=refPara.Style;}catch(ignore){}
    // 中文字体在 NameFarEast，Name 只是西文字体；两者都要跟随，提示里优先说中文字体。
    var name=ref.Font.Name,east=ref.Font.NameFarEast,size=ref.Font.Size,color=ref.Font.Color,bold=ref.Font.Bold,italic=ref.Font.Italic;
    if(!mixed(name))range.Font.Name=name;if(!mixed(east))range.Font.NameFarEast=east;if(!mixed(size))range.Font.Size=size;if(!mixed(color))range.Font.Color=color;
    if(!mixed(bold))range.Font.Bold=bold;if(!mixed(italic))range.Font.Italic=italic;
    var shown=/[\u3400-\u9fff]/.test(String(range.Text))&&!mixed(east)?east:name;
    try{range.ParagraphFormat.LineSpacingRule=refPara.ParagraphFormat.LineSpacingRule;range.ParagraphFormat.LineSpacing=refPara.ParagraphFormat.LineSpacing;range.ParagraphFormat.SpaceAfter=refPara.ParagraphFormat.SpaceAfter;}catch(ignore){}
    var tables=0;try{for(var t=1;t<=range.Tables.Count;t++){range.Tables.Item(t).AutoFitBehavior(2);tables++;}}catch(ignore){}
    return '已按周围正文设置'+(mixed(shown)?'':'字体“'+shown+'”')+(mixed(size)?'':'、'+size+' 号')+(tables?'，并让 '+tables+' 张表格适应页面宽度':'')+'。';
  }
  function replaceWord(a,d,options){
    var selected=wordSelection(a),range=selected.Range,spec=replaceOptions(options);
    if(range.StoryType!==1||(range.Tables&&range.Tables.Count)||(range.Fields&&range.Fields.Count)||d.TrackRevisions)throw Error('快速替换仅支持未启用修订的正文普通文字选区；表格、域、页眉页脚请用 WPS 原生替换。');
    var text=String(range.Text),offset=0,positions=[],target=position(a,'wps',d);
    while((offset=text.indexOf(spec.find,offset))!==-1){positions.push(offset);offset+=spec.find.length;if(positions.length>500)throw Error('一次最多替换 500 处，请缩小选区。');}
    if(!positions.length)return '当前选区没有找到对应文字。';
    if(!L.ask('仅替换当前选区中 '+positions.length+' 处匹配文字，保留选区外内容。是否执行？'))return '已取消。';
    if(position(a,'wps',active(a,'wps'))!==target||String(range.Text)!==text)throw Error('当前文档、选区或内容已变化，请重新执行。');
    var start=range.Start;for(var i=positions.length-1;i>=0;i--)d.Range(start+positions[i],start+positions[i]+spec.find.length).Text=spec.replacement;
    return '已替换当前选区中的 '+positions.length+' 处文字。';
  }
  L.addWordTable=addWordTable;
  T.register({id:'heading1',label:'一级标题',desc:'快速应用章节样式',icon:'heading1',group:'format',hosts:{wps:function(c){var s=wordSelection(c.a);s.Range.Style=-2;return '已将选中段落设为一级标题。';}}});
  T.register({id:'heading2',label:'二级标题',desc:'理清内容层级',icon:'heading2',group:'format',hosts:{wps:function(c){var s=wordSelection(c.a);s.Range.Style=-3;return '已将选中段落设为二级标题。';}}});
  // 正文排版：只整理正文段落——标题（大纲级别、标题样式、明显更大的字、“一、”开头的加粗短行）跳过，
  // 选区里的空行删掉，完成后说清楚改了几段。没选中文字时不动文档（真机：光标停在标题上会把标题改小）。
  function looksLikeTitle(p,r,t,prefs){
    try{var lv=Number(p.OutlineLevel);if(lv>=1&&lv<=9)return true;}catch(e){}
    try{var st=r.Style,name=String(st&&(st.NameLocal||st.Name)||st||'');if(/标题|標題|見出し|Heading|Title/i.test(name))return true;}catch(e){}
    // 2026-11-20 真机：WPS 里整段（含段落标记）的 Bold 常返回“混合”，所以按第一个字判断是否加粗。
    var size=0,bold=false,isB=function(v){return v===true||v===-1||v===1;};try{size=Number(r.Font.Size)||0;bold=isB(r.Font.Bold);if(!bold&&r.Characters&&r.Characters.Item)bold=isB(r.Characters.Item(1).Font.Bold);}catch(e){}
    if(t.length<=40&&size>=prefs.size+3)return true;
    var numbered=/^\s*([一二三四五六七八九十]+[、.．]|第[一二三四五六七八九十\d]+[章节部分])/.test(t);
    // “二、本周计划”这类短的大写序号行，即使没加粗、末尾也没有句号，也按小标题处理。
    return numbered&&(t.length<=30&&bold||t.length<=16&&!/[。；;，,：:！!？?]\s*$/.test(t));
  }
  // 2026-11-21：字体原来只看界面语言，日文界面整理中文文档会变成 Yu Mincho。改成看选中文字：有假名按日文字体，只有汉字按中文字体。
  var TO_JA={'宋体':'Yu Mincho','微软雅黑':'Yu Gothic','黑体':'Yu Gothic'},TO_ZH={'Yu Mincho':'宋体','Yu Gothic':'微软雅黑'};
  function fontForText(font,text){if(/[\u3040-\u30ff]/.test(text))return TO_JA[font]||font;if(/[\u4e00-\u9fff]/.test(text))return TO_ZH[font]||font;return font;}
  function bodyFormat(c){
    var a=c.a,s=a.Selection;
    if(s&&s.Start!=null&&s.End!=null&&Number(s.Start)===Number(s.End))throw Error('请先选中要整理的正文（整篇可按 Ctrl+A 全选），标题会自动跳过。');
    s=wordSelection(a);var prefs=T.settings().text;prefs.font=fontForText(prefs.font,String(s.Text||'').slice(0,3000));
    function apply(x){x.Font.Name=prefs.font;x.Font.NameFarEast=prefs.font;try{x.Font.NameAscii=prefs.font;x.Font.NameOther=prefs.font;}catch(ignore){}x.Font.Size=prefs.size;x.ParagraphFormat.LineSpacingRule=5;x.ParagraphFormat.LineSpacing=12*prefs.lineSpacing;x.ParagraphFormat.SpaceBefore=0;x.ParagraphFormat.SpaceAfter=6;}
    var spec=prefs.font+'、'+prefs.size+' 号、'+prefs.lineSpacing+' 倍行距',ps=s.Paragraphs,n=ps&&Number(ps.Count);
    if(!n){s.Range.Style=-1;apply(s);return '已整理选中文字：'+spec+'。';}
    var body=[],titles=0,empties=[];
    for(var i=1;i<=n;i++){var p=ps.Item(i),r=p.Range,t=String(r.Text||'').replace(/[\r\n\u0007]/g,''),inTable=false;try{inTable=!!r.Information(12);}catch(e){}
      if(!t.replace(/[\s　]/g,'')){if(!inTable)empties.push(r);continue;}
      if(looksLikeTitle(p,r,t,prefs)){titles++;continue;}body.push(r);}
    if(!body.length)throw Error('选中的都是标题或空行，没有需要整理的正文。');
    if(!L.ask('将整理 '+body.length+' 段正文：'+spec+(titles?'；跳过 '+titles+' 个标题':'')+(empties.length?'；删除 '+empties.length+' 个空行':'')+'。\n可以用 Ctrl+Z 撤销。是否执行？'))return '已取消。';
    body.forEach(function(r){try{r.Style=-1;}catch(e){}apply(r);});
    var removed=0;for(var k=empties.length-1;k>=0;k--){try{empties[k].Delete();removed++;}catch(e){}}
    return '已整理 '+body.length+' 段正文（'+spec+'）'+(titles?'，跳过 '+titles+' 个标题':'')+(removed?'，删除 '+removed+' 个空行':'')+'。';
  }
  T.register({id:'bodyFormat',label:'正文排版',desc:'统一字体与行距，标题不动',icon:'paragraph',group:'format',ownConfirm:['wps'],keywords:'格式乱 排版 统一格式 整理格式 字体不一样 弄整齐 美化文档 書式 整える レイアウト',hosts:{wps:bodyFormat}});
  T.register({id:'toc',label:'插入 / 更新目录',desc:'根据标题生成目录',icon:'list',group:'format',hosts:{wps:function(c){var a=c.a,d=c.d;if(d.TablesOfContents.Count){for(var t=1;t<=d.TablesOfContents.Count;t++)d.TablesOfContents.Item(t).Update();return '目录已更新。';}var level=false;for(var i=1;i<=Math.min(d.Paragraphs.Count,2000);i++){if(Number(d.Paragraphs.Item(i).OutlineLevel)<=3){level=true;break;}}if(!level)throw Error('未找到一至三级标题，请先把章节设为标题。');d.TablesOfContents.Add(selectedEnd(a),true,1,3);return '已在选区末尾插入目录。';}}});
  T.register({id:'textTable',label:'文字转表格',desc:'将分隔文字整理成表',icon:'table',group:'format',hosts:{wps:function(c){var s=wordSelection(c.a);if(String(s.Text).indexOf('\t')<0)throw Error('请选中用 Tab 分隔列、换行分隔行的文字。');s.ConvertToTable(1);return '选中文字已转换为表格。';}}});
  T.register({id:'removeBlankParas',label:'清理空段落',desc:'整理选区内多余空行',icon:'clean',group:'format',hosts:{wps:function(c){var a=c.a,d=c.d,selection=wordSelection(a),range=selection.Range;if(range.StoryType!==1||(range.Tables&&range.Tables.Count)||d.TrackRevisions)throw Error('仅支持未启用修订、且不包含表格的正文选区。');var removalTarget=position(a,'wps',d),removals=[];for(var p=range.Paragraphs.Count;p>=1;p--){var paragraph=range.Paragraphs.Item(p).Range;if(paragraph.Start>=range.Start&&paragraph.End<=range.End&&paragraph.End<d.Content.End&&/^[ \t\r\n　]+$/.test(String(paragraph.Text)))removals.push(paragraph);}if(!removals.length)return '选区内没有可删除的完整空段落。';if(!L.ask('将删除选区内 '+removals.length+' 个完整空段落，保留其他内容。是否执行？'))return '已取消。';if(position(a,'wps',active(a,'wps'))!==removalTarget)throw Error('当前文档或选区已切换，请重新执行。');removals.forEach(function(paragraph){paragraph.Delete();});return '已删除 '+removals.length+' 个空段落。';}}});
  T.register({id:'replaceText',label:'查找替换',desc:'只替换选区里的文字',icon:'replace',group:'format',panel:false,hosts:{wps:function(c){return replaceWord(c.a,c.d,c.options);}}});
  T.register({id:'insertTable',label:'快速插入表格',desc:'粘贴内容生成表格',icon:'table',group:'create',panel:false,hosts:{wps:function(c){addWordTable(c.d,selectedEnd(c.a),tableInput(c.options));return '表格已插入选区末尾。';}}});
  T.register({id:'matchFormat',label:'匹配周围格式',desc:'拖拽或粘贴后跟随周围正文',icon:'match',group:'move',hosts:{wps:function(c){return matchWord(c.a,c.d);}}});
  // ---- 图片（整理自猴哥“Word 插入图片统一大小／表格插图不变形／打印身份证”等作品题目，步骤按官方资料补全）----
  var CM=28.3465;
  function pictures(a){var s=a.Selection;if(!s||!s.Range)throw Error('请先选中包含图片的内容。要处理全文，可先按 Ctrl+A。');var list=s.Range.InlineShapes,n=Number(list&&list.Count)||0,out=[];
    for(var i=1;i<=n;i++){var p=list.Item(i);if(Number(p.Width)>0&&Number(p.Height)>0)out.push(p);}
    if(!out.length)throw Error('选区里没有嵌入型图片。浮动图片请先在“环绕方式”里改成“嵌入型”，或逐张调整。');if(out.length>500)throw Error('一次最多处理 500 张图片，请缩小选区。');return out;}
  function resize(p,w){var h=Number(p.Height)*w/Number(p.Width);try{p.LockAspectRatio=-1;}catch(ignore){}p.Width=w;p.Height=h;}
  T.register({id:'unifyPictures',label:'图片统一大小',desc:'以第一张为准，按比例统一宽度',icon:'columns',group:'text',hosts:{wps:function(c){
    var list=pictures(c.a);if(list.length<2)throw Error('选区里只有 1 张图片。请选中至少 2 张图片（可按 Ctrl+A 选全文）。');var w=Number(list[0].Width);
    list.forEach(function(p){resize(p,w);});return '已把 '+list.length+' 张嵌入型图片统一为第一张的宽度（约 '+(w/CM).toFixed(1)+' 厘米），高度按各自比例变化，不会变形。';}}});
  T.register({id:'fitPictures',label:'图片不超出页面',desc:'太宽的图片按比例缩到版心',icon:'wrap',group:'text',hosts:{wps:function(c){
    var d=active(c.a,'wps'),list=pictures(c.a),ps;try{ps=c.a.Selection.Sections.Item(1).PageSetup;}catch(ignore){ps=d.PageSetup;}
    var max=Number(ps.PageWidth)-Number(ps.LeftMargin)-Number(ps.RightMargin);if(!(max>0))throw Error('读不到页面宽度，请检查页面设置。');
    var n=0;list.forEach(function(p){if(Number(p.Width)>max+0.5){resize(p,max);n++;}});
    return n?'已把 '+n+' 张超出版心的图片缩到 '+(max/CM).toFixed(1)+' 厘米宽，比例不变。表格里的图片还受列宽限制，请另外检查。':'选区里的图片都没有超出版心（'+(max/CM).toFixed(1)+' 厘米），没有修改。';}}});
  T.register({id:'idCardSize',label:'证件照按实际尺寸',desc:'身份证等卡片设为 8.56×5.4 厘米',icon:'document',group:'text',hosts:{wps:function(c){
    var list=pictures(c.a);if(list.length>2)throw Error('一次最多处理 2 张（正面和反面）。请只选中证件图片。');
    var bad=list.filter(function(p){var r=Number(p.Width)/Number(p.Height);return Math.abs(r-8.56/5.4)/(8.56/5.4)>0.06;});
    if(bad.length)throw Error('有 '+bad.length+' 张图片的长宽比和身份证不符（应约为 1.59:1）。请先用 WPS 的“裁剪”去掉多余背景，再执行。');
    list.forEach(function(p){try{p.LockAspectRatio=0;}catch(ignore){}p.Width=8.56*CM;p.Height=5.4*CM;});
    return '已把 '+list.length+' 张图片设为 8.56×5.4 厘米。打印时请选“实际大小”，不要选“缩放到纸张”，印出来后用尺子量一下。证件复印件请注明用途。';}}});
})(window);

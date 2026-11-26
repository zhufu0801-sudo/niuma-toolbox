// 办公工具箱 · 文档整理（WPS 文字）：删除空白页、公文 / 日文商务格式、跳页修复、局部横向、
// 表格美化、图片修复、行内标题加粗、选项对齐、西文不断行、去水印、中英文空格。
// 原则：先说明要改什么，确认后再改；整篇文档的操作一律先问；能撤销（合并为一步）。
(function(root){
  'use strict';
  var T=root.OfficeTools,L=T.lib,active=L.active,position=L.position;
  var MM=2.83465,CJK='\\u3040-\\u30ff\\u3400-\\u9fff\\uf900-\\ufaff';
  function rgbHex(h){var m=/^#?([0-9a-f]{6})$/i.exec(String(h||''));if(!m)return 14281463;var n=parseInt(m[1],16);return (n>>16)+((n>>8)&255)*256+(n&255)*65536;}
  function bare(t){return String(t||'').replace(/[\r\n\x07\x0b\x0c\s　 ​]/g,'');}
  // 选区有文字就处理选区，否则（只是光标）问一次，处理全文。
  function scope(a,d,what,preview){var s=a.Selection,has=s&&bare(s.Text).length>0;
    if(!preview&&!L.ask(has?'将对选中的内容执行“'+what+'”。可以按 Ctrl+Z 撤销。是否执行？':'没有选中文字。要对整篇文档执行“'+what+'”吗？'))return null;return {range:has?s.Range:d.Content,all:!has};}
  function paragraphsOf(range){var list=[],ps=range.Paragraphs,n=Number(ps.Count)||0;if(n>20000)throw Error('一次最多处理 2 万个段落，请分段选择。');for(var i=1;i<=n;i++)list.push(ps.Item(i));return list;}
  // 真机发现：WPS 里通过 Paragraphs.Item(i).Range 改字体、段落格式不生效（能读不能写），所以改格式一律用 d.Range(起,止)。
  function R(d,p){var r=p.Range;return d.Range(Number(r.Start),Number(r.End));}
  function pageOf(d,pos){return Number(d.Range(pos,pos).Information(3))||0;}

  // ---------- 删除空白页 ----------
  // 只删“整页没有任何内容”的页：页上只有空段落、分页符。分节符造成的空白页只提示，不删（会影响页眉页码和页面方向）。
  function scanBlank(d){
    var total=Number(d.Content.Information(4))||1,content={},items=[],sectionBreaks=[],n=Number(d.Paragraphs.Count)||0,i;
    if(n>20000)throw Error('文档超过 2 万段，请先拆分后再处理。');
    for(i=1;i<Number(d.Sections.Count);i++)sectionBreaks.push(Number(d.Sections.Item(i).Range.End)-1);
    for(i=1;i<=n;i++){var p=d.Paragraphs.Item(i),r=p.Range,t=String(r.Text),sp=pageOf(d,r.Start),ep=Number(r.Information(3))||sp,inTable=!!r.Information(12),pics=Number(r.InlineShapes&&r.InlineShapes.Count)||0;
      if(bare(t)||pics||inTable){for(var pg=sp;pg<=ep;pg++)content[pg]=1;continue;}
      var sb=sectionBreaks.some(function(x){return x>=r.Start&&x<r.End;});items.push({i:i,start:Number(r.Start),end:Number(r.End),sp:sp,ep:ep,breaks:(t.match(/\f/g)||[]).length,section:sb,last:i===n,prevTable:i>1&&!!d.Paragraphs.Item(i-1).Range.Information(12),nextTable:i<n&&!!d.Paragraphs.Item(i+1).Range.Information(12)});}
    try{for(i=1;i<=Number(d.Shapes.Count);i++){var a=d.Shapes.Item(i).Anchor;if(a)content[pageOf(d,a.Start)]=1;}}catch(ignore){}
    var blank=[];for(var k=1;k<=total;k++)if(!content[k])blank.push(k);
    return {total:total,blank:blank,items:items};
  }
  function touches(it,blank,total){for(var p=it.sp;p<=it.ep;p++)if(blank.indexOf(p)>=0)return true;if(it.breaks){for(var q=it.sp+1;q<=total;q++)if(blank.indexOf(q)<0)return false;return it.sp<total;}return false;}
  function shrink(r){r.Font.Size=1;var f=r.ParagraphFormat;f.SpaceBefore=0;f.SpaceAfter=0;f.LineSpacingRule=4;f.LineSpacing=1;}
  function deleteBlankPages(c){
    var a=c.a,d=c.d;if(d.TrackRevisions)throw Error('文档开着“修订”，请先关闭修订再删除空白页。');
    var scan=scanBlank(d);if(!scan.blank.length)return '没有找到空白页（共 '+scan.total+' 页，每页都有内容）。';
    var del=[],keep=[],fix=null;scan.items.forEach(function(it){if(!touches(it,scan.blank,scan.total))return;
      if(it.section)keep.push(it);else if(it.last){fix=it;if(it.breaks)del.push(Object.assign({},it,{end:it.end-1}));}else if(it.prevTable&&it.nextTable)return;else del.push(it);});
    if(!del.length&&!fix)return '没有可以自动删除的空白页：第 '+scan.blank.join('、')+' 页由分节符造成。为了不打乱页眉、页码和页面方向，请手动检查这些分节符。';
    var target=position(a,'wps',d);
    if(!L.ask('发现空白页：第 '+scan.blank.join('、')+' 页（共 '+scan.total+' 页）。\n将删除这些页上的 '+del.length+' 个空段落或分页符'+(fix?'，并把文末的空段落缩到最小（表格后面那一页就是它造成的）':'')+'。'+(keep.length?'\n其中 '+keep.length+' 处是分节符，不会删除。':'')+'\n可以按 Ctrl+Z 撤销。是否执行？'))return '已取消。';
    if(position(a,'wps',active(a,'wps'))!==target)throw Error('当前文档已切换，请重新执行。');
    del.sort(function(x,y){return y.start-x.start;}).forEach(function(it){d.Range(it.start,it.end).Delete();});
    if(fix){var ps=d.Paragraphs;shrink(R(d,ps.Item(Number(ps.Count))));}
    // 删除后 WPS 异步重新分页，马上再数页会读到旧页数（真机测试发现），所以只报告原来的情况。
    return '已删除 '+del.length+' 处空段落或分页符'+(fix?'，并缩小了文末空段落':'')+'。原来共 '+scan.total+' 页，空白的是第 '+scan.blank.join('、')+' 页；请看一下状态栏的页数确认。';
  }

  // ---------- 公文格式（GB/T 9704-2012）和日文商务文书格式 ----------
  function fontSet(a){var set=null;try{var f=a.FontNames,n=Number(f.Count);if(n>0){set={};for(var i=1;i<=n;i++)set[String(f.Item(i))]=1;}}catch(ignore){}return set;}
  function pick(set,list){if(!set)return list[list.length-1];for(var i=0;i<list.length;i++)if(set[list[i]])return list[i];return list[list.length-1];}
  function setFont(r,name,size,bold){r.Font.NameFarEast=name;r.Font.Name=name;if(size)r.Font.Size=size;if(bold!==undefined)r.Font.Bold=bold?1:0;}
  function pageSetup(d,m){for(var i=1;i<=Number(d.Sections.Count);i++){var ps=d.Sections.Item(i).PageSetup;if(Number(ps.Orientation)!==1){ps.PageWidth=210*MM;ps.PageHeight=297*MM;}ps.TopMargin=m[0]*MM;ps.BottomMargin=m[1]*MM;ps.LeftMargin=m[2]*MM;ps.RightMargin=m[3]*MM;}}
  function footersEmpty(d){for(var i=1;i<=Number(d.Sections.Count);i++)for(var k=1;k<=3;k++){try{if(bare(d.Sections.Item(i).Footers.Item(k).Range.Text))return false;}catch(ignore){}}return true;}
  function pageNumber(d,f,align,text,font,size,indent){var r=f.Range;r.Text=text.replace('#','X');r=f.Range;var at=String(r.Text).indexOf('X');if(at<0)return;
    // 页码位置必须在页脚里取（SetRange 不换文字区）；用 d.Range 会落到正文里，真机上整次修改都会被 WPS 撤回。
    var x=f.Range;x.SetRange(Number(r.Start)+at,Number(r.Start)+at+1);if(String(x.Text)!=='X')throw Error('页脚位置不对');d.Fields.Add(x,33);r=f.Range;setFont(r,font,size,false);r.ParagraphFormat.Alignment=align;try{if(indent==='right')r.ParagraphFormat.CharacterUnitRightIndent=1;if(indent==='left')r.ParagraphFormat.CharacterUnitLeftIndent=1;}catch(ignore){}}
  function lineSpacing(r,pt){r.ParagraphFormat.LineSpacingRule=4;r.ParagraphFormat.LineSpacing=pt;}
  function firstIndent(r,chars,size){try{r.ParagraphFormat.CharacterUnitFirstLineIndent=chars;}catch(e){r.ParagraphFormat.FirstLineIndent=chars*size;}}
  function noIndent(r){try{r.ParagraphFormat.CharacterUnitFirstLineIndent=0;}catch(ignore){}r.ParagraphFormat.FirstLineIndent=0;}
  function isDate(t){return /^(?:(?:令和|平成|昭和)\s*[0-9０-９一二三四五六七八九十元]+|[0-9０-９]{4})\s*年[^。]{0,8}月[^。]{0,8}日$/.test(t)||/^[0-9０-９]{4}[./／-][0-9０-９]{1,2}[./／-][0-9０-９]{1,2}$/.test(t);}
  function isRef(t){return /[〔［\[][^\]〕］]+[〕］\]]\s*[0-9０-９]+\s*号$/.test(t);}
  function titleIndex(ps,japanese){
    var candidates=[],limit=Math.min(ps.length,20);
    for(var i=0;i<limit;i++){var r=ps[i].Range,t=String(r.Text).trim();
      if(!bare(t)||isDate(t)||isRef(t)||t.length>60||/[。．.；;，,：:]$/.test(t)||/^(拝啓|謹啓|前略|記$|以上$)/.test(t)||/(御中|様|各位)$/.test(t)||/(株式会社|有限会社|合同会社|有限公司)/.test(t))continue;
      var f=r.Font||{},semantic=japanese?/(お知らせ|ご案内|お願い|依頼|通知|について)$/.test(t):/(通知|通告|公告|报告|请示|决定|函|纪要|方案|计划)$/.test(t);
      candidates.push({i:i,score:(semantic?100:0)+(Number(f.Bold)!==0&&Number(f.Bold)!==9999999&&Number.isFinite(Number(f.Bold))?20:0)+(Number(f.Size)>0&&Number(f.Size)<100?Number(f.Size):0)});
    }
    candidates.sort(function(a,b){return b.score-a.score||a.i-b.i;});return candidates.length?candidates[0].i:-1;
  }
  function formatCn(c){
    var a=c.a,d=c.d;if(!L.ask('将按国家标准 GB/T 9704-2012 排版整篇文档：A4、页边距上 37 / 下 35 / 左 28 / 右 26 毫米，正文三号仿宋、每页 22 行，标题和各级小标题换成规定字体'+(footersEmpty(d)?'，并加上“— 1 —”式页码':'（已有页脚，不改页码）')+'。\n可以按 Ctrl+Z 撤销。是否执行？'))return '已取消。';
    var fs=fontSet(a),body=pick(fs,['仿宋_GB2312','仿宋']),kai=pick(fs,['楷体_GB2312','楷体']),hei=pick(fs,['黑体']),title=pick(fs,['方正小标宋简体','方正小标宋_GBK','华文中宋','宋体']);
    pageSetup(d,[37,35,28,26]);var ps=paragraphsOf(d.Content),ti=titleIndex(ps),n={title:0,h1:0,h2:0,h3:0,body:0};
    ps.forEach(function(p,i){var r=R(d,p),t=String(r.Text).trim();if(r.Information(12))return;if(!bare(t))return;
      lineSpacing(r,28.95);r.ParagraphFormat.SpaceBefore=0;r.ParagraphFormat.SpaceAfter=0;
      if(i===ti){setFont(r,title,22,false);r.ParagraphFormat.Alignment=1;noIndent(r);n.title++;return;}
      r.ParagraphFormat.Alignment=3;firstIndent(r,2,16);
      if(/^[一二三四五六七八九十]+、/.test(t)){setFont(r,hei,16,false);n.h1++;}
      else if(/^[（(][一二三四五六七八九十]+[）)]/.test(t)){setFont(r,kai,16,false);n.h2++;}
      else if(/^\d{1,2}[\.．]/.test(t)){setFont(r,body,16,true);n.h3++;}
      else{setFont(r,body,16,false);n.body++;}});
    var pn='';if(footersEmpty(d)){try{d.PageSetup.OddAndEvenPagesHeaderFooter=-1;var s1=d.Sections.Item(1);pageNumber(d,s1.Footers.Item(1),2,'— # —','宋体',14,'right');pageNumber(d,s1.Footers.Item(3),0,'— # —','宋体',14,'left');pn='，加了页码（单页居右、双页居左）';}catch(e){pn='（页码没加上：'+e.message+'）';}}
    return '已按公文格式排版：标题 '+n.title+' 个、一级小标题 '+n.h1+' 个、二级 '+n.h2+' 个、三级 '+n.h3+' 个、正文 '+n.body+' 段'+pn+'。使用字体：正文“'+body+'”、标题“'+title+'”。表格未改动。';
  }
  function formatJa(c){
    var a=c.a,d=c.d;if(!L.ask('将按日文商务文书的常用格式排版整篇文档：A4、页边距上 35 / 下 30 / 左右 30 毫米，正文明朝体 10.5 磅，标题哥特体加粗居中，“記”居中、“以上”靠右、开头的日期靠右'+(footersEmpty(d)?'，并在页脚中间加“- 1 -”式页码':'（已有页脚，不改页码）')+'。\n可以按 Ctrl+Z 撤销。是否执行？'))return '已取消。';
    var fs=fontSet(a),mincho=pick(fs,['游明朝','Yu Mincho','ＭＳ 明朝','MS Mincho','宋体']),gothic=pick(fs,['游ゴシック','Yu Gothic','ＭＳ ゴシック','MS Gothic','黑体']);
    pageSetup(d,[35,30,30,30]);var ps=paragraphsOf(d.Content),ti=titleIndex(ps,true),recipientEnd=-1,n={title:0,head:0,body:0};ps.forEach(function(p,i){if(i<ti&&/(御中|様|各位)$/.test(String(p.Range.Text).trim()))recipientEnd=i;});
    ps.forEach(function(p,i){var r=R(d,p),t=String(r.Text).trim();if(r.Information(12))return;if(!bare(t))return;
      lineSpacing(r,18);r.ParagraphFormat.SpaceBefore=0;r.ParagraphFormat.SpaceAfter=0;
      if(i===ti){setFont(r,gothic,14,true);r.ParagraphFormat.Alignment=1;noIndent(r);n.title++;return;}
      if(t==='記'||t==='记'){setFont(r,mincho,10.5,false);r.ParagraphFormat.Alignment=1;noIndent(r);n.body++;return;}
      if(t==='以上'||(i<ti&&isDate(t))){setFont(r,mincho,10.5,false);r.ParagraphFormat.Alignment=2;noIndent(r);n.body++;return;}
      if(i<ti){setFont(r,mincho,10.5,false);r.ParagraphFormat.Alignment=recipientEnd>=0&&i>recipientEnd?2:0;noIndent(r);n.body++;return;}
      if(/^(第[0-9０-９一二三四五六七八九十]+[章条節]|[0-9０-９]{1,2}[\.．、]|[（(][0-9０-９]{1,2}[）)])/.test(t)&&t.length<=40){setFont(r,gothic,10.5,true);r.ParagraphFormat.Alignment=0;noIndent(r);n.head++;return;}
      setFont(r,mincho,10.5,false);r.ParagraphFormat.Alignment=3;firstIndent(r,1,10.5);n.body++;});
    var pn='';if(footersEmpty(d)){try{pageNumber(d,d.Sections.Item(1).Footers.Item(1),1,'- # -',mincho,10.5,'');pn='，加了页码';}catch(e){pn='（页码没加上：'+e.message+'）';}}
    return '已按日文商务文书格式排版：标题 '+n.title+' 个、小标题 '+n.head+' 个、正文 '+n.body+' 段'+pn+'。使用字体：正文“'+mincho+'”、标题“'+gothic+'”。表格未改动。';
  }

  // ---------- 段落乱跳页 ----------
  function fixPageJump(c){
    var a=c.a,d=c.d,sc=scope(a,d,'修复段落乱跳页');if(!sc)return '已取消。';var n={pb:0,kn:0,kt:0,rows:0};
    paragraphsOf(sc.range).forEach(function(p){var f=R(d,p).ParagraphFormat,body=Number(p.OutlineLevel)===10;
      if(f.PageBreakBefore){f.PageBreakBefore=0;n.pb++;}if(body&&f.KeepWithNext){f.KeepWithNext=0;n.kn++;}if(body&&f.KeepTogether){f.KeepTogether=0;n.kt++;}});
    try{for(var t=1;t<=Number(sc.range.Tables.Count);t++){var rows=sc.range.Tables.Item(t).Rows;if(!rows.AllowBreakAcrossPages){rows.AllowBreakAcrossPages=-1;n.rows++;}}}catch(ignore){}
    var sum=n.pb+n.kn+n.kt+n.rows;if(!sum)return '没有发现会造成跳页的设置（段前分页、与下段同页、段中不分页、表格行禁止跨页）。如果还跳页，请打开“显示编辑标记”看看是不是有分页符。';
    return '已清除：段前分页 '+n.pb+' 处、与下段同页 '+n.kn+' 处、段中不分页 '+n.kt+' 处'+(n.rows?'，并允许 '+n.rows+' 张表格的行跨页':'')+'。标题的“与下段同页”保留（防止标题落在页尾）。';
  }

  // ---------- 选中内容改横向 ----------
  function landscape(c){
    var a=c.a,d=c.d,s=a.Selection;if(!s||!bare(s.Text))throw Error('请先选中要横向显示的内容（例如一张宽表格）。');
    var r=s.Range,start=Number(r.Start),end=Number(r.End),docEnd=Number(d.Content.End)-1,sec=r.Sections.Item(1),secR=sec.Range;
    if(Number(r.Sections.Count)===1&&Number(secR.Start)===start&&Number(secR.End)>=end-1){var ps=sec.PageSetup,land=Number(ps.Orientation)===1;if(!L.ask('所选内容已单独成节，将改为'+(land?'纵向':'横向')+'。可以按 Ctrl+Z 撤销。是否执行？'))return '已取消。';ps.Orientation=land?0:1;return land?'这一节原来是横向，已改回纵向。':'所选内容已单独成节，已改为横向。';}
    if(!L.ask('将在所选内容前后各插入一个“下一页”分节符，并把这部分改成横向，前后页面保持纵向。是否执行？'))return '已取消。';
    if(end<docEnd)d.Range(end,end).InsertBreak(2);var shift=0;if(start>0){d.Range(start,start).InsertBreak(2);shift=1;}
    var mid=d.Range(start+shift,start+shift).Sections.Item(1);mid.PageSetup.Orientation=1;
    return '已把所选内容放进单独的一节并改为横向（插入了 '+((end<docEnd?1:0)+shift)+' 个分节符）。页眉页脚沿用前一节；需要不同的页眉时，在页眉里取消“同前节”。';
  }

  // ---------- 表格一键美化 ----------
  function tableBeautify(c){
    var a=c.a,d=c.d,s=a.Selection,tables=s&&s.Range.Tables,list=[],i;
    if(tables&&Number(tables.Count)){if(!L.ask('将美化选区中的 '+tables.Count+' 张表格。可以按 Ctrl+Z 撤销。是否执行？'))return '已取消。';for(i=1;i<=Number(tables.Count);i++)list.push(tables.Item(i));}
    else{if(!Number(d.Tables.Count))return '没有表格。请把光标放进表格里再执行。';if(!L.ask('光标不在表格里。要美化整篇文档的 '+d.Tables.Count+' 张表格吗？'))return '已取消。';for(i=1;i<=Number(d.Tables.Count);i++)list.push(d.Tables.Item(i));}
    if(list.length>200)throw Error('一次最多处理 200 张表格。');
    var prefs=T.settings().table,bg=rgbHex(prefs.headerBackground),pics=0;
    list.forEach(function(t){t.Borders.Enable=1;var head=t.Rows.Item(1);head.Range.Font.Bold=1;try{head.Shading.BackgroundPatternColor=bg;}catch(ignore){}try{head.HeadingFormat=-1;}catch(ignore){}
      try{t.Range.Cells.VerticalAlignment=1;}catch(ignore){}try{t.Rows.HeightRule=0;}catch(ignore){}try{t.AutoFitBehavior(2);}catch(ignore){}try{t.AllowAutoFit=0;}catch(ignore){}
      var ins=t.Range.InlineShapes;for(var k=1;k<=Number(ins.Count);k++){var p=ins.Item(k),w;try{w=Number(p.Range.Cells.Item(1).Width)-8;}catch(e){continue;}if(w>0&&Number(p.Width)>w){var h=Number(p.Height)*w/Number(p.Width);try{p.LockAspectRatio=-1;}catch(ignore){}p.Width=w;p.Height=h;pics++;}}});
    return '已美化 '+list.length+' 张表格：加全框线、表头加粗加底色并在每页重复、内容上下居中、宽度适应页面并固定列宽（之后插图不会撑宽表格）'+(pics?'，'+pics+' 张超出单元格的图片已按比例缩小':'')+'。';
  }

  // ---------- 图片不乱跑 + 只显示一条 ----------
  function pictureFix(c){
    var a=c.a,d=c.d,sc=scope(a,d,'修复图片位置和显示');if(!sc)return '已取消。';var floating=0,lines=0,i;
    var shapes=sc.all?d.Shapes:(sc.range.ShapeRange||null);
    try{for(i=Number(shapes.Count);i>=1;i--){var sh=shapes.Item(i),tp=Number(sh.Type);if(tp===13||tp===11){sh.ConvertToInlineShape();floating++;}}}catch(ignore){}
    var ins=sc.range.InlineShapes;for(i=1;i<=Number(ins.Count);i++){var f=R(d,ins.Item(i).Range.Paragraphs.Item(1)).ParagraphFormat;var rule=Number(f.LineSpacingRule);if(rule===4||rule===3){f.LineSpacingRule=0;lines++;}}
    if(!floating&&!lines)return '没有需要修复的图片：所选范围里没有浮动图片，图片所在段落也没有“固定行距”。';
    return '已把 '+floating+' 张浮动图片改成“嵌入型”（跟着文字走，不再乱跑），并把 '+lines+' 张图片所在段落的固定行距改成单倍行距（不再只显示一条）。';
  }

  // ---------- 批量加粗行内标题 ----------
  var HEAD=/^(\s*(?:\d{1,2}[、．.]|[一二三四五六七八九十]{1,3}[、．.]|[（(][一二三四五六七八九十\d]{1,3}[）)]|[①-⑳]|第[一二三四五六七八九十\d]{1,3}[条章节点项][、\s]?)\s*[^：:。\r\n]{1,30}?[：:])/;
  function boldHeads(c){
    var a=c.a,d=c.d,sc=scope(a,d,'加粗行内标题',true);if(!sc)return '已取消。';var hits=[];
    paragraphsOf(sc.range).forEach(function(p){var r=p.Range,t=String(r.Text),m=HEAD.exec(t);if(!m||m[1].length>=bare(t).length)return;var x=d.Range(Number(r.Start),Number(r.Start)+m[1].length);if(String(x.Text)!==m[1])return;hits.push(x);});
    if(!hits.length)return '没有找到“1、标题：”这类开头带冒号的行内标题。';
    if(!L.ask('找到 '+hits.length+' 处行内标题（例如“'+String(hits[0].Text).trim()+'”），将把冒号前的部分加粗。是否执行？'))return '已取消。';
    hits.forEach(function(x){x.Font.Bold=1;});return '已加粗 '+hits.length+' 处行内标题（到冒号为止，后面的正文不变）。';
  }

  // ---------- 选项 ABCD 对齐 ----------
  var OPT=/^\s*A\s*[\.．、:：]\s*(.+?)\s*B\s*[\.．、:：]\s*(.+?)\s*C\s*[\.．、:：]\s*(.+?)\s*(?:D\s*[\.．、:：]\s*(.+?)\s*)?$/;
  function width(s){var n=0;for(var i=0;i<s.length;i++)n+=/[\u0000-ÿ]/.test(s[i])?0.5:1;return n;}
  function alignOptions(c){
    var a=c.a,d=c.d,sc=scope(a,d,'对齐选择题选项',true);if(!sc)return '已取消。';var found=[];
    paragraphsOf(sc.range).forEach(function(p){var r=p.Range,t=String(r.Text).replace(/[\r\x07]+$/,''),m=OPT.exec(t);if(m&&!/\t/.test(t))found.push({p:p,opts:[m[1],m[2],m[3]].concat(m[4]?[m[4]]:[])});});
    if(!found.length)return '没有找到“A.… B.… C.… D.…”写在同一行的选项。';
    if(!L.ask('找到 '+found.length+' 组选项，将用制表位对齐（短选项一行四个，长选项一行两个）。选项里的局部格式（如下划线）会被去掉。是否执行？'))return '已取消。';
    var ps;try{ps=a.Selection.Sections.Item(1).PageSetup;}catch(e){ps=d.PageSetup;}var w=Number(ps.PageWidth)-Number(ps.LeftMargin)-Number(ps.RightMargin),letters='ABCD';
    found.forEach(function(f){var max=Math.max.apply(null,f.opts.map(width)),cols=max<=6?f.opts.length:2,parts=f.opts.map(function(o,i){return letters[i]+'. '+o.trim();}),text='';
      parts.forEach(function(x,i){text+=(i===0?'':(i%cols===0?'\u000b':'\t'))+x;});var r=f.p.Range,body=d.Range(Number(r.Start),Number(r.End)-1);body.Text=text;
      var tabs=R(d,f.p).ParagraphFormat.TabStops;try{tabs.ClearAll();}catch(ignore){}for(var k=1;k<cols;k++)tabs.Add(w*k/cols,0);});
    return '已对齐 '+found.length+' 组选项。';
  }

  // ---------- 数字、英文不在中间断行 ----------
  function noBreakLatin(c){var a=c.a,d=c.d,sc=scope(a,d,'数字、英文不断行');if(!sc)return '已取消。';var n=0;paragraphsOf(sc.range).forEach(function(p){try{var f=R(d,p).ParagraphFormat;if(f.WordWrap!==0&&f.WordWrap!==false){f.WordWrap=0;n++;}}catch(ignore){}});
    return n?'已关闭 '+n+' 个段落的“允许西文在单词中间换行”，长数字、网址和英文单词不会再被拆到两行。':'所选段落已经不会在单词中间换行，没有修改。';}

  // ---------- 去水印 ----------
  function removeWatermark(c){
    var a=c.a,d=c.d,n=0,bg=false;
    for(var s=1;s<=Number(d.Sections.Count);s++)for(var k=1;k<=3;k++){var h;try{h=d.Sections.Item(s).Headers.Item(k);}catch(e){continue;}var shapes;try{shapes=h.Shapes;}catch(e){continue;}
      for(var i=Number(shapes.Count);i>=1;i--){var sh=shapes.Item(i),name=String(sh.Name||'');if(/watermark|水印|PowerPlusWaterMark/i.test(name)||Number(sh.Type)===15){sh.Delete();n++;}}}
    try{if(d.Background&&d.Background.Fill&&d.Background.Fill.Visible){d.Background.Fill.Visible=0;bg=true;}}catch(ignore){}
    if(!n&&!bg)return '没有找到水印。如果页面上还有，请用“插入 → 水印 → 删除文档中的水印”，或双击页眉后点选水印删除。';
    return '已删除 '+n+' 个水印'+(bg?'，并关闭了页面背景':'')+'。';
  }

  // ---------- 中英文之间加空格 / 去空格 ----------
  function plainSelection(a,d){var s=a.Selection,r=s&&s.Range;if(!r||!bare(r.Text))throw Error('请先选中要处理的文字（全文可按 Ctrl+A）。');
    if(Number(r.StoryType)!==1||Number(r.Fields&&r.Fields.Count)||d.TrackRevisions)throw Error('只支持未开修订、不含域（如目录、页码）的正文选区，请缩小选区。');if(String(r.Text).length>200000)throw Error('一次最多处理 20 万字，请分段选择。');return r;}
  function cjkSpace(add){return function(c){
    var a=c.a,d=c.d,r=plainSelection(a,d),text=String(r.Text),start=Number(r.Start),pos=[],re;
    if(add){re=new RegExp('(?=(['+CJK+'][A-Za-z0-9]|[A-Za-z0-9%]['+CJK+']))','g');var m;while((m=re.exec(text))){pos.push(m.index+1);re.lastIndex=m.index+1;}}
    else{re=new RegExp('(['+CJK+'])([ \\u3000]+)(?=[A-Za-z0-9])|([A-Za-z0-9%])([ \\u3000]+)(?=['+CJK+'])','g');var x;while((x=re.exec(text))){var s0=x.index+(x[1]||x[3]).length;pos.push([s0,s0+(x[2]||x[4]).length]);}}
    if(!pos.length)return add?'所选文字里中文和英文、数字之间已经有空格（或没有中英文相邻的地方）。':'所选文字里中文和英文、数字之间没有空格。';
    if(pos.length>2000)throw Error('要改的地方超过 2000 处，请分段选择。');
    for(var i=pos.length-1;i>=0;i--){if(add)d.Range(start+pos[i],start+pos[i]).InsertAfter(' ');else d.Range(start+pos[i][0],start+pos[i][1]).Delete();}
    return add?'已在 '+pos.length+' 处中文和英文、数字之间加了空格。':'已去掉 '+pos.length+' 处中文和英文、数字之间的空格。';};}

  // Word 转 PPT：按标题层级整理大纲，交给“PPT 版式工作室”生成新的演示文稿并在 WPS 演示里打开；原文档不变。
  var DRAFT='officeToolbox.studio.draft';
  function readOutline(d){var ps=d.Paragraphs,n=Number(ps.Count)||0;if(n>3000)throw Error('文档超过 3000 段：请先把要做成 PPT 的部分复制到新文档再执行。');var list=[];
    for(var i=1;i<=n;i++){var p=ps.Item(i),r=p.Range,inTable=false,style='';try{inTable=Number(r.Tables.Count)>0;}catch(ignore){}if(inTable)continue;try{style=String(p.Style.NameLocal);}catch(ignore){try{style=String(p.Style);}catch(e){}}
      list.push({text:String(r.Text),level:Number(p.OutlineLevel)||10,style:style});}
    return root.OfficeDocOutline.build(list,{name:String(d.Name||'')});}
  function studioTheme(){try{return root.OfficeI18n&&root.OfficeI18n.lang&&root.OfficeI18n.lang()==='ja'?'wa':'midnight';}catch(e){return 'midnight';}}
  function saveDraft(o,theme){try{root.localStorage.setItem(DRAFT,JSON.stringify({title:o.title,subtitle:o.subtitle,outline:o.outline,theme:theme,at:Date.now()}));}catch(ignore){}}
  function wordToPpt(c){var d=c.d,o=readOutline(d),theme=studioTheme();
    if(!L.ask(root.OfficeDocOutline.summary(o)+'\n\n将生成新的 PPT 并在 WPS 演示里打开（标题：'+o.title+'），原文档不变。是否继续？'))return '已取消。';
    saveDraft(o,theme);
    return L.localFetch('/api/session',{}).then(function(r){if(!r.ok)throw Error('本地服务未连接。');return r.json();}).then(function(session){
      return L.localFetch('/api/studio/open',{method:'POST',headers:{'Content-Type':'application/json','X-Coach-Token':session.token},body:JSON.stringify({title:o.title,subtitle:o.subtitle,outline:o.outline,theme:theme})});
    }).then(function(r){return r.json().catch(function(){return {};}).then(function(data){if(!r.ok)throw Error(data.error||'生成失败，请重试。');
      return '已生成 '+data.pages+' 页 PPT 并在 WPS 演示里打开。\n想换样式或改内容：打开“PPT 版式工作室”，大纲已经填好。';});});}
  function reg(id,label,desc,icon,fn){T.register({id:id,label:label,desc:desc,icon:icon,group:'docfix',ownConfirm:['deleteBlankPages','officialFormatCn','businessFormatJa','landscapePages','fixPageJump','pictureFix','noBreakLatin','tableBeautify','boldInlineHeads','alignOptions','wordToPpt'].indexOf(id)>=0?['wps']:[],hosts:{wps:fn}});}
  reg('deleteBlankPages','删除空白页','只删整页没有内容的页，先确认','blankpage',deleteBlankPages);
  reg('officialFormatCn','公文格式（国标）','GB/T 9704 页边距、字体、行数、页码','official',formatCn);
  reg('businessFormatJa','日文商务文书格式','A4、明朝体 10.5 磅、标题居中、記・以上','jadoc',formatJa);
  reg('fixPageJump','段落乱跳页修复','清除段前分页等导致跳页的设置','pagejump',fixPageJump);
  reg('landscapePages','选中内容改横向','自动加分节符，前后页仍是纵向','landscape',landscape);
  reg('tableBeautify','表格一键美化','框线、表头、居中、插图不撑宽','table',tableBeautify);
  reg('pictureFix','图片不乱跑','浮动图片改嵌入，修复只显示一条','picture',pictureFix);
  reg('boldInlineHeads','加粗行内标题','“1、标题：”冒号前的部分加粗','bold',boldHeads);
  reg('alignOptions','选项 ABCD 对齐','用制表位对齐选择题选项','options',alignOptions);
  reg('noBreakLatin','数字英文不断行','长数字、网址、英文单词不拆两行','nobreak',noBreakLatin);
  reg('removeWatermark','一键去水印','删除页眉里的水印和页面背景','watermark',removeWatermark);
  T.register({id:'wordToPpt',label:'Word 转 PPT',desc:'按标题层级一键生成演示文稿，原文档不变',icon:'presentation',group:'docfix',undo:false,ownConfirm:['wps'],hosts:{wps:wordToPpt}});
  reg('addCjkSpace','中英文间加空格','中文和英文、数字之间加半角空格','spacing',cjkSpace(true));
  reg('removeCjkSpace','去掉中英文空格','删除中文和英文、数字之间的空格','spacing',cjkSpace(false));
  T.DocPlus={scanBlank:scanBlank,HEAD:HEAD,OPT:OPT};
})(window);

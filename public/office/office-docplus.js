// 办公工具箱 · Microsoft Word 版“文档整理”（Office.js）。和 WPS 版同名同规则。
// Office.js 读不到页码、浮动图片和制表位，相关部分改用段落 OOXML 处理，或在结果里说明需要手动完成的步骤。
// 需要确认的工具先返回 {confirm:'提示'}，用户确认后带 options.confirmed=true 再执行。尚未在真实 Microsoft Word 中验证。
(function(root){
  'use strict';
  var CJK='\\u3040-\\u30ff\\u3400-\\u9fff\\uf900-\\ufaff',MM=2.83465;
  function need(c,m){if(!c)throw Error(m);}
  function bare(t){return String(t||'').replace(/[\r\n\x07\x0b\x0c\s　 ​]/g,'');}
  function paras(context,all){var ps=(all?context.document.body:context.document.getSelection()).paragraphs;ps.load('items/text,items/tableNestingLevel,items/style');return ps;}
  async function scope(context,o,what){var sel=context.document.getSelection();sel.load('text');await context.sync();if(bare(sel.text))return {all:false};if(!o.confirmed)return {confirm:'没有选中文字。要对整篇文档执行“'+what+'”吗？'};return {all:true};}
  function heading(p){return /^(Heading|标题|見出し)/i.test(String(p.style||''));}
  // OOXML 段落属性修改：只替换确实需要改的段落。
  async function editPpr(context,list,fn){var ox=list.map(function(p){return p.getOoxml();});await context.sync();var n=0;
    ox.forEach(function(x,i){var out=fn(x.value);if(out!==x.value){list[i].insertOoxml(out,'Replace');n++;}});if(n)await context.sync();return n;}

  // ---------- 删除空白页：Word 版看不到页码，只处理确定会造成空白页的三种情况 ----------
  async function deleteBlankPages(api,o){return api.Word.run(async function(context){
    var ps=context.document.body.paragraphs;ps.load('items/text,items/tableNestingLevel');await context.sync();var items=ps.items,n=items.length;need(n,'文档是空的。');
    var ox=items.map(function(p){return p.getOoxml();});await context.sync();
    function kind(i){var t=items[i].text,x=ox[i].value;if(items[i].tableNestingLevel>0||bare(t))return 'content';if(/<w:(drawing|pict|object)\b/.test(x))return 'content';if(/<w:pPr>(?:(?!<\/w:pPr>)[\s\S])*<w:sectPr\b/.test(x))return 'section';if(/<w:br\b[^>]*w:type="page"/.test(x))return 'break';return 'empty';}
    var k=items.map(function(_,i){return kind(i);}),del=[],shrink=false,lastContent=-1;for(var i=0;i<n;i++)if(k[i]==='content')lastContent=i;
    for(i=lastContent+1;i<n-1;i++)if(k[i]!=='section')del.push(i);                                   // 文末多余的空段落、分页符
    for(i=0;i<=lastContent;i++)if(k[i]==='break'){var j=i+1;while(j<=lastContent&&k[j]==='empty')j++;if(j<=lastContent&&k[j]==='break')for(var x=i;x<j;x++)if(del.indexOf(x)<0)del.push(x);} // 连续两个分页符之间
    if(n>1&&k[n-1]!=='content'&&items[n-2].tableNestingLevel>0)shrink=true;                        // 表格后面那一页
    if(!del.length&&!shrink)return '没有找到可以自动删除的空白页原因（文末空段落、连续分页符、表格后的空段落）。Word 版读不到页码，其他空白页请打开“显示编辑标记”手动检查。';
    if(!o.confirmed)return {confirm:'将删除 '+del.length+' 个造成空白页的空段落或分页符'+(shrink?'，并把表格后面的空段落缩到最小':'')+'。可以按 Ctrl+Z 撤销。是否执行？'};
    del.sort(function(a,b){return b-a;}).forEach(function(i){items[i].delete();});
    if(shrink){var last=items[n-1];last.font.size=1;last.lineSpacing=1;last.spaceAfter=0;last.spaceBefore=0;}
    await context.sync();return '已删除 '+del.length+' 处空段落或分页符'+(shrink?'，并缩小了表格后面的空段落':'')+'。Word 版读不到页码，请翻一遍确认。';});}

  // ---------- 公文格式 / 日文商务文书格式 ----------
  async function margins(context,m){try{var secs=context.document.sections;secs.load('items');await context.sync();secs.items.forEach(function(s){var ps=s.pageSetup;ps.topMargin=m[0]*MM;ps.bottomMargin=m[1]*MM;ps.leftMargin=m[2]*MM;ps.rightMargin=m[3]*MM;});await context.sync();return true;}catch(e){return false;}}
  function titleIndex(items){for(var i=0;i<Math.min(items.length,6);i++){var t=items[i].text.trim();if(!bare(t))continue;return t.length<=40&&!/[。．.；;，,]$/.test(t)?i:-1;}return -1;}
  function setP(p,font,size,bold,align,indent,line){p.font.name=font;p.font.size=size;p.font.bold=!!bold;p.alignment=align;p.firstLineIndent=indent;p.lineSpacing=line;p.spaceBefore=0;p.spaceAfter=0;}
  function format(cn){return function(api,o){return api.Word.run(async function(context){
    if(!o.confirmed)return {confirm:cn?'将按国家标准 GB/T 9704-2012 排版整篇文档：页边距上 37 / 下 35 / 左 28 / 右 26 毫米，正文三号仿宋、每页 22 行，标题和各级小标题换成规定字体。Word 版不会自动加页码。可以按 Ctrl+Z 撤销。是否执行？':'将按日文商务文书的常用格式排版整篇文档：页边距上 35 / 下 30 / 左右 30 毫米，正文明朝体 10.5 磅，标题居中，“記”居中、“以上”靠右。Word 版不会自动加页码。可以按 Ctrl+Z 撤销。是否执行？'};
    var ps=paras(context,true);await context.sync();var items=ps.items,ti=titleIndex(items),n={title:0,head:0,body:0};
    items.forEach(function(p,i){var t=p.text.trim();if(p.tableNestingLevel>0||!bare(t))return;
      if(cn){if(i===ti){setP(p,'方正小标宋简体',22,false,'Centered',0,28.95);n.title++;}else if(/^[一二三四五六七八九十]+、/.test(t)){setP(p,'黑体',16,false,'Justified',32,28.95);n.head++;}else if(/^[（(][一二三四五六七八九十]+[）)]/.test(t)){setP(p,'楷体',16,false,'Justified',32,28.95);n.head++;}else if(/^\d{1,2}[\.．]/.test(t)){setP(p,'仿宋',16,true,'Justified',32,28.95);n.head++;}else{setP(p,'仿宋',16,false,'Justified',32,28.95);n.body++;}}
      else{if(i===ti){setP(p,'游ゴシック',14,true,'Centered',0,18);n.title++;}else if(t==='記'||t==='记'){setP(p,'游明朝',10.5,false,'Centered',0,18);n.body++;}else if(t==='以上'||(i<6&&/^(令和|平成|昭和|\d{4}|[０-９]{4})[^。]{0,14}年[^。]{0,6}月[^。]{0,6}日$/.test(t))){setP(p,'游明朝',10.5,false,'Right',0,18);n.body++;}else if(/^(第[0-9０-９一二三四五六七八九十]+[章条節]|[0-9０-９]{1,2}[\.．、]|[（(][0-9０-９]{1,2}[）)])/.test(t)&&t.length<=40){setP(p,'游ゴシック',10.5,true,'Left',0,18);n.head++;}else{setP(p,'游明朝',10.5,false,'Justified',10.5,18);n.body++;}}});
    await context.sync();var m=await margins(context,cn?[37,35,28,26]:[35,30,30,30]);
    return '已排版：标题 '+n.title+' 个、小标题 '+n.head+' 个、正文 '+n.body+' 段。'+(m?'页边距已设置。':'这个 Word 版本不能用加载项改页边距，请在“布局 → 页边距 → 自定义页边距”里设为 '+(cn?'上 3.7、下 3.5、左 2.8、右 2.6 厘米':'上 3.5、下 3、左右 3 厘米')+'。')+(cn?'没有安装“方正小标宋简体”时，标题会用替代字体显示。':'')+'页码请用“插入 → 页码”添加。';});};}

  // ---------- 段落乱跳页、数字英文不断行、图片只显示一条（段落属性） ----------
  async function pprTool(api,o,what,test,fix,ok,none){return api.Word.run(async function(context){var sc=await scope(context,o,what);if(sc.confirm)return sc;
    var ps=paras(context,sc.all);await context.sync();need(ps.items.length<=2000,'一次最多处理 2000 段，请分段选择。');var list=ps.items.filter(function(p){return test(p);});
    var n=await editPpr(context,list,fix);return n?ok(n):none;});}
  function fixPageJump(api,o){return pprTool(api,o,'修复段落乱跳页',function(p){return !heading(p);},function(x){return x.replace(/<w:(pageBreakBefore|keepNext|keepLines)\b(?![^>]*w:val="(0|false)")[^>]*\/>/g,'');},function(n){return '已清除 '+n+' 个正文段落里的“段前分页 / 与下段同页 / 段中不分页”设置。标题的设置保留。';},'没有发现会造成跳页的段落设置。如果还跳页，请打开“显示编辑标记”看看是不是有分页符。');}
  function noBreakLatin(api,o){return pprTool(api,o,'数字、英文不断行',function(){return true;},function(x){if(/<w:wordWrap w:val="(0|false)"\/>/.test(x))return x;return x.replace(/<w:wordWrap\b[^>]*\/>/g,'').replace(/<w:pPr>/g,'<w:pPr><w:wordWrap w:val="0"/>').replace(/<w:p>(?!<w:pPr>)/g,'<w:p><w:pPr><w:wordWrap w:val="0"/></w:pPr>');},function(n){return '已关闭 '+n+' 个段落的“允许西文在单词中间换行”。';},'所选段落已经不会在单词中间换行，没有修改。');}
  function pictureFix(api,o){return pprTool(api,o,'修复图片显示',function(){return true;},function(x){if(!/<w:drawing\b/.test(x))return x;return x.replace(/(<w:spacing\b[^>]*?)w:lineRule="(exact|atLeast)"([^>]*?)\/>/g,function(m,a,b,c){return (a+'w:lineRule="auto"'+c).replace(/w:line="\d+"/,'w:line="240"')+'/>';});},function(n){return '已把 '+n+' 个图片段落的固定行距改成单倍行距（不再只显示一条）。浮动图片请右键 →“环绕方式 → 嵌入型”，Word 版加载项不能直接修改。';},'没有需要修复的图片段落。浮动图片请右键 →“环绕方式 → 嵌入型”。');}

  // ---------- 选中内容改横向（Word 加载项不能插入分节并改方向） ----------
  function landscapePages(){return Promise.reject(Error('Word 版加载项不能自动改页面方向。请在所选内容前后各插入“布局 → 分隔符 → 下一页”，再把中间这一节设为“布局 → 纸张方向 → 横向”。'));}

  // ---------- 表格一键美化 ----------
  function tableBeautify(api,o){return api.Word.run(async function(context){var t=context.document.getSelection().tables;t.load('items');await context.sync();var list=t.items;
    if(!list.length){var all=context.document.body.tables;all.load('items');await context.sync();need(all.items.length,'没有表格。请把光标放进表格里再执行。');if(!o.confirmed)return {confirm:'光标不在表格里。要美化整篇文档的 '+all.items.length+' 张表格吗？'};list=all.items;}
    var bg=(root.WorkbenchSettings&&root.WorkbenchSettings.snapshot()&&root.WorkbenchSettings.snapshot().settings.table.headerBackground)||'#DDEBF7';
    list.forEach(function(tb){try{['Top','Bottom','Left','Right','InsideHorizontal','InsideVertical'].forEach(function(s){var b=tb.getBorder(s);b.type='Single';b.width=0.5;b.color='#000000';});}catch(e){}
      tb.headerRowCount=1;var head=tb.rows.getFirst();head.shadingColor=bg;head.font.bold=true;try{tb.verticalAlignment='Center';}catch(e){}try{tb.autoFitWindow();}catch(e){}});
    await context.sync();return '已美化 '+list.length+' 张表格：全框线、表头加粗加底色并在每页重复、内容上下居中、宽度适应页面。之后插图不撑宽请在“表格属性 → 选项”里取消“自动重调尺寸以适应内容”。';});}

  // ---------- 加粗行内标题 ----------
  var HEAD=/^(\s*(?:\d{1,2}[、．.]|[一二三四五六七八九十]{1,3}[、．.]|[（(][一二三四五六七八九十\d]{1,3}[）)]|[①-⑳]|第[一二三四五六七八九十\d]{1,3}[条章节点项][、\s]?)\s*[^：:。\r\n]{1,30}?[：:])/;
  function boldInlineHeads(api,o){return api.Word.run(async function(context){var sc=await scope(context,o,'加粗行内标题');if(sc.confirm)return sc;var ps=paras(context,sc.all);await context.sync();
    var hits=[];ps.items.forEach(function(p){var m=HEAD.exec(p.text);if(m&&m[1].length<bare(p.text).length)hits.push({p:p,s:m[1].trim()});});if(!hits.length)return '没有找到“1、标题：”这类开头带冒号的行内标题。';
    var rs=hits.map(function(h){var r=h.p.search(h.s,{matchCase:true});r.load('items');return r;});await context.sync();var n=0;rs.forEach(function(r){if(r.items[0]){r.items[0].font.bold=true;n++;}});await context.sync();return '已加粗 '+n+' 处行内标题（到冒号为止，后面的正文不变）。';});}

  // ---------- 选项 ABCD 对齐（用 OOXML 写制表位） ----------
  var OPT=/^\s*A\s*[\.．、:：]\s*(.+?)\s*B\s*[\.．、:：]\s*(.+?)\s*C\s*[\.．、:：]\s*(.+?)\s*(?:D\s*[\.．、:：]\s*(.+?)\s*)?$/;
  function esc(s){return String(s).replace(/[&<>"]/g,function(c){return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c];});}
  function width(s){var n=0;for(var i=0;i<s.length;i++)n+=/[\u0000-ÿ]/.test(s[i])?0.5:1;return n;}
  function optionXml(opts,tw){var max=Math.max.apply(null,opts.map(width)),cols=max<=6?opts.length:2,tabs='',runs='';for(var k=1;k<cols;k++)tabs+='<w:tab w:val="left" w:pos="'+Math.round(tw*k/cols)+'"/>';
    opts.forEach(function(x,i){if(i)runs+=i%cols===0?'<w:r><w:br/></w:r>':'<w:r><w:tab/></w:r>';runs+='<w:r><w:t xml:space="preserve">'+esc('ABCD'[i]+'. '+x.trim())+'</w:t></w:r>';});
    return '<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage"><pkg:part pkg:name="/_rels/.rels" pkg:contentType="application/vnd.openxmlformats-package.relationships+xml"><pkg:xmlData><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships></pkg:xmlData></pkg:part><pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"><pkg:xmlData><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body><w:p><w:pPr><w:tabs>'+tabs+'</w:tabs></w:pPr>'+runs+'</w:p></w:body></w:document></pkg:xmlData></pkg:part></pkg:package>';}
  function alignOptions(api,o){return api.Word.run(async function(context){var sc=await scope(context,o,'对齐选择题选项');if(sc.confirm)return sc;var ps=paras(context,sc.all);await context.sync();
    var found=[];ps.items.forEach(function(p){var t=p.text,m=OPT.exec(t);if(m&&t.indexOf('\t')<0)found.push({p:p,opts:[m[1],m[2],m[3]].concat(m[4]?[m[4]]:[])});});if(!found.length)return '没有找到“A.… B.… C.… D.…”写在同一行的选项。';
    if(!sc.all&&!o.confirmed)return {confirm:'找到 '+found.length+' 组选项，将用制表位对齐（短选项一行四个，长选项一行两个）。选项里的局部格式会被去掉。是否执行？'};
    found.forEach(function(f){f.p.insertOoxml(optionXml(f.opts,9000),'Replace');});await context.sync();return '已对齐 '+found.length+' 组选项。';});}

  // ---------- 去水印（页眉 OOXML） ----------
  function removeWatermark(api){return api.Word.run(async function(context){var secs=context.document.sections;secs.load('items');await context.sync();var hs=[];
    secs.items.forEach(function(s){['Primary','FirstPage','EvenPages'].forEach(function(k){var h=s.getHeader(k);hs.push({h:h,x:h.getOoxml()});});});await context.sync();var n=0;
    hs.forEach(function(e){var x=e.x.value,c=0,out=x.replace(/<w:r\b(?:(?!<\/w:r>)[\s\S])*?(?:PowerPlusWaterMarkObject|WordPictureWatermark|[Ww]ater[Mm]ark|水印)(?:(?!<\/w:r>)[\s\S])*?<\/w:r>/g,function(){c++;return '';});if(c){e.h.insertOoxml(out,'Replace');n+=c;}});
    if(!n)return '没有找到水印。请用“设计 → 水印 → 删除水印”，或双击页眉后点选水印删除。';await context.sync();return '已删除 '+n+' 个水印。';});}

  // ---------- 中英文之间加 / 去空格 ----------
  function cjkSpace(add){return function(api){return api.Word.run(async function(context){var sel=context.document.getSelection();sel.load('text');await context.sync();need(bare(sel.text),'请先选中要处理的文字（全文可按 Ctrl+A）。');
    var pats=add?['[一-龥ぁ-ヿ][A-Za-z0-9]','[A-Za-z0-9%][一-龥ぁ-ヿ]']:['[一-龥ぁ-ヿ][ 　][A-Za-z0-9]','[A-Za-z0-9%][ 　][一-龥ぁ-ヿ]'];
    var rs=pats.map(function(p){var r=sel.search(p,{matchWildcards:true});r.load('items/text');return r;});await context.sync();var hits=[];rs.forEach(function(r){hits=hits.concat(r.items);});
    if(!hits.length)return add?'所选文字里中文和英文、数字之间已经有空格（或没有中英文相邻的地方）。':'所选文字里中文和英文、数字之间没有空格。';need(hits.length<=2000,'要改的地方超过 2000 处，请分段选择。');
    hits.forEach(function(r){var t=r.text;r.insertText(add?t[0]+' '+t.slice(1):t[0]+t.slice(-1),'Replace');});await context.sync();
    return add?'已在 '+hits.length+' 处中文和英文、数字之间加了空格。':'已去掉 '+hits.length+' 处中文和英文、数字之间的空格。';});};}

  // Word 转 PPT：按标题层级整理大纲，交给本机“PPT 版式工作室”生成并打开（装了 WPS 用 WPS 演示，否则用默认程序）。
  function styleLevel(style){var m=String(style||'').match(/^(?:Heading|标题|見出し)\s*([1-9])$/i);return m?Number(m[1]):10;}
  function wordToPpt(api,o){o=o||{};return api.Word.run(async function(context){var ps=context.document.body.paragraphs;ps.load('items/text,items/outlineLevel,items/style,items/tableNestingLevel');await context.sync();
    need(ps.items.length<=3000,'文档超过 3000 段：请先把要做成 PPT 的部分复制到新文档再执行。');
    var list=ps.items.filter(function(p){return !p.tableNestingLevel;}).map(function(p){var lv=Number(p.outlineLevel);if(!(lv>=1&&lv<=9))lv=styleLevel(p.style);return {text:p.text,level:lv,style:p.style};});
    var name='';try{var props=context.document.properties;props.load('title');await context.sync();name=String(props.title||'');}catch(ignore){}
    var D=root.OfficeDocOutline,out=D.build(list,{name:name});
    if(!o.confirmed)return {confirm:D.summary(out)+'\n\n将生成新的 PPT 并打开（标题：'+out.title+'），原文档不变。是否继续？'};
    var theme=root.OfficeI18n&&root.OfficeI18n.lang&&root.OfficeI18n.lang()==='ja'?'wa':'midnight';
    try{root.localStorage.setItem('officeToolbox.studio.draft',JSON.stringify({title:out.title,subtitle:out.subtitle,outline:out.outline,theme:theme,at:Date.now()}));}catch(ignore){}
    var session=await (await root.fetch('/api/session')).json();
    var r=await root.fetch('/api/studio/open',{method:'POST',headers:{'Content-Type':'application/json','X-Coach-Token':session.token},body:JSON.stringify({title:out.title,subtitle:out.subtitle,outline:out.outline,theme:theme})}),data=await r.json().catch(function(){return {};});
    if(!r.ok)throw Error(data.error||'生成失败，请重试。');
    return '已生成 '+data.pages+' 页 PPT 并打开。\n想换样式或改内容：打开“PPT 版式工作室”，大纲已经填好。';});}
  var tools={wordToPpt:wordToPpt,deleteBlankPages:deleteBlankPages,officialFormatCn:format(true),businessFormatJa:format(false),fixPageJump:fixPageJump,landscapePages:landscapePages,tableBeautify:tableBeautify,pictureFix:pictureFix,boldInlineHeads:boldInlineHeads,alignOptions:alignOptions,noBreakLatin:noBreakLatin,removeWatermark:removeWatermark,addCjkSpace:cjkSpace(true),removeCjkSpace:cjkSpace(false)};
  var list=[['deleteBlankPages','删除空白页','只删整页没有内容的页，先确认'],['officialFormatCn','公文格式（国标）','GB/T 9704 页边距、字体、行数、页码'],['businessFormatJa','日文商务文书格式','A4、明朝体 10.5 磅、标题居中、記・以上'],['fixPageJump','段落乱跳页修复','清除段前分页等导致跳页的设置'],['landscapePages','选中内容改横向','自动加分节符，前后页仍是纵向'],['tableBeautify','表格一键美化','框线、表头、居中、插图不撑宽'],['pictureFix','图片不乱跑','浮动图片改嵌入，修复只显示一条'],['boldInlineHeads','加粗行内标题','“1、标题：”冒号前的部分加粗'],['alignOptions','选项 ABCD 对齐','用制表位对齐选择题选项'],['noBreakLatin','数字英文不断行','长数字、网址、英文单词不拆两行'],['removeWatermark','一键去水印','删除页眉里的水印和页面背景'],['wordToPpt','Word 转 PPT','按标题层级一键生成演示文稿，原文档不变'],['addCjkSpace','中英文间加空格','中文和英文、数字之间加半角空格'],['removeCjkSpace','去掉中英文空格','删除中文和英文、数字之间的空格']];
  root.OfficeDocPlus={tools:tools,list:list,HEAD:HEAD,OPT:OPT};
})(typeof window!=='undefined'?window:globalThis);

// 办公工具箱 · Microsoft Office 版工具（Office.js）。和 WPS 版同名同规则，只是改成 Office.js 的“先排队、再 sync”写法。
// 每个工具：run(api, options) → 结果文字；需要确认的工具先返回 {confirm:'提示'}，用户点确认后带 options.confirmed=true 再执行。
// api = {Excel, Word}（测试时传入模拟对象）。尚未在真实 Microsoft Office 中验证。
(function(root){
  'use strict';
  var CM=28.3465;
  function blank(v,f){return (v===undefined||v===null||v==='')&&!(typeof f==='string'&&f.charAt(0)==='=');}
  function isFormula(f){return typeof f==='string'&&f.charAt(0)==='=';}
  function need(cond,msg){if(!cond)throw Error(msg);}
  // ---- Excel ----
  function selected(context,props){var r=context.workbook.getSelectedRange();r.load(props||'rowCount,columnCount,address');return r;}
  function limit(r,max){need(r.rowCount*r.columnCount<=(max||2000),'一次最多处理 '+(max||2000)+' 个单元格，请缩小选区。');}
  var excel={
    formatHeader:function(api){return api.Excel.run(async function(context){var r=selected(context);await context.sync();var h=r.getRow(0);h.format.font.bold=true;h.format.font.color='#1F3864';h.format.fill.color='#DDEBF7';await context.sync();return '已美化选区第一行。';});},
    fitColumns:function(api){return api.Excel.run(async function(context){var r=selected(context);await context.sync();r.format.autofitColumns();await context.sync();return '已按选区内容调整列宽。';});},
    wrapCells:function(api){return api.Excel.run(async function(context){var r=selected(context);await context.sync();r.format.wrapText=true;r.format.autofitRows();await context.sync();return '已开启自动换行并调整行高。';});},
    trimCells:function(api){return api.Excel.run(async function(context){var r=selected(context,'rowCount,columnCount,values,formulas');await context.sync();limit(r);var n=0;
      for(var i=0;i<r.rowCount;i++)for(var j=0;j<r.columnCount;j++){var v=r.values[i][j];if(typeof v!=='string'||isFormula(r.formulas[i][j]))continue;var t=v.replace(/^[\s　]+|[\s　]+$/g,'');if(t!==v){r.getCell(i,j).values=[[t]];n++;}}
      await context.sync();return n?'已清理 '+n+' 个单元格的首尾空格，公式和数值不变。':'没有需要清理的文本单元格。';});},
    centerAcross:function(api){return api.Excel.run(async function(context){var r=selected(context,'rowCount,columnCount,values,formulas');await context.sync();
      need(r.columnCount>=2,'请横向选择标题所在的一行及右侧要跨过的空白单元格，例如 A1:F1。');limit(r);
      for(var i=0;i<r.rowCount;i++){var filled=0;for(var j=0;j<r.columnCount;j++)if(!blank(r.values[i][j],r.formulas[i][j]))filled++;need(filled<=1,'第 '+(i+1)+' 行选区里有多个非空单元格。跨列居中只适合“左边一格标题、右边都空”，不会清空其他内容。');}
      r.format.horizontalAlignment='CenterAcrossSelection';await context.sync();return '已设置跨列居中：标题看起来居中，单元格仍各自独立，可单独选取、排序和筛选。';});},
    shrinkToFit:function(api){return api.Excel.run(async function(context){var r=selected(context);r.format.load('wrapText');await context.sync();var wrapped=r.format.wrapText===true;if(wrapped)r.format.wrapText=false;r.format.shrinkToFit=true;await context.sync();
      return '已开启缩小字体填充，单元格的值不变。'+(wrapped?'原来的自动换行已关闭（两者不能同时生效）。':'')+'文字很长时字会很小，请检查可读性或改用自动换行。';});},
    fillBlanks:function(api,o){return api.Excel.run(async function(context){var r=selected(context,'rowCount,columnCount,values,formulas');await context.sync();limit(r);var fills=[],edge=0;
      for(var j=0;j<r.columnCount;j++){var last;for(var i=0;i<r.rowCount;i++){if(blank(r.values[i][j],r.formulas[i][j])){if(last===undefined)edge++;else fills.push([i,j,last]);}else last=r.values[i][j];}last=undefined;}
      if(!fills.length)return '选区里没有需要填充的真正空白单元格。只含空格的单元格和公式不算空白。';
      if(!(o&&o.confirmed))return {confirm:'将把 '+fills.length+' 个空白单元格填成同列上方最近的内容'+(edge?'（每列顶端的 '+edge+' 个空白没有上一格，会跳过）':'')+'。\n请先确认空白确实表示“同上”，而不是缺失或 0。是否执行？'};
      fills.forEach(function(f){r.getCell(f[0],f[1]).values=[[f[2]]];});await context.sync();return '已填充 '+fills.length+' 个空白单元格（填入的是值，不是公式）。';});},
    removeDuplicates:function(api,o){return api.Excel.run(async function(context){var r=selected(context,'rowCount,columnCount');await context.sync();need(r.rowCount>=2,'请连同表头选择至少两行数据。');limit(r,100000);
      if(!(o&&o.confirmed))return {confirm:'将按整行内容删除选区中的重复行（第一行视为表头，保留第一次出现的行）。重要文件请先另存一份。是否执行？'};
      var cols=[];for(var j=0;j<r.columnCount;j++)cols.push(j);var res=r.removeDuplicates(cols,true);res.load('removed,uniqueRemaining');await context.sync();return '已删除 '+res.removed+' 行重复数据，保留 '+res.uniqueRemaining+' 行。';});},
    filterCells:function(api){return api.Excel.run(async function(context){var r=selected(context);var ws=context.workbook.worksheets.getActiveWorksheet();var af=ws.autoFilter;af.load('enabled');await context.sync();if(af.enabled){af.remove();await context.sync();return '已关闭筛选按钮。';}af.apply(r);await context.sync();return '已打开筛选按钮。';});},
    replaceText:function(api,o){var find=String(o&&o.find||''),rep=String(o&&o.replace||'');if(!find)return Promise.reject(Error('请填写要查找的文字。'));return api.Excel.run(async function(context){var r=selected(context,'rowCount,columnCount,values,formulas');await context.sync();limit(r);var n=0;
      for(var i=0;i<r.rowCount;i++)for(var j=0;j<r.columnCount;j++){var v=r.values[i][j];if(typeof v!=='string'||isFormula(r.formulas[i][j])||v.indexOf(find)<0)continue;var t=v.split(find).join(rep);r.getCell(i,j).values=[[t.charAt(0)==='='?"'"+t:t]];n++;}
      await context.sync();return n?'已替换 '+n+' 个单元格中的文字，公式和数值不变。':'选区里没有找到要替换的文字。';});},
    fitPicturesToCells:function(api){return api.Excel.run(async function(context){var r=selected(context,'rowCount,columnCount');var sheet=context.workbook.worksheets.getActiveWorksheet();var shapes=sheet.shapes;shapes.load('items/type,items/left,items/top,items/width,items/height');await context.sync();limit(r);
      var pics=shapes.items.filter(function(s){return s.type==='Image';});need(pics.length,'当前工作表没有浮动图片。');
      var cells=[];for(var i=0;i<r.rowCount;i++)for(var j=0;j<r.columnCount;j++){var c=r.getCell(i,j);c.load('left,top,width,height');cells.push(c);}await context.sync();
      var done=0,pad=2;pics.forEach(function(s){var box=null;for(var k=0;k<cells.length;k++){var c=cells[k];if(s.left>=c.left-0.5&&s.left<c.left+c.width&&s.top>=c.top-0.5&&s.top<c.top+c.height){box=c;break;}}if(!box)return;
        var bw=box.width-pad*2,bh=box.height-pad*2;if(!(bw>0&&bh>0))return;var k2=Math.min(bw/s.width,bh/s.height),w=s.width*k2,h=s.height*k2;
        try{s.lockAspectRatio=true;}catch(e){}s.width=w;s.height=h;s.left=box.left+(box.width-w)/2;s.top=box.top+(box.height-h)/2;try{s.placement='TwoCell';}catch(e){}done++;});
      await context.sync();return done?'已调整 '+done+' 张图片：按比例缩放到所在单元格内并居中。':'选区里没有左上角落在其中的图片。请连同图片所在的单元格一起选中。';});}
  };
  // ---- Word ----
  function wordPictures(context){var list=context.document.getSelection().inlinePictures;list.load('items/width,items/height');return list;}
  async function pageWidth(context){try{var s=context.document.sections.getFirst();var ps=s.pageSetup;ps.load('pageWidth,leftMargin,rightMargin');await context.sync();var w=ps.pageWidth-ps.leftMargin-ps.rightMargin;if(w>0)return {w:w,exact:true};}catch(e){}return {w:425,exact:false};}
  function resize(p,w){var h=p.height*w/p.width;try{p.lockAspectRatio=true;}catch(e){}p.width=w;p.height=h;}
  function styleTool(style,label){return function(api){return api.Word.run(async function(context){var ps=context.document.getSelection().paragraphs;ps.load('items');await context.sync();need(ps.items.length,'请先在正文里选中要处理的文字。');ps.items.forEach(function(p){p.styleBuiltIn=style;});await context.sync();return '已把 '+ps.items.length+' 段设为'+label+'。';});};}
  var word={
    heading1:styleTool('Heading1','一级标题'),heading2:styleTool('Heading2','二级标题'),bodyText:styleTool('Normal','正文'),
    cleanEmptyParagraphs:function(api){return api.Word.run(async function(context){var ps=context.document.getSelection().paragraphs;ps.load('items/text');await context.sync();need(ps.items.length,'请先在正文里选中要处理的文字。');
      var empty=ps.items.filter(function(p){return !String(p.text||'').replace(/[\s　 ]/g,'');});if(empty.length===ps.items.length)empty=empty.slice(1);empty.forEach(function(p){p.delete();});await context.sync();return empty.length?'已删除 '+empty.length+' 个空段落。':'选区里没有空段落。';});},
    replaceText:function(api,o){var find=String(o&&o.find||''),rep=String(o&&o.replace||'');if(!find)return Promise.reject(Error('请填写要查找的文字。'));need(find.length<=255,'查找文字最多 255 个字符。');return api.Word.run(async function(context){var hits=context.document.getSelection().search(find,{matchCase:true});hits.load('items');await context.sync();
      hits.items.forEach(function(h){h.insertText(rep,'Replace');});await context.sync();return hits.items.length?'已替换 '+hits.items.length+' 处。':'选区里没有找到要替换的文字。';});},
    textToTable:function(api){return api.Word.run(async function(context){var sel=context.document.getSelection();sel.load('text');await context.sync();var lines=String(sel.text||'').split(/\r\n?|\n/).map(function(l){return l.replace(/\s+$/,'');}).filter(Boolean);
      need(lines.length,'请先选中要转换的文字，每行一条，用 Tab 或逗号分隔。');var sep=lines.some(function(l){return l.indexOf('\t')>=0;})?'\t':/[,，]/;var rows=lines.map(function(l){return l.split(sep).map(function(x){return x.trim();});});var cols=Math.max.apply(null,rows.map(function(r){return r.length;}));
      need(rows.length<=200&&cols<=20,'一次最多转换 200 行、20 列。');rows=rows.map(function(r){while(r.length<cols)r.push('');return r;});sel.insertTable(rows.length,cols,'After',rows);await context.sync();return '已在选区后插入 '+rows.length+' 行 × '+cols+' 列表格，原文字保留。';});},
    unifyPictures:function(api){return api.Word.run(async function(context){var list=wordPictures(context);await context.sync();var pics=list.items.filter(function(p){return p.width>0&&p.height>0;});
      need(pics.length,'选区里没有嵌入型图片。要处理全文，可先按 Ctrl+A。');need(pics.length>=2,'选区里只有 1 张图片。请选中至少 2 张图片（可按 Ctrl+A 选全文）。');need(pics.length<=500,'一次最多处理 500 张图片，请缩小选区。');
      var w=pics[0].width;pics.forEach(function(p){resize(p,w);});await context.sync();return '已把 '+pics.length+' 张嵌入型图片统一为第一张的宽度（约 '+(w/CM).toFixed(1)+' 厘米），高度按各自比例变化，不会变形。';});},
    fitPictures:function(api){return api.Word.run(async function(context){var list=wordPictures(context);await context.sync();var pics=list.items;need(pics.length,'选区里没有嵌入型图片。要处理全文，可先按 Ctrl+A。');var pw=await pageWidth(context),n=0;
      pics.forEach(function(p){if(p.width>pw.w+0.5){resize(p,pw.w);n++;}});await context.sync();var note=pw.exact?'':'（读不到页面设置，按 A4、左右边距 3 厘米计算）';
      return n?'已把 '+n+' 张超出版心的图片缩到 '+(pw.w/CM).toFixed(1)+' 厘米宽，比例不变。'+note:'选区里的图片都没有超出版心（'+(pw.w/CM).toFixed(1)+' 厘米），没有修改。'+note;});},
    // 与 WPS 版相同：已有目录就更新，没有就检查一至三级标题后在选区末尾插入（需要 Word 2021 / Microsoft 365 的域接口）。
    toc:function(api){return api.Word.run(async function(context){var body=context.document.body,fields=body.fields;fields.load('items/type');await context.sync();
      var tocs=fields.items.filter(function(f){return String(f.type).toUpperCase()==='TOC';});if(tocs.length){tocs.forEach(function(f){f.updateResult();});await context.sync();return '目录已更新。';}
      var ps=body.paragraphs;ps.load('items/styleBuiltIn');await context.sync();var has=ps.items.slice(0,2000).some(function(p){return /^Heading[123]$/.test(String(p.styleBuiltIn));});need(has,'未找到一至三级标题，请先把章节设为标题。');
      var f=context.document.getSelection().insertField('End','TOC','\\o "1-3" \\h \\z \\u',false);await context.sync();try{f.updateResult();await context.sync();}catch(e){}return '已在选区末尾插入目录。';});},
    // 与 WPS 版相同：参照选区前后最近的正文段落，套用它的样式、字体、字号、颜色和段落间距。
    matchFormat:function(api){return api.Word.run(async function(context){var sel=context.document.getSelection(),ps=sel.paragraphs;ps.load('items');await context.sync();need(ps.items.length,'文档里没有可以参照的周围文字。');
      var around=[ps.items[0].getPreviousOrNullObject(),ps.items[ps.items.length-1].getNextOrNullObject()];around.forEach(function(p){p.load('isNullObject,text,styleBuiltIn,style,lineSpacing,spaceAfter,spaceBefore,firstLineIndent');});await context.sync();
      var ok=around.filter(function(p){return !p.isNullObject;});need(ok.length,'文档里没有可以参照的周围文字。');var ref=ok.filter(function(p){return String(p.text).trim()&&!/^(Heading|Title)/.test(String(p.styleBuiltIn));})[0]||ok[0];
      var font=ref.font;font.load('name,size,color');sel.load('text');await context.sync();
      ps.items.forEach(function(p){try{p.style=ref.style;}catch(e){}p.lineSpacing=ref.lineSpacing;p.spaceAfter=ref.spaceAfter;p.spaceBefore=ref.spaceBefore;p.firstLineIndent=ref.firstLineIndent;});
      if(font.name)sel.font.name=font.name;if(font.size)sel.font.size=font.size;if(font.color)sel.font.color=font.color;await context.sync();
      return '已按周围正文设置'+(font.name?'字体“'+font.name+'”':'')+(font.size?'、'+font.size+' 号':'')+'。';});},
    idCardSize:function(api){return api.Word.run(async function(context){var list=wordPictures(context);await context.sync();var pics=list.items;need(pics.length,'选区里没有嵌入型图片。');need(pics.length<=2,'一次最多处理 2 张（正面和反面）。请只选中证件图片。');
      var bad=pics.filter(function(p){var r=p.width/p.height;return Math.abs(r-8.56/5.4)/(8.56/5.4)>0.06;});need(!bad.length,'有 '+bad.length+' 张图片的长宽比和身份证不符（应约为 1.59:1）。请先裁剪多余背景，再执行。');
      pics.forEach(function(p){try{p.lockAspectRatio=false;}catch(e){}p.width=8.56*CM;p.height=5.4*CM;});await context.sync();return '已把 '+pics.length+' 张图片设为 8.56×5.4 厘米。打印时请选“实际大小”，印出来后用尺子量一下。';});}
  };
  var ppt={
    makeSlides:function(api,o){var text=String(o&&o.text||'').trim();if(!text)return Promise.reject(Error('请先在下面的框里输入大纲：每页第一行是标题，空一行开始下一页。'));
      var pages=text.split(/\n\s*\n/).map(function(b){var l=b.split('\n').map(function(x){return x.trim();}).filter(Boolean);return {title:l[0]||'',body:l.slice(1)};}).filter(function(p){return p.title;});need(pages.length<=50,'一次最多生成 50 页。');
      return api.PowerPoint.run(async function(context){var slides=context.presentation.slides;slides.load('items');await context.sync();var before=slides.items.length;pages.forEach(function(){slides.add();});await context.sync();slides.load('items');await context.sync();
        pages.forEach(function(p,i){var sl=slides.items[before+i],sh=sl.shapes;var t=sh.addTextBox(p.title,{left:48,top:36,width:860,height:70});t.textFrame.textRange.font.size=30;t.textFrame.textRange.font.bold=true;t.textFrame.textRange.font.color='#1F2A44';
          if(p.body.length){var b=sh.addTextBox(p.body.map(function(x){return '• '+x;}).join('\n'),{left:60,top:130,width:840,height:360});b.textFrame.textRange.font.size=20;b.textFrame.textRange.font.color='#2F3B52';}});
        await context.sync();return '已在末尾生成 '+pages.length+' 页幻灯片。';});},
    // 与 WPS 版相同：所有文字框统一为本页最常用的字体，正文统一为最常见字号，标题（唯一最大字号或标题占位符）保持原大小。
    unifyDeckFont:function(api,o){o=o||{};return api.PowerPoint.run(async function(context){var slides=context.presentation.slides;slides.load('items');await context.sync();var all=[];
      slides.items.forEach(function(sl){sl.shapes.load('items/type');});await context.sync();
      slides.items.forEach(function(sl){sl.shapes.items.forEach(function(s){if(TEXT_SHAPES.indexOf(String(s.type))<0)return;var tf=s.textFrame;tf.load('hasText');var font=tf.textRange.font;font.load('name');all.push({tf:tf,font:font});});});await context.sync();
      var frames=all.filter(function(x){return x.tf.hasText;});need(frames.length,'演示文稿里没有文字。');var count={};frames.forEach(function(x){var n=String(x.font.name||'');if(n)count[n]=(count[n]||0)+1;});
      var names=Object.keys(count).sort(function(a,b){return count[b]-count[a];}),want=String(o.font||'').trim()||names[0]||'Microsoft YaHei';
      if(names.length===1&&names[0]===want)return '全部 '+frames.length+' 处文字已经是“'+want+'”，不需要修改。';
      if(!o.confirmed)return {confirm:'全部 '+slides.items.length+' 页共 '+frames.length+' 个文字框，用了 '+names.length+' 种字体：'+names.slice(0,6).map(function(n){return n+' ×'+count[n];}).join('、')+'。\n将全部统一为“'+want+'”，字号和颜色不变（表格和组合里的文字 PowerPoint 插件改不到）。是否执行？'};
      frames.forEach(function(x){x.font.name=want;});await context.sync();return '已把 '+frames.length+' 个文字框统一为“'+want+'”。表格和组合里的文字请用“开始 → 替换 → 替换字体”。';});},
    removeAnimations:function(){return Promise.reject(Error('PowerPoint 插件不能删除动画。手动做法：每页“动画 → 动画窗格”，按 Ctrl+A 全选后按 Delete；切换效果在“切换”里选“无”，再点“应用到全部”。'));},
    arrangeVisuals:function(){return Promise.reject(Error('PowerPoint 插件不能自动排图片。手动做法：按住 Ctrl 选中几张图片 →“图片格式 → 对齐”，先选“顶端对齐”，再选“横向分布”；也可以用“设计 → 设计器”套用推荐版式。'));},
    unifyTitles:function(){return Promise.reject(Error('PowerPoint 插件不能统一标题。手动做法：“视图 → 幻灯片母版”，在最上面的母版里改标题的字体、字号和位置，所有用标题占位符的页会一起变。'));},
    slideNumbers:function(){return Promise.reject(Error('PowerPoint 插件不能加页码。手动做法：“插入 → 页眉和页脚”，勾选“幻灯片编号”和“标题幻灯片中不显示”，点“全部应用”。'));},
    bodySpacing:function(){return Promise.reject(Error('PowerPoint 插件不能统一行距。手动做法：“视图 → 幻灯片母版”，选中正文占位符，在“开始 → 行距”里设置；或在大纲视图按 Ctrl+A 全选后设置行距。'));},
    applyDesign:function(){return Promise.reject(Error('PowerPoint 插件不能套用模板。手动做法：“设计 → 主题”右边的下拉箭头 → “浏览主题”，选择模板文件。'));},
    removeNotes:function(){return Promise.reject(Error('PowerPoint 插件不能删除备注。手动做法：“文件 → 信息 → 检查问题 → 检查文档”，勾选“演示文稿备注”，检查后点“全部删除”（建议先另存一份）。'));},
    formatSlide:function(api){return api.PowerPoint.run(async function(context){var sel=context.presentation.getSelectedSlides();sel.load('items');await context.sync();need(sel.items.length,'请先在左侧点一下要整理的幻灯片。');
      var shapes=sel.items[0].shapes;shapes.load('items/type,items/name');await context.sync();var frames=[];
      var placeholderApi=!!(api.Office&&api.Office.context&&api.Office.context.requirements&&api.Office.context.requirements.isSetSupported('PowerPointApi','1.8'));
      shapes.items.forEach(function(s){if(TEXT_SHAPES.indexOf(String(s.type))<0)return;var tf=s.textFrame;tf.load('hasText');var font=tf.textRange.font;font.load('size,name');var ph=null;if(placeholderApi&&s.type==='Placeholder'){ph=s.placeholderFormat;ph.load('type');}frames.push({shape:s,tf:tf,font:font,placeholder:ph});});
      await context.sync();frames=frames.filter(function(f){return f.tf.hasText;});need(frames.length,'当前页没有可设置的文字框。');
      var sizes=frames.map(function(f){return Number(f.font.size)||0;}),largest=Math.max.apply(null,sizes),titles=frames.map(function(f,k){return (f.placeholder?/^(Title|CenterTitle|VerticalTitle)$/.test(f.placeholder.type):f.shape.type==='Placeholder')||/title|headline|heading|标题|タイトル/i.test(String(f.shape.name))||(frames.length>1&&sizes[k]===largest&&sizes.filter(function(x){return x===largest;}).length===1);});
      var counts={},body=0,names={},font='';sizes.forEach(function(sz,k){if(!titles[k]&&sz)counts[sz]=(counts[sz]||0)+1;});Object.keys(counts).forEach(function(sz){if(!body||counts[sz]>counts[body])body=Number(sz);});
      frames.forEach(function(f){var n=String(f.font.name||'');if(n)names[n]=(names[n]||0)+1;});Object.keys(names).forEach(function(n){if(!font||names[n]>names[font])font=n;});
      frames.forEach(function(f,k){var r=f.tf.textRange.font;if(font)r.name=font;if(!titles[k]&&body)r.size=body;});await context.sync();var kept=titles.filter(Boolean).length;
      return '已统一当前页 '+frames.length+' 个文字框'+(font?'为“'+font+'”':'')+(body?'，正文统一为 '+body+' 号':'')+(kept?'，标题保持原有大小':'')+'。';});},
    duplicateSlide:function(api){return api.PowerPoint.run(async function(context){var sel=context.presentation.getSelectedSlides();sel.load('items/id');await context.sync();need(sel.items.length,'请先在左侧点一下要复制的幻灯片。');
      var slide=sel.items[0],data=slide.exportAsBase64();await context.sync();context.presentation.insertSlidesFromBase64(data.value,{targetSlideId:slide.id,formatting:'KeepSourceFormatting'});await context.sync();return '已在当前页之后插入一份副本，原页不变。';});}
  };
  var TEXT_SHAPES=['GeometricShape','TextBox','Placeholder','Callout'];
  var LIST={
    excel:[['formatHeader','美化表头','突出选区第一行'],['fitColumns','自动列宽','按内容调整宽度'],['wrapCells','自动换行','让单元格内容完整显示'],['trimCells','清理空格','移除文本首尾空白'],['removeDuplicates','删除重复行','重复行只保留第一行'],['centerAcross','跨列居中','标题居中但不合并单元格'],['shrinkToFit','缩小字体填充','不改列宽也能显示完整'],['fillBlanks','空白填上一格','合并拆开后、分类留空时补齐'],['fitPicturesToCells','图片适应单元格','选区内图片缩放居中、随单元格移动'],['filterCells','筛选开关','快速筛出所需数据'],['replaceText','查找替换','只处理选区内的文字','replace'],['backupFile','备份当前文件','复制一份带时间的备份'],['exportPdf','导出 PDF','使用当前版式生成新文件'],['rmbUpper','金额转人民币大写','1234.5 → 壹仟贰佰叁拾肆元伍角'],['idCardInfo','身份证提取生日性别','校验号码，写出生日期和性别'],['phoneCheck','检查手机号','位数、号段是否正确'],['zenHan','全角半角统一','英数字半角、片假名全角'],['toWareki','西历转和历','2026/11/12 → 令和8年11月12日'],['toSeireki','和历转西历','令和8年11月12日 → 2026/11/12'],['postalFormat','邮编格式统一','日本 1234567 → 123-4567'],['furigana','加注假名（读音）','用 PHONETIC 显示读音']],
    word:[['heading1','一级标题','设为标题 1'],['heading2','二级标题','设为标题 2'],['bodyText','正文排版','设为正文样式'],['unifyPictures','图片统一大小','以第一张为准，按比例统一宽度'],['fitPictures','图片不超出页面','太宽的图片按比例缩到版心'],['idCardSize','证件照按实际尺寸','身份证等卡片设为 8.56×5.4 厘米'],['cleanEmptyParagraphs','清理空段落','删除选区内的空行'],['toc','插入 / 更新目录','根据标题生成目录'],['matchFormat','匹配周围格式','拖拽或粘贴后跟随周围正文'],['textToTable','文字转表格','把选中的文字整理成表格'],['replaceText','查找替换','只处理选区内的文字','replace'],['backupFile','备份当前文件','复制一份带时间的备份'],['exportPdf','导出 PDF','使用当前版式生成新文件']],
    powerpoint:[['makeSlides','按大纲生成 PPT','每页第一行是标题，空一行分页','outline'],['formatSlide','统一本页字体','正文统一、标题保持大小'],['unifyDeckFont','全部页统一字体','所有页的文字框用同一种字体，字号不变',null,[{key:'font',label:'字体（不填就用现在最常见的那种）',placeholder:'例如：微软雅黑',required:false}]],['removeAnimations','删除全部动画','删除所有页的动画和切换效果'],['unifyTitles','统一各页标题','各页标题位置、字号、颜色和第一页一样'],['arrangeVisuals','图片图表自动排版','本页的图片、图表按原比例排整齐，间距一致'],['slideNumbers','加页码','右下角页码，调整顺序自动更新'],['bodySpacing','统一正文行距','所有页正文行距一致，读起来不挤'],['applyDesign','套用模板','换上另一个模板的背景、配色和版式'],['removeNotes','删除全部备注','清空所有页的演讲者备注'],['duplicateSlide','复制当前页','保留原页并创建副本'],['backupFile','备份当前文件','复制一份带时间的备份'],['exportPdf','导出 PDF','使用当前版式生成新文件']]
  };
  // 内容卡片：与 WPS 版共用同一个本机卡片库（generated\cards.json），经 https://localhost:39128 转给本机服务。
  var SRC={excel:'et',word:'wps',powerpoint:'wpp'},TEXTS=['GeometricShape','TextBox','Placeholder','Callout'];
  async function body(r){return r.json().catch(function(){return {};});}
  async function token(f){var r;try{r=await f('/api/session',{});}catch(e){r=null;}if(!r||!r.ok)throw Error('本地服务未连接，请先双击桌面上的“办公工具箱”。');return (await body(r)).token;}
  async function cardList(f){var k=await token(f),r=await f('/api/cards',{headers:{'X-Coach-Token':k}}),d=await body(r);if(!r.ok)throw Error(d.error||'内容卡片读取失败。');return d.cards||[];}
  function cell(v){v=v==null?'':String(v);return v;}
  async function capture(host,api){
    if(host==='excel')return api.Excel.run(async function(context){var r=context.workbook.getSelectedRange();r.load('values,rowCount,columnCount');await context.sync();
      if(r.rowCount===1&&r.columnCount===1)return {type:'text',text:cell(r.values[0][0])};need(r.rowCount<=20&&r.columnCount<=10,'卡片里的表格最多 20 行、10 列。');return {type:'table',rows:r.values.map(function(row){return row.map(cell);})};});
    if(host==='word')return api.Word.run(async function(context){var sel=context.document.getSelection(),tables=sel.tables;tables.load('items/values');sel.load('text');await context.sync();
      var text=String(sel.text||'').replace(/\r\n?/g,'\n').replace(/\u0007/g,'');if(tables.items.length===1&&tables.items[0].values.length)return {type:'table',rows:tables.items[0].values.map(function(row){return row.map(cell);})};return {type:'text',text:text};});
    if(host==='powerpoint')return api.PowerPoint.run(async function(context){var shapes=context.presentation.getSelectedShapes();shapes.load('items/type');await context.sync();var ranges=[];
      shapes.items.forEach(function(s){if(TEXTS.indexOf(String(s.type))>=0){var tr=s.textFrame.textRange;tr.load('text');ranges.push(tr);}});await context.sync();
      return {type:'text',text:ranges.map(function(r){return String(r.text||'').replace(/\r\n?/g,'\n');}).filter(function(x){return x.trim();}).join('\n')};});
    throw Error('请在 Excel、Word 或 PowerPoint 中打开。');
  }
  async function saveCard(host,api,o){var block=await capture(host,api);if(block.type==='text')need(String(block.text).trim(),'请先选中要保存的文字或表格。');
    var f=o.fetch,k=await token(f),r=await f('/api/cards',{method:'POST',headers:{'Content-Type':'application/json','X-Coach-Token':k},body:JSON.stringify({block:block,sourceHost:SRC[host]})}),d=await body(r);
    if(!r.ok)throw Error(d.error||'保存失败。');return '已设为内容模块：'+(d.card&&d.card.title||'');}
  // 写进 Excel 时，以 = + - @ 开头的文字加撇号，保持为文字，不会变成公式。
  function literal(v){v=cell(v);return /^[=+\-@]/.test(v)?"'"+v:v;}
  async function insertCard(host,api,o){var card=o.card;need(card&&card.block,'请先选择一张内容卡片。');var b=card.block,rows=b.type==='table'?b.rows:null,text=b.type==='text'?String(b.text||''):'';
    if(host==='excel')return api.Excel.run(async function(context){var data=(rows||text.split('\n').map(function(l){return [l];})).map(function(row){return row.map(literal);}),w=Math.max.apply(null,data.map(function(r){return r.length;}));
      data=data.map(function(r){while(r.length<w)r.push('');return r;});var target=context.workbook.getSelectedRange().getCell(0,0).getResizedRange(data.length-1,w-1);target.load('values,address');await context.sync();
      var used=0;target.values.forEach(function(r){r.forEach(function(v){if(v!==''&&v!=null)used++;});});
      if(used&&!o.confirmed)return {confirm:'目标区域有 '+used+' 个单元格已有内容，插入会覆盖它们。是否继续？'};target.values=data;await context.sync();return '已插入内容卡片：'+card.title;});
    if(host==='word')return api.Word.run(async function(context){var sel=context.document.getSelection();if(rows)sel.insertTable(rows.length,rows[0].length,'After',rows.map(function(r){return r.map(cell);}));else sel.insertText(text,'End');await context.sync();return '已插入内容卡片：'+card.title;});
    if(host==='powerpoint')return api.PowerPoint.run(async function(context){var sel=context.presentation.getSelectedSlides();sel.load('items');await context.sync();need(sel.items.length,'请先在左侧点一下要放入的幻灯片。');var sh=sel.items[0].shapes;
      if(rows){try{sh.addTable(rows.length,rows[0].length,{values:rows.map(function(r){return r.map(cell);}),left:40,top:110});await context.sync();return '已插入内容卡片：'+card.title;}catch(e){text=rows.map(function(r){return r.join('\t');}).join('\n');}}
      sh.addTextBox(text,{left:60,top:110,width:820,height:340});await context.sync();return '已插入内容卡片：'+card.title;});
    throw Error('请在 Excel、Word 或 PowerPoint 中打开。');}
  // 备份当前文件：与 WPS 版一样复制磁盘上的文件到备份文件夹（带时间，不覆盖）；只认本机路径，OneDrive 上的文件请用“另存为”。
  function docUrl(api){var d;try{d=api.Office.context.document;}catch(e){return Promise.resolve('');}
    if(typeof d.getFilePropertiesAsync!=='function')return Promise.resolve(String(d.url||''));
    return new Promise(function(resolve,reject){d.getFilePropertiesAsync(function(r){if(r.status==='succeeded')resolve(String(r.value.url||''));else reject(Error('无法读取当前文件的位置，请重新打开侧栏后再试。'));});});}
  function localPath(u){var p=String(u||'');if(/^file:\/\//i.test(p)){try{p=decodeURIComponent(p.replace(/^file:\/\/(?:localhost)?(?=\/)/i,''));}catch(e){return '';}
      if(/^\/[A-Za-z]:\//.test(p))p=p.slice(1);else if(/^file:\/\//i.test(p))p=p.replace(/^file:\/\//i,'//');}
    return /^[A-Za-z]:[\\/]|^\\\\|^\//.test(p)?p:'';}
  async function post(f,url,data){var k=await token(f),r=await f(url,{method:'POST',headers:{'Content-Type':'application/json','X-Coach-Token':k},body:JSON.stringify(data)}),d=await body(r);if(!r.ok)throw Error(d.error||'操作失败。');return d;}
  async function backupFile(host,api,o){var u=await docUrl(api);need(u,'请先把文件保存到电脑上，再备份。');var p=localPath(u);need(p,'这个文件在 OneDrive 或网络位置上，请用「文件 → 另存为」保存一份副本。');
    var d=await post(o.fetch,'/api/backup',{path:p});return '已备份到：'+d.backupPath+'\n备份的是最近一次保存的内容；有未保存的修改时，请先按 Ctrl+S 再备份。';}
  function slices(api){return new Promise(function(resolve,reject){var O=api.Office;O.context.document.getFileAsync(O.FileType.Pdf,{sliceSize:4194304},function(res){if(res.status!=='succeeded')return reject(res.error||Error('PDF 生成失败。'));var file=res.value,count=file.sliceCount,parts=[],i=0;
      (function next(){if(i>=count){file.closeAsync&&file.closeAsync();return resolve(parts);}file.getSliceAsync(i,function(sr){if(sr.status!=='succeeded'){file.closeAsync&&file.closeAsync();return reject(sr.error||Error('PDF 生成失败。'));}parts.push(sr.value.data);i++;next();});})();});});}
  function base64(parts){var out='',CH=0x8000;parts.forEach(function(p){var a=p instanceof Uint8Array?p:Uint8Array.from(p);for(var i=0;i<a.length;i+=CH)out+=String.fromCharCode.apply(null,Array.prototype.slice.call(a.subarray(i,i+CH)));});return typeof btoa==='function'?btoa(out):b64(out);}
  function b64(x){var A='ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/',o='',i,n;for(i=0;i<x.length;i+=3){n=(x.charCodeAt(i)<<16)|((x.charCodeAt(i+1)||0)<<8)|(x.charCodeAt(i+2)||0);o+=A[n>>18&63]+A[n>>12&63]+(i+1<x.length?A[n>>6&63]:'=')+(i+2<x.length?A[n&63]:'=');}return o;}
  async function exportPdf(host,api,o){need(host!=='excel','Excel 的插件接口不能直接导出 PDF。请用「文件 → 导出 → 创建 PDF/XPS」，文件名里加上日期，不会覆盖原文件。');
    var parts=await slices(api),size=parts.reduce(function(n,p){return n+p.length;},0);need(size<=80*1024*1024,'PDF 超过 80 MB，请在 Office 里用「另存为 PDF」。');
    var u=await docUrl(api),name=(localPath(u)||u).split(/[\\/]/).pop()||'文档',d=await post(o.fetch,'/api/save-pdf',{name:name,data:base64(parts)});return 'PDF 已导出：'+d.path;}
  function plus(host){return host==='word'?root.OfficeDocPlus:host==='excel'?root.OfficeSheetPlus:null;}
  function list(host){var p=plus(host),group=host==='word'?'docfix':'sheetplus';return (LIST[host]||[]).map(function(x){return {id:x[0],label:x[1],desc:x[2],form:x[3]||'',fields:x[4]||null};}).concat(p?p.list.map(function(x){return {id:x[0],label:x[1],desc:x[2],form:'',fields:x[3]||null,group:group};}):[]);}
  // 中日本地化小工具（Excel）：计算用共享的 LocaleEngine，结果写到选区右侧；统一格式类在原处修改并先确认。
  var LOCALE={rmbUpper:1,idCardInfo:1,phoneCheck:1,zenHan:1,toWareki:1,toSeireki:1,postalFormat:1,furigana:1};
  function localeExcel(api,id,o){var E=root.LocaleEngine;if(!E)throw Error('本地化组件没有加载，请刷新侧栏。');
    return api.Excel.run(async function(context){var r=context.workbook.getSelectedRange();r.load('values,formulas,rowCount,columnCount,address');await context.sync();
      if(r.rowCount*r.columnCount>2000)throw Error('一次最多处理 2000 个单元格，请缩小选区。');
      var one=id==='idCardInfo'||id==='furigana';if(one&&r.columnCount!==1)throw Error('请只选一列。');
      var items=[];r.values.forEach(function(row,i){row.forEach(function(v,j){var f=r.formulas[i][j];if(typeof f==='string'&&f.charAt(0)==='=')return;if(v===''||v===null||v===undefined)return;items.push({i:i,j:j,v:v});});});
      if(id==='zenHan'||id==='postalFormat'){var ch=[];items.forEach(function(x){var s=id==='zenHan'?E.zenHan(x.v):E.postal(x.v);if(s!==null)ch.push({i:x.i,j:x.j,s:s,old:x.v});});
        if(!ch.length)return id==='zenHan'?'选区里没有需要统一的全角英数字或半角片假名。':'选区里的邮编格式已经统一，或没有 6～7 位的邮编。';
        if(!o.confirmed)return {confirm:(id==='zenHan'?'将修改 '+ch.length+' 个单元格：英文字母和数字改成半角，半角片假名改成全角。':'将统一 '+ch.length+' 个邮编，并设为文本格式。')+'\n例：'+ch[0].old+' → '+ch[0].s+'\n是否执行？'};
        ch.forEach(function(x){var c=r.getCell(x.i,x.j);if(id==='postalFormat')c.numberFormat=[['@']];c.values=[[x.s]];});await context.sync();return (id==='zenHan'?'已统一 ':'已统一邮编 ')+ch.length+' 个单元格（可按 Ctrl+Z 撤销）。';}
      var rows=[],skip=0,bad=0,width=id==='idCardInfo'?2:1;
      items.forEach(function(x){var out=null;
        if(id==='rmbUpper'){var s=E.rmbUpper(x.v);if(s)out=[s];}
        else if(id==='idCardInfo'){var info=E.idInfo(x.v);if(info.error){bad++;out=[info.error,''];}else out=[info.birth,info.gender];}
        else if(id==='phoneCheck'){var ph=E.phoneCheck(x.v);if(ph){if(!/^有效/.test(ph))bad++;out=[ph];}}
        else if(id==='toWareki'){var w=E.toWareki(x.v);if(w)out=[w];}
        else if(id==='toSeireki'){var se=E.toSeireki(x.v);if(se){if(se.error){bad++;out=[root.OfficeI18n&&root.OfficeI18n.t?root.OfficeI18n.t(se.error):se.error];}else out=[{serial:se.serial}];}}
        else if(id==='furigana'&&typeof x.v==='string')out=[{formula:true}];
        if(out)rows.push({i:x.i,j:x.j,out:out});else skip++;});
      if(!rows.length)throw Error('选区里没有可以识别的内容。');
      var base=r.getOffsetRange(0,r.columnCount),targets=[];rows.forEach(function(x){for(var k=0;k<width;k++){var c=base.getCell(x.i,x.j+k);c.load('values');targets.push(c);}});await context.sync();
      var used=targets.filter(function(c){var v=c.values[0][0];return v!==''&&v!==null&&v!==undefined;}).length;
      if(used&&!o.confirmed)return {confirm:'右侧有 '+used+' 个单元格已有内容，写入会覆盖。是否继续？'};
      rows.forEach(function(x){x.out.forEach(function(v,k){var c=base.getCell(x.i,x.j+k);
        if(v&&v.serial){c.numberFormat=[['yyyy/m/d']];c.values=[[v.serial]];}else if(!(v&&v.formula))c.values=[[v]];});});
      await context.sync();
      if(id==='furigana'){rows.forEach(function(x){var src=r.getCell(x.i,x.j);src.load('address');x.src=src;});await context.sync();rows.forEach(function(x){base.getCell(x.i,x.j).formulas=[['=PHONETIC('+x.src.address.split('!').pop()+')']];});await context.sync();
        return '已在右侧写入 '+rows.length+' 个读音公式。只有用日文输入法输入的文字才有读音信息。';}
      return '已在选区右侧写入 '+rows.length+' 项结果'+(skip?'，'+skip+' 个无法识别的单元格已跳过':'')+(bad?'，'+bad+' 项有问题已标出':'')+'。原数据未改动。';});}
  function run(host,id,api,options){if(host==='excel'&&LOCALE[id])return localeExcel(api,id,options||{});if(id==='saveCard')return saveCard(host,api,options||{});if(id==='backupFile')return backupFile(host,api,options||{});if(id==='exportPdf')return exportPdf(host,api,options||{});if(id==='insertCard')return insertCard(host,api,options||{});var set=host==='excel'?excel:host==='word'?word:host==='powerpoint'?ppt:{};var p=plus(host),fn=set[id]||(p&&p.tools[id]);if(!fn)return Promise.reject(Error('工具无法识别。'));return Promise.resolve().then(function(){return fn(api,options||{});});}
  root.OfficeAddinTools={list:list,run:run,cards:cardList,documentUrl:docUrl,localPath:localPath};
})(typeof window!=='undefined'?window:globalThis);

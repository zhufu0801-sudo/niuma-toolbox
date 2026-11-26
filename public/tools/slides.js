// 办公工具箱 · 演示模块（WPS 演示）。
(function(root){
  'use strict';
  var T=root.OfficeTools,L=T.lib,tableInput=L.tableInput;
  // 演示文稿里已有文字的字体，用来让新页面和原有页面保持一致；读不到时用微软雅黑。
  function deckFont(d){
    try{for(var i=1;i<=Math.min(d.Slides.Count,20);i++){var shapes=d.Slides.Item(i).Shapes;for(var k=1;k<=shapes.Count;k++){var shape=shapes.Item(k);if(shape.HasTextFrame&&shape.TextFrame.HasText){var name=String(shape.TextFrame.TextRange.Font.Name||'');if(name)return name;}}}}catch(ignore){}
    return L.localFont('微软雅黑');
  }
  // WPS 在用户点过某一页之前，View.Slide 可能读不到；再依次尝试选中的页、只有一页时的第 1 页。
  function currentSlide(a,d){
    try{var v=a.ActiveWindow.View.Slide;if(v&&v.SlideIndex)return v;}catch(ignore){}
    try{var r=a.ActiveWindow.Selection.SlideRange;if(r&&r.Count)return r.Item(1);}catch(ignore){}
    try{if(d&&Number(d.Slides.Count)>=1)return d.Slides.Item(1);}catch(ignore){}
    return null;
  }
  // 新页面插在当前页之后，而不是总是放到最后。
  function insertIndex(a,d){var current=currentSlide(a,d);return current&&current.SlideIndex?Number(current.SlideIndex):Number(d.Slides.Count)||0;}
  // 按行列数算出能放进页面的字号：行多、列多时自动缩小，最小 10 号。
  function tableFontSize(rows,cols,width,height){var byHeight=Math.floor((height-110)/rows/1.9),byWidth=Math.floor((width-60)/cols/5.5);return Math.max(10,Math.min(18,byHeight,byWidth));}
  // 版式：只用三种颜色（标题深蓝、正文灰蓝、一条短强调线），版心左对齐，留足页边距，和原有模板放在一起也不突兀。
  // 装饰（强调线、项目符号、行距、底纹）在部分 WPS 版本可能不支持，失败时跳过，不影响文字本身。
  var INK='#1F2A44',BODY='#2F3B52',ACCENT='#2B5797',BAND='#F3F5F1',RULE='#D9DED6';
  function rgb(hex){return parseInt(hex.slice(1,3),16)+parseInt(hex.slice(3,5),16)*256+parseInt(hex.slice(5,7),16)*65536;}
  function soft(fn){try{fn();}catch(ignore){}}
  function textBox(slide,left,top,w,h,text,font,size,color,bold){var box=slide.Shapes.AddTextbox(1,left,top,w,h),range=box.TextFrame.TextRange;range.Text=text;range.Font.Name=font;range.Font.Size=size;if(bold)range.Font.Bold=-1;soft(function(){range.Font.Color.RGB=rgb(color);});soft(function(){box.TextFrame.WordWrap=-1;});return box;}
  function accentBar(slide,left,top,w){soft(function(){var bar=slide.Shapes.AddShape(1,left,top,w,4);soft(function(){bar.Fill.Visible=-1;bar.Fill.Solid();});soft(function(){bar.Fill.Transparency=0;});bar.Fill.ForeColor.RGB=rgb(ACCENT);bar.Line.Visible=0;});}
  // 大纲的一页：只有一行时做成章节页（大字、居中偏上），有正文时做成“标题 + 强调线 + 要点”。
  function outlinePage(slide,page,font,width,height,cover){
    var lines=page.trim().split('\n'),title=lines.shift().replace(/^#{1,6}\s+/,'').trim(),points=lines.map(function(x){return x.replace(/^\s*[-*•·]\s+/,'').trim();}).filter(Boolean),margin=Math.round(width*0.07);
    if(!points.length){var size=Array.from(title).length>18?32:40;textBox(slide,margin,height*0.36,width-margin*2,size*2.6,title,font,size,INK,true);accentBar(slide,margin+7,height*0.36-14,56);return;}
    var titleSize=Array.from(title).length>22?24:28,titleTop=Math.round(height*0.08);
    textBox(slide,margin,titleTop,width-margin*2,titleSize*1.9,title,font,titleSize,INK,true);
    var barTop=titleTop+titleSize*1.9+6;accentBar(slide,margin+7,barTop,48);
    var body=points.join('\r'),total=body.length,bodySize=total>400||points.length>9?16:total>200||points.length>6?18:20,top=barTop+22;
    var box=textBox(slide,margin,top,width-margin*2,height-top-height*0.08,body,font,bodySize,BODY,false),para=box.TextFrame.TextRange;
    if(cover&&points.length===1){soft(function(){para.ParagraphFormat.Bullet.Visible=0;});return;}
    // 实测 WPS 12.1：只设 Visible/Character 会显示成“1. 2. 3.”，要先把类型设成圆点（ppBulletUnnumbered=1）。
    soft(function(){para.ParagraphFormat.Bullet.Visible=-1;});soft(function(){para.ParagraphFormat.Bullet.Type=1;});soft(function(){para.ParagraphFormat.Bullet.Character=8226;});
    soft(function(){para.ParagraphFormat.Bullet.Font.Color.RGB=rgb(ACCENT);});
    soft(function(){para.ParagraphFormat.SpaceWithin=1.2;para.ParagraphFormat.SpaceAfter=Math.round(bodySize*0.45);});
  }
  // 表格页：表头用设置里的“表头颜色”，数据行隔行浅底纹，细分隔线；原内容自带格式时保留原格式。
  function tableLook(grid,values,hasFormat){
    if(hasFormat)return;var conf=T.settings?T.settings().table||{}:{},head=/^#[0-9a-f]{6}$/i.test(conf.headerBackground||'')?conf.headerBackground:'#DDEBF7',headColor=/^#[0-9a-f]{6}$/i.test(conf.headerColor||'')?conf.headerColor:INK;
    var numeric=values[0].map(function(x,c){return values.length>1&&values.slice(1).every(function(row){var v=String(row[c]||'').trim();return !v||/^[-+]?[¥$€£]?[\d,，.]+%?$/.test(v);});});
    grid.forEach(function(row,r){row.forEach(function(box,c){var cell=box.Shape;
      soft(function(){cell.Fill.ForeColor.RGB=rgb(r===0?head:r%2?'#FFFFFF':BAND);});
      soft(function(){cell.TextFrame.TextRange.Font.Color.RGB=rgb(r===0?headColor:BODY);});
      // 整列都是数字（金额、百分比）时右对齐，方便上下比较。
      if(numeric[c])soft(function(){cell.TextFrame.TextRange.ParagraphFormat.Alignment=3;});
      soft(function(){var line=box.Borders.Item(3);line.ForeColor.RGB=rgb(RULE);line.Weight=0.75;});
    });});
  }
  function slides(a,d,block){
    var added=[],width=Number(d.PageSetup.SlideWidth),height=Number(d.PageSetup.SlideHeight),font=deckFont(d),at=insertIndex(a,d);
    function addSlide(){var slide=d.Slides.Add(at+added.length+1,12);added.push(slide);return slide;}
    try{
      if(block.type==='table'){
        var rows=block.rows;if(rows[0].length>10||rows.length>20)throw Error('PPT 快速插入最多 20 行、10 列。');
        // 在创建任何页面前检查文字容量，避免生成后才发现单元格内容被裁掉。
        for(var check=1;check<Math.max(rows.length,2);check+=9){var batch=[rows[0]].concat(rows.slice(check,check+9)),fontSize=tableFontSize(batch.length,batch[0].length,width,height),chars=Math.max(1,Math.floor(((width-60)/batch[0].length-10)/fontSize)),maxLines=Math.max(1,Math.floor((height-100)/batch.length/(fontSize*1.4)));for(var rr=0;rr<batch.length;rr++)for(var cc=0;cc<batch[rr].length;cc++){var needed=String(batch[rr][cc]).split(/\r?\n/).reduce(function(total,line){return total+Math.max(1,Math.ceil(Array.from(line).length/chars));},0);if(needed>maxLines)throw Error('表格单元格文字过长，当前幻灯片无法完整容纳。请拆分模块、缩短文字，或插入到 Excel。');}}
        for(var i=1;i<Math.max(rows.length,2);i+=9){var group=[rows[0]].concat(rows.slice(i,i+9));var size=tableFontSize(group.length,group[0].length,width,height);var slide=addSlide();var shape=slide.Shapes.AddTable(group.length,group[0].length,30,50,width-60,Math.min(height-100,group.length*size*2));var grid=[];for(var r=0;r<group.length;r++)for(var c=0;c<group[r].length;c++){var box=shape.Table.Cell(r+1,c+1);(grid[r]=grid[r]||[])[c]=box;var cell=box.Shape.TextFrame.TextRange;cell.Text=group[r][c];cell.Font.Name=font;cell.Font.Size=size;if(block.format&&L.applyModuleStyle){var original=i+r-1,style=Object.assign({},block.format.base,block.format.cells&&block.format.cells[r===0?0:original]&&block.format.cells[r===0?0:original][c]);L.applyModuleStyle(cell,style,true,true);if(style.background)try{box.Shape.Fill.ForeColor.RGB=L.officeColor(style.background);}catch(ignore){}}if(r===0)cell.Font.Bold=-1;}}tableLook(grid,group,!!block.format);
      }else{
        var paragraphs=block.format&&block.format.paragraphs,lines=String(block.text||'').split('\n'),pages;if(paragraphs&&paragraphs.some(function(p){return p.level;})){pages=[];lines.forEach(function(line,i){if(paragraphs[i]&&paragraphs[i].level||!pages.length)pages.push(line);else pages[pages.length-1]+='\n'+line;});pages=pages.filter(function(p){return p.trim();});}else pages=String(block.text||'').split(/\n\s*\n/).filter(function(p){return p.trim();});if(!pages.length)throw Error('没有可生成幻灯片的内容。');if(pages.length>30)throw Error('一次最多新建 30 页，请拆分内容。');for(var n=0;n<pages.length;n++)if(pages[n].length>1200)throw Error('一页文字超过 1200 字，请用空行分成多页。');
        pages.forEach(function(page,i){outlinePage(addSlide(),page,font,width,height,i===0);});
      }
      return added.length;
    }catch(e){var failed=false;for(var k=added.length-1;k>=0;k--)try{added[k].Delete();}catch(ignore){failed=true;}throw Error(e.message+(added.length?(failed?' 部分页可能已插入，请检查并使用 WPS 撤销。':' 新增页面已回退。'):''));}
  }
  function isTitleShape(shape){try{var t=Number(shape.PlaceholderFormat.Type);return t===1||t===3;}catch(ignore){return false;}}
  // 统一字体：所有文字框用同一种字体；正文统一为最常见的字号，标题保持原有大小，层次不丢失。
  function formatSlide(a){
    var slide=currentSlide(a,a.ActivePresentation);if(!slide)throw Error('请先在左侧点一下要整理的幻灯片。');
    var frames=[];for(var i=1;i<=slide.Shapes.Count;i++){var shape=slide.Shapes.Item(i);if(shape.HasTextFrame&&shape.TextFrame.HasText)frames.push(shape);}
    if(!frames.length)throw Error('当前页没有可设置的文字框。');
    var sizes=frames.map(function(f){return Number(f.TextFrame.TextRange.Font.Size)||0;}),largest=Math.max.apply(null,sizes),titles=frames.map(function(f,k){return isTitleShape(f)||(frames.length>1&&sizes[k]===largest&&sizes.filter(function(x){return x===largest;}).length===1);});
    var counts={},body=0;sizes.forEach(function(sz,k){if(!titles[k]&&sz){counts[sz]=(counts[sz]||0)+1;}});Object.keys(counts).forEach(function(sz){if(!body||counts[sz]>counts[body])body=Number(sz);});
    var font=deckFont(a.ActivePresentation);
    frames.forEach(function(f,k){var range=f.TextFrame.TextRange;range.Font.Name=font;if(!titles[k]&&body)range.Font.Size=body;});
    var kept=titles.filter(Boolean).length;
    return '已统一当前页 '+frames.length+' 个文字框为“'+font+'”'+(body?'，正文统一为 '+body+' 号':'')+(kept?'，标题保持原有大小':'')+'。';
  }
  function matchSlide(a){
    var sel=a.ActiveWindow.Selection,slide=currentSlide(a,a.ActivePresentation),picked=[];
    try{if(sel.Type===2||sel.Type===3)for(var i=1;i<=sel.ShapeRange.Count;i++)picked.push(sel.ShapeRange.Item(i));}catch(ignore){}
    if(!picked.length)throw Error('请先选中刚放进来的文字框或表格。');
    var names=picked.map(function(p){return String(p.Name);}),font='',size=0;
    for(var k=1;k<=slide.Shapes.Count;k++){var other=slide.Shapes.Item(k);if(names.indexOf(String(other.Name))>=0||isTitleShape(other))continue;if(other.HasTextFrame&&other.TextFrame.HasText){font=String(other.TextFrame.TextRange.Font.Name);size=Number(other.TextFrame.TextRange.Font.Size);break;}}
    if(!font)font=deckFont(a.ActivePresentation);
    var width=Number(a.ActivePresentation.PageSetup.SlideWidth),height=Number(a.ActivePresentation.PageSetup.SlideHeight),fitted=0;
    picked.forEach(function(p){
      if(p.HasTable){var rows=p.Table.Rows.Count,cols=p.Table.Columns.Count,ts=size||tableFontSize(rows,cols,width,height);for(var r=1;r<=rows;r++)for(var c=1;c<=cols;c++){var cell=p.Table.Cell(r,c).Shape.TextFrame.TextRange;cell.Font.Name=font;cell.Font.Size=Math.min(ts,tableFontSize(rows,cols,width,height));}}
      else if(p.HasTextFrame){p.TextFrame.TextRange.Font.Name=font;if(size)p.TextFrame.TextRange.Font.Size=size;}
      var maxW=width-60,maxH=height-60;if(p.Width>maxW||p.Height>maxH){var ratio=Math.min(maxW/p.Width,maxH/p.Height);p.Width=p.Width*ratio;p.Height=p.Height*ratio;fitted++;}
      if(p.Left<0||p.Left+p.Width>width)p.Left=(width-p.Width)/2;if(p.Top<0||p.Top+p.Height>height)p.Top=(height-p.Height)/2;
    });
    return '已按本页正文设置字体“'+font+'”'+(size?'、'+size+' 号':'')+(fitted?'，并把 '+fitted+' 个对象缩放进页面':'')+'。';
  }
  L.slides=slides;L.currentSlide=currentSlide;L.tableFontSize=tableFontSize;L.deckFont=deckFont;
  T.register({id:'makeSlides',label:'按大纲制作 PPT',desc:'写好大纲，生成新页面',icon:'presentation',group:'create',panel:false,hosts:{wpp:function(c){return '已在当前页之后新建 '+slides(c.a,c.d,{type:'text',text:c.options.text})+' 页。';}}});
  T.register({id:'insertTable',hosts:{wpp:function(c){return '已在当前页之后新建 '+slides(c.a,c.d,{type:'table',rows:tableInput(c.options)})+' 页表格。';}}});
  T.register({id:'formatSlide',label:'统一本页字体',desc:'正文统一、标题保持大小',icon:'type',group:'slide',hosts:{wpp:function(c){return formatSlide(c.a);}}});
  T.register({id:'duplicateSlide',label:'复制当前页',desc:'保留原页并创建副本',icon:'duplicate',group:'slide',hosts:{wpp:function(c){var a=c.a,d=c.d,current=currentSlide(a,d);if(!current)throw Error('请先在左侧点一下要复制的幻灯片。');var at=Number(current.SlideIndex),dup=d.Slides.Range(at).Duplicate();
    // 部分 WPS 版本会把副本放到别处，这里移回原页之后并选中，方便接着编辑。
    try{var copy=dup&&dup.Item?dup.Item(1):dup;if(copy){if(Number(copy.SlideIndex)!==at+1)copy.MoveTo(at+1);copy.Select();}}catch(ignore){}
    return '已在第 '+at+' 页之后复制一页，保留原页。';}}});
  // ---- 全部页：统一字体、删除动画、删除备注 ----
  // 文字框、表格单元格、组合里的文字框都算。
  function eachText(shape,cb){try{if(Number(shape.Type)===6&&shape.GroupItems){for(var g=1;g<=shape.GroupItems.Count;g++)eachText(shape.GroupItems.Item(g),cb);return;}}catch(ignore){}
    try{if(shape.HasTable){var tb=shape.Table;for(var r=1;r<=tb.Rows.Count;r++)for(var c=1;c<=tb.Columns.Count;c++){var tf=tb.Cell(r,c).Shape.TextFrame;if(tf.HasText)cb(tf.TextRange);}return;}}catch(ignore){}
    try{if(shape.HasTextFrame&&shape.TextFrame.HasText)cb(shape.TextFrame.TextRange);}catch(ignore){}}
  function allText(d,cb){for(var i=1;i<=d.Slides.Count;i++){var shapes=d.Slides.Item(i).Shapes;for(var k=1;k<=shapes.Count;k++)eachText(shapes.Item(k),cb);}}
  function fontOf(range){try{var f=range.Font;return String(f.NameFarEast||f.Name||'');}catch(e){return '';}}
  var KEEP='\n建议先点“备份当前文件”：WPS 演示里这一步可能不能一次撤销。';
  function unifyDeckFont(c){var d=c.a.ActivePresentation,count={},frames=0;allText(d,function(r){frames++;var n=fontOf(r);if(n)count[n]=(count[n]||0)+1;});
    if(!frames)throw Error('演示文稿里没有文字。');
    var names=Object.keys(count).sort(function(x,y){return count[y]-count[x];}),want=String(c.options.font||'').trim()||names[0]||deckFont(d);
    if(names.length===1&&names[0]===want)return '全部 '+frames+' 处文字已经是“'+want+'”，不需要修改。';
    if(!L.ask('全部 '+d.Slides.Count+' 页共 '+frames+' 处文字，现在用了 '+names.length+' 种字体：'+names.slice(0,6).map(function(n){return n+' ×'+count[n];}).join('、')+(names.length>6?'…':'')+'。\n将全部统一为“'+want+'”，字号、颜色、加粗不变。是否执行？'+KEEP))return '已取消。';
    var done=0;allText(d,function(r){try{r.Font.Name=want;}catch(ignore){}try{r.Font.NameFarEast=want;}catch(ignore){}if(fontOf(r)===want)done++;});
    return '已把 '+done+' 处文字统一为“'+want+'”'+(done<frames?'（'+(frames-done)+' 处没能修改，可能是图片里的文字或艺术字）':'')+'。';}
  function animOf(s){var n=0;try{n+=Number(s.TimeLine.MainSequence.Count)||0;}catch(ignore){}try{var iq=s.TimeLine.InteractiveSequences;for(var k=1;k<=iq.Count;k++)n+=Number(iq.Item(k).Count)||0;}catch(ignore){}return n;}
  function transOf(s){try{var e=Number(s.SlideShowTransition.EntryEffect);return e&&e!==0?1:0;}catch(e){return 0;}}
  function removeAnimations(c){var d=c.a.ActivePresentation,anims=0,pages=0,trans=0;for(var i=1;i<=d.Slides.Count;i++){var s=d.Slides.Item(i),n=animOf(s);if(n){anims+=n;pages++;}trans+=transOf(s);}
    if(!anims&&!trans)return '没有找到动画和切换效果。';
    if(!L.ask('将删除 '+pages+' 页上的 '+anims+' 个动画'+(trans?'，以及 '+trans+' 页的切换效果':'')+'。文字和图片本身不动。是否执行？'+KEEP))return '已取消。';
    for(var j=1;j<=d.Slides.Count;j++){var sl=d.Slides.Item(j);try{var m=sl.TimeLine.MainSequence;for(var x=Number(m.Count);x>=1;x--)m.Item(x).Delete();}catch(ignore){}
      try{var q=sl.TimeLine.InteractiveSequences;for(var y=Number(q.Count);y>=1;y--){var seq=q.Item(y);for(var z=Number(seq.Count);z>=1;z--)seq.Item(z).Delete();}}catch(ignore){}
      try{if(transOf(sl))sl.SlideShowTransition.EntryEffect=0;}catch(ignore){}}
    var left=0,leftT=0;for(var k2=1;k2<=d.Slides.Count;k2++){left+=animOf(d.Slides.Item(k2));leftT+=transOf(d.Slides.Item(k2));}
    if(left)throw Error('还有 '+left+' 个动画没删掉（可能是 WPS 不允许用插件删除的类型），请在“动画 → 动画窗格”里手动删除。');
    return '已删除 '+anims+' 个动画'+(trans?(leftT?'；切换效果有 '+leftT+' 页没能去掉，请在“切换”里选“无”':'和 '+trans+' 页的切换效果'):'')+'。';}
  function notesBody(s){try{var shapes=s.NotesPage.Shapes;for(var k=1;k<=shapes.Count;k++){var sh=shapes.Item(k);try{if(Number(sh.PlaceholderFormat.Type)===2&&sh.HasTextFrame)return sh.TextFrame.TextRange;}catch(ignore){}}}catch(ignore){}return null;}
  function removeNotes(c){var d=c.a.ActivePresentation,pages=[];for(var i=1;i<=d.Slides.Count;i++){var r=notesBody(d.Slides.Item(i));if(r&&String(r.Text).replace(/\s/g,''))pages.push(i);}
    if(!pages.length)return '没有找到备注。';
    if(!L.ask('将清空 '+pages.length+' 页的备注（第 '+pages.slice(0,10).join('、')+(pages.length>10?'…':'')+' 页）。发给别人前常用；演讲时就看不到备注了。是否执行？'+KEEP))return '已取消。';
    pages.forEach(function(i){var r=notesBody(d.Slides.Item(i));if(r)r.Text='';});
    var left=pages.filter(function(i){var r=notesBody(d.Slides.Item(i));return r&&String(r.Text).replace(/\s/g,'');}).length;
    if(left)throw Error('有 '+left+' 页的备注没能清空，请在“视图 → 备注页”里手动删除。');
    return '已清空 '+pages.length+' 页的备注。';}
  // ---- 美化已有演示：统一各页标题、加页码、统一正文行距、套用模板 ----
  function titleOf(slide){var shapes=slide.Shapes;for(var k=1;k<=shapes.Count;k++){var sh=shapes.Item(k);try{var t=Number(sh.PlaceholderFormat.Type);if((t===1||t===3)&&sh.HasTextFrame&&String(sh.TextFrame.TextRange.Text||'').trim())return {shape:sh,center:t===3};}catch(ignore){}}return null;}
  function unifyTitles(c){var d=c.a.ActivePresentation,n=d.Slides.Count,list=[],ref=null;if(n<2)throw Error('只有一页，不需要统一。');
    for(var i=1;i<=n;i++){var t=titleOf(d.Slides.Item(i));if(t&&!t.center)list.push({i:i,sh:t.shape});}
    if(list.length<2)throw Error('有标题占位符的内容页不到两页。用文本框写的标题认不出来：请在“版式”里换成带标题的版式。');
    var r0=list[0].sh,f0=r0.TextFrame.TextRange.Font;ref={l:r0.Left,t:r0.Top,w:r0.Width,h:r0.Height,name:fontOf(r0.TextFrame.TextRange),size:Number(c.options&&c.options.size)||Number(f0.Size),bold:f0.Bold,color:f0.Color.RGB,align:r0.TextFrame.TextRange.ParagraphFormat.Alignment};
    if(!(ref.size>=10&&ref.size<=96))throw Error('标题字号请填 10～96。');
    if(!L.ask('将以第 '+list[0].i+' 页的标题为准（'+(ref.name||'当前字体')+'、'+ref.size+' 磅），统一其余 '+(list.length-1)+' 页标题的位置、大小、字体、颜色和对齐。封面和节标题页不动。是否执行？'+KEEP))return '已取消。';
    var done=0;list.forEach(function(x){var sh=x.sh,tr=sh.TextFrame.TextRange;try{sh.Left=ref.l;sh.Top=ref.t;sh.Width=ref.w;sh.Height=ref.h;if(ref.name){tr.Font.Name=ref.name;try{tr.Font.NameFarEast=ref.name;}catch(ignore){}}tr.Font.Size=ref.size;tr.Font.Bold=ref.bold;tr.Font.Color.RGB=ref.color;tr.ParagraphFormat.Alignment=ref.align;done++;}catch(ignore){}});
    return '已统一 '+done+' 页的标题（位置、字号、字体、颜色、对齐），以第 '+list[0].i+' 页为准。';}
  function slideNumbers(c){var d=c.a.ActivePresentation,n=d.Slides.Count,skip=!/^(否|不|不要|no|n|いいえ|表示)$/i.test(String(c.options&&c.options.skipFirst||'').trim());if(!n)throw Error('演示里还没有页面。');
    if(!L.ask('将给'+(skip?'除封面外的 '+(n-1):'全部 '+n)+' 页加上页码（右下角，调整页面顺序后会自动更新）。是否执行？'))return '已取消。';
    var on=0,boxes=0;for(var i=1;i<=n;i++){var s=d.Slides.Item(i),want=!(skip&&i===1);try{s.HeadersFooters.SlideNumber.Visible=want;if(want)on++;continue;}catch(ignore){}
      if(!want)continue;try{var w=d.PageSetup.SlideWidth,h=d.PageSetup.SlideHeight,box=s.Shapes.AddTextbox(1,w-90,h-40,70,26);box.Name='办公工具箱页码';var tr=box.TextFrame.TextRange;tr.Text=String(i);tr.Font.Size=12;tr.Font.Color.RGB=8421504;tr.ParagraphFormat.Alignment=3;boxes++;}catch(ignore){}}
    return on?'已给 '+on+' 页打开页码'+(skip?'，封面不显示':'')+'。页码位置和样式跟着母版，可以在“视图 → 幻灯片母版”里调整。':'这个版本不支持自动页码，已在 '+boxes+' 页右下角加了页码文本框（调整页面顺序后需要重新加）。';}
  function bodySpacing(c){var d=c.a.ActivePresentation,o=c.options||{},line=Number(o.line)||1.2;if(!(line>=0.8&&line<=3))throw Error('行距请填 0.8～3 倍。');var frames=0,paras=0;
    if(!L.ask('将把所有页正文文字框（不含标题）的行距统一为 '+line+' 倍，段后留 6 磅。是否执行？'+KEEP))return '已取消。';
    for(var i=1;i<=d.Slides.Count;i++){var shapes=d.Slides.Item(i).Shapes;for(var k=1;k<=shapes.Count;k++){var sh=shapes.Item(k),isTitle=false;try{var t=Number(sh.PlaceholderFormat.Type);isTitle=t===1||t===3;}catch(ignore){}if(isTitle)continue;try{if(!sh.HasTextFrame||!String(sh.TextFrame.TextRange.Text||'').trim())continue;var pf=sh.TextFrame.TextRange.ParagraphFormat;pf.LineRuleWithin=true;pf.SpaceWithin=line;pf.LineRuleAfter=false;pf.SpaceAfter=6;frames++;paras+=Number(sh.TextFrame.TextRange.Paragraphs().Count)||0;}catch(ignore){}}}
    return '已统一 '+frames+' 个文字框的行距（'+line+' 倍，段后 6 磅）。';}
  // Generated editorial decks contain slide-local artwork, so ApplyTemplate cannot reproduce their theme safely.
  function themeInk(hex){var v=hex.match(/../g).map(function(x){var n=parseInt(x,16)/255;return n<=.04045?n/12.92:Math.pow((n+.055)/1.055,2.4);}),l=v[0]*.2126+v[1]*.7152+v[2]*.0722;return l>.179?'000000':'FFFFFF';}
  function themeShape(shape,font,background){
    if(Number(shape.Type)===6&&shape.GroupItems){for(var g=1;g<=shape.GroupItems.Count;g++)themeShape(shape.GroupItems.Item(g),font,background);return;}
    if(shape.HasTable){var tb=shape.Table;for(var r=1;r<=tb.Rows.Count;r++)for(var c=1;c<=tb.Columns.Count;c++)themeShape(tb.Cell(r,c).Shape,font,background);return;}
    if(!shape.HasTextFrame)return;
    var color=background;
    // Background/inherited fills expose a stale ForeColor in WPS; only a visible solid fill is a real surface.
    try{if(shape.Fill.Visible&&Number(shape.Fill.Type)===1&&Number(shape.Fill.Transparency)<.5){var n=Number(shape.Fill.ForeColor.RGB);if(n>=0)color=[n&255,(n>>8)&255,(n>>16)&255].map(function(x){return ('0'+x.toString(16)).slice(-2);}).join('');}}catch(ignore){}
    var f=shape.TextFrame.TextRange.Font;f.Name=font;soft(function(){f.NameFarEast=font;});f.Color.RGB=rgb('#'+themeInk(color));
  }
  async function applyLibraryDesign(c){var a=c.a,d=a.ActivePresentation,identity=String(d.FullName),response=await L.localFetch('/templates/ppt/index.json',{});if(!response.ok)throw Error('模板列表暂时无法读取，请重新打开工具箱。');var items=await response.json(),item=items.filter(function(x){return x.id===c.options.templateId;})[0];if(!item||!item.colors||!/^[0-9a-f]{6}$/i.test(item.colors.background))throw Error('模板不存在，请刷新后重试。');if(String(a.ActivePresentation.FullName)!==identity)throw Error('当前文档已改变，请重新执行。');
    if(!L.ask('将套用模板配色和字体，保留文字、图片及其位置。图片和图表颜色保持原样。是否继续？'+KEEP))return '已取消。';
    var font=item.lang==='ja'?'Yu Gothic':'微软雅黑',bg=item.colors.background;
    for(var i=1;i<=d.Slides.Count;i++){var slide=d.Slides.Item(i);slide.FollowMasterBackground=0;slide.Background.Fill.Solid();slide.Background.Fill.ForeColor.RGB=rgb('#'+bg);for(var k=1;k<=slide.Shapes.Count;k++)themeShape(slide.Shapes.Item(k),font,bg);}
    return '已套用模板配色和字体，内容和位置保持不变。请逐页核对图片与文字的对比。';
  }
  async function applyDesign(c){var a=c.a,d=a.ActivePresentation,p;if(c.options&&c.options.templateId){return applyLibraryDesign(c);}else{var dlg=a.FileDialog(3);dlg.Title='选择要套用的模板或演示文稿';dlg.AllowMultiSelect=false;try{dlg.Filters.Clear();dlg.Filters.Add('模板和演示文稿','*.potx;*.pptx;*.dpt;*.dps;*.ppt;*.pot');}catch(ignore){}if(dlg.Show()!==-1)return '已取消。';p=String(dlg.SelectedItems.Item(1));}
    if(!L.ask('将把“'+p.split(/[\\\/]/).pop()+'”的母版（背景、配色、字体、版式）套用到全部 '+d.Slides.Count+' 页。各页的文字和图片不动，但位置可能跟着新版式变化。是否执行？'+KEEP))return '已取消。';
    try{d.ApplyTemplate(p);}catch(e){throw Error('没能套用这个模板：'+String(e&&e.message||e).slice(0,80)+'。可以改用“设计 → 导入模板”。');}
    return '已套用模板，共 '+d.Slides.Count+' 页。不满意可以关闭不保存，或用备份恢复。';}
  // ---- 图片和图表自动排版：按网格排整齐、大小协调、间距一致；页上只有一段文字时排成“左文右图” ----
  function isVisual(sh){var t=-1;try{t=Number(sh.Type);}catch(e){}if(t===13||t===11||t===3||t===24||t===16)return true;
    if(t===14||t===7){try{if(sh.HasChart===true||sh.HasChart===-1)return true;}catch(e){}try{var ct=Number(sh.PlaceholderFormat.ContainedType);if(ct===13||ct===3||ct===11)return true;}catch(e){}}return false;}
  function isTitleShape(sh){try{var t=Number(sh.PlaceholderFormat.Type);return t===1||t===3||t===5;}catch(e){return false;}}
  function hasText(sh){try{return !!(sh.HasTextFrame&&String(sh.TextFrame.TextRange.Text||'').trim());}catch(e){return false;}}
  // 选列数：每种列数都算一遍，让图片按原比例放大后占的面积最大。
  function bestGrid(items,aw,ah,gap,cols){var best=null;for(var c=1;c<=items.length;c++){if(cols&&c!==cols)continue;var r=Math.ceil(items.length/c),cw=(aw-gap*(c-1))/c,ch=(ah-gap*(r-1))/r;if(cw<20||ch<20)continue;var area=0;items.forEach(function(it){var k=Math.min(cw/it.w,ch/it.h);area+=it.w*k*it.h*k;});if(!best||area>best.area+0.5)best={c:c,r:r,cw:cw,ch:ch,area:area};}return best;}
  function layoutSlide(d,slide,cols){var W=Number(d.PageSetup.SlideWidth),H=Number(d.PageSetup.SlideHeight),m=W*0.05,gap=W*0.02,shapes=slide.Shapes,vis=[],texts=[],title=null;
    for(var k=1;k<=shapes.Count;k++){var sh=shapes.Item(k);if(isVisual(sh))vis.push({sh:sh,w:Math.max(1,Number(sh.Width)),h:Math.max(1,Number(sh.Height)),t:Number(sh.Top),l:Number(sh.Left)});else if(isTitleShape(sh)){if(hasText(sh))title=sh;}else if(hasText(sh))texts.push(sh);}
    if(!vis.length)return null;
    var top=title?Number(title.Top)+Number(title.Height)+H*0.03:H*0.08,bottom=H-H*0.06,left=m,right=W-m,mode='grid';
    if(texts.length===1&&vis.length<=4){var tx=texts[0];tx.Left=m;tx.Top=top;tx.Width=W*0.42;tx.Height=bottom-top;left=W*0.5;mode='side';}
    else if(texts.length){var low=0;texts.forEach(function(x){low=Math.max(low,Number(x.Top)+Number(x.Height));});if(bottom-(low+gap)>=(bottom-top)*0.4)top=Math.max(top,low+gap);}
    vis.sort(function(x,y){return Math.abs(x.t-y.t)>20?x.t-y.t:x.l-y.l;});
    var g=bestGrid(vis,right-left,bottom-top,gap,cols)||bestGrid(vis,right-left,bottom-top,gap,0);if(!g)return null;
    vis.forEach(function(it,i){var r=Math.floor(i/g.c),c=i%g.c,inRow=Math.min(g.c,vis.length-r*g.c),rowW=inRow*g.cw+(inRow-1)*gap,x0=left+(right-left-rowW)/2,k=Math.min(g.cw/it.w,g.ch/it.h),w=it.w*k,h=it.h*k;
      try{it.sh.LockAspectRatio=0;}catch(e){}it.sh.Width=w;it.sh.Height=h;it.sh.Left=x0+c*(g.cw+gap)+(g.cw-w)/2;it.sh.Top=top+r*(g.ch+gap)+(g.ch-h)/2;try{it.sh.LockAspectRatio=-1;}catch(e){}});
    return {count:vis.length,cols:g.c,rows:g.r,mode:mode};}
  function arrangeVisuals(c){var a=c.a,d=a.ActivePresentation,o=c.options||{},all=/^(全部|所有|all|すべて|全て)$/i.test(String(o.scope||'').trim()),cols=Number(o.cols)||0;if(cols&&!(cols>=1&&cols<=6))throw Error('列数请填 1～6。');
    if(!all){var s=currentSlide(a,d);if(!s)throw Error('请先点一下要排版的那一页。');var n=0;for(var k=1;k<=s.Shapes.Count;k++)if(isVisual(s.Shapes.Item(k)))n++;if(!n)throw Error('这一页没有图片或图表。');
      if(!L.ask('将把第 '+s.SlideIndex+' 页的 '+n+' 张图片 / 图表按原比例排整齐（大小协调、间距一致）。是否执行？'+KEEP))return '已取消。';
      var r=layoutSlide(d,s,cols);return '已排好第 '+s.SlideIndex+' 页的 '+r.count+' 张图片 / 图表（'+(r.mode==='side'?'左边文字、右边图片，':'')+r.rows+' 行 × '+r.cols+' 列）。';}
    var pages=0,total=0;for(var i=1;i<=d.Slides.Count;i++){var sl=d.Slides.Item(i);for(var q=1;q<=sl.Shapes.Count;q++)if(isVisual(sl.Shapes.Item(q))){pages++;break;}}if(!pages)throw Error('演示里没有图片或图表。');
    if(!L.ask('将把 '+pages+' 页里的图片 / 图表按原比例排整齐。没有图片的页不动。是否执行？'+KEEP))return '已取消。';
    pages=0;for(var j=1;j<=d.Slides.Count;j++){var x=layoutSlide(d,d.Slides.Item(j),cols);if(x){pages++;total+=x.count;}}
    return '已排好 '+pages+' 页，共 '+total+' 张图片 / 图表。没有图片的页没有改动。';}
  T.register({id:'unifyDeckFont',label:'全部页统一字体',desc:'所有页的文字框和表格用同一种字体，字号不变',icon:'type',group:'slide',undo:false,fields:[{key:'font',label:'字体（不填就用现在最常见的那种）',placeholder:'例如：微软雅黑',required:false}],hosts:{wpp:unifyDeckFont}});
  T.register({id:'removeAnimations',label:'删除全部动画',desc:'删除所有页的动画和切换效果',icon:'trash',group:'slide',undo:false,hosts:{wpp:removeAnimations}});
  T.register({id:'unifyTitles',label:'统一各页标题',desc:'各页标题位置、字号、颜色和第一页一样',icon:'heading1',group:'slide',undo:false,fields:[{key:'size',label:'标题字号（不填按第一页）',placeholder:'例如：32',required:false}],hosts:{wpp:unifyTitles}});
  T.register({id:'slideNumbers',label:'加页码',desc:'右下角页码，调整顺序自动更新',icon:'list',group:'slide',undo:false,fields:[{key:'skipFirst',label:'封面不显示页码（填“否”则显示）',placeholder:'是',required:false}],hosts:{wpp:slideNumbers}});
  T.register({id:'bodySpacing',label:'统一正文行距',desc:'所有页正文行距一致，读起来不挤',icon:'spacing',group:'slide',undo:false,fields:[{key:'line',label:'行距（倍）',placeholder:'不填为 1.2',required:false}],hosts:{wpp:bodySpacing}});
  T.register({id:'arrangeVisuals',label:'图片图表自动排版',desc:'本页的图片、图表按原比例排整齐，间距一致',icon:'picture',group:'slide',undo:false,fields:[{key:'scope',label:'处理范围（填“全部”排所有页）',placeholder:'不填只排当前页',required:false},{key:'cols',label:'列数（不填自动）',placeholder:'例如：2',required:false}],hosts:{wpp:arrangeVisuals}});
  T.register({id:'applyDesign',label:'套用模板',desc:'换上另一个模板的背景、配色和版式',icon:'presentation',group:'slide',undo:false,hosts:{wpp:applyDesign}});
  T.register({id:'removeNotes',label:'删除全部备注',desc:'清空所有页的演讲者备注，发给别人前用',icon:'clean',group:'slide',undo:false,hosts:{wpp:removeNotes}});
  T.register({id:'matchFormat',labels:{wpp:'匹配本页格式'},descs:{wpp:'粘贴的文字或表格跟随本页'},hosts:{wpp:function(c){return matchSlide(c.a);}}});
})(window);

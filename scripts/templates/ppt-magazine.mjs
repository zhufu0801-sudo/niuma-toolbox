// Native, editable editorial layouts. Illustrations are composed from PowerPoint shapes.
export function magazineSlides(ppt,t,lang,gradient){
 const ja=lang==='ja',font=ja?'Yu Gothic':'Microsoft YaHei',q=(cn,jp)=>ja?jp:cn;
 const W=13.333,H=7.5,C={ink:t.bg,paper:t.id==='year-review'?'F3E9DF':'F3F2E9',accent:t.accent,secondary:t.secondary,muted:'667576'};
 const luminance=c=>c.match(/../g).map(x=>parseInt(x,16)/255).map(x=>x<=.04045?x/12.92:((x+.055)/1.055)**2.4).reduce((sum,x,i)=>sum+x*[.2126,.7152,.0722][i],0);
 const contrast=(a,b)=>(Math.max(luminance(a),luminance(b))+.05)/(Math.min(luminance(a),luminance(b))+.05);
 const onColor=c=>contrast(C.ink,c)>=4.5?C.ink:contrast('FFFFFF',c)>=4.5?'FFFFFF':'000000';
 const secondaryText=contrast(C.secondary,C.paper)>=4.5?C.secondary:C.ink;
 const names=['cover','agenda','chapter','points','metric','timeline','compare','picture','quote','closing'];
 const themes={
  'work-report':{issue:q('增长特刊','成長特集'),cover:q('工作汇报','業務報告'),sub:q('把进展，变成下一次突破。','進捗を、次の飛躍へ。'),tag:q('观察 / 行动 / 证据','観察 / 行動 / 根拠'),small:q('团队协作的四个切面','チーム連携の4つの視点'),a:q('效率','効率'),b:q('协作','連携'),c:q('体验','体験')},
  'year-review':{issue:q('年度档案','年間アーカイブ'),cover:q('年终总结','年間活動報告'),sub:q('把一年，读成新的可能。','一年の歩みから、次の可能性へ。'),tag:q('足迹 / 成果 / 远见','歩み / 成果 / 展望'),small:q('这一年，值得记录的变化','この一年、記録したい変化'),a:q('积累','蓄積'),b:q('突破','飛躍'),c:q('成长','成長')},
  pitch:{issue:q('创新专刊','イノベーション特集'),cover:q('路演融资','事業提案'),sub:q('下一种可能，现在发生。','次の可能性は、ここから。'),tag:q('洞察 / 产品 / 增长','洞察 / 製品 / 成長'),small:q('从真实需求，到规模化价值','実際の需要から、広がる価値へ'),a:q('需求','需要'),b:q('产品','製品'),c:q('增长','成長')}
 };const copy=themes[t.id]||{issue:t[lang],cover:t[lang],sub:q('从清晰的想法，到具体的行动。','明確な考えから、具体的な行動へ。'),tag:q('背景 / 重点 / 下一步','背景 / ポイント / 次の一歩'),small:q('用事实、图表与案例说明观点','事実・図表・事例で考えを伝える'),a:q('发现','発見'),b:q('选择','選択'),c:q('行动','行動')},motions=[];
 for(const [n,name] of names.entries())ppt.defineSlideMaster({title:name,background:{color:C.paper},objects:[{placeholder:{options:{name:'heading',objectName:'headline',type:'title',x:.55,y:n===0?1.75:n===9?1.15:1.0,w:n===0?7.2:11.9,h:n===0?1.15:1.05,fontFace:font,fontSize:n===0?(ja?47:62):n===8?27:n===9?43:34,bold:true,color:[0,2,4,8,9].includes(n)?C.paper:C.ink,align:'left',margin:0},text:''}}]});
 for(let i=0;i<10;i++){
  const dark=[0,2,4,8,9].includes(i),ink=dark?C.paper:C.ink,bg=dark?C.ink:C.paper,muted=dark?t.muted:C.muted;
  const s=ppt.addSlide({masterName:names[i]});s.background={color:bg};if(t.gradient&&dark)s.background={data:'image/png;base64,'+gradient.toString('base64')};
  const rect=(x,y,w,h,c,opts={})=>s.addShape(ppt.ShapeType.rect,{x,y,w,h,fill:{color:c},line:{color:c,transparency:100},...opts});
  const line=(x,y,w,c=ink,width=1,h=0)=>s.addShape(ppt.ShapeType.line,{x,y,w,h,line:{color:c,width}});
  const txt=(v,x,y,w,h,size=18,color=ink,opts={})=>s.addText(v,{x,y,w,h,fontFace:font,fontSize:size,color,margin:0,align:'left',vertAnchor:'mid',breakLine:false,...opts});
  const ellipse=(x,y,w,h,c,opts={})=>s.addShape(ppt.ShapeType.ellipse,{x,y,w,h,fill:{color:c},line:{color:c,transparency:100},...opts});
  const heading=(v,opts={})=>s.addText(v,{placeholder:'heading',objectName:'headline',fontFace:font,fontSize:34,bold:true,color:ink,align:'left',margin:0,...opts});
  const tag=(v,x,y,w,c=C.accent)=>{rect(x,y,w,.34,c);txt(v,x+.10,y+.04,w-.2,.22,10,onColor(c),{bold:true});};
  const caption=(v,x,y,w)=>txt(v,x,y,w,.4,10,muted);
  const circleNote=(v,x,y,c=C.accent)=>{ellipse(x,y,1.0,1.0,c);txt(v,x+.08,y+.1,.84,.8,12,C.ink,{bold:true,align:'center'});};
  // Collage visual: three editable, staggered panels with charts, metrics and annotations.
  function collage(x,y,scale=1,variant=0){
   const z=(n)=>n*scale;
   rect(x+z(.25),y+z(.18),z(4.7),z(3.9),C.secondary,{rotate:4,transparency:0});
   rect(x,y,z(4.65),z(3.9),C.paper,{rotate:358});
   txt(q('观察笔记','観察ノート'),x+z(.22),y+z(.2),z(2.8),z(.3),12*scale,C.ink,{bold:true});
   line(x+z(.22),y+z(.7),z(4.15),C.ink,1);
   if(variant===0){
    [1.05,1.75,2.3,2.8].forEach((h,k)=>{rect(x+z(.4+k*.93),y+z(3.55-h),z(.64),z(h),k===3?C.secondary:C.ink);txt(['Q1','Q2','Q3','Q4'][k],x+z(.35+k*.93),y+z(3.62),z(.8),z(.2),9*scale,C.ink,{align:'center'});});
    rect(x+z(2.8),y+z(.88),z(2.3),z(.8),C.accent);txt('+32%',x+z(3.0),y+z(1.0),z(1.9),z(.5),29*scale,C.ink,{bold:true});
   }else if(variant===1){
    ellipse(x+z(.55),y+z(1),z(2.3),z(2.3),C.secondary);ellipse(x+z(1),y+z(1.45),z(1.4),z(1.4),C.paper);
    txt('2026',x+z(2.45),y+z(1.1),z(2.05),z(.8),29*scale,C.ink,{bold:true});
    [q('关键时刻','重要な瞬間'),q('共同成果','チームの成果'),q('未来方向','次の方向性')].forEach((v,k)=>{line(x+z(2.85),y+z(2.25+k*.48),z(1.4),C.ink);txt(v,x+z(2.85),y+z(2.05+k*.48),z(1.6),z(.23),9*scale,C.ink);});
   }else{
    rect(x+z(.35),y+z(.95),z(3.95),z(2.5),C.ink);rect(x+z(.55),y+z(1.13),z(.65),z(2.1),C.secondary);
    [0,1,2].forEach(k=>rect(x+z(1.4+k*.9),y+z(1.28),z(.72),z(.55),k===0?C.accent:'BBA1D7'));
    [1,.65,1.35,1.05,1.65].forEach((h,k)=>rect(x+z(1.45+k*.49),y+z(3.15-h*.65),z(.32),z(h*.65),C.accent));
   }
   rect(x-z(.42),y+z(2.85),z(2.55),z(1.45),C.accent,{rotate:357});
   txt(q('现场观察','現場から'),x-z(.18),y+z(3.02),z(1.95),z(.3),12*scale,C.ink,{bold:true});
   txt(q('把变化看见\n把价值讲清','変化を捉える\n価値を伝える'),x-z(.18),y+z(3.5),z(1.95),z(.62),14*scale,C.ink,{bold:true});
   s.addShape(ppt.ShapeType.chevron,{x:x+z(3.8),y:y-z(.35),w:z(1.25),h:z(.65),fill:{color:C.accent},line:{transparency:100,color:C.accent},rotate:335,objectName:'focusGraphic'});
  }
  // Editorial furniture shared across all layouts.
  txt('OFFICE / '+(ja?'編集可能なテンプレート':'可编辑模板'),.55,.2,7,.22,9,muted,{charSpacing:1.8});
  txt('2026   ·   '+copy.issue,9,.2,3.8,.22,9,muted,{align:'right'});line(.55,.62,12.23,ink,.7);
  line(.55,7.04,12.23,ink,.6);txt(copy.issue+' / '+t[lang],.55,7.15,7,.18,8,muted);txt(String(i+1).padStart(2,'0')+' / 10',11.3,7.13,1.48,.23,9,muted,{align:'right'});
  if(i===0){
   tag(q('封面故事 · 示例','巻頭特集 · サンプル'),.55,1.05,2.9);heading(copy.cover,{x:.55,y:1.75,w:7.2,h:1.15,fontSize:ja?47:62});
   txt(copy.sub,.6,3.13,6.5,1.0,27,ink,{bold:true});line(.6,4.5,1.1,C.accent,5);
   txt(copy.small,.6,4.86,5.6,.55,20,ink);txt(copy.tag,.6,5.6,5.4,.5,14,muted);txt(q('2026 年度 / 团队名称 / 汇报人','2026年度 / チーム名 / 発表者'),.6,6.45,5.8,.25,11,muted);
   collage(7.7,1.78,1, t.variant);
   caption(q('图示为可编辑示例，非真实业绩。','図は編集可能なサンプルです。'),8.2,6.35,4.4);
  }
  if(i===1){
   heading(q('这一期，我们讨论什么','今号で取り上げること'));txt('01—04',.55,2.35,4.3,1.1,53,C.ink,{bold:true});
   txt(q('先看全局，\n再读关键细节。','全体を捉え、\n大切な細部へ。'),.6,3.8,4.05,1.3,29,C.ink,{bold:true});
   rect(.55,5.65,4.1,.8,C.secondary);txt(copy.tag,.75,5.87,3.7,.3,13,onColor(C.secondary));
   const items=[[q('方向','方向性'),q('目标、背景与判断依据','目標・背景・判断の根拠')],[q('证据','エビデンス'),q('关键成果与趋势数据','主要な成果とデータの推移')],[q('选择','選択肢'),q('比较方案，找到取舍','案を比較し、優先順位を決める')],[q('行动','アクション'),q('把下一步交给具体的人','次の行動と担当者を決める')]];
   items.forEach(([a,b],k)=>{const y=2.25+k*1.08;txt('0'+(k+1),5.1,y,1,.6,30,secondaryText,{bold:true});txt(a,6.3,y,5.7,.46,23,C.ink,{bold:true});txt(b,6.3,y+.52,5.9,.35,13,C.muted);line(5.1,y+.94,7.4,C.muted,.5);});
  }
  if(i===2){
   heading(q('从一个好问题开始','良い問いから始める'),{x:.55,y:1.05,w:7.5,h:1});
   txt('01',.42,2.1,5,2.55,156,C.accent,{bold:true});
   txt(q('变化发生在哪里？\n我们为什么现在行动？','変化はどこで起きているか。\nなぜ今、行動するのか。'),5.35,2.4,7.2,1.4,29,ink,{bold:true});
   [q('背景：正在发生的变化','背景：起きている変化'),q('约束：时间、资源和能力','制約：時間・資源・能力'),q('标准：怎样才算做得更好','基準：何をもって改善とするか')].forEach((v,k)=>{line(5.4,4.4+k*.6,6.9,muted,.6);txt(v,5.4,4.5+k*.6,6.8,.36,15,muted);});
   tag(q('章节页 / 用问题建立语境','章扉 / 問いで背景を共有'),.6,6.35,4.7);
  }
  if(i===3){
   heading(q('三条线索，拼出完整答案','3つの視点から、答えを描く'));
   const titles=[copy.a,copy.b,copy.c],desc=[q('把重复工作变成清晰流程。\n让时间用在真正重要的地方。','繰り返す作業を整理する。\n重要なことに時間を使う。'),q('用共同目标连接不同角色。\n把反馈变成持续改进的依据。','共通の目標で役割をつなぐ。\n改善にフィードバックを生かす。'),q('让每一次使用都更容易。\n用具体的结果验证价值。','毎回の利用をもっと簡単に。\n具体的な結果で価値を確かめる。')];
   [0,1,2].forEach(k=>{const x=.55+k*4.18,y=k===1?2.6:2.2;rect(x,y,3.85,3.9,k===1?C.ink:k===0?C.accent:'E6DAD0');const ci=k===1?C.paper:C.ink;
    txt('0'+(k+1),x+.24,y+.2,1.4,.6,35,ci,{bold:true});
    if(k===0)[.4,.8,1.15].forEach((h,j)=>rect(x+2+j*.43,y+1.6-h,.28,h,C.ink));
    if(k===1){ellipse(x+2.1,y+.45,.7,.7,C.accent);ellipse(x+2.7,y+.8,.7,.7,C.secondary);line(x+2.3,y+1.15,.6,C.paper,2);}
    if(k===2){s.addShape(ppt.ShapeType.chevron,{x:x+2,y:y+.5,w:1.2,h:.85,fill:{color:C.secondary},line:{color:C.secondary,transparency:100}});}
    txt(titles[k],x+.24,y+1.85,3.3,.52,27,ci,{bold:true});txt(desc[k],x+.24,y+2.7,3.3,.95,15,ci,{paraSpaceAfterPt:8});});
  }
  if(i===4){
   heading(q('让成果可见，让趋势可读','成果を見せ、推移を伝える'));
   txt('+32%',.55,2.35,6.2,1.65,96,C.accent,{bold:true});txt(q('核心指标的同比提升','主要指標の前年比'),.65,4.15,5.1,.55,23,ink,{bold:true});
   line(.65,5.05,4.9,muted,.7);txt(q('变化来自：\n流程优化 × 团队协作 × 持续迭代','変化の要因：\n業務改善 × チーム連携 × 継続的な改善'),.65,5.35,5.3,1.0,17,muted);
   rect(6.45,2.15,6.33,4.42,C.paper);txt(q('季度观察','四半期レビュー'),6.7,2.4,3.5,.36,17,C.ink,{bold:true});
   s.addChart(ppt.ChartType.bar,[{name:q('示例指标','サンプル指標'),labels:['Q1','Q2','Q3','Q4'],values:[48,62,78,96]}],{x:6.7,y:2.95,w:5.7,h:2.85,barDir:'col',catAxisLabelFontFace:font,catAxisLabelFontSize:12,catAxisLabelColor:C.ink,valAxisLabelColor:C.muted,valAxisLabelFontSize:10,showLegend:false,showTitle:false,showValue:true,dataLabelColor:C.ink,dataLabelPosition:'outEnd',dataLabelFormatCode:'0',chartColors:[C.secondary],showBorder:false,catAxisLineShow:false,valAxisLineShow:false,valGridLine:{color:'D6D6CF',width:.6},showShadow:false,gapSize:70});
   txt(q('数据均为示例，请替换为实际业绩。','数値は例です。実績に差し替えてください。'),6.75,6.02,5.5,.4,10,C.muted);
  }
  if(i===5){
   heading(q('四个时刻，连接下一段旅程','4つの節目で、次の歩みをつなぐ'));
   rect(.55,2.3,2.75,4.22,C.ink);txt('2026',.8,2.6,2.2,.7,37,C.paper,{bold:true});txt(q('行动日历','行動計画'),.8,3.8,2.2,1,27,C.accent,{bold:true});txt(q('负责人 / 交付物\n每个节点都可追踪','担当者 / 成果物\n節目ごとに確認する'),.8,5.3,2.2,.8,13,C.paper);
   const labels=[q('确定方向','方針を決める'),q('小步试点','小さく試す'),q('扩大验证','検証を広げる'),q('复盘迭代','振り返り改善')];
   for(let k=0;k<4;k++){const x=3.7+k*2.3,y=k%2?3.0:2.4;line(x+.16,5.95,1.95,C.secondary,2);ellipse(x,5.8,.3,.3,C.secondary);txt('Q'+(k+1),x,y,2,.65,35,C.ink,{bold:true});txt(labels[k],x,y+.9,2.05,.8,20,C.ink,{bold:true});txt(q('目标 / 产出\n填写具体里程碑','目標 / 成果\n具体的な節目を記入'),x,y+1.95,2.05,.85,12,C.muted);}
  }
  if(i===6){
   heading(q('把取舍，放在同一张桌面上','同じ基準で、選択肢を比べる'));
   const rows=[q('投入资源','投入資源'),q('验证周期','検証期間'),q('预期收益','期待する効果'),q('关键风险','主なリスク')];
   rect(3.4,2.25,4.3,.75,C.accent);rect(8.15,2.25,4.63,.75,C.ink);txt(q('A / 稳步改进','A / 段階的な改善'),3.65,2.45,3.8,.3,18,C.ink,{bold:true});txt(q('B / 集中突破','B / 集中的な改革'),8.4,2.45,4.1,.3,18,C.paper,{bold:true});
   rows.forEach((v,k)=>{const y=3.25+k*.72;txt('0'+(k+1),.6,y,.55,.35,13,secondaryText,{bold:true});txt(v,1.3,y,1.9,.35,15,C.ink,{bold:true});txt(q(['现有团队','短周期试点','稳定改善','收益增幅有限'][k],['既存チーム','短期間で試行','安定した改善','効果の幅が限定的'][k]),3.65,y,3.8,.4,16,C.ink);txt(q(['专项投入','分阶段推进','规模化增长','执行复杂度较高'][k],['専任で投入','段階的に推進','規模の拡大','実行の複雑さ'][k]),8.4,y,4.1,.4,16,C.ink);line(.6,y+.56,12.15,'C3C9C3',.6);});
   tag(q('选择依据：收益、成本与可逆性','判断基準：効果・費用・やり直しやすさ'),.6,6.45,7.4,C.secondary);
  }
  if(i===7){
   heading(q('从现场，读懂真实价值','利用場面から、本当の価値へ'));
   collage(1.0,2.43,.86,t.variant);
   tag(q('案例观察 / 可替换图片区域','事例 / 画像に差し替え可能'),7,2.28,5.7,C.accent);
   txt(q('改变发生在\n每一个细节里。','変化は、\n細部から生まれる。'),7,3,5.7,1.35,32,C.ink,{bold:true});
   txt(q('谁在使用？在什么场景？\n用两句话说明实际问题，\n再用图片和数据给出证据。','誰が、どこで使うのか。\n実際の課題を短く説明し、\n画像とデータで根拠を示します。'),7,4.72,5.6,1.25,18,C.muted);
   caption(q('原生图解可编辑，也可替换成现场照片。','図は編集可能。現場の写真にも差し替えられます。'),.65,6.55,6.2);
  }
  if(i===8){
   heading(q('留下一句话，带走一个方向','ひとつの言葉で、次の方向へ'),{fontSize:27});
   rect(.7,2.35,1.05,3.8,C.accent);txt('“',2.1,1.95,1.6,1.2,96,C.accent,{bold:true});
   txt(q('好的计划，\n让每个人都知道\n下一步往哪里走。','良い計画は、\n一人ひとりの\n次の一歩を照らす。'),3.1,2.7,8.9,2.6,42,ink,{bold:true});
   line(3.15,5.9,2,C.secondary,3);txt(q('团队共识 / 用自己的观点替换','チームの共通認識 / 自分たちの言葉に差し替え'),3.15,6.2,9,.4,13,muted);
  }
  if(i===9){
   heading(q('下一页，由我们一起写。','次のページを、ともに。'),{x:.55,y:1.15,w:11.9,h:1.1,fontSize:43});
   txt(q('感谢聆听','ご清聴ありがとうございました'),.6,2.8,9,.75,ja?30:44,C.accent,{bold:true});
   const cards=[q('01 / 明确优先级','01 / 優先順位を決める'),q('02 / 确定负责人','02 / 担当者を決める'),q('03 / 约定下一次','03 / 次の予定を決める')];
   cards.forEach((v,k)=>{const fill=k===1?C.secondary:C.paper;rect(.6+k*4.13,4.25,3.83,1.25,fill);txt(v,.8+k*4.13,4.65,3.43,.45,17,onColor(fill),{bold:true});});
   txt(q('姓名 / 团队 / 联系方式','氏名 / チーム / 連絡先'),.65,6.35,9,.35,14,muted);
  }
  s.addNotes(q('设计杂志风可编辑样板。所有数据、年份与业务描述均为示例。图解由原生形状组成。替换为真实内容后请检查换行、图表数据与对比度。','編集可能な雑誌風サンプルです。数値・年・業務内容は例です。図は標準図形で構成されています。差し替え後は改行・データ・コントラストを確認してください。'));
  motions.push({transition:{type:'fade',speed:'fast'},items:[{name:'headline',effect:'fade',dur:400},...(i===0?[{name:'focusGraphic',effect:'zoom',dur:600,delay:120}]:[])]});
 }
 return motions;
}

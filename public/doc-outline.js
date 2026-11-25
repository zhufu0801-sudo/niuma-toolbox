// 办公工具箱 · Word 转 PPT 的纯计算部分（WPS 文字版和 Microsoft Word 版共用）。
// 输入段落列表 [{text, level(1-9 是标题级别，10 是正文), style}]，按标题层级整理成“PPT 版式工作室”的大纲格式。
(function(root){
  'use strict';
  var MAX_SECTIONS=12,MAX_POINTS=10,POINT_LEN=60;
  function clean(s){return String(s==null?'':s).replace(/[\r\x07\x0b\x0c]/g,' ').replace(/[ \t　 ]+/g,' ').trim();}
  function isToc(style){return /^(TOC|目录|目次)\s*\d/i.test(String(style||''));}
  function isTitleStyle(style){return /^(Title|标题|表題|タイトル)$/i.test(String(style||'').trim());}
  // 没有设置标题样式时，按编号认标题：“一、”“第一章”是 1 级，“（一）”是 2 级，“1.”“1、”是 3 级。
  function numbered(t){
    if(t.length>40||/[。！？；!?;]$/.test(t))return 0;
    if(/^第[一二三四五六七八九十百\d０-９]+[章节部分篇節]/.test(t)||/^[一二三四五六七八九十]+[、．.]\s*\S/.test(t))return 1;
    if(/^[（(][一二三四五六七八九十]+[)）]\s*\S/.test(t))return 2;
    if(/^[0-9０-９]{1,2}[、．.](?![0-9０-９])\s*\S/.test(t)||/^[■●◆]\s?\S/.test(t))return 3;
    return 0;
  }
  function shorten(t){var first=(t.match(/^[\s\S]*?[。！？!?](?=\s*\S)/)||[t])[0].trim();if(first.length<t.length&&first.length>=8)t=first;return t.length>POINT_LEN?t.slice(0,POINT_LEN-1)+'…':t;}
  function strip(h){return h.replace(/^(第[一二三四五六七八九十百\d０-９]+[章节部分篇節]|[一二三四五六七八九十]+[、．.]|[（(][一二三四五六七八九十]+[)）]|[0-9０-９]{1,2}[、．.]|[■●◆])\s*/,'').trim()||h;}
  function build(paras,o){
    o=o||{};var list=[];(paras||[]).forEach(function(p){var t=clean(p.text);if(!t||isToc(p.style))return;var lv=Number(p.level)||10;list.push({text:t,level:lv>=1&&lv<=9?lv:10,style:String(p.style||'')});});
    if(!list.length)throw Error('文档里没有文字。');
    var title='',i0=0;
    var ti=list.findIndex(function(p){return isTitleStyle(p.style);});
    if(ti>=0){title=list[ti].text;list.splice(ti,1);}
    var styled=list.some(function(p){return p.level<=9;});
    if(!styled)list.forEach(function(p){var n=numbered(p.text);if(n)p.level=n;});
    var levels=list.filter(function(p){return p.level<=9;}).map(function(p){return p.level;});
    if(!levels.length)throw Error('没有找到标题：请先给各节标题设置“标题 1”“标题 2”样式，或用“一、”“（一）”“1.”这样的编号开头。');
    var top=Math.min.apply(null,levels);
    // 第一个标题之前：短的第一段当标题（没有标题样式时），其余第一段当副标题。
    var first=list.findIndex(function(p){return p.level===top;}),before=list.slice(0,first);
    if(!title&&before.length&&before[0].text.length<=40){title=before.shift().text;}
    var subtitle=before.length?shorten(before[0].text).slice(0,150):'';
    if(!title)title=clean(o.name||'').replace(/\.[^.]+$/,'')||'演示文稿';
    if(title.length>80)title=title.slice(0,79)+'…';
    var sections=[],cur=null,skipped=0;
    list.slice(first).forEach(function(p){
      if(p.level===top){cur={heading:strip(p.text).slice(0,60),points:[]};sections.push(cur);return;}
      if(!cur||cur.points.length>=MAX_POINTS)return;
      var t=p.level<=9?strip(p.text):shorten(p.text);t=t.replace(/^>+\s*/,'');if(t&&cur.points.indexOf(t)<0)cur.points.push(t);
    });
    var kept=sections.filter(function(s){if(s.points.length)return true;skipped++;return false;});
    if(!kept.length)throw Error('标题下面都没有内容，没法生成页面。请在各标题下写上要点或正文。');
    var cut=Math.max(0,kept.length-MAX_SECTIONS);kept=kept.slice(0,MAX_SECTIONS);
    var outline=kept.map(function(s){return s.heading+'\n'+s.points.map(function(x){return '- '+x;}).join('\n');}).join('\n\n');
    return {title:title,subtitle:subtitle,sections:kept,outline:outline,skipped:skipped,cut:cut,numbered:!styled};
  }
  function summary(r){return '整理出 '+r.sections.length+' 节：'+r.sections.slice(0,5).map(function(s){return '“'+s.heading+'”';}).join('、')+(r.sections.length>5?'…':'')+'。'+(r.skipped?'\n'+r.skipped+' 个标题下面没有内容，已跳过。':'')+(r.cut?'\n超过 '+MAX_SECTIONS+' 节，后面 '+r.cut+' 节没有放进去。':'');}
  root.OfficeDocOutline={build:build,summary:summary,numbered:numbered,MAX_SECTIONS:MAX_SECTIONS};
})(typeof window!=='undefined'?window:globalThis);

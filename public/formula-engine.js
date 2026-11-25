/* Deterministic formula planner. No model, eval, macro execution or network. */
(function(root){
  'use strict';
  var TYPES={lookup:'跨表查找（兼容 VLOOKUP / INDEX）',xlookup:'跨表查找（XLOOKUP）',sumif:'条件求和',countif:'条件计数',averageif:'条件平均',if:'条件判断',sum:'区域求和',average:'区域平均',count:'统计非空',trim:'清理多余空格',left:'提取左边文字',right:'提取右边文字',round:'四舍五入',year:'提取年份',month:'提取月份',day:'提取日期中的日',concat:'合并两列文字',datediff:'计算相差天数',reverse:'反向查找（结果列在左边）',multilookup:'多条件查找',onetomany:'一对多查找（列出全部结果）',sumifs:'多条件求和',distinctcount:'去重计数',rank:'排名（不打乱顺序）',age:'计算年龄 / 工龄（整年）',weekday:'日期转星期'};
  function col(n){var s='';while(n>0){n--;s=String.fromCharCode(65+n%26)+s;n=Math.floor(n/26);}return s;}
  function number(s){var n=0;for(var i=0;i<s.length;i++)n=n*26+s.charCodeAt(i)-64;return n;}
  function quoted(s){return '"'+String(s==null?'':s).replace(/"/g,'""')+'"';}
  function sheetName(s){if(!s||s.length>31||/[\[\]:*?\/\\\x00-\x1f]/.test(s))throw Error('工作表名称无效。');return "'"+s.replace(/'/g,"''")+"'";}
  function ref(text,defaultSheet){
    text=String(text||'').trim();var sheet=defaultSheet,area=text,match;
    if(text[0]==="'"){match=text.match(/^'((?:[^']|'')+)'!(.+)$/);if(!match)throw Error('区域格式应为：\'员工表\'!A2:D100。');sheet=match[1].replace(/''/g,"'");area=match[2];}
    else if(text.indexOf('!')>=0){var parts=text.split('!');if(parts.length!==2)throw Error('区域中只能有一个工作表引用。');sheet=parts[0];area=parts[1];}
    sheetName(sheet);match=area.toUpperCase().match(/^\$?([A-Z]{1,3})\$?([1-9][0-9]{0,6})(?::\$?([A-Z]{1,3})\$?([1-9][0-9]{0,6}))?$/);
    if(!match)throw Error('请填写具体区域，例如 A2:A100；暂不使用整列、名称或跨文件引用。');
    var out={sheet:sheet,c1:number(match[1]),r1:Number(match[2]),c2:number(match[3]||match[1]),r2:Number(match[4]||match[2])};
    if(out.c2<out.c1||out.r2<out.r1||out.c2>16384||out.r2>1048576)throw Error('区域边界无效。');
    out.rows=out.r2-out.r1+1;out.cols=out.c2-out.c1+1;return out;
  }
  function address(r,absolute){var p=absolute===false?'':'$';return p+col(r.c1)+p+r.r1+(r.c1===r.c2&&r.r1===r.r2?'':':'+p+col(r.c2)+p+r.r2);}
  function full(r){return sheetName(r.sheet)+'!'+address(r);}
  function cell(r,i){return sheetName(r.sheet)+'!$'+col(r.c1)+(r.r1+i);}
  function slice(r,c){return {sheet:r.sheet,r1:r.r1,r2:r.r2,c1:r.c1+c-1,c2:r.c1+c-1};}
  function intersects(a,b){return a.sheet.toLowerCase()===b.sheet.toLowerCase()&&a.r1<=b.r2&&a.r2>=b.r1&&a.c1<=b.c2&&a.c2>=b.c1;}
  function integer(v,min,max,label){v=Number(v);if(!Number.isInteger(v)||v<min||v>max)throw Error(label+'应为 '+min+'～'+max+' 的整数。');return v;}
  function literal(v){v=String(v==null?'':v);if(v.length>250)throw Error('条件或返回文字最多 250 字。');return /^-?(?:0|[1-9]\d*)(?:\.\d+)?$/.test(v)&&Number.isFinite(Number(v))?v:quoted(v);}
  function infer(text){
    text=String(text||'').trim().toLowerCase();var kinds=[];
    var rules=[['multilookup',/多条件.*(查找|查|匹配)|两个条件.*(查|找)/],['onetomany',/一对多|全部结果|所有结果|列出.*全部/],['sumifs',/多条件.*(求和|合计)|sumifs|两个条件.*(求和|合计)/],['distinctcount',/去重.*(计数|个数|统计)|不重复.*(个数|计数)/],['rank',/排名|名次|rank/],['age',/年龄|工龄|周岁|入职.*年/],['weekday',/星期|周几|礼拜|曜日/],['reverse',/反向查找|从右往左|左边.*列.*查/],['xlookup',/xlookup/],['lookup',/vlookup|unlook|查找|查出|匹配|找工资|找价格|找出|带出|关联/],['averageif',/条件.*平均|按.*平均/],['sumif',/条件.*求和|按.*(合计|求和|金额)|sumif/],['countif',/countif|条件.*计数|统计.*(人数|次数|出现)|出现.*次/],['if',/判断|及格|合格|如果|是否/],['trim',/空格|trim/],['round',/四舍五入|小数|round/],['left',/左边|前\s*\d+\s*(个|位)|left/],['right',/右边|后\s*\d+\s*(个|位)|right/],['datediff',/相差.*天|间隔.*天|天数/],['year',/年份|提取年/],['month',/月份|提取月/],['day',/提取日/],['concat',/合并.*(列|文字)|拼接|连接.*列/],['average',/平均/],['sum',/求和|合计|sum/],['count',/非空|有内容.*数量/]];
    rules.forEach(function(rule){if(rule[1].test(text))kinds.push(rule[0]);});
    if(!kinds.length&&/(?:按.*查|从.*表.*查|查.*(?:工资|单价|价格|姓名|部门))/.test(text))kinds.push('lookup');
    if(kinds.indexOf('xlookup')>=0||kinds.indexOf('multilookup')>=0||kinds.indexOf('onetomany')>=0||kinds.indexOf('reverse')>=0)kinds=kinds.filter(function(k){return k!=='lookup';});if(kinds.indexOf('sumifs')>=0)kinds=kinds.filter(function(k){return k!=='sumif';});if(kinds.indexOf('distinctcount')>=0)kinds=kinds.filter(function(k){return k!=='countif'&&k!=='count';});if(kinds.indexOf('sumif')>=0)kinds=kinds.filter(function(k){return k!=='sum';});if(kinds.indexOf('averageif')>=0)kinds=kinds.filter(function(k){return k!=='average';});
    var options={},m=text.match(/(?:大于等于|不小于|>=)\s*(-?\d+(?:\.\d+)?)/);if(m){options.operator='>=';options.threshold=m[1];}else if((m=text.match(/(?:小于等于|不大于|<=)\s*(-?\d+(?:\.\d+)?)/))){options.operator='<=';options.threshold=m[1];}else if((m=text.match(/(?:大于|超过|>)\s*(-?\d+(?:\.\d+)?)/))){options.operator='>';options.threshold=m[1];}else if((m=text.match(/(?:小于|低于|<)\s*(-?\d+(?:\.\d+)?)/))){options.operator='<';options.threshold=m[1];}
    m=text.match(/(\d+)\s*位小数/);if(m)options.digits=Number(m[1]);m=text.match(/(?:前|后)\s*(\d+)\s*(?:个|位)/);if(m)options.length=Number(m[1]);
    return {kind:kinds[0]||null,candidates:kinds,options:options,message:kinds.length?'识别为“'+TYPES[kinds[0]]+'”。请核对类型并选择区域。':'暂未识别，请从公式类型中选择；不会猜测列名或数据位置。'};
  }
  function plan(input){
    if(!input||!Object.prototype.hasOwnProperty.call(TYPES,input.kind))throw Error('请选择公式类型。');
    var kind=input.kind,source=ref(input.source,input.sheet),target=ref(input.target,input.sheet),deps=[source],table=null,second=null,rows=source.rows,formula=[],warnings=[],lookup=kind==='lookup'||kind==='xlookup'||kind==='reverse',multi=kind==='multilookup'||kind==='sumifs',conditional=['sumif','countif','averageif'].indexOf(kind)>=0,aggregate=['sum','average','count','distinctcount'].indexOf(kind)>=0;
    if(source.rows*source.cols>100000)throw Error('来源最多 10 万个单元格，请缩小区域。');
    if(aggregate)rows=1;else if(source.cols!==1)throw Error('待处理内容请选一列，结果会逐行生成。');
    if(rows>1000)throw Error('一次最多生成 1000 行公式。');
    if(target.cols!==1||(target.rows!==1&&target.rows!==rows))throw Error('结果位置请选择一个起始单元格，或与结果行数一致的一列。');
    target={sheet:target.sheet,c1:target.c1,c2:target.c1,r1:target.r1,r2:target.r1+rows-1,rows:rows,cols:1};if(target.r2>1048576)throw Error('结果超出工作表行数。');
    var keyCol,returnCol,key2Col;
    if(lookup||conditional||multi||kind==='onetomany'){table=ref(input.table,input.sheet);if(table.rows*table.cols>100000)throw Error('对照表最多 10 万个单元格。');deps.push(table);keyCol=integer(input.keyColumn||1,1,table.cols,'匹配列序号');if(kind!=='countif')returnCol=integer(input.returnColumn||2,1,table.cols,'返回/金额列序号');if(multi)key2Col=integer(input.key2Column||2,1,table.cols,'第二个条件列序号');}
    if(kind==='concat'||kind==='datediff'||multi){second=ref(input.second,input.sheet);if(second.cols!==1||second.rows!==source.rows)throw Error(multi?'第二个条件需要选一列，行数与第一个条件相同。':'第二列需要与第一列行数相同。');deps.push(second);}
    deps.forEach(function(r){if(intersects(r,target))throw Error('结果区域与来源重叠，会造成循环引用。请选择空白的其他列。');});
    var fallback=quoted(input.fallback==null?'未找到':input.fallback),operator=input.operator||'>=';if(!/^(?:=|<>|>=|<=|>|<)$/.test(operator))throw Error('比较方式无效。');
    for(var i=0;i<rows;i++){
      var v=cell(source,i),f;
      if(lookup){
        // Escape Excel wildcards for exact identifier matching; keep numeric keys numeric.
        var exact='IF(ISTEXT('+v+'),SUBSTITUTE(SUBSTITUTE(SUBSTITUTE('+v+',"~","~~"),"*","~*"),"?","~?"),'+v+')';
        if(kind==='reverse')f='IFNA(INDEX('+full(slice(table,returnCol))+',MATCH('+exact+','+full(slice(table,keyCol))+',0)),'+fallback+')';
        else if(kind==='xlookup')f='XLOOKUP('+v+','+full(slice(table,keyCol))+','+full(slice(table,returnCol))+','+fallback+',0)';
        else if(returnCol>=keyCol){var area={sheet:table.sheet,r1:table.r1,r2:table.r2,c1:table.c1+keyCol-1,c2:table.c2};f='IFNA(VLOOKUP('+exact+','+full(area)+','+(returnCol-keyCol+1)+',FALSE),'+fallback+')';}
        else f='IFNA(INDEX('+full(slice(table,returnCol))+',MATCH('+exact+','+full(slice(table,keyCol))+',0)),'+fallback+')';
        f='IF('+v+'="","",'+f+')';
      }else if(conditional){var criterion='"="&SUBSTITUTE(SUBSTITUTE(SUBSTITUTE('+v+',"~","~~"),"*","~*"),"?","~?")';f=(kind==='sumif'?'SUMIF':kind==='countif'?'COUNTIF':'AVERAGEIF')+'('+full(slice(table,keyCol))+','+criterion+(kind==='countif'?'':','+full(slice(table,returnCol)))+')';if(kind==='averageif')f='IFERROR('+f+',"")';f='IF('+v+'="","",'+f+')';}
      else if(kind==='multilookup'){var w=cell(second,i);f='IF(OR('+v+'="",'+w+'=""),"",IFERROR(INDEX('+full(slice(table,returnCol))+',MATCH(1,INDEX(('+full(slice(table,keyCol))+'='+v+')*('+full(slice(table,key2Col))+'='+w+'),0),0)),'+fallback+'))';}
      else if(kind==='sumifs'){var w2=cell(second,i);f='IF(OR('+v+'="",'+w2+'=""),"",SUMIFS('+full(slice(table,returnCol))+','+full(slice(table,keyCol))+','+v+','+full(slice(table,key2Col))+','+w2+'))';}
      else if(kind==='onetomany')f='IF('+v+'="","",TEXTJOIN('+quoted(input.separator==null?'、':input.separator)+',TRUE,IF('+full(slice(table,keyCol))+'='+v+','+full(slice(table,returnCol))+',"")))';
      else if(kind==='distinctcount')f='SUMPRODUCT(('+full(source)+'<>"")/COUNTIF('+full(source)+','+full(source)+'&""))';
      else if(kind==='rank')f='IF(ISNUMBER('+v+'),RANK('+v+','+full(source)+',0),"")';
      else if(kind==='age')f='IF('+v+'="","",DATEDIF('+v+',TODAY(),"Y"))';
      else if(kind==='weekday')f='IF('+v+'="","",TEXT('+v+',"aaaa"))';
      else if(aggregate)f=({sum:'SUM',average:'AVERAGE',count:'COUNTA'})[kind]+'('+full(source)+')';
      else if(kind==='if')f='IF('+v+'="","",IF('+v+operator+literal(input.threshold==null?'60':input.threshold)+','+literal(input.yes==null?'是':input.yes)+','+literal(input.no==null?'否':input.no)+'))';
      else if(kind==='trim')f='TRIM('+v+')';
      else if(kind==='left'||kind==='right')f=kind.toUpperCase()+'('+v+','+integer(input.length||4,1,32767,'提取长度')+')';
      else if(kind==='round')f='IF('+v+'="","",ROUND('+v+','+integer(input.digits==null?2:input.digits,-15,15,'小数位数')+'))';
      else if(kind==='concat')f=v+'&'+quoted(input.separator==null?' ':input.separator)+'&'+cell(second,i);
      else if(kind==='datediff')f='IF(OR('+v+'="",'+cell(second,i)+'=""),"",INT('+cell(second,i)+')-INT('+v+'))';
      else f='IF('+v+'="","",'+kind.toUpperCase()+'('+v+'))';
      formula.push('='+f);
    }
    if(kind==='multilookup')warnings.push('两个条件都相同才算找到；有多条时返回第一条。');
    if(kind==='onetomany')warnings.push('需要 TEXTJOIN（WPS、Excel 2019 及以后）。旧版本 Excel 请按 Ctrl+Shift+Enter 确认，或改用筛选。');
    if(kind==='sumifs')warnings.push('两个条件都相同的行才会加进合计；对照表不要包含标题行。');
    if(kind==='distinctcount')warnings.push('空白不计；数字 1 和文字“1”算同一个。区域很大（上万行）时计算会变慢。');
    if(kind==='rank')warnings.push('数值最大的为第 1 名，并列名次相同，原表顺序不变。');
    if(kind==='weekday')warnings.push('结果是“星期一”这样的文字；单元格必须是真正的日期。只要“周一”可以把公式里的 aaaa 改成 aaa。');
    if(kind==='age')warnings.push('按今天计算整年，每天打开会自动更新；单元格必须是真正的日期（不是文字）。');
    if(lookup)warnings.push('精确匹配；重复键返回第一条。编号的数字/文本类型应一致。');
    if(kind==='xlookup')warnings.push('XLOOKUP 需要当前 Excel/WPS 支持；旧版本请选兼容查找。');
    if(conditional)warnings.push('按键值精确匹配；对照表不要包含标题行。');
    warnings.push('来源区域固定；修改区域内的值会重新计算，新增到区域之外的行需重新选择范围。');
    return {kind:kind,title:TYPES[kind],target:target,dependencies:deps,formulas:formula,warnings:warnings,summary:'将在 '+full(target)+' 写入 '+rows+' 行公式。'};
  }
  root.OfficeFormula={types:TYPES,infer:infer,plan:plan,ref:ref,address:address,full:full,col:col,intersects:intersects};
})(typeof window==='undefined'?globalThis:window);

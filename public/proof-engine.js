// 办公工具箱 · 错别字与标点检查（2026-11-24）。只做规则检查，不联网、不用 AI。
// check(text) 返回问题列表：{start,end,kind,text,fix,msg}。fix 为 null 的只提示不自动改。
// kind：typo 常见错别字、repeat 重复字、punct 标点、bracket 括号引号不成对、width 全角半角。
(function(root){'use strict';
// 常见错别字：只收“错的写法基本不会是正确用法”的词，避免误报。
var TYPOS=[['按装','安装'],['既使','即使'],['再接再励','再接再厉'],['一股作气','一鼓作气'],['迫不急待','迫不及待'],['穿流不息','川流不息'],['甘败下风','甘拜下风'],['默守成规','墨守成规'],['谈笑风声','谈笑风生'],['关健','关键'],['布署','部署'],['帐号','账号'],['帐户','账户'],['渡假','度假'],['松驰','松弛'],['宣泄','宣泄'],['九洲','九州'],['辐射','辐射'],['一愁莫展','一筹莫展'],['走头无路','走投无路'],['名列前矛','名列前茅'],['原形毕露','原形毕露'],['出奇不意','出其不意'],['金榜提名','金榜题名'],['滥芋充数','滥竽充数'],['心心相映','心心相印'],['不径而走','不胫而走'],['人情事故','人情世故'],['一如继往','一如既往'],['英雄倍出','英雄辈出'],['入场卷','入场券'],['优惠卷','优惠券'],['代金卷','代金券'],['重新装璜','重新装潢'],['装璜','装潢'],['蛰伏','蛰伏'],['磬竹难书','罄竹难书'],['自抱自弃','自暴自弃'],['精兵减政','精兵简政'],['仗义直言','仗义执言'],['貌和神离','貌合神离'],['迄今为止','迄今为止'],['截止目前','截至目前'],['截止到','截至'],['截止今日','截至今日'],['截止本月','截至本月'],['截止年底','截至年底'],['按排','安排'],['安装费用','安装费用'],['以经','已经'],['己经','已经'],['已后','以后'],['即然','既然'],['副作用','副作用'],['坐阵','坐镇'],['震憾','震撼'],['竟然','竟然'],['沉缅','沉湎'],['气慨','气概'],['烦燥','烦躁'],['急燥','急躁'],['干躁','干燥'],['暴燥','暴躁'],['寒喧','寒暄'],['妨碍','妨碍'],['蜂涌','蜂拥'],['针贬','针砭'],['宁静致远','宁静致远'],['不落巢臼','不落窠臼'],['声名雀起','声名鹊起'],['委屈求全','委曲求全'],['发韧','发轫'],['扶养费','抚养费'],['报筹','报酬'],['筹码','筹码'],['题纲','提纲'],['题名','题名'],['做为','作为'],['反应问题','反映问题'],['反应情况','反映情况'],['反应了','反映了'],['合格证书','合格证书'],['计画','计划'],['凭心而论','平心而论'],['一付','一副'],['挖墙角','挖墙脚'],['明信片','明信片'],['美仑美奂','美轮美奂'],['再所不惜','在所不惜'],['再所难免','在所难免'],['不能自己','不能自已'],['悬梁刺骨','悬梁刺股'],['破斧沉舟','破釜沉舟'],['鬼鬼崇崇','鬼鬼祟祟'],['趋之若骛','趋之若鹜'],['好高鹜远','好高骛远'],['甜言密语','甜言蜜语'],['一诺千斤','一诺千金'],['草管人命','草菅人命'],['哄堂大笑','哄堂大笑'],['莫名其妙','莫名其妙'],['莫明其妙','莫名其妙'],['直接了当','直截了当'],['姿意妄为','恣意妄为'],['黄梁美梦','黄粱美梦'],['怨天忧人','怨天尤人'],['绵里藏针','绵里藏针'],['瑕不掩玉','瑕不掩瑜'],['相形见拙','相形见绌'],['越俎代疱','越俎代庖'],['张惶失措','张皇失措'],['一张一驰','一张一弛'],['源远流常','源远流长'],
  // 日文：常见的误变换
  ['見ずらい','見づらい'],['分かりずらい','分かりづらい'],['使いずらい','使いづらい'],['読みずらい','読みづらい'],['言いずらい','言いづらい'],['気ずく','気づく'],['少しずつ','少しずつ'],['以外と','意外と'],['確立が高い','確率が高い'],['御社宛','御社宛']
].filter(function(p){return p[0]!==p[1];});
// 重复字：这些字连写两次几乎都是多打了一个；“不了了之”“是是非非”等例外。
var REPEAT=/([的了是在和与及或也都就将把被对])\1/g,REPEAT_OK=/不了了之|是是非非|了了分明|小了了|明明|天天|的的确确|在在/;
var JA_REPEAT=/([がをにでへ])\1(?![ぁ-ん])/g;
var CJK='\\u3400-\\u9fff\\uf900-\\ufaff',KANA='\\u3040-\\u30ff';
var HALF_PUNCT=new RegExp('(['+CJK+'])([,;:?!])(?=['+CJK+KANA+'“”‘’「」（(]|\\s*$)','g'),HALF_DOT=new RegExp('(['+CJK+'])\\.(?=\\s*$|['+CJK+'])','g');
var FULL={',':'，',';':'；',':':'：','?':'？','!':'！','.':'。'};
var DOUBLE=/([，。、；：])\1+|，。|。，|、，|，、/g;
var SPACE_BEFORE=/[ 　\t]+(?=[，。、；：！？）》」』】])/g;
var HALF_PAREN=new RegExp('\\(([^()\\n]{0,40}['+CJK+KANA+'][^()\\n]{0,40})\\)','g');
var DOTS=/。{3,}|(?:\.\s?){3,}(?=[㐀-鿿぀-ヿ]|\s*$)/g;
var HALF_KANA=/[ｦ-ﾟ]+/g;
var PAIRS=[['（','）'],['“','”'],['《','》'],['【','】'],['「','」'],['『','』'],['‘','’']];
function hasKana(t){return /[぀-ヿ]/.test(t);}
function hasCjk(t){return /[㐀-鿿]/.test(t);}
function push(out,start,end,kind,text,fix,msg){if(Array.isArray(msg))msg=msg[LANG==='ja'?1:0];out.push({start:start,end:end,kind:kind,text:text,fix:fix,msg:msg});}
var LANG='zh';
// opts.lang='ja' 时提示写成日文（写进批注给日本客户看）。
function check(text,opts){LANG=opts&&opts.lang==='ja'?'ja':'zh';text=String(text||'');var out=[],m,ja=hasKana(text);
  TYPOS.forEach(function(p){var i=0;while((i=text.indexOf(p[0],i))>=0){push(out,i,i+p[0].length,'typo',p[0],p[1],['“'+p[0]+'”应为“'+p[1]+'”','「'+p[0]+'」は「'+p[1]+'」の誤りではありませんか']);i+=p[0].length;}});
  REPEAT.lastIndex=0;while((m=REPEAT.exec(text))){var ctx=text.slice(Math.max(0,m.index-2),m.index+4);if(REPEAT_OK.test(ctx))continue;push(out,m.index,m.index+2,'repeat',m[0],m[1],['“'+m[0]+'”多打了一个字','「'+m[0]+'」：同じ文字が重なっています']);}
  if(ja){JA_REPEAT.lastIndex=0;while((m=JA_REPEAT.exec(text)))push(out,m.index,m.index+2,'repeat',m[0],m[1],['“'+m[0]+'”多打了一个字','「'+m[0]+'」：同じ文字が重なっています']);}
  if(hasCjk(text)){
    HALF_PUNCT.lastIndex=0;while((m=HALF_PUNCT.exec(text))){var at=m.index+m[1].length;if(ja&&(m[2]===','||m[2]==='.'))continue;push(out,at,at+1,'punct',m[2],FULL[m[2]],['中文里用了半角“'+m[2]+'”，应为“'+FULL[m[2]]+'”','半角の「'+m[2]+'」があります。全角の「'+FULL[m[2]]+'」にしてください']);}
    if(!ja){HALF_DOT.lastIndex=0;while((m=HALF_DOT.exec(text))){var at2=m.index+m[1].length;push(out,at2,at2+1,'punct','.','。',['句末用了英文句点，应为“。”','文末が半角ピリオドです。「。」にしてください']);}}
    HALF_PAREN.lastIndex=0;while((m=HALF_PAREN.exec(text)))push(out,m.index,m.index+m[0].length,'punct',m[0],'（'+m[1]+'）',['中文内容用了半角括号，应为全角“（）”','半角かっこです。全角の「（）」にしてください']);
  }
  DOUBLE.lastIndex=0;while((m=DOUBLE.exec(text))){if(/^。+$/.test(m[0])&&m[0].length>=3)continue;push(out,m.index,m.index+m[0].length,'punct',m[0],m[0].charAt(0),['标点重复：“'+m[0]+'”','句読点が重なっています：「'+m[0]+'」']);}
  SPACE_BEFORE.lastIndex=0;while((m=SPACE_BEFORE.exec(text)))push(out,m.index,m.index+m[0].length,'punct',m[0],'',['标点前面多了空格','句読点の前に余分なスペースがあります']);
  DOTS.lastIndex=0;while((m=DOTS.exec(text))){if(!hasCjk(text.slice(Math.max(0,m.index-3),m.index)))continue;push(out,m.index,m.index+m[0].length,'punct',m[0],ja?'…':'……',['省略号应为“'+(ja?'…':'……')+'”','省略記号は「'+(ja?'…':'……')+'」にしてください']);}
  HALF_KANA.lastIndex=0;while((m=HALF_KANA.exec(text))){var full=m[0].normalize('NFKC');if(full!==m[0])push(out,m.index,m.index+m[0].length,'width',m[0],full,['半角片假名“'+m[0]+'”应为全角“'+full+'”','半角カタカナ「'+m[0]+'」は全角「'+full+'」にしてください']);}
  PAIRS.forEach(function(p){var open=0,first=-1,last=-1;for(var i=0;i<text.length;i++){var ch=text.charAt(i);if(ch===p[0]){if(open===0)first=i;open++;}else if(ch===p[1]){if(open===0){push(out,i,i+1,'bracket',ch,null,['“'+ch+'”前面没有对应的“'+p[0]+'”','「'+ch+'」に対応する「'+p[0]+'」がありません']);}else open--;last=i;}}if(open>0)push(out,first,first+1,'bracket',p[0],null,['“'+p[0]+'”后面没有对应的“'+p[1]+'”','「'+p[0]+'」に対応する「'+p[1]+'」がありません']);});
  // 同一位置只保留一条（重叠时保留先出现、范围更大的）。
  out.sort(function(a,b){return a.start-b.start||(b.end-b.start)-(a.end-a.start);});var kept=[],end=-1;out.forEach(function(x){if(x.start>=end||x.kind==='bracket'&&x.end<=x.start+1&&x.start>=end){kept.push(x);end=Math.max(end,x.end);}});return kept;}
// 按问题自动改：只改有 fix 的，从后往前改，返回新文字。
function apply(text,issues){text=String(text||'');issues.filter(function(x){return x.fix!=null;}).sort(function(a,b){return b.start-a.start;}).forEach(function(x){text=text.slice(0,x.start)+x.fix+text.slice(x.end);});return text;}
var KIND_NAMES={typo:'错别字',repeat:'重复字',punct:'标点',bracket:'括号引号不成对',width:'全角半角'},KIND_JA={typo:'誤字',repeat:'文字の重複',punct:'句読点',bracket:'かっこの対応',width:'全角・半角'};
function summary(issues,lang){var c={},J=lang==='ja';issues.forEach(function(x){c[x.kind]=(c[x.kind]||0)+1;});return Object.keys(KIND_NAMES).filter(function(k){return c[k];}).map(function(k){return (J?KIND_JA:KIND_NAMES)[k]+' '+c[k]+(J?' 件':' 处');}).join('、');}
root.OfficeProof={check:check,apply:apply,summary:summary,TYPOS:TYPOS,KIND_NAMES:KIND_NAMES};
})(typeof window!=='undefined'?window:globalThis);

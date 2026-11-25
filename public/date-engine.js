// 日期引擎：识别各种写法的日期、换成统一格式、和暦、中国 / 日本的节假日与工作日。
// 中国：每年的放假和调休由国务院办公厅公布，这里内置 2025–2027 年（来源：新华网刊登的国务院办公厅通知）。
//   没有数据的年份只按周六周日算，并在结果里提示。以后每年 11 月左右公布下一年的安排，补进 CN 表即可。
// 日本：按《国民の祝日に関する法律》计算（含振替休日、国民の休日，以及 2019 年、2020–2021 年的特例），适用 2000–2099 年。
(function(root){'use strict';
var pad=function(n){return (n<10?'0':'')+n;};
function ymd(d){return d.getFullYear()+'-'+pad(d.getMonth()+1)+'-'+pad(d.getDate());}
function day(y,m,d){return new Date(y,m-1,d);}
function valid(y,m,d){var x=day(y,m,d);return x.getFullYear()===y&&x.getMonth()===m-1&&x.getDate()===d?x:null;}

// ---- 和暦 ----
var ERAS=[{name:'令和',short:'R',start:day(2019,5,1)},{name:'平成',short:'H',start:day(1989,1,8)},{name:'昭和',short:'S',start:day(1926,12,25)},{name:'大正',short:'T',start:day(1912,7,30)},{name:'明治',short:'M',start:day(1868,1,25)}];
function eraOf(d){for(var i=0;i<ERAS.length;i++)if(d>=ERAS[i].start)return ERAS[i];return null;}
function warekiYear(d){var e=eraOf(d);if(!e)return null;return {era:e,year:d.getFullYear()-e.start.getFullYear()+1};}
var ERA_BY={};ERAS.forEach(function(e){ERA_BY[e.name]=e;ERA_BY[e.short]=e;ERA_BY[e.short.toLowerCase()]=e;});

// ---- 识别 ----
// 支持：2026/11/21、2026-11-21、2026.11.21、2026年11月21日、20261121、令和8年11月21日、R8.11.21、令和元年、
//       11月21日 / 11/21（没写年份时用 defaultYear）、Excel 日期序号（数字，只有 allowSerial 时）。
function z2h(s){return String(s).replace(/[０-９]/g,function(c){return String.fromCharCode(c.charCodeAt(0)-0xFEE0);}).replace(/[／－．]/g,function(c){return {'／':'/','－':'-','．':'.'}[c];});}
function fromSerial(n){if(!(n>=1&&n<2958466))return null;var d=new Date(Date.UTC(1899,11,30)+Math.floor(n)*86400000);return day(d.getUTCFullYear(),d.getUTCMonth()+1,d.getUTCDate());}
function toSerial(d){return Math.round((Date.UTC(d.getFullYear(),d.getMonth(),d.getDate())-Date.UTC(1899,11,30))/86400000);}
function parse(input,opts){opts=opts||{};
  if(input instanceof Date)return isNaN(input)?null:day(input.getFullYear(),input.getMonth()+1,input.getDate());
  if(typeof input==='number')return opts.allowSerial?fromSerial(input):null;
  var s=z2h(String(input==null?'':input)).trim().replace(/\s+/g,'').replace(/[（(][日月火水木金土一二三四五六天][）)]$/,'').replace(/(星期|周|週)[一二三四五六日天]$/,'');
  if(!s)return null;var m;
  if((m=s.match(/^(令和|平成|昭和|大正|明治|[RHSTMrhstm])\.?(元|\d{1,2})[年.\/-](\d{1,2})[月.\/-](\d{1,2})日?$/))){var e=ERA_BY[m[1]],y=e.start.getFullYear()+(m[2]==='元'?1:Number(m[2]))-1;return valid(y,+m[3],+m[4]);}
  if((m=s.match(/^(\d{4})[年\/.-](\d{1,2})[月\/.-](\d{1,2})日?$/)))return valid(+m[1],+m[2],+m[3]);
  if((m=s.match(/^(\d{4})(\d{2})(\d{2})$/)))return valid(+m[1],+m[2],+m[3]);
  if((m=s.match(/^(\d{1,2})[月\/](\d{1,2})日?$/))&&opts.defaultYear)return valid(opts.defaultYear,+m[1],+m[2]);
  if(opts.allowSerial&&/^\d+(\.\d+)?$/.test(s))return fromSerial(Number(s));
  return null;}

// ---- 格式 ----
var WEEK_JA='日月火水木金土',WEEK_ZH='日一二三四五六';
var FORMATS=[
  {id:'slash',label:'2026/11/21',excel:'yyyy/mm/dd'},
  {id:'iso',label:'2026-11-21',excel:'yyyy-mm-dd'},
  {id:'cn',label:'2026年11月21日',excel:'yyyy"年"m"月"d"日"'},
  {id:'cnWeek',label:'2026年11月21日 星期六'},
  {id:'ja',label:'2026年11月21日（土）'},
  {id:'wareki',label:'令和8年11月21日'},
  {id:'warekiShort',label:'R8.11.21'}
];
function format(d,id){var y=d.getFullYear(),m=d.getMonth()+1,dd=d.getDate(),w;
  switch(id){
    case 'iso':return y+'-'+pad(m)+'-'+pad(dd);
    case 'cn':return y+'年'+m+'月'+dd+'日';
    case 'cnWeek':return y+'年'+m+'月'+dd+'日 星期'+WEEK_ZH[d.getDay()];
    case 'ja':return y+'年'+m+'月'+dd+'日（'+WEEK_JA[d.getDay()]+'）';
    case 'wareki':w=warekiYear(d);return w?w.era.name+(w.year===1?'元':w.year)+'年'+m+'月'+dd+'日':y+'年'+m+'月'+dd+'日';
    case 'warekiShort':w=warekiYear(d);return w?w.era.short+w.year+'.'+m+'.'+dd:y+'.'+m+'.'+dd;
    default:return y+'/'+pad(m)+'/'+pad(dd);
  }}

// ---- 中国：放假与调休 ----
var CN={
  2025:{off:[['2025-01-01','2025-01-01','元旦'],['2025-01-28','2025-02-04','春节'],['2025-04-04','2025-04-06','清明节'],['2025-05-01','2025-05-05','劳动节'],['2025-05-31','2025-06-02','端午节'],['2025-10-01','2025-10-08','国庆节、中秋节']],work:['2025-01-26','2025-02-08','2025-04-27','2025-09-28','2025-10-11']},
  2026:{off:[['2026-01-01','2026-01-03','元旦'],['2026-02-15','2026-02-23','春节'],['2026-04-04','2026-04-06','清明节'],['2026-05-01','2026-05-05','劳动节'],['2026-06-19','2026-06-21','端午节'],['2026-09-25','2026-09-27','中秋节'],['2026-10-01','2026-10-07','国庆节']],work:['2026-01-04','2026-02-14','2026-02-28','2026-05-09','2026-09-20','2026-10-10']},
  2027:{off:[['2027-01-01','2027-01-03','元旦'],['2027-02-05','2027-02-13','春节'],['2027-04-03','2027-04-05','清明节'],['2027-05-01','2027-05-05','劳动节'],['2027-06-09','2027-06-09','端午节'],['2027-09-15','2027-09-15','中秋节'],['2027-10-01','2027-10-07','国庆节']],work:['2027-01-31','2027-02-14','2027-05-08','2027-09-26','2027-10-09']}
};
var cnIndex=null;
function cnTable(){if(cnIndex)return cnIndex;cnIndex={off:{},work:{}};Object.keys(CN).forEach(function(y){CN[y].off.forEach(function(r){var a=parse(r[0]),b=parse(r[1]);for(var d=a;d<=b;d=day(d.getFullYear(),d.getMonth()+1,d.getDate()+1))cnIndex.off[ymd(d)]=r[2];});CN[y].work.forEach(function(w){cnIndex.work[w]=true;});});return cnIndex;}
function cnHas(y){return !!CN[y];}

// ---- 日本：国民の祝日 ----
function nthMonday(y,m,n){var first=day(y,m,1),off=(8-first.getDay())%7;return day(y,m,1+off+(n-1)*7);}
function equinox(y,spring){var base=spring?20.8431:23.2488;return Math.floor(base+0.242194*(y-1980)-Math.floor((y-1980)/4));}
var jpCache={};
function jpHolidays(y){if(jpCache[y])return jpCache[y];var h={},add=function(d,name){h[ymd(d)]=name;};
  add(day(y,1,1),'元日');add(nthMonday(y,1,2),'成人の日');add(day(y,2,11),'建国記念の日');
  if(y>=2020)add(day(y,2,23),'天皇誕生日');else if(y<=2018)add(day(y,12,23),'天皇誕生日');
  add(day(y,3,equinox(y,true)),'春分の日');add(day(y,4,29),y>=2007?'昭和の日':'みどりの日');add(day(y,5,3),'憲法記念日');
  if(y>=2007)add(day(y,5,4),'みどりの日');add(day(y,5,5),'こどもの日');
  if(y===2020)add(day(2020,7,23),'海の日');else if(y===2021)add(day(2021,7,22),'海の日');else add(nthMonday(y,7,3),'海の日');
  if(y===2020)add(day(2020,8,10),'山の日');else if(y===2021)add(day(2021,8,8),'山の日');else if(y>=2016)add(day(y,8,11),'山の日');
  add(nthMonday(y,9,3),'敬老の日');add(day(y,9,equinox(y,false)),'秋分の日');
  if(y===2020)add(day(2020,7,24),'スポーツの日');else if(y===2021)add(day(2021,7,23),'スポーツの日');else add(nthMonday(y,10,2),y>=2020?'スポーツの日':'体育の日');
  add(day(y,11,3),'文化の日');add(day(y,11,23),'勤労感謝の日');
  if(y===2019){add(day(2019,5,1),'天皇の即位の日');add(day(2019,10,22),'即位礼正殿の儀の行われる日');}
  // 国民の休日：前後都是祝日的平日
  Object.keys(h).forEach(function(k){var d=parse(k),next=day(d.getFullYear(),d.getMonth()+1,d.getDate()+2),mid=day(d.getFullYear(),d.getMonth()+1,d.getDate()+1);if(h[ymd(next)]&&!h[ymd(mid)]&&mid.getDay()!==0)h[ymd(mid)]='国民の休日';});
  // 振替休日：祝日が日曜なら、その後の最初の祝日でない日
  Object.keys(h).sort().forEach(function(k){var d=parse(k);if(d.getDay()!==0)return;var n=day(d.getFullYear(),d.getMonth()+1,d.getDate()+1);while(h[ymd(n)])n=day(n.getFullYear(),n.getMonth()+1,n.getDate()+1);h[ymd(n)]='振替休日';});
  jpCache[y]=h;return h;}

// region：cn | jp；opts.yearEnd：日本的年末年始（12/29–1/3）也算休息。
function dayInfo(d,region,opts){opts=opts||{};var key=ymd(d),w=d.getDay();
  if(region==='cn'){var t=cnTable();if(t.off[key])return {work:false,name:t.off[key]};if(t.work[key])return {work:true,name:'调休上班'};return {work:w!==0&&w!==6,name:w===0||w===6?'周末':'',known:cnHas(d.getFullYear())};}
  var h=jpHolidays(d.getFullYear());if(h[key])return {work:false,name:h[key]};
  if(opts.yearEnd){var m=d.getMonth()+1,dd=d.getDate();if(m===12&&dd>=29||m===1&&dd<=3)return {work:false,name:'年末年始'};}
  return {work:w!==0&&w!==6,name:w===0||w===6?'週末':''};}
function isWorkday(d,region,opts){return dayInfo(d,region,opts).work;}
// 包含起止两天。
function countWorkdays(a,b,region,opts){if(a>b){var t=a;a=b;b=t;}var n=0,holidays=0,makeup=0,weekend=0,missing={};
  for(var d=a;d<=b;d=day(d.getFullYear(),d.getMonth()+1,d.getDate()+1)){var i=dayInfo(d,region,opts),w=d.getDay();if(region==='cn'&&!cnHas(d.getFullYear()))missing[d.getFullYear()]=1;
    if(i.work){n++;if(w===0||w===6)makeup++;}else if(w===0||w===6)weekend++;else holidays++;}
  return {workdays:n,weekend:weekend,holidays:holidays,makeup:makeup,missingYears:Object.keys(missing).map(Number)};}
// 从 d 往后数 n 个工作日（不含 d 当天）；n 为负数时往前数。
function addWorkdays(d,n,region,opts){var step=n<0?-1:1,left=Math.abs(n),x=d;while(left>0){x=day(x.getFullYear(),x.getMonth()+1,x.getDate()+step);if(isWorkday(x,region,opts))left--;}return x;}
function period(d,kind,lang){var y=d.getFullYear(),m=d.getMonth()+1;if(kind==='year')return {key:String(y),label:y+'年'};if(kind==='quarter'){var q=Math.ceil(m/3);return {key:y+'-Q'+q,label:lang==='ja'?y+'年 第'+q+'四半期':y+'年第'+q+'季度'};}return {key:y+'-'+pad(m),label:y+'年'+m+'月'};}

var api={parse:parse,format:format,FORMATS:FORMATS,toSerial:toSerial,fromSerial:fromSerial,warekiYear:warekiYear,jpHolidays:jpHolidays,dayInfo:dayInfo,isWorkday:isWorkday,countWorkdays:countWorkdays,addWorkdays:addWorkdays,cnYears:function(){return Object.keys(CN).map(Number);},period:period,ymd:ymd};
root.OfficeDates=api;if(typeof module!=='undefined'&&module.exports)module.exports=api;
})(typeof window!=='undefined'?window:globalThis);

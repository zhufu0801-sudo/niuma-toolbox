// 办公工具箱 · 号码校验（2026-11-24）：统一社会信用代码、身份证号、发票号码、日本法人番号 / インボイス登録番号、银行卡号。
// 只按公开的编码规则算校验位，不联网查询，不能证明号码真实存在。
(function(root){'use strict';
var USCC_CHARS='0123456789ABCDEFGHJKLMNPQRTUWXY',USCC_W=[1,3,9,27,19,26,16,17,20,29,25,13,8,24,10,30,28];
var ID_W=[7,9,10,5,8,4,2,1,6,3,7,9,10,5,8,4,2],ID_CHECK='10X98765432';
var NAMES={uscc:'统一社会信用代码',id:'身份证号',invoice:'发票号码',invoiceCode:'发票代码',jpcorp:'法人番号',jpinvoice:'インボイス登録番号',bank:'银行卡号'},NAMES_JA={uscc:'統一社会信用コード（中国）',id:'身分証番号（中国）',invoice:'発票番号（中国）',invoiceCode:'発票コード（中国）',jpcorp:'法人番号',jpinvoice:'インボイス登録番号',bank:'銀行カード番号'},LANG='zh';
// 消息写进表格给客户看：opts.lang='ja' 时用日文。
function M(zh,ja){return LANG==='ja'?ja:zh;}
function clean(v){return String(v==null?'':v).normalize('NFKC').replace(/[\s\-‐－ー—]/g,'').toUpperCase();}
function usccCheck(s){var sum=0;for(var i=0;i<17;i++){var k=USCC_CHARS.indexOf(s.charAt(i));if(k<0)return null;sum+=k*USCC_W[i];}var c=31-sum%31;return USCC_CHARS.charAt(c===31?0:c);}
function idCheck(s){var sum=0;for(var i=0;i<17;i++)sum+=Number(s.charAt(i))*ID_W[i];return ID_CHECK.charAt(sum%11);}
function idBirth(s){var y=+s.slice(6,10),m=+s.slice(10,12),d=+s.slice(12,14),dt=new Date(y,m-1,d);return y>=1900&&dt.getFullYear()===y&&dt.getMonth()===m-1&&dt.getDate()===d&&dt<=new Date();}
function jpCheck(base12){var sum=0;for(var n=1;n<=12;n++){var p=Number(base12.charAt(12-n));sum+=p*(n%2?1:2);}return String(9-sum%9);}
function luhn(s){var sum=0,alt=false;for(var i=s.length-1;i>=0;i--){var d=Number(s.charAt(i));if(alt){d*=2;if(d>9)d-=9;}sum+=d;alt=!alt;}return sum%10===0;}
function nm(type){return (LANG==='ja'?NAMES_JA:NAMES)[type];}
function ok(type,extra){return {ok:true,type:type,name:nm(type),msg:'✓ '+nm(type)+(extra?'（'+extra+'）':'')};}
function bad(type,why){return {ok:false,type:type,name:nm(type),msg:'✗ '+(nm(type)?nm(type)+'：':'')+why};}
var CHECKS={
  uscc:function(s){if(s.length!==18)return bad('uscc',M('应为 18 位，现在 '+s.length+' 位','18 桁のはずですが '+s.length+' 桁です'));if(/[IOZSV]/.test(s))return bad('uscc',M('不能含有字母 I、O、Z、S、V','I・O・Z・S・V は使えません'));var c=usccCheck(s);if(c==null)return bad('uscc',M('含有不允许的字符','使えない文字が含まれています'));return c===s.charAt(17)?ok('uscc'):bad('uscc',M('最后一位校验码应为 '+c,'最後のチェック文字は '+c+' のはずです'));},
  id:function(s){if(!/^\d{17}[\dX]$/.test(s))return bad('id',s.length!==18?M('应为 18 位，现在 '+s.length+' 位','18 桁のはずですが '+s.length+' 桁です'):M('前 17 位必须是数字','先頭 17 桁は数字のはずです'));if(!idBirth(s))return bad('id',M('出生日期 '+s.slice(6,14)+' 不对','生年月日 '+s.slice(6,14)+' が正しくありません'));var c=idCheck(s);return c===s.charAt(17)?ok('id'):bad('id',M('最后一位校验码应为 '+c,'最後のチェック文字は '+c+' のはずです'));},
  invoice:function(s){if(/^\d{20}$/.test(s))return ok('invoice',M('20 位数电发票号码格式正确；真伪请到税务局查验平台查','20 桁の電子発票番号の形式です。真偽は税務局の照会サイトで確認してください'));if(/^\d{8}$/.test(s))return ok('invoice',M('8 位发票号码格式正确；真伪请到税务局查验平台查','8 桁の発票番号の形式です。真偽は税務局の照会サイトで確認してください'));if(/^\d{10}$|^\d{12}$/.test(s))return {ok:true,type:'invoiceCode',name:nm('invoiceCode'),msg:'✓ '+nm('invoiceCode')+M('（'+s.length+' 位，格式正确）','（'+s.length+' 桁、形式は正しい）')};return bad('invoice',M('应为 8 位（纸质发票）或 20 位（数电发票）数字，现在 '+s.length+' 位','8 桁（紙の発票）か 20 桁（電子発票）の数字のはずですが '+s.length+' 桁です'));},
  jpcorp:function(s){var t=/^T/.test(s),b=t?s.slice(1):s;if(!/^\d{13}$/.test(b))return bad(t?'jpinvoice':'jpcorp',M((t?'T 后面':'')+'应为 13 位数字，现在 '+b.length+' 位',(t?'T の後ろは':'')+'13 桁の数字のはずですが '+b.length+' 桁です'));var c=jpCheck(b.slice(1));return c===b.charAt(0)?ok(t?'jpinvoice':'jpcorp'):bad(t?'jpinvoice':'jpcorp',M('校验位（第 1 位）应为 '+c,'チェックデジット（1 桁目）は '+c+' のはずです'));},
  bank:function(s){if(!/^\d{12,19}$/.test(s))return bad('bank',M('应为 12～19 位数字','12～19 桁の数字のはずです'));return luhn(s)?ok('bank',M('校验位正确','チェックデジット OK')):bad('bank',M('校验位不对，可能有一位写错','チェックデジットが合いません。どこか 1 桁違う可能性があります'));}
};
// 自动判断号码种类：先看形状，再看哪种规则说得通。
function guess(s){if(/^T\d{13}$/.test(s))return 'jpcorp';if(/^\d{17}[\dX]$/.test(s)){if(idBirth(s)||/^[1-6]\d{5}(?:19|20)\d{9}[\dX]$/.test(s))return 'id';return 'uscc';}if(s.length===18&&/^[0-9A-HJ-NP-RT-UW-Y]+$/.test(s))return 'uscc';if(/^\d{13}$/.test(s))return 'jpcorp';if(/^\d{20}$|^\d{8}$|^\d{10}$|^\d{12}$/.test(s))return 'invoice';if(/^\d{16,19}$/.test(s))return 'bank';return '';}
function check(value,type,opts){LANG=opts&&opts.lang==='ja'?'ja':'zh';var s=clean(value);if(!s)return null;var t=type&&CHECKS[type]?type:guess(s);if(!t)return {ok:false,type:'',name:'',msg:M('? 认不出是哪种号码（'+s.length+' 位）','? 何の番号か判別できません（'+s.length+' 桁）')};return CHECKS[t](s);}
root.OfficeIdCheck={check:check,guess:guess,clean:clean,NAMES:NAMES,usccCheck:usccCheck,idCheck:idCheck,jpCheck:jpCheck,luhn:luhn};
})(typeof window!=='undefined'?window:globalThis);

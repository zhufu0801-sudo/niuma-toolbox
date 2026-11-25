(function(root){
 'use strict';
 function clean(s){return String(s).replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f\u007f\u200b\u200c\u200d\u2060\ufeff]/g,'').replace(/\u00a0/g,' ');}
 function unique(values){var seen=new Set();return values.filter(function(v){if(v===null||v===undefined||v==='')return false;var key=typeof v+':'+String(v);if(seen.has(key))return false;seen.add(key);return true;});}
 function choices(text){var out=unique(String(text||'').split(/\r?\n/).map(function(s){return s.trim();}));if(!out.length)throw Error('请填写下拉选项，每行一个。');if(out.some(function(s){return /[,;\r\n\x00-\x1f]/.test(s)||s[0]==='=';}))throw Error('选项不能以等号开头，也不能包含逗号、分号或控制字符。');if(out.join(',').length>255)throw Error('下拉选项合计最多 255 个字符，请减少选项。');return out;}
 function stats(values){var n=0,sum=0,filled=0,min=null,max=null;values.forEach(function(row){row.forEach(function(v){if(v!==null&&v!==undefined&&v!=='')filled++;if(typeof v==='number'&&isFinite(v)){n++;sum+=v;min=min===null?v:Math.min(min,v);max=max===null?v:Math.max(max,v);}});});return {sum:sum,count:n,filled:filled,average:n?sum/n:null,min:min,max:max};}
 function summary(rows,keyCol,valueCol){var map=new Map(),result=[];rows.forEach(function(row){var v=row[keyCol],key=typeof v+':'+String(v==null?'':v),r=map.get(key);if(!r){r=[v==null?'':v,0,0];map.set(key,r);result.push(r);}r[1]++;var amount=row[valueCol];if(typeof amount==='number'&&isFinite(amount))r[2]+=amount;});return result;}
 function chord(raw){var parts=String(raw||'').replace(/\s/g,'').split('+'),key=(parts.pop()||'').toUpperCase(),set=new Set(parts.map(function(x){return x.toLowerCase();}));if(!key)return '';if(!/^[A-Z0-9]$/.test(key)||parts.length!==set.size||[...set].some(function(x){return ['ctrl','alt','shift'].indexOf(x)<0;})||!set.has('ctrl')||!set.has('alt'))throw Error('使用 Ctrl+Alt+字母/数字，可再加 Shift，例如 Ctrl+Alt+Q。');return 'Ctrl+Alt+'+(set.has('shift')?'Shift+':'')+key;}
 function fromEvent(e){if(e.isComposing||e.repeat||e.metaKey||!e.ctrlKey||!e.altKey||!e.key||!/^[a-z0-9]$/i.test(e.key))return '';return 'Ctrl+Alt+'+(e.shiftKey?'Shift+':'')+e.key.toUpperCase();}
 function literal(v){return typeof v==='string'?"'"+v:(v==null?'':v);}
 root.QuickEngine={clean:clean,unique:unique,choices:choices,stats:stats,summary:summary,chord:chord,fromEvent:fromEvent,literal:literal};
})(window);

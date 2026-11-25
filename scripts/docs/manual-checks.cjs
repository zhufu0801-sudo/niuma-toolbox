const pairs={'“':'”','「':'」','（':'）'},closing=new Set(Object.values(pairs));
function checkLine(line){const stack=[];for(const c of line){if(pairs[c])stack.push(pairs[c]);else if(closing.has(c)&&stack.pop()!==c)return false;}return stack.length===0;}
function validateManual(text,label='manual'){
 const errors=[];text.split(/\r?\n/).forEach((line,i)=>{if(!line.trimStart().startsWith('|')&&!checkLine(line))errors.push(`${label}:${i+1}: unbalanced quotes or parentheses`);});
 if(errors.length)throw Error(errors.join('\n'));
}
// Only separators outside quoted UI paths start a new instruction.
function splitSteps(text){const out=[],stack=[];let start=0;for(let i=0;i<text.length;i++){const c=text[i];if(pairs[c])stack.push(pairs[c]);else if(closing.has(c)&&stack.at(-1)===c)stack.pop();if(!stack.length&&text.startsWith(' → ',i)){out.push(text.slice(start,i));start=i+3;i+=2;}}out.push(text.slice(start));return out;}
module.exports={validateManual,splitSteps};

// 读 PDF 文字用的 pdf.js 在加载时就要 DOMMatrix。Windows 安装包带了 @napi-rs/canvas 会自动补上；
// 其他系统（Mac、Linux）没有时，补一个只做二维仿射变换的简化版，够提取文字用（不画图）。
export async function ensurePdfEnv(){
  if(typeof globalThis.DOMMatrix==='function')return 'native';
  try{const c=await import('@napi-rs/canvas');const m=c.DOMMatrix||(c.default&&c.default.DOMMatrix);if(m){globalThis.DOMMatrix=m;return 'canvas';}}catch{}
  class Matrix2D{
    constructor(init){const v=Array.isArray(init)&&init.length>=6?init:[1,0,0,1,0,0];[this.a,this.b,this.c,this.d,this.e,this.f]=v.slice(0,6).map(Number);}
    get m11(){return this.a;}get m12(){return this.b;}get m21(){return this.c;}get m22(){return this.d;}get m41(){return this.e;}get m42(){return this.f;}
    get is2D(){return true;}get isIdentity(){return this.a===1&&this.b===0&&this.c===0&&this.d===1&&this.e===0&&this.f===0;}
    multiplySelf(o){const m=o instanceof Matrix2D?o:new Matrix2D(o&&[o.a,o.b,o.c,o.d,o.e,o.f]);const a=this.a*m.a+this.c*m.b,b=this.b*m.a+this.d*m.b,c=this.a*m.c+this.c*m.d,d=this.b*m.c+this.d*m.d,e=this.a*m.e+this.c*m.f+this.e,f=this.b*m.e+this.d*m.f+this.f;Object.assign(this,{a,b,c,d,e,f});return this;}
    preMultiplySelf(o){const m=new Matrix2D([o.a,o.b,o.c,o.d,o.e,o.f]);m.multiplySelf(this);Object.assign(this,{a:m.a,b:m.b,c:m.c,d:m.d,e:m.e,f:m.f});return this;}
    multiply(o){return new Matrix2D([this.a,this.b,this.c,this.d,this.e,this.f]).multiplySelf(o);}
    translateSelf(x=0,y=0){return this.multiplySelf(new Matrix2D([1,0,0,1,x,y]));}translate(x,y){return this.multiply(new Matrix2D([1,0,0,1,x||0,y||0]));}
    scaleSelf(x=1,y=x){return this.multiplySelf(new Matrix2D([x,0,0,y,0,0]));}scale(x=1,y=x){return this.multiply(new Matrix2D([x,0,0,y,0,0]));}
    invertSelf(){const det=this.a*this.d-this.b*this.c;if(!det){Object.assign(this,{a:NaN,b:NaN,c:NaN,d:NaN,e:NaN,f:NaN});return this;}const {a,b,c,d,e,f}=this;Object.assign(this,{a:d/det,b:-b/det,c:-c/det,d:a/det,e:(c*f-d*e)/det,f:(b*e-a*f)/det});return this;}
    inverse(){return new Matrix2D([this.a,this.b,this.c,this.d,this.e,this.f]).invertSelf();}
    transformPoint(p={}){const x=p.x||0,y=p.y||0;return {x:this.a*x+this.c*y+this.e,y:this.b*x+this.d*y+this.f,z:0,w:1};}
  }
  globalThis.DOMMatrix=Matrix2D;return 'basic';
}

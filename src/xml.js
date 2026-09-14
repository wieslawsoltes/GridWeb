export const xmlEscape=value=>String(value??'').replace(/[\x00-\x08\x0b\x0c\x0e-\x1f]/g,'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&apos;'}[c]));
export function xmlDecode(value){
 return value.replace(/&(?:#x[\da-fA-F]+|#\d+|[A-Za-z]+);|&/g,token=>{
  const map={'&amp;':'&','&lt;':'<','&gt;':'>','&quot;':'"','&apos;':"'"};if(Object.hasOwn(map,token))return map[token];
  if(!/^&#(?:x[\da-fA-F]+|\d+);$/.test(token))throw new Error('Unknown or malformed XML entity');
  const n=token[2]==='x'?parseInt(token.slice(3,-1),16):parseInt(token.slice(2,-1),10);
  if(!Number.isFinite(n)||n>0x10ffff||(n>=0xd800&&n<=0xdfff)||n===0xfffe||n===0xffff||(n<32&&![9,10,13].includes(n)))throw new Error('Invalid XML character');
  return String.fromCodePoint(n);
 });
}
/** Bounded, non-validating XML parser; no DTD, external entity resolution or executable content. */
export function parseXML(text){
 if(typeof text!=='string'||text.length>16*1024*1024)throw new RangeError('XML part too large');
 if(/<!DOCTYPE|<!ENTITY/i.test(text))throw new Error('DTD/entity declarations are not accepted');
 if(/[\x00-\x08\x0b\x0c\x0e-\x1f]/.test(text))throw new Error('Invalid XML character');
 const root={name:'#document',attrs:Object.create(null),children:[],content:[],text:''},stack=[root];let nodes=0,offset=0;
 const tokenizer=/<!--[\s\S]*?-->|<\?[\s\S]*?\?>|<!\[CDATA\[[\s\S]*?\]\]>|<(?:[^"'<>]|"[^"]*"|'[^']*')+>|[^<]+/gy;
 const addText=value=>{stack.at(-1).text+=value;stack.at(-1).content.push(value);};
 while(offset<text.length){
  tokenizer.lastIndex=offset;const match=tokenizer.exec(text);if(!match)throw new Error('Malformed XML token');const token=match[0];offset=tokenizer.lastIndex;
  if(token.startsWith('<?'))continue;
  if(token.startsWith('<!--')){if(token.slice(4,-3).includes('--'))throw new Error('Malformed XML comment');continue;}
  if(token.startsWith('<![CDATA[')){if(stack.length===1)throw new Error('CDATA outside document element');addText(token.slice(9,-3));continue;}
  if(token.startsWith('</')){const name=token.slice(2,-1).trim();if(stack.length<=1||stack.at(-1).fullName!==name)throw new Error('Mismatched XML tags');stack.pop();}
  else if(token.startsWith('<')){
   const head=/^<([A-Za-z_][\w:.-]*)/.exec(token);if(!head)throw new Error('Invalid XML tag');const attrs=Object.create(null),end=token.endsWith('/>')?token.length-2:token.length-1;let pos=head[0].length;
   while(pos<end){const white=/^\s+/.exec(token.slice(pos));if(!white)throw new Error('XML attributes need whitespace');pos+=white[0].length;if(pos===end)break;
    const attr=/^([A-Za-z_][\w:.-]*)\s*=\s*(?:"([^"<]*)"|'([^'<]*)')/.exec(token.slice(pos,end));if(!attr)throw new Error('Malformed XML attribute');if(Object.hasOwn(attrs,attr[1]))throw new Error('Duplicate XML attribute');attrs[attr[1]]=xmlDecode(attr[2]??attr[3]);pos+=attr[0].length;
   }
   const node={fullName:head[1],name:head[1].split(':').at(-1),attrs,children:[],content:[],text:''};stack.at(-1).children.push(node);stack.at(-1).content.push(node);
   if(++nodes>1500000)throw new RangeError('XML node limit');if(!token.endsWith('/>'))stack.push(node);if(stack.length>128)throw new RangeError('XML depth limit');
  }else{if(token.includes(']]>'))throw new Error('CDATA closing delimiter in text');addText(xmlDecode(token));}
 }
 if(stack.length!==1||root.children.length!==1||root.text.trim())throw new Error('Incomplete XML document');return root.children[0];
}
export const children=(node,name)=>node?.children.filter(n=>n.name===name)??[];
export const child=(node,name)=>node?.children.find(n=>n.name===name);
export function descendants(node,name){const result=[];const walk=n=>{if(n.name===name)result.push(n);for(const c of n.children)walk(c);};if(node)walk(node);return result;}
export const allText=node=>node?(node.content??[node.text,...node.children]).map(item=>typeof item==='string'?item:allText(item)).join(''):'';

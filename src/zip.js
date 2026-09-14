const enc=new TextEncoder(),dec=new TextDecoder('utf-8',{fatal:true});
const crcTable=Uint32Array.from({length:256},(_,n)=>{for(let k=0;k<8;k++)n=n&1?0xedb88320^(n>>>1):n>>>1;return n>>>0;});
export function crc32(bytes){let crc=0xffffffff;for(const b of bytes)crc=crcTable[(crc^b)&255]^(crc>>>8);return(crc^0xffffffff)>>>0;}
const validPath=name=>name.length>0&&name.length<512&&!name.startsWith('/')&&!name.includes('\\')&&!name.includes('\0')&&!name.split('/').includes('..');
export function writeZip(entries){
  const parts=[],central=[];let offset=0,count=0,total=0;
  for(const[name,value]of Object.entries(entries)){
    if(!validPath(name))throw new TypeError('Unsafe ZIP path');const filename=enc.encode(name),data=typeof value==='string'?enc.encode(value):new Uint8Array(value);if((total+=data.length)>64*1024*1024)throw new RangeError('ZIP output limit');const crc=crc32(data);
    const local=new Uint8Array(30+filename.length),l=new DataView(local.buffer);l.setUint32(0,0x04034b50,true);l.setUint16(4,20,true);l.setUint16(6,0x800,true);l.setUint16(12,33,true);l.setUint32(14,crc,true);l.setUint32(18,data.length,true);l.setUint32(22,data.length,true);l.setUint16(26,filename.length,true);local.set(filename,30);parts.push(local,data);
    const header=new Uint8Array(46+filename.length),c=new DataView(header.buffer);c.setUint32(0,0x02014b50,true);c.setUint16(4,20,true);c.setUint16(6,20,true);c.setUint16(8,0x800,true);c.setUint16(14,33,true);c.setUint32(16,crc,true);c.setUint32(20,data.length,true);c.setUint32(24,data.length,true);c.setUint16(28,filename.length,true);c.setUint32(42,offset,true);header.set(filename,46);central.push(header);offset+=local.length+data.length;count++;
  }
  if(count>3000)throw new RangeError('ZIP entry limit');const size=central.reduce((n,a)=>n+a.length,0),end=new Uint8Array(22),e=new DataView(end.buffer);e.setUint32(0,0x06054b50,true);e.setUint16(8,count,true);e.setUint16(10,count,true);e.setUint32(12,size,true);e.setUint32(16,offset,true);
  const output=new Uint8Array(offset+size+22);let p=0;for(const a of [...parts,...central,end]){output.set(a,p);p+=a.length;}return output;
}
export async function readZip(input,{maxBytes=64*1024*1024,maxFileBytes=16*1024*1024,maxEntries=3000}={}){
  const bytes=input instanceof Uint8Array?input:new Uint8Array(input);if(bytes.length>32*1024*1024||bytes.length<22)throw new RangeError('ZIP input must be 22 bytes–32 MiB');const v=new DataView(bytes.buffer,bytes.byteOffset,bytes.byteLength);let end=-1;
  for(let i=bytes.length-22;i>=Math.max(0,bytes.length-65557);i--)if(v.getUint32(i,true)===0x06054b50&&i+22+v.getUint16(i+20,true)===bytes.length){end=i;break;}
  if(end<0)throw new Error('Not a supported ZIP archive');if(v.getUint16(end+4,true)||v.getUint16(end+6,true))throw new Error('Multi-disk ZIP is unsupported');const count=v.getUint16(end+10,true),start=v.getUint32(end+16,true),size=v.getUint32(end+12,true);if(count>maxEntries||start+size>end)throw new RangeError('Invalid ZIP directory');
  const entries=new Map();let p=start,total=0;
  for(let i=0;i<count;i++){
    if(p+46>start+size||v.getUint32(p,true)!==0x02014b50)throw new Error('Invalid ZIP entry');const flags=v.getUint16(p+8,true),method=v.getUint16(p+10,true),crc=v.getUint32(p+16,true),compressed=v.getUint32(p+20,true),uncompressed=v.getUint32(p+24,true),nl=v.getUint16(p+28,true),el=v.getUint16(p+30,true),cl=v.getUint16(p+32,true),local=v.getUint32(p+42,true);
    if(flags&1)throw new Error('Encrypted workbooks are not supported');if(![0,8].includes(method))throw new Error('Unsupported ZIP compression');if(uncompressed>maxFileBytes||(total+=uncompressed)>maxBytes)throw new RangeError('Uncompressed ZIP size limit');if(p+46+nl+el+cl>start+size)throw new Error('Invalid ZIP name');const name=dec.decode(bytes.subarray(p+46,p+46+nl));if(!validPath(name)||entries.has(name))throw new Error('Unsafe or duplicate ZIP path');
    if(local+30>start||v.getUint16(local+6,true)!==flags||v.getUint32(local,true)!==0x04034b50||v.getUint16(local+8,true)!==method)throw new Error('Invalid local ZIP header');const lnl=v.getUint16(local+26,true),lel=v.getUint16(local+28,true),dataStart=local+30+lnl+lel;if(dataStart+compressed>start||dec.decode(bytes.subarray(local+30,local+30+lnl))!==name)throw new Error('Invalid ZIP entry bounds');
    let data=bytes.subarray(dataStart,dataStart+compressed);
    if(method===8){
      if(typeof process!=='undefined'&&process.versions?.node){const{inflateRawSync}=await import('node:zlib');data=new Uint8Array(inflateRawSync(data,{maxOutputLength:Math.min(maxFileBytes,uncompressed+1)}));}
      else {const stream=new Blob([data]).stream().pipeThrough(new DecompressionStream('deflate-raw')),reader=stream.getReader(),chunks=[];let length=0;try{while(true){const r=await reader.read();if(r.done)break;length+=r.value.length;if(length>uncompressed||length>maxFileBytes)throw new RangeError('Inflated ZIP size limit');chunks.push(r.value);}}catch(e){await reader.cancel().catch(()=>{});throw e;}data=new Uint8Array(length);let at=0;for(const c of chunks){data.set(c,at);at+=c.length;}}
    }
    if(data.length!==uncompressed||crc32(data)!==crc)throw new Error('ZIP checksum or size mismatch');entries.set(name,data);p+=46+nl+el+cl;
  }
  if(p!==start+size)throw new Error('Unexpected ZIP directory data');return entries;
}

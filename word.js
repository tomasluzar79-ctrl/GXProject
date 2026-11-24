"use strict";
// ZIP with stored entries. Word packages do not require compression.
function packageZip(files){
 const enc=new TextEncoder(),parts=[],directory=[];let offset=0;
 function header(size,signature){const bytes=new Uint8Array(size);const d=new DataView(bytes.buffer);d.setUint32(0,signature,true);return {bytes,u16:(i,n)=>d.setUint16(i,n,true),u32:(i,n)=>d.setUint32(i,n,true)};}
 const crcTable=Uint32Array.from({length:256},(_,n)=>{for(let j=0;j<8;j++)n=(n>>>1)^((n&1)?0xedb88320:0);return n;});
 for(const [path,raw] of files){const name=enc.encode(path),data=typeof raw==="string"?enc.encode(raw):raw;
  let crc=0xffffffff;for(const b of data)crc=(crc>>>8)^crcTable[(crc^b)&255];crc=(crc^0xffffffff)>>>0;
  const h=header(30,0x04034b50);h.u16(4,20);h.u16(12,33);h.u32(14,crc);h.u32(18,data.length);h.u32(22,data.length);h.u16(26,name.length);parts.push(h.bytes,name,data);
  const c=header(46,0x02014b50);c.u16(4,20);c.u16(6,20);c.u16(14,33);c.u32(16,crc);c.u32(20,data.length);c.u32(24,data.length);c.u16(28,name.length);c.u32(42,offset);directory.push(c.bytes,name);offset+=30+name.length+data.length;
 }
 const end=header(22,0x06054b50);end.u16(8,files.length);end.u16(10,files.length);end.u32(12,directory.reduce((n,x)=>n+x.length,0));end.u32(16,offset);
 return new Blob([...parts,...directory,end.bytes],{type:"application/vnd.openxmlformats-officedocument.wordprocessingml.document"});
}
async function wordFile(pages,number){
 const ns="http://schemas.openxmlformats.org",office=ns+"/officeDocument/2006/relationships",draw=ns+"/drawingml/2006/",files=[],rels=[];
 const xml=s=>esc(s).replace(/[\x00-\x08\x0b\x0c\x0e-\x1f]/g,"");
 const paragraph=(s,size=22,bold=false,pageBreak=false)=>'<w:p><w:pPr>'+(pageBreak?'<w:pageBreakBefore/>':'')+'<w:spacing w:after="120"/></w:pPr><w:r><w:rPr>'+(bold?'<w:b/><w:color w:val="17485A"/>':'')+'<w:sz w:val="'+size+'"/></w:rPr>'+s.split("\n").map(line=>'<w:t xml:space="preserve">'+xml(line)+'</w:t>').join('<w:br/>')+'</w:r></w:p>';
 let body="";
 for(const [i,page] of pages.entries()){
  body+=paragraph(page.title,36,true,i>0);
  for(const b of page.blocks){
   if(!b.src){body+=paragraph(b.text,b.heading?26:22,!!b.heading);continue;}
   const img=new Image();img.src=b.src;await img.decode();
   const n=rels.length+1,path="media/image"+n+(b.src.startsWith("data:image/png")?".png":".jpg"),scale=Math.min(5900000/img.width,6400000/img.height);
   const cx=Math.round(img.width*scale),cy=Math.round(img.height*scale),name="Obrazek"+n;
   files.push(["word/"+path,Uint8Array.from(atob(b.src.split(",")[1]),c=>c.charCodeAt(0))]);
   rels.push('<Relationship Id="img'+n+'" Type="'+office+'/image" Target="'+path+'"/>');
   body+='<w:p><w:r><w:drawing><wp:inline><wp:extent cx="'+cx+'" cy="'+cy+'"/><wp:docPr id="'+n+'" name="'+name+'"/><a:graphic><a:graphicData uri="'+draw+'picture"><pic:pic><pic:nvPicPr><pic:cNvPr id="'+n+'" name="'+name+'"/><pic:cNvPicPr/></pic:nvPicPr><pic:blipFill><a:blip r:embed="img'+n+'"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill><pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="'+cx+'" cy="'+cy+'"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>';
   if(b.caption)body+=paragraph(b.caption,20);
  }
 }
 const wn=ns+"/wordprocessingml/2006/main";
 const relationships=arr=>'<Relationships xmlns="'+ns+'/package/2006/relationships">'+arr.join("")+'</Relationships>';
 files.push(["[Content_Types].xml",'<Types xmlns="'+ns+'/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="jpg" ContentType="image/jpeg"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/footer.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/></Types>'],
 ["_rels/.rels",relationships(['<Relationship Id="doc" Type="'+office+'/officeDocument" Target="word/document.xml"/>'])],
 ["word/_rels/document.xml.rels",relationships([...rels,'<Relationship Id="foot" Type="'+office+'/footer" Target="footer.xml"/>'])],
 ["word/footer.xml",'<w:ftr xmlns:w="'+wn+'"><w:p><w:r><w:t xml:space="preserve">GX SOLUTIONS BOHEMIA, s.r.o.   '+xml(number)+'   </w:t></w:r><w:fldSimple w:instr="PAGE"/></w:p></w:ftr>'],
 ["word/document.xml",'<w:document xmlns:w="'+wn+'" xmlns:r="'+office+'" xmlns:wp="'+draw+'wordprocessingDrawing" xmlns:a="'+draw+'main" xmlns:pic="'+draw+'picture"><w:body>'+body+'<w:sectPr><w:footerReference w:type="default" r:id="foot"/><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="980" w:right="980" w:bottom="980" w:left="980" w:header="460" w:footer="460" w:gutter="0"/></w:sectPr></w:body></w:document>']);
 return packageZip(files);
}
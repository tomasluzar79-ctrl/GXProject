"use strict";
const $ = id => document.getElementById(id);
const tabs = ["Údaje projektu","Situace","Knihovna řešení","Stránky dokumentace","Náhled a export"];
const esc = value => String(value).replace(/[&<>"']/g, c => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));
const safeImage = s => typeof s === "string" && /^data:image\/(png|jpeg);base64,[A-Za-z0-9+/]+=*$/.test(s);
let p = initialProject(), tab = 0, database, timer, generation = 0;
const msg = s => $("status").textContent = s;
const view = $("view");
const field = (key,label,value,long=false) => '<label class="field">'+esc(label)+(long?'<textarea data-key="'+key+'">'+esc(value)+'</textarea>':'<input data-key="'+key+'" value="'+esc(value)+'">')+'</label>';
const button = (cmd,label,attrs="") => '<button data-cmd="'+cmd+'" '+attrs+'>'+esc(label)+'</button>';
const imageTag = (value,label) => safeImage(value)?'<img class="thumb" src="'+value+'" alt="'+esc(label)+'">':"";
const upload = (key,label,multiple=false) => '<label class="field">'+label+'<input type="file" data-upload="'+key+'" accept="image/jpeg,image/png,image/webp" '+(multiple?"multiple":"")+'></label>';
const sitName = (s,i) => 'Situace '+(s.numbers||String(i+1))+(s.title?' · '+s.title:'');
function owner(el){
 const box=el.closest('[data-group]');
 return box?p[box.dataset.group][Number(box.dataset.index)]:p;
}
function persist(){
 clearTimeout(timer);timer=0;
 if(!database){msg("Pro uchování změn použij Uložit projekt.");return;}
 const rev=++generation;
 msg("Ukládám…");
 const tx=database.transaction("draft","readwrite");
 tx.objectStore("draft").put(p,"current");
 tx.oncomplete=()=>{if(rev===generation) msg("Uloženo v tomto prohlížeči");};
 tx.onabort=tx.onerror=()=>msg("Místní uložení selhalo. Stáhni projekt tlačítkem Uložit projekt.");
}
function schedule(){generation++;msg("Změny k uložení");clearTimeout(timer);timer=setTimeout(persist,350);}
function projectView(){
 return '<section class="card"><h2>Identifikace zakázky</h2><div class="grid">'+Object.entries({number:"Číslo projektové dokumentace",customer:"Zadavatel",address:"Adresa instalace",contact:"Kontaktní osoba zákazníka",author:"Zpracovatel a kontakt GX"}).map(([k,t])=>field(k,t,p[k])).join("")+'</div></section><section class="card"><h2>Obrazové podklady</h2><div class="grid">'+[["logo","Logo zákazníka"],["map","Schéma lokalizace situací"]].map(([k,t])=>'<div>'+upload(k,t)+imageTag(p[k],t)+(p[k]?button("clear","Odebrat obrázek",'data-image="'+k+'"'):"")+'</div>').join("")+'</div></section>';
}
function situationView(group){
 const lib=group==="solutions", items=p[group];
 const top=lib?"Přidat vzor":"Přidat situaci";
 return '<div class="actions">'+button("add",top,'class="primary"')+(lib?"":'<select aria-label="Vzor řešení" id="preset" style="width:auto"><option value="">Vyber vzor řešení</option>'+p.solutions.map((s,i)=>'<option value="'+i+'">'+esc(s.title||"Bez názvu")+'</option>').join("")+'</select>'+button("from","Vložit ze vzoru"))+'</div>'+(!items.length?'<p class="hint">Projekt zatím nemá žádnou situaci.</p>':"")+items.map((s,i)=>'<details data-group="'+group+'" data-index="'+i+'"><summary>'+esc(lib?(s.title||"Nový vzor"):sitName(s,i))+'</summary>'+field("title","Název",s.title)+(lib?"":field("numbers","Označení na mapě (např. 1, 2, 3)",s.numbers))+Object.entries(fields).map(([k,t])=>field(k,t,s[k],true)).join("")+(lib?"":upload("photos","Fotografie a vizualizace",true)+s.photos.map((ph,n)=>'<div>'+imageTag(ph.src,ph.caption||"Vizualizace")+'<label class="field">Popisek<input data-caption="'+n+'" value="'+esc(ph.caption)+'"></label>'+button("removePhoto","Odebrat fotografii",'data-photo="'+n+'"')+'</div>').join(""))+'<div class="actions">'+button("up","Výše",i?"":"disabled")+button("down","Níže",i<items.length-1?"":"disabled")+button("copy","Duplikovat")+(lib?"":button("makePreset","Uložit jako vzor"))+button("remove","Odstranit")+'</div></details>').join("");
}
function sectionView(){
 return '<section class="card"><h2>Volitelné technické stránky</h2>'+p.extras.map((s,i)=>'<details data-group="extras" data-index="'+i+'"><summary>'+esc(s.title||"Nová stránka")+(s.enabled?" · zařazeno":"")+'</summary><label><input type="checkbox" data-key="enabled" '+(s.enabled?"checked":"")+'> Zařadit do dokumentace</label>'+field("title","Název stránky",s.title)+field("body","Text stránky",s.body,true)+button("remove","Odstranit stránku")+'</details>').join("")+button("extra","Přidat technickou stránku")+'</section><section class="card"><h2>Stálé části dokumentace</h2>'+Object.entries({intro:"Úvod",glossary:"Slovník pojmů",operation:"Charakter a provoz systému",maintenance:"Údržba a kontrola"}).map(([k,t])=>field(k,t,p[k],true)).join("")+'</section>';
}
// Shared document structure for preview, print and Word.
function documentPages(){
 const para=text=>({text}), heading=text=>({text,heading:true}), pic=(src,caption)=>({src,caption});
 const pages=[{title:"Projektová dokumentace",blocks:[heading("Zadávací část"),para("Aplikace systému GX ActiveSafety pro zákazníka"),para(p.customer),...(p.logo?[pic(p.logo,"")]:[]),para(p.number)]},
 {title:"Identifikace projektu",blocks:[para([["Číslo dokumentace",p.number],["Zadavatel",p.customer],["Adresa instalace",p.address],["Kontaktní osoba",p.contact]].filter(([,value])=>value).map(([label,value])=>label+" · "+value).join("\n")),heading("Vyhotovitel"),para("GX SOLUTIONS BOHEMIA, s.r.o.\n17. listopadu 907, 738 01 Frýdek-Místek\n"+p.author),para(p.intro)]},
 {title:"Obsah zadávací části",blocks:[]},{title:"Slovník pojmů",blocks:[para(p.glossary)]},
 {title:"Schéma lokalizace řešených situací v provozu",blocks:p.map?[pic(p.map,"")]:[]},
 {title:"Popis a návrh řešení situací v provozu",blocks:[para("Přehled řešených míst a jejich označení v provozu."),...p.situations.map((s,i)=>para(sitName(s,i)))]}];
 p.situations.forEach((s,i)=>{
  pages.push({title:sitName(s,i),blocks:Object.entries(fields).flatMap(([k,t])=>[heading(t),para(s[k])])});
  s.photos.forEach(ph=>pages.push({title:"Vizualizace řešení · "+sitName(s,i),blocks:[pic(ph.src,ph.caption)]}));
 });
 p.extras.filter(x=>x.enabled).forEach(x=>pages.push({title:x.title,blocks:[para(x.body)]}));
 pages.push({title:"Ustanovení o charakteru a provozu systému GX ActiveSafety",blocks:[para(p.operation)]},{title:"Údržba a kontrola systému GX ActiveSafety",blocks:[para(p.maintenance)]});
 pages[2].blocks=pages.slice(3).map((x,i)=>({text:x.title,target:i+3}));
 return pages;
}
function reportHTML(pages,links=false){
 return pages.map((page,i)=>'<section class="sheet '+(i===0?"cover":"")+'" '+(links?'id="sheet-'+i+'"':'')+'><small>GX ActiveSafety</small><h2>'+esc(page.title)+'</h2>'+page.blocks.map(b=>b.src?'<figure><img src="'+b.src+'" alt="'+esc(b.caption||page.title)+'"><figcaption>'+esc(b.caption)+'</figcaption></figure>':b.heading?'<h3>'+esc(b.text)+'</h3>':links&&b.target?'<p><a href="#sheet-'+b.target+'">'+esc(b.text)+'</a></p>':'<p>'+esc(b.text)+'</p>').join("")+'<footer>GX SOLUTIONS BOHEMIA, s.r.o. · '+esc(p.number)+'</footer></section>').join("");
}
function render(expand){
 $("nav").innerHTML=tabs.map((t,i)=>'<button data-tab="'+i+'" aria-current="'+(tab===i)+'">'+esc(t)+'</button>').join("");
 $("title").textContent=tabs[tab];$("projectName").textContent=p.customer||"Nový projekt";
 view.innerHTML=tab===0?projectView():tab===1?situationView("situations"):tab===2?situationView("solutions"):tab===3?sectionView():'<div class="actions">'+button("print","Tisk / uložit PDF",'class="primary"')+button("word","Stáhnout Word (.docx)")+'</div>'+reportHTML(documentPages(),true);
 if(expand!==undefined){const el=view.querySelector('details[data-index="'+expand+'"]');if(el)el.open=true;}
}
function fileDownload(blob,name){
 const url=URL.createObjectURL(blob),a=document.createElement("a");a.href=url;a.download=name;a.click();setTimeout(()=>URL.revokeObjectURL(url),30000);
}
const filename=()=>(p.number||p.customer||"GXProject").replace(/[<>:"/\\|?*\x00-\x1F]/g,"-");
function validate(data){
 const string=(x)=>typeof x==="string";
 const picture=x=>x===""||safeImage(x);
 const sit=s=>s&&["numbers","title",...Object.keys(fields)].every(k=>string(s[k]))&&Array.isArray(s.photos)&&s.photos.every(ph=>ph&&safeImage(ph.src)&&string(ph.caption));
 if(!data||data.version!==1||!["number","customer","address","contact","author","intro","glossary","operation","maintenance"].every(k=>string(data[k]))||!picture(data.logo)||!picture(data.map)||!["situations","solutions"].every(k=>Array.isArray(data[k])&&data[k].every(sit))||!Array.isArray(data.extras)||!data.extras.every(x=>x&&string(x.title)&&string(x.body)&&typeof x.enabled==="boolean"))throw Error("Soubor nemá platný formát projektu GXProject.");
 return data;
}
async function photo(file){
 const url=URL.createObjectURL(file);
 try{const img=new Image();img.src=url;await img.decode();const ratio=Math.min(1,1800/Math.max(img.width,img.height));const canvas=document.createElement("canvas");canvas.width=Math.max(1,Math.round(img.width*ratio));canvas.height=Math.max(1,Math.round(img.height*ratio));const ctx=canvas.getContext("2d");ctx.fillStyle="white";ctx.fillRect(0,0,canvas.width,canvas.height);ctx.drawImage(img,0,0,canvas.width,canvas.height);return canvas.toDataURL("image/jpeg",0.87);}finally{URL.revokeObjectURL(url);}
}
view.addEventListener("input",e=>{
 const el=e.target,obj=owner(el),key=el.dataset.key;
 if(key){obj[key]=el.type==="checkbox"?el.checked:el.value;const box=el.closest("details");if(box&&["title","numbers","enabled"].includes(key))box.querySelector("summary").textContent=box.dataset.group==="situations"?sitName(obj,Number(box.dataset.index)):(obj.title||"Bez názvu")+(obj.enabled?" · zařazeno":"");if(key==="customer")$("projectName").textContent=obj.customer||"Nový projekt";schedule();}
 else if(el.dataset.caption!==undefined){obj.photos[Number(el.dataset.caption)].caption=el.value;schedule();}
});
view.addEventListener("change",async e=>{
 const el=e.target,key=el.dataset.upload;if(!key||!el.files.length)return;
 const obj=owner(el),project=p,section=tab;el.disabled=true;
 try{const files=Array.from(el.files),imgs=[];for(const f of files)imgs.push(await photo(f));if(p!==project)return;if(key==="photos")obj.photos.push(...imgs.map((src,i)=>({src,caption:files[i].name.replace(/\.[^.]+$/,"")})));else obj[key]=imgs[0];persist();if(section===tab)render(p.situations.indexOf(obj));}
 catch{msg("Obrázek se nepodařilo načíst. Zkus JPG nebo PNG.");}finally{el.disabled=false;}
});
view.addEventListener("click",async e=>{
 const b=e.target.closest("[data-cmd]");if(!b)return;
 const box=b.closest("details"),group=box?.dataset.group||(tab===2?"solutions":"situations"),items=p[group],i=Number(box?.dataset.index),obj=items[i];let idx=i;
 switch(b.dataset.cmd){
 case "add":idx=items.push(blankSituation())-1;break;
 case "from":{const v=$("preset").value;if(v===""){msg("Vyber vzor řešení.");return;}idx=items.push(structuredClone(p.solutions[Number(v)]))-1;break;}
 case "copy":items.splice(i+1,0,structuredClone(obj));idx=i+1;break;
 case "up":case "down":idx=i+(b.dataset.cmd==="up"?-1:1);[items[i],items[idx]]=[items[idx],items[i]];break;
 case "makePreset":p.solutions.push({...structuredClone(obj),numbers:"",photos:[]});persist();msg("Vzor přidán do knihovny řešení.");return;
 case "remove":if(!confirm("Odstranit tuto položku?"))return;items.splice(i,1);break;
 case "extra":idx=p.extras.push({title:"Nová technická stránka",body:"",enabled:true})-1;break;
 case "removePhoto":obj.photos.splice(Number(b.dataset.photo),1);break;
 case "clear":p[b.dataset.image]="";break;
 case "print":window.print();return;
 case "word":try{b.disabled=true;fileDownload(await wordFile(documentPages(),p.number),filename()+".docx");}catch{msg("Export do Wordu se nezdařil.");}finally{b.disabled=false;}return;
 default:return;
 }persist();render(idx);
});
$("nav").onclick=e=>{if(e.target.dataset.tab!==undefined){tab=Number(e.target.dataset.tab);render();}};
$("save").onclick=()=>fileDownload(new Blob([JSON.stringify(p)],{type:"application/json"}),filename()+".gxproject.json");
$("open").onchange=async e=>{try{if(!e.target.files.length)return;const next=validate(JSON.parse(await e.target.files[0].text()));if(!confirm("Otevřít uložený projekt místo aktuálního?"))return;p=next;tab=0;persist();render();}catch{msg("Tento soubor nelze otevřít jako projekt GXProject.");}finally{e.target.value="";}};
$("fresh").onclick=()=>{if(confirm("Založit nový projekt? Aktuální projekt si můžeš předem stáhnout. Vzory a technické stránky se přenesou.")){p={...initialProject(),solutions:structuredClone(p.solutions),extras:p.extras.map(x=>({...x,enabled:false}))};tab=0;persist();render();}};
window.addEventListener("beforeprint",()=>$("report").innerHTML=reportHTML(documentPages()));
window.addEventListener("afterprint",()=>$("report").replaceChildren());
document.addEventListener("visibilitychange",()=>{if(document.hidden&&timer)persist();});
async function start(){
 try{database=await new Promise((resolve,reject)=>{const r=indexedDB.open("GXProject-v1",1);r.onupgradeneeded=()=>r.result.createObjectStore("draft");r.onsuccess=()=>resolve(r.result);r.onerror=()=>reject(r.error);});const stored=await new Promise((resolve,reject)=>{const r=database.transaction("draft").objectStore("draft").get("current");r.onsuccess=()=>resolve(r.result);r.onerror=()=>reject(r.error);});if(stored)p=validate(stored);msg(stored?"Obnoven místní projekt":"Projekt se ukládá v tomto prohlížeči");}
 catch{msg("Pro uchování projektu použij Uložit projekt.");}
 render();
}
start();

#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Этап 18Б, шаг 3: встроенный импорт заполненного Excel-калькулятора № 9 (interop).
Вставляет в модель:
  1) модуль xlsx18 (ZIP-читатель + DEFLATE-распаковщик + разбор OOXML через DOMParser),
  2) importExcel9 / importPreview18 / importCard18 / importShow18 (+ хелперы x18Num/Txt/Date, deriveConsensus18),
  3) карточку «📥 Импорт…» в конце renderCheck (перед заметкой о legacy31),
  4) wiring поля input[type=file] #impXlsx18 после строки q18Md,
  5) экспансии в niokrAPI.
Только ОДИН прогон против НЕпропатченного step3 HTML (якоря проверяются assert'ами).
"""
import io, re

MODEL = 'output/Аудит_НИОКР_интерактивная_модель.html'

s = io.open(MODEL, encoding='utf-8').read()
assert 'function importExcel9' not in s, 'шаг 3 уже применён — не перезапускать!'
assert s.count('const Q18=') == 1

# ---------------- JS-модуль ----------------
JS_MODULE = r'''
/* ===== xlsx18 — встроенный читатель .xlsx: ZIP (store/deflate) + DEFLATE + OOXML через DOMParser ===== */
const xlsx18=(function(){
 var utf8dec=new TextDecoder('utf-8');
 function inflateRaw(src,outSize){
  var out=new Uint8Array(outSize>0?outSize:(1<<20)),op=0,cap=out.length;
  function need(n){
   if(op+n<=cap)return;
   while(cap<op+n)cap=cap*2||65536;
   var nb=new Uint8Array(cap);nb.set(out.subarray(0,op));out=nb;cap=nb.length;}
  var ip=0,bBuf=0,bCnt=0;
  function bits(n){
   while(bCnt<n){bBuf|=(src[ip++]<<bCnt);bCnt+=8;}
   var v=bBuf&((1<<n)-1);bBuf>>>=n;bCnt-=n;return v;}
  function Huff(lens){
   var max=0,i;
   for(i=0;i<lens.length;i++){if(lens[i]>max)max=lens[i];}
   this.max=max;
   this.count=new Int32Array(max+1);
   this.symbol=new Int32Array(lens.length);
   for(i=0;i<lens.length;i++){if(lens[i])this.count[lens[i]]++;}
   var offs=new Int32Array(max+1),sum=0;
   for(i=1;i<=max;i++){offs[i]=sum;sum+=this.count[i];}
   for(i=0;i<lens.length;i++){if(lens[i])this.symbol[offs[lens[i]]++]=i;}}
  function decode(h){
   var code=0,first=0,index=0;
   for(var len=1;len<=h.max;len++){
    code|=bits(1);
    if(code-first<h.count[len])return h.symbol[index+(code-first)];
    index+=h.count[len];
    first=(first+h.count[len])<<1;
    code<<=1;}
   throw new Error('inflate: повреждённый поток');}
  var i,gl=new Int32Array(288);
  for(i=0;i<144;i++)gl[i]=8;for(;i<256;i++)gl[i]=9;for(;i<280;i++)gl[i]=7;for(;i<288;i++)gl[i]=8;
  var FL=new Huff(gl),gd=new Int32Array(30);for(i=0;i<30;i++)gd[i]=5;var FD=new Huff(gd);
  var LENB=[3,4,5,6,7,8,9,10,11,13,15,17,19,23,27,31,35,43,51,59,67,83,99,115,131,163,195,227,258];
  var LENX=[0,0,0,0,0,0,0,0,1,1,1,1,2,2,2,2,3,3,3,3,4,4,4,4,5,5,5,5,0];
  var DSB=[1,2,3,4,5,7,9,13,17,25,33,49,65,97,129,193,257,385,513,769,1025,1537,2049,3073,4097,6145,8193,12289,16385,24577];
  var DSX=[0,0,0,0,1,1,2,2,3,3,4,4,5,5,6,6,7,7,8,8,9,9,10,10,11,11,12,12,13,13];
  function block(lit,dis){
   for(;;){
    var sym=decode(lit);
    if(sym<256){need(1);out[op++]=sym;}
    else if(sym===256)return;
    else{
     sym=sym-257;
     var len=LENB[sym]+bits(LENX[sym]);
     var d=decode(dis);
     var dist=DSB[d]+bits(DSX[d]);
     if(dist>op)throw new Error('inflate: недопустимая дистанация');
     need(len);
     for(var k=0;k<len;k++){out[op]=out[op-dist];op++;}}}}
  for(;;){
   var last=bits(1),type=bits(2);
   if(type===0){
    bBuf=0;bCnt=0;
    var l=src[ip]|(src[ip+1]<<8);ip+=4;
    need(l);
    for(var c2=0;c2<l;c2++)out[op++]=src[ip++];}
   else if(type===1){block(FL,FD);}
   else if(type===2){
    var hlit=bits(5)+257,hdist=bits(5)+1,hclen=bits(4)+4;
    var order=[16,17,18,0,8,7,9,6,10,5,11,4,12,3,13,2,14,1,15];
    var cl=new Int32Array(19);
    for(var z=0;z<hclen;z++)cl[order[z]]=bits(3);
    var HC=new Huff(cl),lensL=new Int32Array(hlit+hdist);z=0;
    while(z<hlit+hdist){
     var sv=decode(HC);
     if(sv<16){lensL[z++]=sv;}
     else if(sv===16){var prev=lensL[z-1],n2=3+bits(2);while(n2-- >0)lensL[z++]=prev;}
     else if(sv===17){n2=3+bits(3);while(n2-- >0)lensL[z++]=0;}
     else{n2=11+bits(7);while(n2-- >0)lensL[z++]=0;}}
    var L=new Huff(lensL.subarray(0,hlit)),D=new Huff(lensL.subarray(hlit));
    block(L,D);}
   else throw new Error('inflate: неизвестный тип блока');
   if(last)break;}
  return out.subarray(0,op);}
 function unzip(u8){
  if(u8.length<22)throw new Error('zip: файл слишком мал (это не .xlsx?)');
  var dv=new DataView(u8.buffer,u8.byteOffset,u8.byteLength);
  var eocd=-1;
  for(var i=u8.length-22;i>=0&&i>u8.length-66350;i--){
   if(dv.getUint32(i,true)===0x06054b50){eocd=i;break;}}
  if(eocd<0)throw new Error('zip: не найден конец центрального каталога');
  var n=dv.getUint16(eocd+10,true),p=dv.getUint32(eocd+16,true),files={};
  for(var e=0;e<n;e++){
   if(dv.getUint32(p,true)!==0x02014b50)break;
   var method=dv.getUint16(p+10,true),csz=dv.getUint32(p+20,true),usz=dv.getUint32(p+24,true);
   var nl=dv.getUint16(p+28,true),el=dv.getUint16(p+30,true),kl=dv.getUint16(p+32,true);
   var lo=dv.getUint32(p+42,true);
   var name=utf8dec.decode(u8.subarray(p+46,p+46+nl));
   var ln=dv.getUint16(lo+26,true),le=dv.getUint16(lo+28,true);
   var ds=lo+30+ln+le;
   var raw=u8.subarray(ds,ds+csz);
   if(name.indexOf('__MACOSX')===0||name.charAt(name.length-1)==='/'){}
   else files[name]=(method===0)?raw:inflateRaw(raw,usz);
   p+=46+nl+el+kl;}
  return files;}
 function parse(buf){
  var u8=(buf instanceof Uint8Array)?buf:new Uint8Array(buf);
  var files=unzip(u8);
  if(typeof DOMParser==='undefined')throw new Error('xlsx18: нет DOMParser');
  var dp=new DOMParser();
  function doc(path){var f=files[path];if(!f)return null;return dp.parseFromString(utf8dec.decode(f),'application/xml');}
  var wb=doc('xl/workbook.xml');
  if(!wb)throw new Error('xlsx: нет xl/workbook.xml — файл не похож на книгу Excel');
  var rels=doc('xl/_rels/workbook.xml.rels'),rel={};
  if(rels){var rl=rels.getElementsByTagName('Relationship');
   for(var i=0;i<rl.length;i++)rel[rl[i].getAttribute('Id')]=rl[i].getAttribute('Target');}
  var NSOO='http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  var shList=wb.getElementsByTagName('sheet'),pathOf={},order=[];
  for(i=0;i<shList.length;i++){
   var sh=shList[i],nm=sh.getAttribute('name');
   var rid=null;
   try{rid=sh.getAttributeNS(NSOO,'id');}catch(e0){}
   if(!rid)rid=sh.getAttribute('r:id');
   if(!nm||!rid)continue;
   var tg=String(rel[rid]||'').replace(/^\/+/,'');
   if(tg.indexOf('xl/')!==0)tg='xl/'+tg;
   pathOf[nm]=tg;order.push(nm);}
  var sst=[];var ss=doc('xl/sharedStrings.xml');
  if(ss){var si=ss.getElementsByTagName('si');
   for(i=0;i<si.length;i++){var tt=si[i].getElementsByTagName('t'),txt='';
    for(var j=0;j<tt.length;j++)txt+=tt[j].textContent;
    sst.push(txt);}}
  var cells={};
  function readSheet(nm){
   var d=doc(pathOf[nm]),cs={};
   if(!d)return cs;
   var cc=d.getElementsByTagName('c');
   for(var i2=0;i2<cc.length;i2++){
    var el=cc[i2],addr=el.getAttribute('r');
    if(!addr)continue;
    var ty=el.getAttribute('t')||'n',v=null,vn=el.getElementsByTagName('v')[0];
    if(ty==='s'){v=(vn!=null)?sst[parseInt(vn.textContent,10)]:null;}
    else if(ty==='str'){v=(vn!=null)?vn.textContent:null;}
    else if(ty==='inlineStr'){var iso=el.getElementsByTagName('is')[0];v=iso?(iso.textContent||''):null;}
    else if(ty==='b'){v=(vn!=null)?(vn.textContent==='1'):null;}
    else if(ty==='e'){v=null;}
    else{if(vn!=null){var pv=parseFloat(vn.textContent);v=isFinite(pv)?pv:null;}}
    if(v!=null&&v!=='')cs[addr.toUpperCase()]=v;}
   return cs;}
  for(i=0;i<order.length;i++)cells[order[i]]=readSheet(order[i]);
  return {sheetNames:order,cells:cells,
   get:function(nm,addr){var cs=cells[nm];if(!cs)return null;var v2=cs[String(addr).toUpperCase()];return (v2==null||v2==='')?null:v2;}};}
 return {parse:parse,inflate:inflateRaw};})();

function x18Num(v){if(v==null||v==='')return null;var n=parseFloat(String(v).replace(',','.'));return isFinite(n)?n:null;}
function x18Txt(v){return v==null?'':String(v).trim();}
function x18Esc(s){return String(s==null?'':s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');}
function x18Date(v){var n=x18Num(v);if(n==null||n<1||n>80000)return null;var d=new Date(Math.round((n-25569)*86400000));return isNaN(d.getTime())?null:d;}
function x18Dfmt(d){if(!d)return '';return ('0'+d.getDate()).slice(-2)+'.'+('0'+(d.getMonth()+1)).slice(-2)+'.'+d.getFullYear();}
/* консенсус экспертов по образцу формулы M листа № 9: числовые — средневзвешенное (ROUND 2); НП побеждает, если его больше, чем числовых; иначе НП/НД/'' */
function deriveConsensus18(vals,wts){
 var nums=[],np=0,nd=0,i;
 for(i=0;i<vals.length;i++){
  var t=x18Txt(vals[i]);
  if(t==='НП'){np++;continue;}
  if(t==='НД'){nd++;continue;}
  var n=x18Num(vals[i]);
  if(n!=null)nums.push(n);}
 if(!nums.length){if(np>0)return 'НП';if(nd>0)return 'НД';return '';}
 if(np>nums.length)return 'НП';
 if(!wts||wts.length!==3)wts=[1,1,1];
 var sw=0,ss=0;
 for(i=0;i<vals.length;i++){var n2=x18Num(vals[i]);if(n2!=null){ss+=n2*(wts[i]||0);sw+=(wts[i]||0);}}
 var m=sw>0?ss/sw:(nums.reduce(function(a,b){return a+b;},0)/nums.length);
 return String(Math.round(m*100)/100);}
function x18SheetOk(wbk){
 return !!wbk.cells && (wbk.cells['Паспорт_НИОКР']||wbk.cells['Краткий_15']||wbk.cells['Средний_30']||wbk.cells['Полный_60']);}
function importPreview18(wbk){
 var need=['Паспорт_НИОКР','Краткий_15','Средний_30','Полный_60','Стоп_факторы','Зрелость_M0_M9','Экономика_LCC','AHP','Эксперты'];
 var found=[],miss=[];
 for(var i=0;i<need.length;i++){(wbk.cells[need[i]]?found:miss).push(need[i]);}
 var g=function(sn,a2){return wbk.get(sn,a2);};
 var out={found:found,miss:miss,q:{K:0,S:0,P:0},lvl:null,name:'',passport:false,sf2:0,sf1:0,mat:'',econ:'',ahp:0,ub:false,total:0,notes:[]};
 if(wbk.cells['Паспорт_НИОКР']){
  out.name=x18Txt(g('Паспорт_НИОКР','C6'));
  out.passport=!!(out.name||x18Txt(g('Паспорт_НИОКР','C7'))||x18Txt(g('Паспорт_НИОКР','C14')));
  var mm={'Краткий':'K','Средний':'S','Полный':'P'};
  out.lvl=mm[x18Txt(g('Паспорт_НИОКР','C17'))]||null;
  if(!out.lvl&&x18Txt(g('Паспорт_НИОКР','C17')))out.notes.push('уровень «'+x18Txt(g('Паспорт_НИОКР','C17'))+'» не распознан');}
 function cnt(sn,lk,rows0,rows1){
  if(!wbk.cells[sn])return;
  var a=0;
  for(var rw=rows0;rw<=rows1;rw++){
   var mv=g(sn,'M'+rw),hv=g(sn,'H'+rw),nv=g(sn,'N'+rw),ov=g(sn,'O'+rw);
   if((mv!=null&&mv!=='')||(hv!=null&&hv!=='')||(nv!=null&&nv!=='')||(ov!=null&&ov!==''))a++;}
  out.q[lk]=a;out.total+=a;}
 cnt('Краткий_15','K',6,20);cnt('Средний_30','S',6,35);cnt('Полный_60','P',6,65);
 if(wbk.cells['Стоп_факторы']){
  for(var s2=6;s2<=15;s2++){
   var st=x18Txt(g('Стоп_факторы','B'+s2));
   if(st==='подтверждено')out.sf2++;
   else if(st==='выявлена неопределённость')out.sf1++;}}
 if(wbk.cells['Зрелость_M0_M9']){
  var c=x18Txt(g('Зрелость_M0_M9','C17')),p=x18Txt(g('Зрелость_M0_M9','C18')),t=x18Txt(g('Зрелость_M0_M9','C22'));
  out.mat='M'+(c||'—')+' → M'+(p||'—')+(t?', испытания: '+String(t).toLowerCase():'');}
 if(wbk.cells['Экономика_LCC']){
  var nf=0,nl=0,na=0;
  for(var f2=0;f2<11;f2++){if(g('Экономика_LCC','B'+(21+f2))!=null||g('Экономика_LCC','C'+(21+f2))!=null)nf++;}
  for(var l2=0;l2<8;l2++){if(g('Экономика_LCC','D'+(9+l2))!=null)nl++;}
  for(var a2=0;a2<9;a2++){for(var cix=0;cix<4;cix++){if(g('Экономика_LCC',String.fromCharCode(67+cix)+(43+a2))!=null)na++;}}
  var rv=x18Txt(g('Экономика_LCC','C4')),tv=x18Txt(g('Экономика_LCC','C5'));
  out.econ='r='+(rv||'—')+', T='+(tv||'—')+' лет, потоки '+nf+'/11, LCC '+nl+'/8, альтернативы '+na+'/36';}
 if(wbk.cells['AHP']){
  var cellsAHP=[['dir','C6'],['dir','D6'],['dir','D7'],
   ['A','C18'],['A','D18'],['A','E18'],['A','D19'],['A','E19'],['A','E20'],
   ['B','C30'],['B','D30'],['B','E30'],['B','D31'],['B','E31'],['B','E32'],
   ['V','C42'],['V','D42'],['V','E42'],['V','D43'],['V','E43'],['V','E44']];
  for(var q2=0;q2<cellsAHP.length;q2++){
   var av=x18Num(g('AHP',cellsAHP[q2][1]));
   if(av!=null&&av!==1)out.ahp++;}
  if(wbk.cells['Полный_60'])out.ub=(x18Txt(g('Полный_60','C82'))==='Да');}
 return out;}
function importExcel9(wbk,r){
 var rep={answers:{K:0,S:0,P:0},passport:0,sf:0,mat:0,econ:0,ahp:0,lvl:r.lvl,notes:[]};
 if(!x18SheetOk(wbk))throw new Error('Файл не похож на калькулятор № 9: нет листов Паспорт_НИОКР / Краткий_15 / Средний_30 / Полный_60');
 var g=function(sn,a2){return wbk.get(sn,a2);};
 var has=function(sn){return !!wbk.cells[sn];};
 /* --- Паспорт --- */
 if(has('Паспорт_НИОКР')){
  var pm=Q18.import.passport;
  var nm=x18Txt(g('Паспорт_НИОКР',pm.name));if(nm){r.name=nm;rep.passport++;}
  var shf=x18Txt(g('Паспорт_НИОКР',pm.code));if(shf){r.shifr=shf;rep.passport++;}
  var ct=x18Txt(g('Паспорт_НИОКР',pm.contract));if(ct){r.contract=ct;rep.passport++;}
  var cu=x18Txt(g('Паспорт_НИОКР',pm.customer));if(cu){r.customer=cu;rep.passport++;}
  var ex=x18Txt(g('Паспорт_НИОКР',pm.exec));if(ex){r.executor=ex;rep.passport++;}
  var co=x18Txt(g('Паспорт_НИОКР',pm.coop));if(co)r.coop=co;
  var d1=x18Date(g('Паспорт_НИОКР',pm.start)),d2=x18Date(g('Паспорт_НИОКР',pm.end));
  if(d1||d2){r.dates=x18Dfmt(d1)+' — '+x18Dfmt(d2);rep.passport++;}
  if(d2)r.endDate=d2.getFullYear()+'-'+('0'+(d2.getMonth()+1)).slice(-2)+'-'+('0'+d2.getDate()).slice(-2);
  var cst=x18Txt(g('Паспорт_НИОКР',pm.cost));if(cst){r.cost=cst;rep.passport++;}
  var stg=x18Txt(g('Паспорт_НИОКР',pm.stage));if(stg)r.stage=stg;
  var ch=x18Txt(g('Паспорт_НИОКР',pm.chair));if(ch){r.head=ch;rep.passport++;}
  var sec=x18Txt(g('Паспорт_НИОКР',pm.secr));if(sec)r.secretary=sec;
  var mplan=x18Txt(g('Паспорт_НИОКР',pm.mplan));if(mplan)r.mplan=mplan;
  var ad=x18Date(g('Паспорт_НИОКР',pm.adate));if(ad)r.auditDate=x18Dfmt(ad);
  var mm2={'Краткий':'K','Средний':'S','Полный':'P'};
  var lv=x18Txt(g('Паспорт_НИОКР',pm.level));
  if(mm2[lv]){r.lvl=mm2[lv];rep.lvl=mm2[lv];}
  else if(lv)rep.notes.push('уровень в паспорте не распознан: «'+lv+'» (оставлен '+Q18.levels[r.lvl].name+')');}
 /* --- веса экспертов --- */
 var wts=[1,1,1];
 if(has('Эксперты')){
  for(var w2=0;w2<3;w2++){var wv=x18Num(g('Эксперты','D'+(5+w2)));if(wv!=null)wts[w2]=wv;}}
 /* --- опросники --- */
 var row30rev={};
 for(var cd in Q18.import.row30)row30rev[Q18.import.row30[cd]]=cd;
 var qmaps=[['K','Краткий_15',6,20],['S','Средний_30',6,35],['P','Полный_60',6,65]];
 for(var qi=0;qi<qmaps.length;qi++){
  var lk=qmaps[qi][0],sn=qmaps[qi][1],rw0=qmaps[qi][2],rw1=qmaps[qi][3];
  if(!has(sn))continue;
  if(!r.q)r.q={K:{},S:{},P:{}};
  if(!r.q[lk])r.q[lk]={};
  var qs=Q18.levels[lk].q,byCode={},z2;
  for(z2=0;z2<qs.length;z2++)byCode[String(qs[z2].code)]=qs[z2];
  for(var rw=rw0;rw<=rw1;rw++){
   var qd=lk==='S'?byCode[String(row30rev[rw]||'')]:(qs[rw-rw0]||null);
   if(!qd)continue;
   var mv=g(sn,'M'+rw),a='';
   var mvT=x18Txt(mv);
   if(mvT==='НД'||mvT==='НП')a=mvT;
   else{var mn=x18Num(mv);if(mn!=null)a=String(mn);}
   if(a==='')a=deriveConsensus18([g(sn,'H'+rw),g(sn,'I'+rw),g(sn,'J'+rw)],wts);
   var jv=x18Txt(g(sn,'N'+rw)),ev=x18Txt(g(sn,'O'+rw));
   if(a===''&&!jv&&!ev)continue;
   var old=r.q[lk][qd.code]||{a:'',j:'',e:''};
   r.q[lk][qd.code]={a:a,j:jv||old.j||'',e:ev||old.e||''};
   rep.answers[lk]++;}}
 /* --- стоп-факторы --- */
 if(has('Стоп_факторы')){
  if(!r.sf)r.sf=Array(10).fill('не выявлено');
  if(!r.sfc)r.sfc=Array(5).fill('');
  for(var si=0;si<10;si++){
   var stt=x18Txt(g('Стоп_факторы','B'+(6+si)));
   if(stt==='не выявлено'||stt==='выявлена неопределённость'||stt==='подтверждено'){
    if(r.sf[si]!==stt)rep.sf++;
    r.sf[si]=stt;}
   else if(stt)rep.notes.push('Стоп_факторы B'+(6+si)+': неизвестный статус «'+stt+'»');}
  for(var ci=0;ci<5;ci++){
   var txt=x18Txt(g('Стоп_факторы','A'+(11+ci)));
   if(txt&&!/^Иное обстоятельство \d+ \(вписать\):?$/i.test(txt))r.sfc[ci]=txt;}}
 /* --- зрелость --- */
 if(has('Зрелость_M0_M9')){
  if(!r.mat)r.mat={cur:null,plan:null,tests:'Нет'};
  var cv=x18Num(g('Зрелость_M0_M9','C17'));if(cv!=null){r.mat.cur=Math.round(cv);rep.mat++;}
  var pvv=x18Num(g('Зрелость_M0_M9','C18'));if(pvv!=null){r.mat.plan=Math.round(pvv);rep.mat++;}
  var tv2=x18Txt(g('Зрелость_M0_M9','C22'));if(tv2){r.mat.tests=tv2;rep.mat++;}
  var note2=x18Txt(g('Зрелость_M0_M9','C20'));if(note2)r.mat.note=note2;}
 /* --- экономика --- */
 if(has('Экономика_LCC')){
  if(!r.econ)r.econ={rate:0.1,T:5,lcc:Array.from({length:8},function(){return ['',''];}),flows:Array.from({length:11},function(){return ['',''];}),altW:Array(9).fill(1),alt:{c:Array(9).fill(''),d:Array(9).fill(''),e:Array(9).fill(''),f:Array(9).fill('')}};
  var rv2=x18Num(g('Экономика_LCC','C4'));if(rv2!=null)r.econ.rate=rv2;
  var tv3=x18Num(g('Экономика_LCC','C5'));if(tv3!=null)r.econ.T=Math.round(tv3);
  var touched=0;
  for(var li=0;li<8;li++){
   var dval=g('Экономика_LCC','D'+(9+li)),eval2=g('Экономика_LCC','E'+(9+li));
   if(dval!=null||eval2!=null)touched++;
   r.econ.lcc[li]=[dval==null?'':String(dval),eval2==null?'':String(eval2)];}
  for(var fi=0;fi<11;fi++){
   var bv=g('Экономика_LCC','B'+(21+fi)),cvv=g('Экономика_LCC','C'+(21+fi));
   if(bv!=null||cvv!=null)touched++;
   r.econ.flows[fi]=[bv==null?'':String(bv),cvv==null?'':String(cvv)];}
  for(var ai=0;ai<9;ai++){
   var aw=g('Экономика_LCC','B'+(43+ai));
   r.econ.altW[ai]=aw==null?1:String(aw);
   ['c','d','e','f'].forEach(function(key,idx4){
    var av2=g('Экономика_LCC',String.fromCharCode(67+idx4)+(43+ai));
    r.econ.alt[key][ai]=av2==null?'':String(av2);
    if(av2!=null)touched++;});}
  if(touched)rep.econ=touched;}
 /* --- AHP --- */
 if(has('AHP')){
  if(!r.ahp18)r.ahp18={dir:{a12:'1',a13:'1',a23:'1'},A:{a12:'1',a13:'1',a14:'1',a23:'1',a24:'1',a34:'1'},B:{a12:'1',a13:'1',a14:'1',a23:'1',a24:'1',a34:'1'},V:{a12:'1',a13:'1',a14:'1',a23:'1',a24:'1',a34:'1'},useBlocks:false};
  var amap=[['dir',[['a12','C6'],['a13','D6'],['a23','D7']]],
   ['A',[['a12','C18'],['a13','D18'],['a14','E18'],['a23','D19'],['a24','E19'],['a34','E20']]],
   ['B',[['a12','C30'],['a13','D30'],['a14','E30'],['a23','D31'],['a24','E31'],['a34','E32']]],
   ['V',[['a12','C42'],['a13','D42'],['a14','E42'],['a23','D43'],['a24','E43'],['a34','E44']]]];
  for(var mi=0;mi<amap.length;mi++){
   var grp=amap[mi][0],pairs=amap[mi][1];
   for(var pi=0;pi<pairs.length;pi++){
    var av3=x18Num(g('AHP',pairs[pi][1]));
    if(av3!=null&&av3>0){r.ahp18[grp][pairs[pi][0]]=String(Math.round(av3*100)/100);rep.ahp++;}}}
  if(has('Полный_60'))r.ahp18.useBlocks=(x18Txt(g('Полный_60','C82'))==='Да');}
 return rep;}
function importCard18(){
 var h='<div class="card"><h3>📥 Импорт из Excel-калькулятора № 9 (офлайн-ответы комиссии)</h3>';
 h+='<div class="mini">Файл <b>9_Калькулятор_аудита_НИОКР_v2.xlsx</b>, заполненный комиссией без модели. Загружаются: паспорт НИОКР и уровень опроса, ответы всех трёх опросников (колонка «ИТОГ», при её отсутствии — консенсус оценок экспертов Э1–Э3 с весами листа «Эксперты»), обоснования (N) и доказательства (O), стоп-факторы, зрелость M0–M9, экономика (r, T, LCC, потоки, альтернативы), матрицы AHP и переключатель весов блоков. Перед применением показывается предпросмотр.</div>';
 if(!canPass())return h+'<div class="mini" style="color:#c62828">Роль «Наблюдатель» — импорт данных недоступен.</div></div>';
 h+='<div style="margin-top:8px"><input type="file" id="impXlsx18" accept=".xlsx"></div>';
 h+='<div class="mini" style="margin-top:4px">Совет: сначала выберите целевую НИОКР во вкладке «Реестр» — ответы будут записаны в неё.</div>';
 return h+'</div>';}
function importShow18(wbk,r,name){
 var pv=importPreview18(wbk);
 var h='<div style="font-size:13px">';
 h+='<div><b>'+x18Esc(name)+'</b> → НИОКР <b>'+x18Esc(r?r.name:'—')+'</b></div>';
 h+='<div class="mini" style="margin:4px 0">Распознаны листы: '+x18Esc(pv.found.join(', ')||'—')+'</div>';
 if(pv.miss.length)h+='<div class="mini">Нет в файле: '+x18Esc(pv.miss.join(', '))+'</div>';
 function ln2(lbl,val){h+='<div style="display:flex;justify-content:space-between;gap:12px;border-bottom:1px solid #eee;padding:3px 0"><span class="mini">'+lbl+'</span><span><b>'+val+'</b></span></div>';}
 ln2('Вопросы заполнены — Краткий_15',pv.q.K+' из 15');
 ln2('· Средний_30',pv.q.S+' из 30');
 ln2('· Полный_60',pv.q.P+' из 60');
 ln2('Уровень аудита (Паспорт)',pv.lvl?x18Esc(Q18.levels[pv.lvl].name)+' — станет активным':'— останется '+x18Esc(Q18.levels[(r&&r.lvl)?r.lvl:'S'].name));
 ln2('Паспорт НИОКР',pv.passport?x18Esc(pv.name||'заполнен'):'пустой');
 ln2('Стоп-факторы','подтверждено: '+pv.sf2+' · неопределённость: '+pv.sf1);
 if(pv.mat)ln2('Зрелость M0–M9',x18Esc(pv.mat));
 if(pv.econ)ln2('Экономика',x18Esc(pv.econ));
 if(pv.found.indexOf('AHP')>=0)ln2('AHP',pv.ahp+' парных сравнений ≠ 1 · веса блоков: '+(pv.ub?'Да':'Нет'));
 pv.notes.forEach(function(nt){h+='<div class="mini" style="color:#b26a00">⚠ '+x18Esc(nt)+'</div>';});
 if(!pv.total)h+='<div class="mini" style="color:#c62828">Ответов экспертов не найдено — будут импортированы только паспорт, стоп-факторы, зрелость, экономика и AHP.</div>';
 h+='</div>';
 showModal('<h3>📥 Импорт из Excel-калькулятора № 9</h3>'+h,[
  {label:'Отмена'},
  {label:'✅ Импортировать',cls:'ok',fn:function(){
   try{
    var rep=importExcel9(wbk,r);
    save();renderAll();
    logEvent(r,'Импорт Excel № 9 «'+name+'»: ответов К/'+rep.answers.K+' · С/'+rep.answers.S+' · П/'+rep.answers.P+'; SF-статусов '+rep.sf+'; AHP-пар '+rep.ahp+'; эконом. ячеек '+rep.econ);
    var m18=calc(r)||{};
    notify('ok','Импорт выполнен',Q18.levels[r.lvl].name+': '+rep.answers[r.lvl]+' ответов · I_А '+(m18.iA==null?'н/д':Math.round(m18.iA*10)/10+'%')+(rep.notes&&rep.notes.length?' · замечания см. в журнале':''));
   }catch(err){notify('err','Импорт не выполнен',(err&&err.message)||String(err));}}}]);}
'''

ins_at = '\nfunction evidStats(r){'
assert s.count(ins_at) == 1, 'якорь evidStats не найден/не уникален'
s = s.replace(ins_at, JS_MODULE + ins_at)

# ---------------- карточка в renderCheck ----------------
card_anchor = 'if(r.legacy31)h+=\''
assert s.count(card_anchor) == 1, 'якорь legacy31'
s = s.replace(card_anchor, 'h+=importCard18();\n ' + card_anchor, 1)

# ---------------- wiring ----------------
wire_anchor = "const bm=document.getElementById('q18Md');if(bm)bm.onclick=()=>exportMD18dl(cur());"
assert s.count(wire_anchor) == 1, 'якорь q18Md'
s = s.replace(wire_anchor, wire_anchor + r'''
  const fx=document.getElementById('impXlsx18');
  if(fx)fx.onchange=function(ev2){
   var fl=ev2.target&&ev2.target.files&&ev2.target.files[0];
   if(!fl)return;
   var fr=new FileReader();
   fr.onload=function(){
    try{
     var wbk=xlsx18.parse(fr.result);
     importShow18(wbk,cur(),fl.name);
    }catch(err){notify('err','Импорт не выполнен',(err&&err.message)||String(err));}};
   fr.onerror=function(){notify('err','Импорт не выполнен','файл не читается');};
   fr.readAsArrayBuffer(fl);
   ev2.target.value='';};''', 1)

# ---------------- niokrAPI ----------------
api_anchor = 'get Q18(){return Q18;},'
assert s.count(api_anchor) == 1, 'якорь niokrAPI Q18'
s = s.replace(api_anchor, api_anchor + 'xlsx18,importExcel9,importPreview18,importShow18,deriveConsensus18,', 1)

io.open(MODEL, 'w', encoding='utf-8').write(s)
print('patch18_import: OK,', len(s), 'байт (симв.), модель записана')

# проверка синтаксиса js-блока
import subprocess
m = re.search(r'<script>([\s\S]*?)</script>', s)
io.open('/tmp/model18_check.js', 'w', encoding='utf-8').write(m.group(1))
r2 = subprocess.run(['node', '--check', '/tmp/model18_check.js'], capture_output=True, text=True)
print('node --check:', 'OK' if r2.returncode == 0 else r2.stderr[:800])

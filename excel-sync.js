/* ================================================================
   EXCEL SYNC — the Main Log, out and back in
   ----------------------------------------------------------------
   Drop this beside the tracker page and add one line before </body>:

       <script src="excel-sync.js"></script>

   Nothing else changes. Two buttons appear under "More": one writes
   the whole log as a workbook with the same seventy-seven columns in
   the same order, the other reads that workbook back and shows what
   would change before anything is touched.

   No library is fetched. A workbook is a zip of XML, and the browser
   already knows how to inflate a zip stream, so the reader is the
   browser's own DecompressionStream and the writer stores its entries
   uncompressed. A site network that blocks every CDN cannot stop it.

   The row is the record. Every one of the seventy-seven cells is kept
   on the material exactly as it was read, so a column this page has no
   opinion about — mock-ups, storage, the installer's name — survives
   the trip untouched and comes back out where it went in. The columns
   the page does understand are written back from the page, so an edit
   made here shows up in the sheet.
   ================================================================ */
(function(){
'use strict';

/* ---------------------------------------------------------------
   1. THE COLUMNS
   In the order the sheet has them. This list is the contract: it
   decides what is read, what is written, and where each cell lands.
   --------------------------------------------------------------- */
var COLS=[
 'Item Description','Material Category','Common Package (Yes/No)','Discipline',
 'Finishing (Internal/External)','Package (Lump Sum / Provisional Sum / Prime Cost)',
 'Sub-contractor Name','Supply & Install / Supply / Install Only','Manufacturer',
 'Country of Origin of Manufacture',
 'PQD Number','PQD Revision','PQD Status','PQD Submittal Date',
 'MAT Number','MAT Submittal Date','MAT Revision','MAT Status',
 'PA Tentative Date','PA Document Number','PA Date',
 '3rd Party Assessment Done','Client Assessment Done','PMC/LDC Assessment Done',
 'Contractor Assessment Done','Assessment Result',
 'Sample','Mock-up Delivered Date','Mock-up First in Place','Mock-up Approved by PMC',
 'Mock-up per Client DLA',
 'Purchase Order Issued (Yes/No)','PO Number','PO Date',
 'Method Statement Number','MES Incl. ITP (Yes/No)','MES Revision','MES Status',
 'ITP Number','ITP Submittal Date','ITP Revision','ITP Status',
 'PID Number','PID Submittal Date','PID Revision','PID Status',
 'Pre-Fabrication Meeting Date','3rd Party Assigned (Yes/No)','3rd Party Service Provider Name',
 'Design Verification/Calculation Status',
 'Fabrication Planned Date','Fabrication 1st Batch Started Date',
 'Fabrication Percentage Completion (%)',
 'FAT Package/Procedure Number/ITP','FAT Package Status','FAT Planned Date',
 'FAT Location / City','FAT/TPI Results',
 '1st Batch Delivery To Site Planned Date','1st Batch Delivery To Site Actual Date',
 'Storage (Site/Offsite)',
 'MIR Number','MIR Approval Date','MIR Status',
 'Installation Planned Date','Installation Actual Date','Installer Name',
 'WIR Number','WIR Approval Date','WIR Status',
 'Total Quantity','Total Quantity Unit','Delivered No','Delivered Unit',
 'Remaining','Remaining Unit','Delivered %'
];
var IDX={};COLS.forEach(function(c,i){IDX[c]=i;});

/* A cell that carries a date is written back as a real date rather
   than as text, so the sheet still sorts and filters. One that carries
   several dates at once — and a few of them do — stays as it was. */
var DATE_COLS={};
COLS.forEach(function(c){if(/date$/i.test(c.trim()))DATE_COLS[c]=1;});

/* The columns this page owns. On the way out these are written from
   the record, not from what came in, so work done here reaches the
   sheet. Everything else is passed through as it arrived. */
var OWNED=['Item Description','Material Category','Discipline','Sub-contractor Name',
 'Manufacturer','Country of Origin of Manufacture',
 'PQD Number','PQD Status','PQD Submittal Date',
 'MAT Number','MAT Submittal Date','MAT Status',
 'ITP Number','ITP Submittal Date','ITP Status',
 'PID Number','PID Submittal Date','PID Status',
 'Pre-Fabrication Meeting Date','FAT Package/Procedure Number/ITP','FAT Package Status',
 'FAT Planned Date','FAT/TPI Results','3rd Party Service Provider Name',
 'MIR Number','MIR Approval Date','MIR Status',
 'Total Quantity','Total Quantity Unit','Delivered No','Delivered Unit',
 'Remaining','Remaining Unit','Delivered %'];

/* ---------------------------------------------------------------
   2. SMALL THINGS
   --------------------------------------------------------------- */
function pad2(n){return ('0'+n).slice(-2);}
function esc(s){return String(s==null?'':s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;');}
/* the host page has one of these too, but a module that borrows a
   global it did not define breaks quietly the day the global moves */
function attr(s){return String(s==null?'':s).replace(/"/g,'&quot;');}
function xml(s){return String(s==null?'':s)
  .replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')
  .replace(/"/g,'&quot;').replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g,'');}
function trim(v){return String(v==null?'':v).replace(/\u00a0/g,' ').trim();}

/* Excel counts days from the last day of 1899, and believes 1900 had
   a 29th of February. Both quirks live in this one line. */
function serialToIso(n){
  n=Math.round(Number(n));
  if(!isFinite(n)||n<1||n>80000)return '';
  var d=new Date(Date.UTC(1899,11,30)+n*864e5);
  return d.getUTCFullYear()+'-'+pad2(d.getUTCMonth()+1)+'-'+pad2(d.getUTCDate());
}
function isoToSerial(iso){
  var m=/^(\d{4})-(\d{2})-(\d{2})$/.exec(iso);if(!m)return null;
  return Math.round((Date.UTC(+m[1],+m[2]-1,+m[3])-Date.UTC(1899,11,30))/864e5);
}
var MON={jan:1,feb:2,mar:3,apr:4,may:5,jun:6,jul:7,aug:8,sep:9,oct:10,nov:11,dec:12};
/* The sheet has 19-Ict-24 and 09-Ju-25 in it. A month spelt wrong is
   still a month, so the first three letters decide, and a name that
   matches nothing at all is left alone rather than guessed at. */
function textToIso(t){
  t=trim(t);if(!t)return '';
  var m;
  if((m=/^(\d{4})-(\d{1,2})-(\d{1,2})/.exec(t)))return mk(+m[1],+m[2],+m[3]);
  if((m=/^(\d{1,2})[-\/ ]([A-Za-z]{2,9})[-\/ ](\d{2,4})$/.exec(t))){
    var nm=m[2].toLowerCase().slice(0,3), mo=MON[nm];
    if(!mo)mo=near(nm);
    if(!mo)return '';
    var y=m[3].length===2?2000+ +m[3]:+m[3];
    return mk(y,mo,+m[1]);
  }
  if((m=/^(\d{1,2})[-\/.](\d{1,2})[-\/.](\d{2,4})$/.exec(t))){
    var y2=m[3].length===2?2000+ +m[3]:+m[3];
    return mk(y2,+m[2],+m[1]);
  }
  return '';
  function mk(y,mo,d){
    if(!(y>=1990&&y<=2100&&mo>=1&&mo<=12&&d>=1&&d<=31))return '';
    return y+'-'+pad2(mo)+'-'+pad2(d);
  }
  /* One letter wrong is a typo — Ict is October and the sheet has it.
     Two letters wrong is a different word. And "Ju" sits one letter
     from both June and July, so it is nobody's month: an ambiguous
     stub is reported rather than decided, because picking the wrong
     one of the two moves a deadline by thirty days in silence. */
  function near(k){
    var best=0,bd=2,tie=false;
    Object.keys(MON).forEach(function(n){
      var d=dist(k,n);
      if(d<bd){bd=d;best=MON[n];tie=false;}
      else if(d===bd&&MON[n]!==best)tie=true;
    });
    return tie?0:best;
  }
}
function dist(a,b){
  var m=a.length,n=b.length,prev=[],cur=[],i,j;
  for(j=0;j<=n;j++)prev[j]=j;
  for(i=1;i<=m;i++){cur[0]=i;
    for(j=1;j<=n;j++)cur[j]=Math.min(prev[j]+1,cur[j-1]+1,prev[j-1]+(a[i-1]===b[j-1]?0:1));
    prev=cur.slice();}
  return prev[n];
}
function anyDate(v){
  if(v==null||v==='')return '';
  if(typeof v==='number')return serialToIso(v);
  var t=trim(v);
  if(/^\d+(\.\d+)?$/.test(t)&&+t>20000&&+t<80000)return serialToIso(+t);
  return textToIso(t);
}
function showDate(iso){
  var m=/^(\d{4})-(\d{2})-(\d{2})$/.exec(iso||'');
  return m?(m[3]+'/'+m[2]+'/'+m[1]):'';
}

/* The log spells approval six ways. A word within one letter of the
   right one is the right one; a cell holding several outcomes takes
   the strongest, which is what the sheet's own note already says. */
function normStatus(v){
  var raw=trim(v);
  if(!raw||raw==='-')return '';
  var parts=raw.split(/[\n\r]+/).map(trim).filter(Boolean);
  var best='',rank={'Rejected':1,'Resubmit':2,'Pending':3,'Approved with comments':4,'Approved':5};
  parts.forEach(function(p){
    var one=oneStatus(p);
    if(one&&(!best||(rank[one]||0)>(rank[best]||0)))best=one;
  });
  return best;
}
function oneStatus(p){
  var k=p.toLowerCase().replace(/[^a-z& ]/g,' ').replace(/\s+/g,' ').trim();
  if(!k)return '';
  if(/as noted|with comment|conditional/.test(k))return 'Approved with comments';
  if(/revise|resubmit/.test(k))return 'Resubmit';
  if(/reject|fail/.test(k))return 'Rejected';
  if(/under review|^ur$|^u r$|in review/.test(k))return 'Pending';
  if(/^avl$|^alv$|approved vendor/.test(k))return 'Approved';
  var w=k.split(' ')[0];
  if(dist(w,'approved')<=2||dist(w,'approve')<=1)return 'Approved';
  return '';
}
function normCat(v){
  var t=trim(v).toUpperCase().replace(/CATEGORY/,'').replace(/[^C0-9]/g,'');
  var m=/C([0-3])/.exec(t);
  return m?('C'+m[1]):'';
}
function yesNo(v){
  var k=trim(v).toLowerCase();
  if(!k||k==='-')return '';
  if(k==='n'||k==='mo'||k==='no')return 'No';
  if(k[0]==='y')return 'Yes';
  if(k[0]==='n')return 'No';
  return trim(v);
}
function K(x){return String(x||'').toLowerCase().replace(/\s+/g,' ').replace(/[.,]/g,'').trim();}

/* ---------------------------------------------------------------
   3. READING A WORKBOOK
   A zip read backwards from its end directory, then four XML files.
   --------------------------------------------------------------- */
function dv(buf){return new DataView(buf);}
async function inflate(bytes){
  if(typeof DecompressionStream==='undefined')
    throw new Error('This browser cannot open a zip on its own. Chrome, Edge, or Safari 16.4 and later can.');
  var s=new Blob([bytes]).stream().pipeThrough(new DecompressionStream('deflate-raw'));
  return new Uint8Array(await new Response(s).arrayBuffer());
}
async function unzip(buf){
  var d=dv(buf),n=buf.byteLength,end=-1;
  for(var i=n-22;i>=0&&i>n-66000;i--){if(d.getUint32(i,true)===0x06054b50){end=i;break;}}
  if(end<0)throw new Error('That file is not a workbook — no zip directory in it.');
  var count=d.getUint16(end+10,true), off=d.getUint32(end+16,true), out={}, p=off;
  for(var k=0;k<count;k++){
    if(d.getUint32(p,true)!==0x02014b50)break;
    var method=d.getUint16(p+10,true), csize=d.getUint32(p+20,true),
        nlen=d.getUint16(p+28,true), elen=d.getUint16(p+30,true), clen=d.getUint16(p+32,true),
        lho=d.getUint32(p+42,true);
    var name=new TextDecoder().decode(new Uint8Array(buf,p+46,nlen));
    var lnlen=d.getUint16(lho+26,true), lelen=d.getUint16(lho+28,true);
    var start=lho+30+lnlen+lelen;
    out[name]={method:method,bytes:new Uint8Array(buf,start,csize)};
    p+=46+nlen+elen+clen;
  }
  return out;
}
async function textOf(entry){
  if(!entry)return '';
  var b=entry.method===0?entry.bytes:await inflate(entry.bytes);
  return new TextDecoder('utf-8').decode(b);
}
function parse(t){
  var doc=new DOMParser().parseFromString(t,'application/xml');
  var bad=doc.getElementsByTagName('parsererror');
  if(bad&&bad.length)throw new Error('A part of the workbook could not be read.');
  return doc;
}
function colNum(ref){
  var m=/^([A-Z]+)/.exec(ref||'');if(!m)return 0;
  var n=0;for(var i=0;i<m[1].length;i++)n=n*26+(m[1].charCodeAt(i)-64);
  return n-1;
}
/* A workbook opened once, its sheets readable by name. Both readers in
   this file need the same twenty lines of zip and XML, and the register
   export is a different shape from the log — so the opening is separated
   from the understanding. */
async function openBook(file){
  var zip=await unzip(await file.arrayBuffer());
  var wbDoc=parse(await textOf(zip['xl/workbook.xml']));
  var relDoc=parse(await textOf(zip['xl/_rels/workbook.xml.rels']));
  var rels={};
  Array.prototype.forEach.call(relDoc.getElementsByTagName('Relationship'),function(r){
    rels[r.getAttribute('Id')]=r.getAttribute('Target').replace(/^\/?xl\//,'').replace(/^\//,'');
  });
  var sheets=[];
  Array.prototype.forEach.call(wbDoc.getElementsByTagName('sheet'),function(s){
    var id=s.getAttribute('r:id')||s.getAttributeNS(
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships','id');
    sheets.push({name:s.getAttribute('name'),path:'xl/'+(rels[id]||'')});
  });

  /* the strings live in one shared table, referenced by number */
  var shared=[];
  var ssTxt=await textOf(zip['xl/sharedStrings.xml']);
  if(ssTxt){
    var ss=parse(ssTxt);
    Array.prototype.forEach.call(ss.getElementsByTagName('si'),function(si){
      var t='';
      Array.prototype.forEach.call(si.getElementsByTagName('t'),function(n){t+=n.textContent;});
      shared.push(t);
    });
  }

  async function rowsOf(sh){
    var doc=parse(await textOf(zip[sh.path])), out=[];
    Array.prototype.forEach.call(doc.getElementsByTagName('row'),function(r){
      var line=[];
      Array.prototype.forEach.call(r.getElementsByTagName('c'),function(c){
        var i=colNum(c.getAttribute('r')), t=c.getAttribute('t'), v='';
        if(t==='inlineStr'){
          var is=c.getElementsByTagName('t');
          for(var q=0;q<is.length;q++)v+=is[q].textContent;
        }else{
          var vn=c.getElementsByTagName('v')[0];
          v=vn?vn.textContent:'';
          if(t==='s')v=shared[+v]||'';
          else if(v!==''&&/^-?\d+(\.\d+)?([eE][-+]?\d+)?$/.test(v))v=Number(v);
        }
        line[i]=v;
      });
      out.push(line);
    });
    return out;
  }
  return {sheets:sheets,names:sheets.map(function(s){return s.name;}),rows:rowsOf};
}

async function readMainLog(file){
  var book=await openBook(file);
  var want=book.sheets.filter(function(s){return K(s.name)==='main log';})[0];
  if(want)return {rows:await book.rows(want),name:want.name,sheets:book.names};
  /* no sheet by that name: take the widest one, and say so */
  var best=null,bw=0;
  for(var i=0;i<book.sheets.length;i++){
    var r=await book.rows(book.sheets[i]);
    var w=r.reduce(function(a,x){return Math.max(a,x.length);},0);
    if(w>bw){bw=w;best={rows:r,name:book.sheets[i].name};}
  }
  if(!best||bw<20)throw new Error('No sheet in that file looks like the Main Log.');
  best.sheets=book.names;best.guessed=true;
  return best;
}

/* ---------------------------------------------------------------
   4. WRITING A WORKBOOK
   Stored, not deflated. A zip entry may legally be uncompressed, and
   a hundred rows of text is a file small enough that nobody minds.
   --------------------------------------------------------------- */
var CRC=(function(){
  var t=new Uint32Array(256);
  for(var n=0;n<256;n++){var c=n;
    for(var k=0;k<8;k++)c=(c&1)?(0xEDB88320^(c>>>1)):(c>>>1);
    t[n]=c>>>0;}
  return t;
})();
function crc32(b){
  var c=0xFFFFFFFF;
  for(var i=0;i<b.length;i++)c=CRC[(c^b[i])&0xFF]^(c>>>8);
  return (c^0xFFFFFFFF)>>>0;
}
function zipUp(files){
  var enc=new TextEncoder(), parts=[], central=[], offset=0;
  var now=new Date();
  var time=((now.getHours()<<11)|(now.getMinutes()<<5)|(now.getSeconds()>>1))&0xFFFF;
  var date=(((now.getFullYear()-1980)<<9)|((now.getMonth()+1)<<5)|now.getDate())&0xFFFF;
  files.forEach(function(f){
    var name=enc.encode(f.name), body=enc.encode(f.text), c=crc32(body);
    var lh=new Uint8Array(30+name.length), v=new DataView(lh.buffer);
    v.setUint32(0,0x04034b50,true);v.setUint16(4,20,true);v.setUint16(6,0x0800,true);
    v.setUint16(8,0,true);v.setUint16(10,time,true);v.setUint16(12,date,true);
    v.setUint32(14,c,true);v.setUint32(18,body.length,true);v.setUint32(22,body.length,true);
    v.setUint16(26,name.length,true);v.setUint16(28,0,true);
    lh.set(name,30);
    parts.push(lh,body);
    var ch=new Uint8Array(46+name.length), w=new DataView(ch.buffer);
    w.setUint32(0,0x02014b50,true);w.setUint16(4,20,true);w.setUint16(6,20,true);
    w.setUint16(8,0x0800,true);w.setUint16(10,0,true);w.setUint16(12,time,true);
    w.setUint16(14,date,true);w.setUint32(16,c,true);
    w.setUint32(20,body.length,true);w.setUint32(24,body.length,true);
    w.setUint16(28,name.length,true);w.setUint32(42,offset,true);
    ch.set(name,46);
    central.push(ch);
    offset+=lh.length+body.length;
  });
  var cdSize=central.reduce(function(a,c){return a+c.length;},0);
  var eocd=new Uint8Array(22), e=new DataView(eocd.buffer);
  e.setUint32(0,0x06054b50,true);
  e.setUint16(8,files.length,true);e.setUint16(10,files.length,true);
  e.setUint32(12,cdSize,true);e.setUint32(16,offset,true);
  return new Blob(parts.concat(central,[eocd]),
    {type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'});
}
function colName(i){
  var s='';i++;
  while(i>0){var r=(i-1)%26;s=String.fromCharCode(65+r)+s;i=(i-r-1)/26;}
  return s;
}
/* a cell knows what it is: a date, a number, or words */
function cellXml(ref,val,head){
  if(val==null||val==='')return '';
  if(head)return '<c r="'+ref+'" t="inlineStr" s="2"><is><t xml:space="preserve">'+xml(val)+'</t></is></c>';
  if(typeof val==='object'&&val.date!=null){
    var s=isoToSerial(val.date);
    if(s!=null)return '<c r="'+ref+'" s="1"><v>'+s+'</v></c>';
    val=showDate(val.date)||'';
    if(!val)return '';
  }
  if(typeof val==='number'&&isFinite(val))return '<c r="'+ref+'"><v>'+val+'</v></c>';
  var t=String(val);
  if(/^-?\d+(\.\d+)?$/.test(t)&&t.length<15&&!/^0\d/.test(t))
    return '<c r="'+ref+'"><v>'+t+'</v></c>';
  return '<c r="'+ref+'" t="inlineStr"><is><t xml:space="preserve">'+xml(t)+'</t></is></c>';
}
function sheetXml(rows,widths){
  var body=rows.map(function(row,r){
    var cells=row.map(function(v,c){
      return cellXml(colName(c)+(r+1),v,r===0||(row.head&&c===0));
    }).join('');
    return cells?('<row r="'+(r+1)+'">'+cells+'</row>'):'';
  }).join('');
  var cols=widths?('<cols>'+widths.map(function(w,i){
    return '<col min="'+(i+1)+'" max="'+(i+1)+'" width="'+w+'" customWidth="1"/>';}).join('')+'</cols>'):'';
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
   +'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
   +'<sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>'
   +cols+'<sheetData>'+body+'</sheetData></worksheet>';
}
function workbook(sheets){
  var files=[
    {name:'[Content_Types].xml',text:'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      +'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      +'<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      +'<Default Extension="xml" ContentType="application/xml"/>'
      +'<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
      +sheets.map(function(_,i){return '<Override PartName="/xl/worksheets/sheet'+(i+1)+'.xml" '
        +'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>';}).join('')
      +'<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
      +'</Types>'},
    {name:'_rels/.rels',text:'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      +'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      +'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
      +'</Relationships>'},
    {name:'xl/workbook.xml',text:'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      +'<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
      +'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>'
      +sheets.map(function(s,i){return '<sheet name="'+xml(s.name)+'" sheetId="'+(i+1)+'" r:id="rId'+(i+1)+'"/>';}).join('')
      +'</sheets></workbook>'},
    {name:'xl/_rels/workbook.xml.rels',text:'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      +'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      +sheets.map(function(_,i){return '<Relationship Id="rId'+(i+1)+'" '
        +'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" '
        +'Target="worksheets/sheet'+(i+1)+'.xml"/>';}).join('')
      +'<Relationship Id="rId'+(sheets.length+1)+'" '
      +'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
      +'</Relationships>'},
    {name:'xl/styles.xml',text:'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      +'<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
      +'<numFmts count="1"><numFmt numFmtId="164" formatCode="dd/mm/yyyy"/></numFmts>'
      +'<fonts count="2"><font><sz val="11"/><name val="Calibri"/></font>'
      +'<font><b/><sz val="11"/><name val="Calibri"/></font></fonts>'
      +'<fills count="2"><fill><patternFill patternType="none"/></fill>'
      +'<fill><patternFill patternType="gray125"/></fill></fills>'
      +'<borders count="1"><border/></borders>'
      +'<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
      +'<cellXfs count="3">'
      +'<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
      +'<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
      +'<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>'
      +'</cellXfs>'
      +'<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
      +'</styleSheet>'}
  ];
  sheets.forEach(function(s,i){
    files.push({name:'xl/worksheets/sheet'+(i+1)+'.xml',text:sheetXml(s.rows,s.widths)});
  });
  return zipUp(files);
}

/* ---------------------------------------------------------------
   5. A ROW BECOMES A RECORD
   --------------------------------------------------------------- */
function rowToRaw(line){
  var raw={};
  COLS.forEach(function(c,i){
    var v=line[i];
    if(v==null||v==='')return;
    if(DATE_COLS[c]){
      var iso=anyDate(v);
      raw[c]=iso||trim(v);            /* unreadable dates keep their words */
      return;
    }
    raw[c]=(typeof v==='number')?v:trim(v);
  });
  return raw;
}
/* the identity of a row. The MAT number is the one thing that is
   filled on every line and never repeats. */
/* A row is known by its reference, and only by its description when it
   has none. The description is the thing a person edits — merging two
   entries, fixing a spelling — and a key that changes when someone
   tidies the file is not a key. */
function idOf(raw){
  var n=trim(raw['MAT Number']);
  if(n)return 'mat:'+K(n);
  var alt='';
  ['ITP Number','Method Statement Number','PID Number','MIR Number','WIR Number']
    .some(function(c){var v=splitRefs(raw[c])[0];if(v){alt=v;return true;}return false;});
  return alt?('ref:'+K(alt)):('name:'+K(raw['Item Description']));
}

/* Identifiers have to be unique or the database rejects the whole batch,
   and the page's own uid() is the millisecond plus a random number —
   which collides the moment a hundred records are made inside the same
   millisecond. Two thousand of them collide as good as certainly. So a
   run gets a counter of its own that knows what is already taken and
   never repeats. */
function idMaker(){
  var used={};
  ['mats','mfrs','people','events'].forEach(function(k){
    (DB[k]||[]).forEach(function(o){used[String(o.id)]=1;});
  });
  var n=Date.now()*1000;
  return function(){
    while(used[String(n)])n++;
    used[String(n)]=1;
    return n++;
  };
}

/* the vendor a row points at, made once and shared afterwards */
/* The vendor is whoever made the material. The subcontractor column
   holds the trade doing the installation — FIRST FIX against half the
   log — and folding the two together would hang fifty materials off
   one company and let its pre-qualification, which the sheet records
   per row, overwrite forty rows that never named it. So a row with no
   manufacturer gets no vendor, which is the truth and is exactly the
   gap the board is built to show. */
/* The sub-contractor column names the company that brought the maker
   onto the project — FIRST FIX brought twenty-two of them in this file.
   It is a company in its own right, so it gets a record of its own and
   the maker remembers who brought it. */
function subOf(name,made){
  var n=trim(name);
  if(!n||n==='-')return '';
  var have=(DB.mfrs||[]).filter(function(v){return K(v.name)===K(n);})[0];
  if(!have){
    have={id:made.id(),name:n,kind:'sub',cat:'',country:'',site:'',scope:'',
      steps:{},pq:{},added:today()};
    DB.mfrs.push(have);made.vendors.push(n);
  }
  return have.name;
}

function vendorFor(raw,made){
  var subName=trim(raw['Sub-contractor Name']);
  var name=trim(raw['Manufacturer']);
  if(!name||name==='-'){
    if(subName)subOf(subName,made);      /* the sub is worth keeping even alone */
    return null;
  }
  var found=(DB.mfrs||[]).filter(function(v){return K(v.name)===K(name);})[0];
  if(!found){
    found={id:made.id(),name:name,kind:'maker',
      cat:normCat(raw['Material Category']),country:trim(raw['Country of Origin of Manufacture'])||'',
      site:'',scope:trim(raw['Discipline'])||'',steps:{},pq:{},added:today()};
    DB.mfrs.push(found);made.vendors.push(found.name);
  }
  if(subName&&!found.by&&K(subName)!==K(found.name))found.by=subOf(subName,made);
  /* the pre-qualification lives on the vendor, so it is written there */
  var pq=trim(raw['PQD Number']), st=normStatus(raw['PQD Status']);
  if(pq||st){
    found.steps=found.steps||{};
    var have=found.steps.pqd||{};
    if(!have.ref&&!have.status)
      found.steps.pqd={ref:pq&&pq!=='-'?pq:'',date:anyDate(raw['PQD Submittal Date'])||'',
        status:st||'Pending'};
  }
  if(!found.country&&trim(raw['Country of Origin of Manufacture']))
    found.country=trim(raw['Country of Origin of Manufacture']);
  return found;
}

/* what the page understands, drawn out of the row */
function applyRaw(m,raw,made){
  m.raw=raw;
  m.name=trim(raw['Item Description'])||m.name||'Untitled';
  m.cat=normCat(raw['Material Category'])||m.cat||'';
  m.catFrom=m.cat?'from the Main Log':'';
  m.catSure=true;
  m.disc=trim(raw['Discipline'])||'';
  m.sub=trim(raw['Sub-contractor Name'])||'';
  m.ref=trim(raw['MAT Number'])||'';
  var v=vendorFor(raw,made);
  if(v)m.mfr=String(v.id);
  m.steps=m.steps||{};
  put('mts','MAT Number','MAT Submittal Date','MAT Status');
  put('itp','ITP Number','ITP Submittal Date','ITP Status');
  put('pid','PID Number','PID Submittal Date','PID Status');
  var pfm=anyDate(raw['Pre-Fabrication Meeting Date']);
  if(pfm)m.steps.pfm={date:pfm,ref:'',status:'Issued'};
  var fatRef=trim(raw['FAT/TPI Results'])||trim(raw['FAT Package/Procedure Number/ITP']);
  var fatSt=normStatus(raw['FAT Package Status']);
  var fatDt=anyDate(raw['FAT Planned Date']);
  if(fatRef||fatSt||fatDt)m.steps.fat={ref:fatRef&&fatRef!=='-'?fatRef:'',date:fatDt||'',
    by:trim(raw['3rd Party Service Provider Name'])||'',
    status:fatSt==='Approved'?'Passed':fatSt==='Approved with comments'?'Passed with comments':(fatSt||'Pending')};
  /* the quantities, and each delivery the sheet has recorded */
  if(raw['Total Quantity']!=null&&raw['Total Quantity']!==''){
    m.qty=String(raw['Total Quantity']);
    m.unit=trim(raw['Total Quantity Unit'])||'';
  }
  var mir=trim(raw['MIR Number']);
  if(mir&&mir!=='-'){
    m.dels=m.dels||[];
    var seen=m.dels.filter(function(d){return K(d.ref)===K(mir);})[0];
    var rec={id:seen?seen.id:(made.id()),qty:raw['Delivered No']!=null?String(raw['Delivered No']):'',
      date:anyDate(raw['MIR Approval Date'])||today(),ref:mir,
      status:normStatus(raw['MIR Status'])==='Approved'?'Received':(normStatus(raw['MIR Status'])||'Pending'),
      note:'from the Main Log'};
    if(seen)Object.assign(seen,rec);else m.dels.push(rec);
  }
  function put(k,nCol,dCol,sCol){
    var ref=trim(raw[nCol]), st=normStatus(raw[sCol]), dt=anyDate(raw[dCol]);
    if(!ref&&!st&&!dt)return;
    m.steps[k]={ref:(ref&&ref!=='-')?ref:'',date:dt||'',status:st||'Pending'};
  }
}

/* ---------------------------------------------------------------
   6. THE PLAN
   Nothing is written until it has been shown. A row is new, changed,
   the same, or missing from the file — and the fourth is the one that
   deserves a decision rather than an assumption.
   --------------------------------------------------------------- */
function planFrom(rows){
  var head=rows[0]||[];
  var map={},shift=0;
  /* the header must be the one we know, or the columns land wrong */
  var seen=head.map(trim);
  var miss=COLS.filter(function(c,i){return K(seen[i]||'')!==K(c);});
  var out={add:[],change:[],same:[],gone:[],skipped:0,dividers:0,header:miss.length};
  var byId={},byRef={};
  (DB.mats||[]).forEach(function(m){
    if(m.raw){
      byId[idOf(m.raw)]=m;
      REF_FIELDS.forEach(function(c){
        splitRefs(m.raw[c]).forEach(function(r){if(!byRef[K(r)])byRef[K(r)]=m;});
      });
    }
  });
  var hit={};

  rows.slice(1).forEach(function(line,n){
    var desc=trim(line[0]);
    if(!desc)return;
    /* A section divider carries a heading and nothing else. Reading it
       off the category column alone was wrong the moment materials
       without a category existed — and the register brings in more than
       a hundred of those, every one of which was being skipped. */
    var lone=true;
    for(var q=1;q<COLS.length&&lone;q++)if(trim(line[q])!=='')lone=false;
    if(lone){out.dividers++;return;}
    var raw=rowToRaw(line), id=idOf(raw);
    var have=byId[id]||byRef[K(trim(raw['MAT Number']))]
      ||(DB.mats||[]).filter(function(m){
        return !m.raw&&K(m.name)===K(desc);})[0];
    if(!have){out.add.push({raw:raw,row:n+2});return;}
    hit[id]=1;
    var diff=COLS.filter(function(c){
      return String((have.raw||{})[c]==null?'':(have.raw||{})[c])!==String(raw[c]==null?'':raw[c]);
    });
    if(diff.length)out.change.push({raw:raw,row:n+2,rec:have,diff:diff});
    else out.same.push({rec:have});
  });
  (DB.mats||[]).forEach(function(m){
    /* A document has no row in a log shaped one-per-material, so its
       absence from the file means nothing. Counting it as missing would
       put sixteen hundred records under a heading that invites deleting
       them. */
    if(isDoc(m))return;
    if(m.raw&&!hit[idOf(m.raw)])out.gone.push(m);
  });

  /* Two things in the file are worth saying out loud before anything
     is written: a date nobody can read, and one manufacturer holding
     two different pre-qualifications. Neither stops the import — the
     cell is kept exactly as typed — but both are the kind of thing
     that is invisible in a sheet of seventy-seven columns. */
  var seenRows=out.add.concat(out.change);
  var byVendor={};
  out.badDates=[];out.clash=[];
  seenRows.forEach(function(r){
    var desc=trim(r.raw['Item Description']);
    COLS.forEach(function(c){
      if(!DATE_COLS[c])return;
      var v=r.raw[c];
      if(!v)return;
      /* a dash, AVL, or N/A is a deliberate "does not apply", not a
         date somebody mistyped, so it is left in peace */
      if(/^(-+|n\/?a|avl|alv|tbd|tba|na)$/i.test(trim(v)))return;
      if(!/^\d{4}-\d{2}-\d{2}$/.test(String(v))&&out.badDates.length<40)
        out.badDates.push({row:r.row,col:c,val:String(v),what:desc,
          many:/[\n\r]/.test(String(v))||(String(v).match(/\d{1,2}[-\/ ][A-Za-z]{3}/g)||[]).length>1});
    });
    var man=trim(r.raw['Manufacturer']), pq=trim(r.raw['PQD Number']);
    if(man&&man!=='-'&&pq&&pq!=='-'){
      var b=byVendor[K(man)]=byVendor[K(man)]||{name:man,refs:{},shown:{}};
      b.shown[K(pq)]=pq;
      (b.refs[K(pq)]=b.refs[K(pq)]||[]).push(r.row);
    }
  });
  Object.keys(byVendor).forEach(function(k){
    var b=byVendor[k], refs=Object.keys(b.refs);
    if(refs.length>1)out.clash.push({name:b.name,
      refs:refs.map(function(r){return b.shown[r];}),
      rows:refs.map(function(r){return b.refs[r][0];})});
  });
  return out;
}
function applyPlan(p,dropGone){
  var made={n:1,vendors:[],id:idMaker()};
  p.add.forEach(function(a){
    var m={id:made.id(),name:'',cat:'',ref:'',mfr:'',qty:'',unit:'',
      steps:{},dels:[],ncrs:[],added:today()};
    applyRaw(m,a.raw,made);
    DB.mats.push(m);
  });
  p.change.forEach(function(c){applyRaw(c.rec,c.raw,made);});
  if(dropGone)DB.mats=DB.mats.filter(function(m){return p.gone.indexOf(m)<0;});
  touch();rList();rPane();
  return made;
}

/* ---------------------------------------------------------------
   7. THE RECORD BECOMES A ROW AGAIN
   --------------------------------------------------------------- */
function rawOut(m){
  var raw={};
  COLS.forEach(function(c){var v=(m.raw||{})[c];if(v!=null&&v!=='')raw[c]=v;});
  var v=m.mfr?mfr(m.mfr):null, s=m.steps||{};
  /* A record born in this page has no row behind it, so every column
     it can fill, it fills. A record that came from the sheet is a
     guest in someone else's document: the page writes back the columns
     it owns outright, and for the ones it merely infers — the vendor's
     qualification, which the sheet keeps per row and the page keeps
     once per vendor — it only refreshes cells that already had
     something in them. Filling ninety blanks with the same reference
     would look like ninety edits nobody made. */
  var fresh=!m.raw||!Object.keys(m.raw).length;
  function fill(c,val){
    if(!fresh&&(raw[c]==null||raw[c]===''))return;
    if(K(raw[c])===K(val))return;      /* a stray double space is not an edit */
    set(c,val);
  }
  set('Item Description',m.name);
  set('Material Category',m.cat?('Category '+m.cat):'');
  set('Discipline',m.disc||'');
  set('Sub-contractor Name',m.sub||'');
  if(v){
    if(v.kind==='sub'){if(!raw['Sub-contractor Name'])set('Sub-contractor Name',v.name);}
    else fill('Manufacturer',v.name);
    if(v.country)fill('Country of Origin of Manufacture',v.country);
    /* The sheet records a pre-qualification on every row; this page
       records it once on the vendor. Where the two rows of one vendor
       disagree — and in the log two do — the row is left alone rather
       than quietly made to agree with its neighbour. The import screen
       names the disagreement instead, which is a thing a person can
       settle and a program cannot. */
    var pq=(v.steps||{}).pqd||{};
    var agrees=!raw['PQD Number']||K(raw['PQD Number'])===K(pq.ref);
    if(fresh&&pq.ref)set('PQD Number',pq.ref);
    if(agrees){
      if(pq.status&&pq.status!=='Pending'&&normStatus(raw['PQD Status'])!==pq.status)fill('PQD Status',pq.status);
      if(pq.date)fill('PQD Submittal Date',pq.date);
    }
  }
  step('mts','MAT Number','MAT Submittal Date','MAT Status');
  step('itp','ITP Number','ITP Submittal Date','ITP Status');
  step('pid','PID Number','PID Submittal Date','PID Status');
  if(s.pfm&&s.pfm.date)set('Pre-Fabrication Meeting Date',s.pfm.date);
  if(s.fat){
    if(s.fat.ref)fill(raw['FAT/TPI Results']!=null&&raw['FAT/TPI Results']!==''
      ?'FAT/TPI Results':'FAT Package/Procedure Number/ITP',s.fat.ref);
    if(s.fat.date)set('FAT Planned Date',s.fat.date);
    if(s.fat.status&&s.fat.status!=='Pending')
      setStatus('FAT Package Status',/Passed/.test(s.fat.status)?'Approved':s.fat.status);
    if(s.fat.by)fill('3rd Party Service Provider Name',s.fat.by);
  }
  /* deliveries: the newest one fills the inspection-request columns,
     and the running total fills the quantities */
  var dels=(m.dels||[]).slice().sort(function(a,b){return String(b.date||'').localeCompare(String(a.date||''));});
  if(dels.length){
    set('MIR Number',dels[0].ref||'');
    set('MIR Approval Date',dels[0].date||'');
    setStatus('MIR Status',dels[0].status==='Received'?'Approved':(dels[0].status||''));
  }
  var ord=parseFloat(m.qty), got=(m.dels||[]).filter(function(d){
    return ['Received','Approved','Approved with comments'].indexOf(d.status)>=0;})
    .reduce(function(a,d){return a+(parseFloat(d.qty)||0);},0);
  if(m.qty){set('Total Quantity',m.qty);set('Total Quantity Unit',m.unit||'');}
  if(m.dels&&m.dels.length){
    set('Delivered No',got);set('Delivered Unit',m.unit||'');
    if(isFinite(ord)&&ord>0){
      set('Remaining',Math.max(0,ord-got));set('Remaining Unit',m.unit||'');
      set('Delivered %',Math.round(got/ord*100)/100);
    }
  }
  return raw;
  function set(c,val){if(val==null||val==='')return;raw[c]=val;}
  /* The log spells approval six ways and all six are correct enough.
     Rewriting "Approve" as "Approved" would mark ninety rows as changed
     on the next read for no gain, so a word that already means the
     right thing is left exactly as its author typed it. */
  function setStatus(c,val){
    if(!val)return;
    if(normStatus(raw[c])===val)return;
    raw[c]=val;
  }
  function step(k,nCol,dCol,sCol){
    var d=s[k];if(!d)return;
    if(d.ref)set(nCol,d.ref);
    if(d.status&&d.status!=='Pending')setStatus(sCol,d.status);
    if(d.date)set(dCol,d.date);
  }
}
function logRows(){
  var groups={},order=[];
  (DB.mats||[]).forEach(function(m){
    var g=(m.disc||'Uncategorised').trim();
    if(!groups[g]){groups[g]=[];order.push(g);}
    groups[g].push(m);
  });
  var rows=[COLS.slice()];
  order.forEach(function(g,gi){
    if(gi>0){var div=new Array(COLS.length);div[0]=g;div.head=true;rows.push(div);}
    groups[g].forEach(function(m){
      var raw=rawOut(m);
      rows.push(COLS.map(function(c){
        var v=raw[c];
        if(v==null||v==='')return '';
        if(DATE_COLS[c]&&/^\d{4}-\d{2}-\d{2}$/.test(String(v)))return {date:String(v)};
        return v;
      }));
    });
  });
  return rows;
}
/* the same figures the dashboard sheet carried, recomputed */
function summaryRows(){
  var mats=(DB.mats||[]);
  function n(f){return mats.filter(f).length;}
  function ok(m,k){var d=(m.steps||{})[k]||{};return /^Approved/.test(d.status||'');}
  var rows=[
    ['Material Submittals — Summary'],
    ['Written by the tracker on '+showDate(today())+'. Every figure below is counted from the Main Log sheet beside it.'],
    [],
    ['1. OVERVIEW'],
    ['Total items tracked',mats.length],
    ['Disciplines covered',Object.keys(mats.reduce(function(a,m){if(m.disc)a[m.disc]=1;return a;},{})).length],
    ['Vendors on record',(DB.mfrs||[]).length],
    ['Items with MAT approved',n(function(m){return ok(m,'mts');})],
    ['Purchase orders issued',n(function(m){return /^yes$/i.test(trim((m.raw||{})['Purchase Order Issued (Yes/No)']));})],
    ['Open non-conformances',mats.reduce(function(a,m){
      return a+((m.ncrs||[]).filter(function(x){return x.status!=='Closed';}).length);},0)],
    [],
    ['2. BY STAGE'],
    ['Stage','Approved','Pending or returned','Not recorded','Total','% approved']
  ];
  [['MAT (material approval)','mts'],['ITP (inspection and test plan)','itp'],
   ['PID (pre-inspection dossier)','pid'],['FAT (final inspection)','fat']].forEach(function(p){
    var a=n(function(m){var d=(m.steps||{})[p[1]]||{};return /^(Approved|Passed)/.test(d.status||'');});
    var b=n(function(m){var d=(m.steps||{})[p[1]]||{};return d.status&&!/^(Approved|Passed)/.test(d.status);});
    rows.push([p[0],a,b,mats.length-a-b,mats.length,mats.length?Math.round(a/mats.length*100)/100:0]);
  });
  rows.push([]);
  rows.push(['3. BY DISCIPLINE']);
  rows.push(['Discipline','Items','MAT approved','ITP approved','Vendors']);
  var byD={};
  mats.forEach(function(m){var g=m.disc||'Uncategorised';(byD[g]=byD[g]||[]).push(m);});
  Object.keys(byD).forEach(function(g){
    var list=byD[g], vs={};
    list.forEach(function(m){if(m.mfr)vs[m.mfr]=1;});
    rows.push([g,list.length,list.filter(function(m){return ok(m,'mts');}).length,
      list.filter(function(m){return ok(m,'itp');}).length,Object.keys(vs).length]);
  });
  rows.push([]);
  rows.push(['4. BY CATEGORY']);
  rows.push(['Category','Items','MAT approved','ITP approved']);
  ['C0','C1','C2','C3',''].forEach(function(c){
    var list=mats.filter(function(m){return (m.cat||'')===c;});
    if(!list.length&&c)return;
    rows.push([c||'No category set',list.length,
      list.filter(function(m){return ok(m,'mts');}).length,
      list.filter(function(m){return ok(m,'itp');}).length]);
  });
  return rows;
}
function download(blob,name){
  var u=URL.createObjectURL(blob), a=document.createElement('a');
  a.href=u;a.download=name;document.body.appendChild(a);a.click();document.body.removeChild(a);
  setTimeout(function(){URL.revokeObjectURL(u);},900);
}
function widths(){
  return COLS.map(function(c,i){return i===0?62:Math.min(30,Math.max(12,c.length*0.95));});
}
window.excelOut=function(){
  try{
    var name=(DB.project||'Project Materials').replace(/[^\w \-]/g,'').trim();
    download(workbook([
      {name:'Main Log',rows:logRows(),widths:widths()},
      {name:'Summary',rows:summaryRows()}
    ]),name+' — Live Tracking '+today()+'.xlsx');
    toast('Workbook written — '+(DB.mats||[]).length+' items');
  }catch(e){toast('Could not write the workbook — '+(e.message||e));}
};

/* ---------------------------------------------------------------
   8. THE SCREENS
   --------------------------------------------------------------- */
var PLAN=null;
window.excelPick=function(){document.getElementById('xl-file').click();};
window.excelRead=async function(ev){
  var f=ev.target.files[0];ev.target.value='';
  if(!f)return;
  if(typeof busy==='function')busy(true,'Reading the workbook');
  try{
    var got=await readMainLog(f);
    PLAN=planFrom(got.rows);
    PLAN.file=f.name;PLAN.sheet=got.name;PLAN.guessed=!!got.guessed;
    if(typeof busy==='function')busy(false);
    showPlan();
  }catch(e){
    if(typeof busy==='function')busy(false);
    sheet('That file could not be read','<div style="font-size:14px;line-height:1.75">'
      +esc(e.message||String(e))
      +'<br><br>The tracker expects the workbook it wrote, or the original Main Log sheet with its '
      +COLS.length+' columns in their original order.</div>');
  }
};
function showPlan(){
  var p=PLAN;
  var lines=[];
  function block(title,items,body){
    if(!items.length)return '';
    return '<div class="sec">'+esc(title)+'</div><div class="panel"><div class="panel-b">'+body+'</div></div>';
  }
  var warn='';
  if(p.header)warn+='<div class="panel" style="border-color:var(--now);background:var(--now-b);margin-bottom:12px">'
    +'<div class="panel-b"><b>'+p.header+' of the '+COLS.length+' column headings do not match.</b> '
    +'Cells are read by position, so a moved or renamed column lands in the wrong field. '
    +'Check the heading row before applying.</div></div>';
  if(p.guessed)warn+='<div class="panel" style="border-color:var(--now);background:var(--now-b);margin-bottom:12px">'
    +'<div class="panel-b">No sheet is called <b>Main Log</b>, so the widest one — <b>'+esc(p.sheet)+'</b> — was read instead.</div></div>';

  var body=warn
   +'<div class="grid" style="margin-bottom:18px">'
   +stat(p.add.length,'New')+stat(p.change.length,'Changed')
   +stat(p.same.length,'Unchanged')+stat(p.gone.length,'Not in the file')
   +'</div>';

  body+=block('New — these will be added',p.add,
    p.add.slice(0,40).map(function(a){
      return '<div class="line"><span class="tag t-ok">row '+a.row+'</span>'
        +'<div class="line-m"><div>'+esc(trim(a.raw['Item Description']))+'</div>'
        +'<div class="dim" style="font-size:12.5px;margin-top:2px">'
        +esc([normCat(a.raw['Material Category']),trim(a.raw['Discipline']),
             trim(a.raw['Manufacturer'])||trim(a.raw['Sub-contractor Name'])].filter(Boolean).join(' · '))
        +'</div></div></div>';
    }).join('')+more(p.add.length,40));

  body+=block('Changed — these will be updated',p.change,
    p.change.slice(0,40).map(function(c){
      return '<div class="line"><span class="tag t-wait">'+c.diff.length+' field'+(c.diff.length===1?'':'s')+'</span>'
        +'<div class="line-m"><div>'+esc(c.rec.name)+'</div>'
        +'<div class="dim" style="font-size:12.5px;margin-top:2px">'+esc(c.diff.slice(0,6).join(', '))
        +(c.diff.length>6?(' and '+(c.diff.length-6)+' more'):'')+'</div></div>'
        +'<span class="meta">row '+c.row+'</span></div>';
    }).join('')+more(p.change.length,40));

  body+=block('In the tracker but not in the file',p.gone,
    '<div class="dim" style="font-size:13.5px;margin-bottom:12px">These are kept unless you say otherwise. '
    +'A row deleted from the sheet and a row never in it look the same from here, so nothing is removed by accident.</div>'
    +p.gone.slice(0,40).map(function(m){
      return '<div class="line"><span class="tag t-na">kept</span>'
        +'<div class="line-m"><div>'+esc(m.name)+'</div>'
        +'<div class="dim" style="font-size:12.5px;margin-top:2px">'+esc(m.ref||'')+'</div></div></div>';
    }).join('')+more(p.gone.length,40));

  body+=block('Dates the tracker cannot use',p.badDates||[],
    '<div class="dim" style="font-size:13.5px;margin-bottom:12px">Every one of these cells is kept exactly as it is '
    +'typed and will come back out unchanged. What they cannot do is fall due, because a cell holding two dates — or '
    +'a month spelt in a way that could be either of two — does not say when something is expected.</div>'
    +(p.badDates||[]).slice(0,20).map(function(b){
      return '<div class="line"><span class="tag t-'+(b.many?'na':'now')+'">'
        +(b.many?'several':'row '+b.row)+'</span>'
        +'<div class="line-m"><div><span class="mono">'+esc(b.val)+'</span> '
        +'<span class="dim">in '+esc(b.col)+'</span></div>'
        +'<div class="dim" style="font-size:12.5px;margin-top:2px">'+esc(b.what)+'</div></div></div>';
    }).join('')+more((p.badDates||[]).length,20));

  body+=block('One manufacturer, two pre-qualifications',p.clash||[],
    '<div class="dim" style="font-size:13.5px;margin-bottom:12px">Qualification is held once per vendor here, '
    +'so these rows disagree with each other. Each row keeps its own reference; the vendor takes the first one read.</div>'
    +(p.clash||[]).map(function(c){
      return '<div class="line"><span class="tag t-bad">'+c.refs.length+' refs</span>'
        +'<div class="line-m"><div>'+esc(c.name)+'</div>'
        +'<div class="dim" style="font-size:12.5px;margin-top:2px">'
        +c.refs.map(function(r,i){return esc(r)+' (row '+c.rows[i]+')';}).join(' · ')+'</div></div></div>';
    }).join(''));

  body+='<div class="f-act" style="margin-top:22px">'
   +'<button class="btn btn-p" onclick="excelApply(false)">Apply — '
   +(p.add.length+p.change.length)+' record'+((p.add.length+p.change.length)===1?'':'s')+'</button>'
   +(p.gone.length?('<button class="btn btn-d btn-s" onclick="excelApply(true)">Apply and delete the '
      +p.gone.length+' missing</button>'):'')
   +'<button class="btn-q" onclick="closeSheet()">Cancel</button></div>';

  sheet('From '+p.file,'<div class="dim" style="font-size:13.5px;margin-bottom:16px">'
    +'Read from the <b>'+esc(p.sheet)+'</b> sheet'+(p.dividers?(', past '+p.dividers+' section rows'):'')
    +'. Nothing has been changed yet.</div>'+body);
}
function stat(v,l){
  return '<div class="stat"><div class="stat-v">'+v+'</div><div class="stat-l">'+esc(l)+'</div></div>';
}
function more(n,cap){
  return n>cap?('<div class="dim" style="font-size:13px;padding-top:10px">and '+(n-cap)+' more</div>'):'';
}
window.excelApply=function(dropGone){
  if(!PLAN)return;
  var p=PLAN,made=applyPlan(p,dropGone);
  closeSheet();
  var msg=p.add.length+' added, '+p.change.length+' updated';
  if(dropGone&&p.gone.length)msg+=', '+p.gone.length+' deleted';
  if(made.vendors.length)msg+=', '+made.vendors.length+' new vendor'+(made.vendors.length===1?'':'s');
  toast(msg);
  PLAN=null;
};

/* ================================================================
   9. THE ACONEX REGISTER
   ----------------------------------------------------------------
   The register is not the transmittal log. It carries one row per
   document and no history, so a reference appears exactly once and
   the question of which occurrence is the live one never arises.
   Three thousand rows of it are read in a moment, which is why this
   belongs here rather than in a sheet of formulas.

   Nothing is created from it. A document already named somewhere in
   the tracker has its outcome, its revision and its date brought up
   to date; a document that is not named anywhere is shown and left
   alone, because no column in the register says which material it
   belongs to and guessing that is worse than asking.
   ================================================================ */

/* the register's own words, and what each one means here. The first
   four are a verdict. Aconex writes the second of them two ways —
   2411 rows with spaces around the dash and 36 without — so the dash
   and the spaces are stripped before the comparison rather than both
   spellings being listed and a third one catching us out later. */
function verdictOf(s){
  var k=K(s).replace(/\s*-\s*/g,' ').replace(/\s+/g,' ').trim();
  if(!k||k==='none')return '';
  if(/^a approved$|^approved$/.test(k))return 'Approved';
  if(/^b approved with comments$/.test(k))return 'Approved as Noted';
  if(/^c revise & resubmit$|^c revise and resubmit$/.test(k))return 'Revise & Resubmit';
  if(/^d rejected$|^rejected$/.test(k))return 'Rejected';
  if(k==='terminated')return 'Terminated';
  if(/under workflow review|for approval|^pending$|under review/.test(k))return 'Under Review';
  return '';                       /* Closed, Reviewed: a step, not a verdict */
}

function delMeans(s){
  var k=K(s);
  if(k==='received')return 'Approved';
  if(k==='approved with comments')return 'Approved as Noted';
  if(k==='pending')return 'Under Review';
  return s;
}

/* where a reference lives decides where its outcome is written. The
   column the tracker actually holds it in is more reliable than the
   type the register gives it, because a transmittal number sitting in
   the dossier column is still a dossier as far as this file goes. */
var STATUS_OF={
 'PQD Number':{s:'PQD Status',r:'PQD Revision',d:'PQD Submittal Date',step:null},
 'MAT Number':{s:'MAT Status',r:'MAT Revision',d:'MAT Submittal Date',step:'mts'},
 'Method Statement Number':{s:'MES Status',r:'MES Revision',d:null,step:null},
 'ITP Number':{s:'ITP Status',r:'ITP Revision',d:'ITP Submittal Date',step:'itp'},
 'PID Number':{s:'PID Status',r:'PID Revision',d:'PID Submittal Date',step:'pid'},
 'MIR Number':{s:'MIR Status',r:null,d:'MIR Approval Date',step:null},
 'WIR Number':{s:'WIR Status',r:null,d:'WIR Approval Date',step:null},
 'FAT/TPI Results':{s:'FAT Package Status',r:null,d:null,step:'fat'},
 'FAT Package/Procedure Number/ITP':{s:'FAT Package Status',r:null,d:null,step:'fat'}
};
var REF_FIELDS=Object.keys(STATUS_OF).concat(['PA Document Number','PO Number']);

/* the category the procedure assigns, as written into the title. The
   register has no column for it, and 658 of the rows are C0 or C1 —
   material this tracker does not follow — so reading it off the title
   is what keeps those from burying the rest. */
function catOfTitle(t){
  var m=/CATEGOR\w*[\s:\-]*\(?"?\s*C?\s*([0-3])/i.exec(String(t||''));
  return m?('C'+m[1]):'';
}
function splitRefs(v){
  if(v==null||v==='')return [];
  return String(v).split(/[\r\n]+|\s{2,}/).map(trim)
    .filter(function(x){return x&&x!=='-';});
}
function trim(s){return String(s==null?'':s).replace(/\u00a0/g,' ').trim();}

/* every reference the tracker holds, and where it holds it */
function refIndex(){
  var idx={};
  (DB.mats||[]).forEach(function(m){
    var here={};
    REF_FIELDS.forEach(function(c){
      var all=splitRefs((m.raw||{})[c]);
      all.forEach(function(ref){
        here[K(ref)]=1;
        (idx[K(ref)]=idx[K(ref)]||[]).push({m:m,col:c,shared:all.length});
      });
    });
    /* An inspection request moved onto the material it belongs to lives
       in a consignment, not in a column, and would otherwise look new on
       every upload for ever. But while it is still in a column as well,
       the column is its home — indexing both would have one reference
       answering twice and disagreeing with itself. */
    (m.dels||[]).forEach(function(d){
      if(d.ref&&!here[K(d.ref)])
        (idx[K(d.ref)]=idx[K(d.ref)]||[]).push({m:m,col:'MIR Number',shared:1,del:d});
    });
    (m.ncrs||[]).forEach(function(n){
      if(n.no)(idx[K(n.no)]=idx[K(n.no)]||[]).push({m:m,col:'',shared:1,ncr:n});
    });
  });
  (DB.mfrs||[]).forEach(function(v){
    var pq=(v.steps||{}).pqd||{};
    if(pq.ref)(idx[K(pq.ref)]=idx[K(pq.ref)]||[]).push({v:v,col:'PQD Number',shared:1});
    (v.pq2||[]).forEach(function(x){
      (idx[K(x.ref)]=idx[K(x.ref)]||[]).push({v:v,col:'PQD Number',shared:1,second:true});
    });
  });
  return idx;
}

async function readRegister(file){
  var book=await openBook(file);
  var head=-1,rows=null,name='';
  for(var i=0;i<book.sheets.length&&head<0;i++){
    var r=await book.rows(book.sheets[i]);
    for(var j=0;j<Math.min(r.length,40);j++){
      if((r[j]||[]).some(function(c){return K(c)==='document no';})){
        head=j;rows=r;name=book.sheets[i].name;break;
      }
    }
  }
  if(head<0)throw new Error(
    'No sheet in that file has a "Document No" heading. This reads the register '
    +'export from Aconex — Document No, Title, Type, Status, Review Status, Discipline.');

  var col={};
  (rows[head]||[]).forEach(function(h,i){col[K(h)]=i;});
  function at(line,n){var i=col[n];return i==null?'':line[i];}

  /* the lines above the heading carry who ran it and when */
  var about={};
  rows.slice(0,head).forEach(function(line){
    var a=trim(line[0]),b=trim(line[1]);
    if(/^generated on/i.test(a))about.on=b;
    else if(/^generated by/i.test(a))about.by=b;
    else if(/^filter/i.test(a)){about.filter=b;about._f=true;}
    else if(about._f&&!a&&b)about.filter+=' · '+b;
    else if(a)about._f=false;
    else if(/^project/i.test(a))about.project=b;
  });

  var out=[];
  rows.slice(head+1).forEach(function(line){
    var no=trim(at(line,'document no'));
    if(!no)return;
    out.push({
      no:no,
      rev:trim(at(line,'revision')),
      title:trim(at(line,'title')),
      type:trim(at(line,'type')),
      status:trim(at(line,'status')),
      review:trim(at(line,'review status')),
      disc:trim(at(line,'discipline')),
      date:anyDate(at(line,'revision date'))||'',
      moved:anyDate(at(line,'date modified'))||'',
      tran:trim(at(line,'transmittal in'))
    });
  });
  if(!out.length)throw new Error('That sheet has a heading but no document rows under it.');
  return {rows:out,sheet:name,about:about,cols:Object.keys(col)};
}

/* ---------------------------------------------------------------
   What the register says against what the tracker holds.
   --------------------------------------------------------------- */
function planRegister(reg){
  var idx=refIndex();
  var p={moved:[],ended:[],locked:[],same:[],newC23:[],newPlain:[],newC01:[],unknown:0};
  reg.forEach(function(d){
    var hits=idx[K(d.no)];
    /* Review Status is the reviewer's word and Status the document's own;
       where they disagree the review is the later of the two. */
    var want=verdictOf(d.review)||verdictOf(d.status);
    if(!hits||!hits.length){
      var cat=catOfTitle(d.title);
      d.cat=cat;d.want=want;
      if(cat==='C2'||cat==='C3')p.newC23.push(d);
      else if(!cat)p.newPlain.push(d);
      else p.newC01.push(d);
      return;
    }
    hits.forEach(function(h){
      var where=STATUS_OF[h.col];
      var rec=h.m||h.v;
      var now=h.del?trim(h.del.status)
             :h.m?trim((h.m.raw||{})[where?where.s:''])
                 :trim(h.second?(((h.v.pq2||[]).filter(function(x){
                     return K(x.ref)===K(d.no);})[0]||{}).status)
                   :(((h.v.steps||{}).pqd||{}).status));
      var item={d:d,h:h,rec:rec,where:where,now:now,want:want};
      if(!where||!want){p.same.push(item);return;}
      /* Both sides are read through the same reduction before being
         compared. The file writes Approve, Aprrove and Approved as Noted;
         the register writes B - Approved with Comments. Comparing the
         words themselves would report every document as moved, and
         would report it again after it had been set. */
      /* A consignment says Received where a column says Approved. Both
         mean the material arrived and was accepted, so the comparison is
         made on the meaning rather than on either word. */
      var mine=h.del?delMeans(now):now;
      if(want!=='Terminated'&&normStatus(mine)===normStatus(want)){p.same.push(item);return;}
      if(want==='Terminated'&&K(now)==='terminated'){p.same.push(item);return;}
      /* a cell holding four references cannot be given one status */
      if(h.shared>1){p.locked.push(item);return;}
      if(want==='Terminated')p.ended.push(item);
      else p.moved.push(item);
    });
  });
  /* read in a settled order: type first, then number, so a block of a
     thousand rows is something a person can work down */
  function tidy(list){
    list.sort(function(a,b){
      var d=String(a.type||'').localeCompare(String(b.type||''));
      return d||String(a.no).localeCompare(String(b.no));
    });
  }
  tidy(p.newC23);tidy(p.newPlain);tidy(p.newC01);
  return p;
}

function applyRegister(p,alsoEnded){
  var n=0;
  function write(it){
    var w=it.where,d=it.d;
    if(it.h.v){                                   /* a vendor's qualification */
      it.h.v.steps=it.h.v.steps||{};
      var pq=it.h.v.steps.pqd||{};
      pq.status=it.want==='Approved as Noted'?'Approved with comments':it.want;
      if(d.date)pq.date=d.date;
      it.h.v.steps.pqd=pq;n++;return;
    }
    if(it.h.del){                                 /* it lives in a consignment */
      it.h.del.status=(it.want==='Approved as Noted')?'Approved with comments'
        :(it.want==='Approved')?'Received':it.want;
      if(d.date)it.h.del.date=d.date;
      n++;return;
    }
    var m=it.h.m;
    m.raw=m.raw||{};
    m.raw[w.s]=it.want;
    if(w.r&&d.rev!=='')m.raw[w.r]=d.rev;
    if(w.d&&d.date)m.raw[w.d]=d.date;
    /* the page's own road, so the board moves with the sheet */
    if(w.step){
      m.steps=m.steps||{};
      var st=m.steps[w.step]||{};
      st.ref=st.ref||d.no;
      if(d.date)st.date=d.date;
      st.status=(it.want==='Approved as Noted')?'Approved with comments'
               :(it.want==='Revise & Resubmit')?'Resubmit'
               :(it.want==='Under Review')?'Pending':it.want;
      m.steps[w.step]=st;
    }
    n++;
  }
  p.moved.forEach(write);
  if(alsoEnded)p.ended.forEach(write);
  touch();rList();rPane();
  return n;
}

/* ---------------------------------------------------------------
   The screen. Everything is shown — three thousand rows of it — but
   in the order the day should be spent: what moved, what was ended,
   then what is new, heaviest category first.
   --------------------------------------------------------------- */
var REG=null;
window.regPick=function(){document.getElementById('xl-reg').click();};
window.regRead=async function(ev){
  var f=ev.target.files[0];ev.target.value='';
  if(!f)return;
  if(typeof busy==='function')busy(true,'Reading the register');
  try{
    var reg=await readRegister(f);
    REG=planRegister(reg.rows);
    REG.file=f.name;REG.about=reg.about;REG.count=reg.rows.length;REG.sheet=reg.sheet;
    if(typeof busy==='function')busy(false);
    showRegister();
  }catch(e){
    if(typeof busy==='function')busy(false);
    sheet('That register could not be read',
      '<div style="font-size:14px;line-height:1.75">'+esc(e.message||String(e))+'</div>');
  }
};
function regRow(it){
  var d=it.d;
  var go=it.h.m?("jump('mat',"+it.h.m.id+")"):("jump('mfr',"+it.h.v.id+")");
  return '<div class="line row-a" onclick="'+go+'">'
    +'<span class="tag t-'+(it.want==='Terminated'?'bad':it.want==='Rejected'?'bad'
        :it.want==='Approved'?'ok':'now')+'" style="min-width:118px;text-align:center;flex-shrink:0">'
    +esc(it.now||'nothing yet')+' → '+esc(it.want)+'</span>'
    +'<div class="line-m"><div>'+esc(it.rec.name)+'</div>'
    +'<div class="dim" style="font-size:12.5px;margin-top:2px"><span class="mono">'+esc(d.no)+'</span>'
    +' · '+esc(it.where?it.where.s:'')+(d.rev!==''?(' · rev '+esc(d.rev)):'')+'</div></div></div>';
}
function newRow(d){
  return '<div class="line">'
    +'<span class="tag t-'+(d.cat==='C3'||d.cat==='C2'?'now':'na')+'" '
    +'style="min-width:52px;text-align:center;flex-shrink:0">'+esc(d.cat||'—')+'</span>'
    +'<div class="line-m"><div>'+esc(d.title||'(no title)')+'</div>'
    +'<div class="dim" style="font-size:12.5px;margin-top:2px"><span class="mono">'+esc(d.no)+'</span>'
    +' · '+esc(d.type)+' · '+esc(d.disc)+(d.want?(' · '+esc(d.want)):'')+'</div></div></div>';
}
/* A thousand of the new rows are inspection requests and carry no
   category in their title, so a single number for "no category" says
   nothing. The types are counted under the heading instead. */
function byType(list){
  var n={},order=[];
  list.forEach(function(x){
    var t=(x.d?x.d.type:x.type)||'—';
    if(!(t in n)){n[t]=0;order.push(t);}
    n[t]++;
  });
  order.sort(function(a,b){return n[b]-n[a];});
  return order.map(function(t){return n[t]+' '+t;}).join(' · ');
}
function regBlock(title, list, draw, cap){
  if(!list.length)return '';
  cap=cap||40;
  return '<div class="sec">'+esc(title)+' <span class="dim">'+list.length+'</span></div>'
    +'<div class="panel"><div class="panel-b">'
    +'<div class="dim" style="font-size:12.5px;margin-bottom:10px">'+esc(byType(list))+'</div>'
    +list.slice(0,cap).map(draw).join('')
    +(list.length>cap?('<div class="dim" style="font-size:13px;padding-top:10px">and '
      +(list.length-cap)+' more — the full list is in the spreadsheet below</div>'):'')
    +'</div></div>';
}
function showRegister(){
  var p=REG;
  var body='<div class="grid" style="margin-bottom:18px">'
   +stat(p.moved.length,'Outcome moved')
   +stat(p.ended.length,'Terminated')
   +stat(p.newC23.length,'New · C2 and C3')
   +stat(p.newPlain.length,'New · no category')
   +stat(p.newC01.length,'New · C0 and C1')
   +stat(p.same.length,'Already matching')
   +'</div>';

  if(p.locked.length)body+='<div class="panel" style="border-color:var(--now);'
    +'background:var(--now-b);margin-bottom:12px"><div class="panel-b">'
    +'<b>'+p.locked.length+' could not be set on their own.</b> Each shares a cell with other '
    +'references in the Main Log, and one cell cannot hold two different outcomes. '
    +'They are listed at the end.</div></div>';

  body+=regBlock('The outcome moved',p.moved,regRow);
  body+=regBlock('Terminated — the document is dead, so nothing should be waiting on it',
                 p.ended,regRow);
  body+=regBlock('New · category 2 and 3 — these belong in your file',p.newC23,newRow);
  body+=regBlock('New · no category named in the title',p.newPlain,newRow,25);
  body+=regBlock('New · category 0 and 1 — this tracker does not follow these',p.newC01,newRow,10);
  body+=regBlock('Shares a cell with other references',p.locked,regRow,15);

  var fresh=p.newC23.length+p.newPlain.length+p.newC01.length;
  body+='<div class="f-act" style="margin-top:22px">'
   +((p.moved.length||fresh)?('<button class="btn btn-p" onclick="regAll()">Do it all — '
      +p.moved.length+' updated, '+fresh+' brought in</button>'):'')
   +(p.moved.length?('<button class="btn" onclick="regApply(false)">Only the '
      +p.moved.length+' that moved</button>'):'')
   +(p.ended.length?('<button class="btn" onclick="regApply(true)">Apply those and mark the '
      +p.ended.length+' terminated</button>'):'')
   +(fresh?('<button class="btn" onclick="regAddAll()">Bring in all '+fresh
      +' new for review</button>'):'')
   +'<button class="btn" onclick="regCSV()">Export the whole list</button>'
   +'<button class="btn-q" onclick="closeSheet()">Close</button></div>'
   +(fresh?('<div class="dim" style="font-size:13px;margin-top:10px;line-height:1.7">'
     +'Bringing them in creates '+fresh+' records at once and marks every one of them '
     +'unreviewed, so they can be told apart from what was already here — and the whole '
     +'upload can be taken back in one action if it turns out wrong. '
     +'Download a backup first.</div>'):'');

  var a=p.about||{};
  sheet('From '+p.file,
    '<div class="dim" style="font-size:13.5px;margin-bottom:16px">'
    +p.count+' documents on the <b>'+esc(p.sheet)+'</b> sheet'
    +(a.on?(' · generated '+esc(a.on)):'')+(a.by?(' · by '+esc(a.by.split(',')[0])):'')
    +'. Nothing has been changed yet, and nothing new is created — a new document is shown, '
    +'not added, because the register does not say which material it belongs to.</div>'+body);
}
window.regApply=function(alsoEnded){
  if(!REG)return;
  var n=applyRegister(REG,alsoEnded);
  closeSheet();
  toast(n+' document'+(n===1?'':'s')+' brought up to date');
  REG=null;
};
/* The morning, in one press: the outcomes that moved are written, the
   terminated ones are marked, and everything the file has never seen is
   brought in for review. The separate buttons stay for the times when
   only half of that is wanted. */
window.regAll=function(){
  if(!REG)return;
  var p=REG;
  var moved=p.moved.length+p.ended.length;
  var fresh=p.newC23.length+p.newPlain.length+p.newC01.length;
  sheet('Do it all',
    '<div style="font-size:14px;line-height:1.8">This will, in order:'
    +'<br><br><b>1.</b> Update '+p.moved.length+' document'+(p.moved.length===1?'':'s')
    +' whose outcome moved'
    +(p.ended.length?(', and mark '+p.ended.length+' terminated'):'')+'.'
    +'<br><b>2.</b> Bring in '+fresh+' new record'+(fresh===1?'':'s')+', every one marked '
    +'unreviewed so it can be told from what was already here.'
    +'<br><br>The whole intake can be taken back afterwards from '
    +'<b>More \u2192 Waiting to be reviewed</b>, as long as you have not marked it reviewed. '
    +'Saving '+fresh+' records takes a moment; the percentage beside the project name says '
    +'how far it has got.</div>'
    +'<div class="f-act" style="margin-top:22px">'
    +'<button class="btn btn-p" onclick="regAllYes()">Go ahead</button>'
    +'<button class="btn" onclick="regCSV()">Export the list first</button>'
    +'<button class="btn-q" onclick="showRegister()">Back</button></div>');
};
window.regAllYes=function(){
  if(!REG)return;
  var p=REG;
  var n=applyRegister(p,true);
  var tag=(REG.file||'a register')+' \u00b7 '+new Date().toLocaleString('en-GB',
    {day:'2-digit',month:'short',hour:'2-digit',minute:'2-digit'});
  var made=createFromRegister(p,tag);
  closeSheet();
  toast(n+' updated, '+made.mats+' materials and '+made.vendors+' vendors brought in');
  REG=null;
  setTimeout(regReview,600);
};
window.regAddAll=function(){
  if(!REG)return;
  var p=REG;
  var tag=(REG.file||'a register')+' · '+new Date().toLocaleString('en-GB',
    {day:'2-digit',month:'short',hour:'2-digit',minute:'2-digit'});
  var n=createFromRegister(p,tag);
  closeSheet();
  toast(n.mats+' materials and '+n.vendors+' vendors brought in — all marked for review');
  setTimeout(regReview,400);
};
window.regCSV=function(){
  if(!REG)return;
  var p=REG,lines=[];
  function add(group,d,extra){
    lines.push([group,d.no,d.type,d.disc,d.cat||'',d.title,d.rev,d.status,d.review,
      d.date?show(d.date):'',extra||''].map(function(c){
        return '"'+String(c==null?'':c).replace(/"/g,'""')+'"';}).join(','));
  }
  p.moved.forEach(function(it){add('outcome moved',it.d,it.now+' → '+it.want+' · '+it.rec.name);});
  p.ended.forEach(function(it){add('terminated',it.d,it.rec.name);});
  p.locked.forEach(function(it){add('shares a cell',it.d,it.rec.name);});
  p.newC23.forEach(function(d){add('new C2/C3',d);});
  p.newPlain.forEach(function(d){add('new, no category',d);});
  p.newC01.forEach(function(d){add('new C0/C1',d);});
  p.same.forEach(function(it){add('already matching',it.d,it.rec.name);});
  var head=['Group','Document No','Type','Discipline','Category','Title','Revision',
            'Status','Review Status','Revision Date','Note'].join(',');
  var blob=new Blob(['\ufeff'+[head].concat(lines).join('\r\n')],{type:'text/csv;charset=utf-8'});
  var u=URL.createObjectURL(blob),a=document.createElement('a');
  a.href=u;a.download=(DB.project||'register')+' — against the register '+today()+'.csv';
  document.body.appendChild(a);a.click();document.body.removeChild(a);
  setTimeout(function(){URL.revokeObjectURL(u);},900);
  toast('Exported '+lines.length+' rows');
};

/* ================================================================
   THE REGISTER, BROUGHT IN WHOLE
   ----------------------------------------------------------------
   Everything the register holds becomes a record, and the sorting out
   happens afterwards inside the tracker rather than in a spreadsheet
   beforehand. That is a deliberate trade: it is faster to get started
   and it puts three thousand rows in front of a person who can judge
   them, at the cost of a file that is briefly untidy.

   Two things make the untidiness survivable. Every record created this
   way is stamped with the upload it came from and marked unreviewed,
   so it can always be told from what was there before. And an upload
   can be taken back whole, so a bad import is a button rather than an
   afternoon.
   ================================================================ */

/* the register writes its disciplines with a code in front */
function plainDisc(s){return trim(String(s||'').replace(/^[A-Z]{2,4}\s*-\s*/,''));}

/* Which columns of the Main Log a document of this type belongs in.
   A document that lands here is the only thing on its row, so the
   number, the date, the revision and the outcome all go together. */
var NEW_AS={
 'material submittal':          {n:'MAT Number',d:'MAT Submittal Date',r:'MAT Revision',s:'MAT Status'},
 'inspection & test plan':      {n:'ITP Number',d:'ITP Submittal Date',r:'ITP Revision',s:'ITP Status'},
 'method statement':            {n:'Method Statement Number',d:null,r:'MES Revision',s:'MES Status'},
 'material inspection request': {n:'MIR Number',d:'MIR Approval Date',r:null,s:'MIR Status'}
};

/* The company a pre-qualification is about, dug out of its title.
   The titles are written by hand and no two agree: "P4- Makkah-
   Prequalification-Sodamco-Concrete Admixtures…", "PRQ for Dar
   Al-Rokham - Marble Cladding work", "P4-Makkah-Prequalification-
   Supplier- JAZEERA PAINTS…". This gets most of them and will get
   some of them wrong, which is why the whole title is kept beside the
   name and the record is marked for review rather than trusted. */
function companyOf(title){
  var t=' '+String(title||'').replace(/\s+/g,' ').trim()+' ';
  t=t.replace(/^\s*P4\s*-?\s*/i,' ')
     .replace(/\bmakkah\b/i,' ')
     .replace(/\b(pre[- ]?qualification|prequalification|PRQ)\b/ig,' ')
     .replace(/\bfor\b/i,' ')
     .replace(/\b(supplier|sub[- ]?contractor|subcontractor|manufacturer|vendor)\b\s*[-:]?/ig,' ')
     .replace(/^[\s\-:–,]+/,'');
  /* the name runs until the work it is qualified for starts */
  var cut=t.split(/\s[-–]\s|\s*[-–]\s|\(|,|\u2013/)[0];
  cut=trim(cut).replace(/[\s\-:]+$/,'');
  if(cut.length<3||cut.length>60)cut=trim(t).slice(0,60);
  return cut||'(name not in the title)';
}
function kindOfTitle(t){
  var k=K(t);
  if(/sub[- ]?contractor/.test(k))return 'sub';
  if(/supplier|agency letter|distributor/.test(k))return 'supplier';
  if(/third party|inspection agency|tuv|bureau veritas/.test(k))return 'agency';
  return 'maker';
}

/* a raw row of the Main Log, built from one line of the register */
/* What a reference is, read off the reference itself. The register has
   a Type column, but a record can also arrive from the Main Log where
   there is no such column — and the three letters in the middle of an
   Aconex number say it either way. */
function docKind(ref,type){
  var t=K(type);
  if(t==='material submittal')return '';
  if(t==='pre-qualification')return 'PQD';
  if(t==='method statement')return 'MES';
  if(t==='inspection & test plan')return 'ITP';
  if(t==='material inspection request')return 'MIR';
  var r=String(ref||'').toUpperCase();
  if(/-MES-/.test(r))return 'MES';
  if(/-ITP-/.test(r))return 'ITP';
  if(/-MIR-/.test(r))return 'MIR';
  if(/-WIR-/.test(r))return 'WIR';
  if(/-TRN-/.test(r))return 'PID';
  if(/-PRQ-/.test(r))return 'PQD';
  if(/-REP-|-RPT-/.test(r))return 'Report';
  if(/-PRO-/.test(r))return 'Procedure';
  return '';
}

function rawFromDoc(d){
  var raw={};
  raw['Item Description']=d.title||d.no;
  if(d.cat)raw['Material Category']='Category '+d.cat;
  if(d.disc)raw['Discipline']=plainDisc(d.disc);
  var w=NEW_AS[K(d.type)];
  if(w){
    raw[w.n]=d.no;
    if(w.d&&d.date)raw[w.d]=d.date;
    if(w.r&&d.rev!=='')raw[w.r]=d.rev;
    if(w.s&&d.want)raw[w.s]=d.want;
  }else{
    /* a type with no home of its own still keeps its number somewhere
       it can be found again */
    raw['MAT Number']=d.no;
    if(d.want)raw['MAT Status']=d.want;
    if(d.date)raw['MAT Submittal Date']=d.date;
  }
  return raw;
}

function createFromRegister(p,tag){
  var made={n:1,vendors:[],id:idMaker()},mats=0,vends=0;
  var all=p.newC23.concat(p.newPlain,p.newC01);
  all.forEach(function(d){
    if(K(d.type)==='pre-qualification'){
      var name=companyOf(d.title);
      var twin=(DB.mfrs||[]).filter(function(v){return K(v.name)===K(name);})[0];
      var v=twin||{id:made.id(),name:name,kind:kindOfTitle(d.title),cat:d.cat||'',
        country:'',site:'',scope:'',steps:{},pq:{},added:today()};
      v.scope=v.scope||d.title;
      v.steps=v.steps||{};
      if(!(v.steps.pqd&&v.steps.pqd.ref)){
        v.steps.pqd={ref:d.no,date:d.date||'',
          status:d.want==='Approved as Noted'?'Approved with comments':(d.want||'Pending')};
      }else if(K(v.steps.pqd.ref)!==K(d.no)){
        /* the same company qualified twice. The page holds one
           qualification per vendor, so the second is kept beside it
           rather than dropped — losing it would make this document
           look new again on the next upload, for ever. */
        v.pq2=v.pq2||[];
        if(!v.pq2.some(function(x){return K(x.ref)===K(d.no);}))
          v.pq2.push({ref:d.no,date:d.date||'',status:d.want||'',title:d.title||''});
      }
      if(!twin){v.reg=tag;v.review=1;DB.mfrs.push(v);vends++;}
      return;
    }
    var m={id:made.id(),name:'',cat:'',ref:'',mfr:'',qty:'',unit:'',
      steps:{},dels:[],ncrs:[],added:today()};
    applyRaw(m,rawFromDoc(d),made);
    m.reg=tag;m.review=1;
    var kind=docKind(d.no,d.type);
    if(kind)m.doc=kind;          /* a document, not a material of its own */
    DB.mats.push(m);mats++;
  });
  touch();rList();rPane();
  return {mats:mats,vendors:vends};
}

/* Records brought in before this existed carry no label, so one is
   worked out from the reference they hold. Run once, it costs nothing
   and it is what the Documents tab reads. */
function labelDocuments(){
  var n=0;
  (DB.mats||[]).forEach(function(m){
    if(m.doc!=null)return;
    var ref=m.ref||'';
    if(!ref)REF_FIELDS.some(function(c){
      var v=splitRefs((m.raw||{})[c])[0];
      if(v){ref=v;return true;}
      return false;
    });
    var kind=docKind(ref,'');
    if(kind){m.doc=kind;n++;}
  });
  if(n)touch();
  return n;
}
window.labelDocuments=labelDocuments;

/* ---------------------------------------------------------------
   The review pass. Everything brought in this way is listed here
   until it is looked at, whatever tab it ended up in.
   --------------------------------------------------------------- */
function pending(){
  var out=[];
  (DB.mats||[]).forEach(function(m){if(m.review)out.push({k:'mat',r:m});});
  (DB.mfrs||[]).forEach(function(v){if(v.review)out.push({k:'mfr',r:v});});
  return out;
}
function batches(){
  var b={};
  pending().forEach(function(x){
    var t=x.r.reg||'(unknown upload)';
    b[t]=(b[t]||0)+1;
  });
  return b;
}
window.regReview=function(){
  var list=pending();
  if(!list.length)return sheet('Nothing waiting',
    '<div class="dim" style="padding:26px 0;text-align:center">'
    +'Nothing is waiting to be reviewed. Records brought in from a register appear here '
    +'until you have been through them.</div>');
  var b=batches();
  var cap=60;
  sheet('Waiting to be reviewed',
    '<div class="dim" style="font-size:13.5px;margin-bottom:16px">'
    +list.length+' record'+(list.length===1?'':'s')+' came in from a register and have not '
    +'been looked at. They sit in Materials and Vendors like any other, and this list is '
    +'only a way of finding them again.</div>'
    +'<div class="panel"><div class="panel-b">'
    +Object.keys(b).map(function(t){
      return '<div class="line"><span class="tag t-na">'+b[t]+'</span>'
        +'<div class="line-m">'+esc(t)+'</div>'
        +'<button class="btn btn-s btn-d" onclick="regUndo(\''+attr(t)+'\')">Take this upload back</button>'
        +'</div>';}).join('')
    +'</div></div>'
    +'<div class="sec">The records</div><div class="panel"><div class="panel-b">'
    +list.slice(0,cap).map(function(x){
      return '<div class="line row-a" onclick="closeSheet();jump(\''+(x.k==='mat'?'mat':'mfr')
        +'\','+x.r.id+')">'
        +'<span class="tag t-'+(x.k==='mat'?'na':'wait')+'" style="min-width:64px;text-align:center">'
        +(x.k==='mat'?'material':'vendor')+'</span>'
        +'<div class="line-m"><div>'+esc(x.r.name||'(no name)')+'</div>'
        +'<div class="dim" style="font-size:12.5px;margin-top:2px">'
        +esc([x.r.cat,x.r.disc,x.r.ref].filter(Boolean).join(' · '))+'</div></div>'
        +'<button class="btn-q" onclick="event.stopPropagation();regDone('+x.r.id+')">Reviewed</button>'
        +'</div>';}).join('')
    +(list.length>cap?('<div class="dim" style="font-size:13px;padding-top:10px">and '
      +(list.length-cap)+' more</div>'):'')
    +'</div></div>'
    +'<div class="f-act" style="margin-top:20px">'
    +'<button class="btn" onclick="regDoneAll()">Mark all '+list.length+' reviewed</button>'
    +'<button class="btn-q" onclick="closeSheet()">Close</button></div>');
};
window.regDone=function(id){
  (DB.mats||[]).concat(DB.mfrs||[]).forEach(function(r){
    if(String(r.id)===String(id)){delete r.review;}
  });
  touch();rList();regReview();
};
window.regDoneAll=function(){
  pending().forEach(function(x){delete x.r.review;});
  touch();rList();rPane();closeSheet();toast('All marked reviewed');
};
window.regUndo=function(tag){
  var n=pending().filter(function(x){return (x.r.reg||'(unknown upload)')===tag;}).length;
  sheet('Take back '+n+' record'+(n===1?'':'s')+'?',
    '<div style="font-size:14px;line-height:1.75">Everything that came in from '
    +'<b>'+esc(tag)+'</b> and has not yet been reviewed will be deleted. Anything you '
    +'have already marked reviewed stays, and nothing that was in the file before the '
    +'upload is touched.</div>'
    +'<div class="f-act" style="margin-top:20px">'
    +'<button class="btn btn-d" onclick="regUndoYes(\''+attr(tag)+'\')">Delete them</button>'
    +'<button class="btn-q" onclick="regReview()">Keep them</button></div>');
};
window.regUndoYes=function(tag){
  function drop(r){return r.review&&(r.reg||'(unknown upload)')===tag;}
  var before=(DB.mats||[]).length+(DB.mfrs||[]).length;
  DB.mats=(DB.mats||[]).filter(function(m){return !drop(m);});
  DB.mfrs=(DB.mfrs||[]).filter(function(v){return !drop(v);});
  var gone=before-(DB.mats.length+DB.mfrs.length);
  if(SEL.mat&&!mat(SEL.mat))SEL.mat=DB.mats.length?DB.mats[0].id:null;
  if(SEL.mfr&&!mfr(SEL.mfr))SEL.mfr=DB.mfrs.length?DB.mfrs[0].id:null;
  touch();rList();rPane();closeSheet();
  toast(gone+' record'+(gone===1?'':'s')+' taken back');
};

/* ================================================================
   DOCUMENTS, AND THE MATERIALS THEY SERVE
   ----------------------------------------------------------------
   A method statement is not a material. It arrived as one because the
   register has no column saying which material it belongs to, and
   guessing was worse than asking. So it keeps its own record and its
   own tab, and the link is made by hand from the material — which is
   the thing everything else hangs off.

   One document serves several materials: in this project one
   inspection plan covers twelve. The link is therefore a list on the
   material, and a document knows which materials point at it only by
   being looked for. With a few thousand records that costs nothing,
   and it keeps the material as the single place a link is edited.
   ================================================================ */

/* The Materials machinery draws three lists now. Inspection requests
   are twelve hundred of the fifteen hundred documents, so leaving them
   in with the rest would bury everything else — they get a tab of their
   own, and the remaining kinds are sorted out inside Documents. */
var VIEW='mat';                            /* mat | mir | doc */
var DOCKIND='';                            /* which kind, inside Documents */
var DOC_KINDS=['MES','ITP','PID','PQD','Report','Procedure','WIR'];

function isDoc(m){return !!(m&&m.doc);}
function inView(m){
  if(VIEW==='mat')return !isDoc(m);
  if(VIEW==='mir')return m.doc==='MIR';
  return isDoc(m)&&m.doc!=='MIR'&&(!DOCKIND||m.doc===DOCKIND);
}
function docsOf(m){
  var ids=(m&&m.docs)||[];
  return ids.map(function(id){
    return (DB.mats||[]).filter(function(x){return String(x.id)===String(id);})[0];
  }).filter(Boolean);
}
function servedBy(doc){
  return (DB.mats||[]).filter(function(m){
    return (m.docs||[]).some(function(id){return String(id)===String(doc.id);});
  });
}

/* ---------------------------------------------------------------
   The tab bar moves out of the rail and across the top, because four
   tabs fitted down the side and seven do not.
   --------------------------------------------------------------- */
function liftTabs(){
  var tabs=document.querySelector('.side .tabs');
  var main=document.getElementById('pane');
  var app=document.querySelector('.app');
  if(!tabs||!main||!app||document.querySelector('.mainwrap'))return;

  var css=document.createElement('style');
  css.textContent=
   '.mainwrap{flex:1;min-width:0;min-height:0;display:flex;flex-direction:column;overflow:hidden}'
  +'.main{flex:1;min-height:0}'
  +'.topbar{display:flex;gap:2px;padding:8px 14px 0;background:var(--card);'
  +'border-bottom:1px solid var(--line);flex-shrink:0;flex-wrap:wrap;overflow-x:auto}'
  +'.topbar .tab{flex:0 0 auto;color:var(--ink-3);border-radius:8px 8px 0 0;min-height:42px;'
  +'padding:10px 15px;border-bottom:2px solid transparent;background:none}'
  +'.topbar .tab:hover{background:var(--hover);color:var(--ink)}'
  +'.topbar .tab[aria-selected=true]{background:none;color:var(--ink);font-weight:550;'
  +'border-bottom-color:var(--accent)}'
  +'.topbar .tab .n{color:var(--ink-4);font-family:var(--mono);font-size:11px}'
  +'@media(max-width:880px){.mainwrap{overflow:visible}.topbar{padding:8px 10px 0}}';
  document.head.appendChild(css);

  var col=document.createElement('div');
  col.className='mainwrap';
  app.insertBefore(col,main);
  tabs.classList.add('topbar');
  col.appendChild(tabs);
  col.appendChild(main);

  /* Inspections and Documents sit beside Materials, not inside it */
  var after=document.getElementById('tab-mat');
  [['mir','Inspections'],['doc','Documents']].forEach(function(pair){
    var b=document.createElement('button');
    b.className='tab';b.id='tab-'+pair[0];b.setAttribute('role','tab');
    b.setAttribute('aria-selected','false');
    b.appendChild(document.createTextNode(pair[1]+' '));
    var cnt=document.createElement('span');
    cnt.className='n';cnt.id='n-'+pair[0];cnt.textContent='0';
    b.appendChild(cnt);
    b.onclick=function(){setTab(pair[0]);};
    if(after&&after.parentNode===tabs){tabs.insertBefore(b,after.nextSibling);after=b;}
    else tabs.appendChild(b);
  });
  var rep=document.createElement('button');
  rep.className='tab';rep.id='tab-rep';rep.setAttribute('role','tab');
  rep.setAttribute('aria-selected','false');
  rep.appendChild(document.createTextNode('Reports'));
  rep.onclick=function(){setTab('rep');};
  tabs.appendChild(rep);
}

function paintTabs(){
  var n={mat:0,mir:0,doc:0};
  (DB.mats||[]).forEach(function(m){
    if(!isDoc(m))n.mat++;
    else if(m.doc==='MIR')n.mir++;
    else n.doc++;
  });
  ['mat','mir','doc'].forEach(function(k){
    var c=document.getElementById('n-'+k);if(c)c.textContent=n[k];
    var t=document.getElementById('tab-'+k);
    if(t)t.setAttribute('aria-selected',String(TAB==='mat'&&VIEW===k));
  });
  var r=document.getElementById('tab-rep');
  if(r)r.setAttribute('aria-selected',String(TAB==='rep'));
}

/* ---------------------------------------------------------------
   The two views share the Materials machinery and differ only in
   which records they are given, so the list is drawn by the page's
   own code over a filtered set rather than reimplemented here.
   --------------------------------------------------------------- */
function withSubset(fn){
  var all=DB.mats;
  DB.mats=all.filter(inView);
  try{return fn();}finally{DB.mats=all;}
}

/* inside Documents the kinds are separated, because a method statement
   and a transmittal are not the same errand */
function kindChips(){
  if(TAB!=='mat'||VIEW!=='doc')return;
  var box=document.getElementById('filters');
  if(!box)return;
  var n={};
  (DB.mats||[]).forEach(function(m){
    if(isDoc(m)&&m.doc!=='MIR')n[m.doc]=(n[m.doc]||0)+1;
  });
  var total=Object.keys(n).reduce(function(a,k){return a+n[k];},0);
  var html='<button class="fchip" aria-pressed="'+(!DOCKIND)+'" onclick="setKind(\'\')">'
    +'All<span class="fn">'+total+'</span></button>';
  DOC_KINDS.forEach(function(k){
    if(!n[k])return;
    html+='<button class="fchip" aria-pressed="'+(DOCKIND===k)+'" onclick="setKind(\''+k+'\')">'
      +esc(k)+'<span class="fn">'+n[k]+'</span></button>';
  });
  box.innerHTML=html+box.innerHTML;
}
window.setKind=function(k){DOCKIND=(DOCKIND===k?'':k);rList();rPane();};

function install2(){
  if(window.__docs)return;
  window.__docs=true;
  liftTabs();

  var origSetTab=window.setTab;
  window.setTab=function(t){
    if(t==='mir'||t==='doc'){VIEW=t;if(t!=='doc')DOCKIND='';origSetTab('mat');}
    else if(t==='rep'){
      VIEW='mat';DOCKIND='';
      origSetTab('home');          /* borrows the shape of a page with no list */
      window.TAB='rep';
      var sb=document.getElementById('side-body');
      if(sb)sb.classList.add('hidden');
      var add=document.getElementById('add-btn');
      if(add)add.style.display='none';
      rPane();
    }
    else{VIEW='mat';DOCKIND='';origSetTab(t);}
    paintTabs();
  };

  var origList=window.rList;
  window.rList=function(){
    if(TAB!=='mat')
      {origList();paintTabs();return;}
    withSubset(origList);
    kindChips();
    paintTabs();
  };

  var origPane=window.rPane;
  window.rPane=function(){
    if(TAB==='rep'){
      var el=document.getElementById('pane');
      if(el)el.innerHTML=reportsPane();
      return;
    }
    if(TAB!=='mat')return origPane();
    withSubset(origPane);
  };

  /* the vendor's page gains the one about who brought them */
  var origMfr=window.mfrPane;
  window.mfrPane=function(v){
    var html=origMfr(v);
    var i=html.lastIndexOf('</div></div>');
    if(i<0)return html+broughtPanel(v);
    return html.slice(0,i)+broughtPanel(v)+html.slice(i);
  };

  /* the material's page gains the panel where the linking happens */
  var origMat=window.matPane;
  window.matPane=function(m){
    var html=origMat(m);
    var i=html.lastIndexOf('</div></div>');
    if(i<0)return html+linkPanel(m);
    return html.slice(0,i)+linkPanel(m)+html.slice(i);
  };
}

function linkPanel(m){
  if(isDoc(m)){
    var on=servedBy(m);
    return '<div class="sec">The materials this '+esc(m.doc)+' serves</div>'
      +'<div class="panel"><div class="panel-b">'
      +(on.length?on.map(function(x){
          return '<div class="line row-a" onclick="jump(\'mat\','+x.id+')">'
            +'<span class="tag t-na">'+esc(x.cat||'—')+'</span>'
            +'<div class="line-m">'+esc(x.name)+'</div></div>';}).join('')
        :'<span class="dim">Not linked to any material yet. Open the material and link it '
         +'from there — the material is where a link is made and unmade.</span>')
      +'</div></div>';
  }
  var list=docsOf(m);
  return '<div class="sec">Documents</div><div class="panel">'
    +'<div class="panel-h"><div class="panel-t">Linked to this material</div>'
    +'<button class="btn btn-s no-print" onclick="linkPick('+m.id+')">Link a document</button></div>'
    +'<div class="panel-b">'
    +(list.length?list.map(function(d){
        var st=(d.raw&&d.raw['MAT Status'])||'';
        return '<div class="line">'
          +'<span class="tag t-na" style="min-width:54px;text-align:center">'+esc(d.doc)+'</span>'
          +'<div class="line-m"><div>'+esc(d.name)+'</div>'
          +'<div class="dim mono" style="font-size:12.5px;margin-top:2px">'
          +esc(refOf(d))+(st?(' · '+esc(st)):'')+'</div></div>'
          +'<button class="btn-q" onclick="jump(\'mat\','+d.id+')">Open</button>'
          +'<button class="btn-q" onclick="unlink('+m.id+','+d.id+')">Unlink</button></div>';
      }).join('')
      :'<span class="dim">Nothing linked yet. A method statement or an inspection plan that '
      +'belongs to this material is attached here, and one document can serve many materials.</span>')
    +'</div></div>';
}
function refOf(d){
  if(d.ref)return d.ref;
  var out='';
  REF_FIELDS.some(function(c){
    var v=splitRefs((d.raw||{})[c])[0];
    if(v){out=v;return true;}
    return false;
  });
  return out;
}

window.unlink=function(matId,docId){
  var m=mat(matId);if(!m)return;
  m.docs=(m.docs||[]).filter(function(x){return String(x)!==String(docId);});
  touch();rPane();
};
window.linkPick=function(matId,q){
  var m=mat(matId);if(!m)return;
  var has={};(m.docs||[]).forEach(function(id){has[String(id)]=1;});
  var need=K(q||'');
  var all=(DB.mats||[]).filter(isDoc);
  var rows=all.filter(function(d){
    if(has[String(d.id)])return false;
    if(!need)return K(d.disc||'')===K(m.disc||'');   /* start with its own trade */
    return K(d.name+' '+refOf(d)+' '+(d.doc||'')).indexOf(need)>=0;
  });
  sheet('Link a document to '+m.name,
     '<div class="dim" style="font-size:13.5px;margin-bottom:14px">'
    +(need?('Searching all '+all.length+' documents.')
          :('Showing the '+rows.length+' in '+esc(m.disc||'no discipline')
            +' — search to see the other '+(all.length-rows.length)+'.'))
    +' A document can serve many materials; linking it here does not take it from anywhere else.'
    +'</div>'
    +'<div class="f" style="margin-bottom:14px"><label for="lk">Search by number, title or kind</label>'
    +'<input id="lk" value="'+attr(q||'')+'" autocomplete="off" '
    +'onkeydown="if(event.key===\'Enter\'){event.preventDefault();linkPick('+matId+',this.value);}">'
    +'<span class="dim" style="font-size:12px">Press Enter to search</span></div>'
    +'<div class="panel"><div class="panel-b">'
    +(rows.length?rows.slice(0,60).map(function(d){
        return '<div class="line row-a" onclick="linkAdd('+matId+','+d.id+')">'
          +'<span class="tag t-na" style="min-width:54px;text-align:center">'+esc(d.doc)+'</span>'
          +'<div class="line-m"><div>'+esc(d.name)+'</div>'
          +'<div class="dim mono" style="font-size:12.5px;margin-top:2px">'+esc(refOf(d))+'</div></div>'
          +'</div>';}).join('')
        +(rows.length>60?('<div class="dim" style="font-size:13px;padding-top:10px">and '
          +(rows.length-60)+' more — narrow the search</div>'):'')
      :'<span class="dim">Nothing matches.</span>')
    +'</div></div>');
  var f=document.getElementById('lk');
  if(f){f.focus();f.setSelectionRange(f.value.length,f.value.length);}
};
window.linkAdd=function(matId,docId){
  var m=mat(matId);if(!m)return;
  m.docs=m.docs||[];
  if(m.docs.indexOf(docId)<0&&!m.docs.some(function(x){return String(x)===String(docId);}))
    m.docs.push(docId);
  touch();closeSheet();rPane();
  var d=mat(docId);
  toast('Linked '+(d?d.doc:'the document')+' — it still serves '
    +(d?servedBy(d).length:1)+' material'+((d&&servedBy(d).length!==1)?'s':''));
};

/* ================================================================
   WHO BROUGHT WHOM
   ----------------------------------------------------------------
   A pre-qualification qualifies a company, and the company is often
   the subcontractor rather than the factory: 4MAKA08-...-ME-PRQ-00024
   is Faisal Abdullah Awad Binladen Contracting, and the three
   manufacturers the file hangs off it — GRUNDFOS, PROMINENT/ITC,
   SOLICO — are the makers that subcontractor brought.

   In this project FIRST FIX brought twenty-two manufacturers and not
   one manufacturer was brought by two subcontractors, so this is one
   field and not a list. If a maker later works through somebody else,
   the question this answers is still who brought them first.
   ================================================================ */
function broughtBy(v){
  if(!v||!v.by)return null;
  return (DB.mfrs||[]).filter(function(x){return K(x.name)===K(v.by);})[0]||{name:v.by};
}
function brought(sub){
  return (DB.mfrs||[]).filter(function(v){return v.by&&K(v.by)===K(sub.name);});
}

window.pickBroughtBy=function(id,q){
  var v=mfr(id);if(!v)return;
  var need=K(q||'');
  var subs=(DB.mfrs||[]).filter(function(x){
    return String(x.id)!==String(v.id)&&(!need||K(x.name).indexOf(need)>=0);
  }).sort(function(a,b){
    var A=(a.kind==='sub')?0:1,B=(b.kind==='sub')?0:1;
    return A-B||String(a.name).localeCompare(String(b.name));
  });
  sheet('Who brought '+v.name+'?',
     '<div class="dim" style="font-size:13.5px;margin-bottom:14px">'
    +'The subcontractor that first brought this company onto the project. '
    +'Subcontractors are listed first; anyone on the vendor list can be chosen.</div>'
    +'<div class="f" style="margin-bottom:14px"><label for="bb">Search</label>'
    +'<input id="bb" value="'+attr(q||'')+'" autocomplete="off" '
    +'onkeydown="if(event.key===\'Enter\'){event.preventDefault();pickBroughtBy('+id+',this.value);}">'
    +'</div>'
    +'<div class="panel"><div class="panel-b">'
    +subs.slice(0,60).map(function(x){
        return '<div class="line row-a" onclick="setBroughtBy('+id+','+x.id+')">'
          +'<span class="tag t-na" style="min-width:96px;text-align:center">'
          +esc(KINDS[kindOf(x)].l.toLowerCase())+'</span>'
          +'<div class="line-m">'+esc(x.name)+'</div>'
          +(v.by&&K(v.by)===K(x.name)?'<span class="tag t-wait">current</span>':'')+'</div>';
      }).join('')
    +(subs.length>60?('<div class="dim" style="font-size:13px;padding-top:10px">and '
      +(subs.length-60)+' more — narrow the search</div>'):'')
    +'</div></div>'
    +'<div class="f-act" style="margin-top:18px">'
    +'<button class="btn btn-p" onclick="newSubFor('+id+')">Add a subcontractor</button>'
    +(v.by?'<button class="btn-q" onclick="setBroughtBy('+id+',0)">Nobody — clear it</button>':'')
    +'</div>');
  var f=document.getElementById('bb');
  if(f){f.focus();f.setSelectionRange(f.value.length,f.value.length);}
};
window.setBroughtBy=function(id,subId){
  var v=mfr(id);if(!v)return;
  if(!subId){delete v.by;}
  else{var s=mfr(subId);if(!s)return;v.by=s.name;}
  touch();closeSheet();rPane();rList();
};
window.newSubFor=function(id){
  var box=document.getElementById('sheet-b');if(!box)return;
  box.insertAdjacentHTML('beforeend',
    '<div class="panel" style="margin-top:14px"><div class="panel-b">'
    +'<div class="f"><label for="nsub">New subcontractor</label>'
    +'<input id="nsub" placeholder="company name" autocomplete="off"></div>'
    +'<div class="pop-f"><button class="btn btn-p btn-s" onclick="createSubFor('+id+')">'
    +'Add and use</button></div></div></div>');
  var i=document.getElementById('nsub');
  if(i){i.focus();i.addEventListener('keydown',function(e){
    if(e.key==='Enter'){e.preventDefault();createSubFor(id);}});}
};
window.createSubFor=function(id){
  var n=(document.getElementById('nsub')||{value:''}).value.trim();
  if(!n)return toast('Give it a name');
  var twin=(DB.mfrs||[]).filter(function(x){return K(x.name)===K(n);})[0];
  var s=twin||{id:idMaker()(),name:n,kind:'sub',cat:'',country:'',site:'',scope:'',
    steps:{},pq:{},added:today()};
  if(!twin)DB.mfrs.push(s);
  setBroughtBy(id,s.id);
};

function broughtPanel(v){
  var mine=brought(v);
  var by=broughtBy(v);
  var out='<div class="sec">Who brought them</div><div class="panel"><div class="panel-b">'
    +'<div class="line"><div class="line-m">'
    +(by?('Brought onto the project by <b>'+esc(by.name)+'</b>')
        :'<span class="dim">Nobody recorded. A manufacturer usually comes onto a project '
         +'through a subcontractor, and that is worth knowing when something goes wrong.</span>')
    +'</div>'
    +(by&&by.id?('<button class="btn-q" onclick="jump(\'mfr\','+by.id+')">Open</button>'):'')
    +'<button class="btn btn-s" onclick="pickBroughtBy('+v.id+')">'+(by?'Change':'Set it')
    +'</button></div></div></div>';
  if(mine.length)out+='<div class="sec">Companies this one brought <span class="dim">'
    +mine.length+'</span></div><div class="panel"><div class="panel-b">'
    +mine.map(function(x){
      return '<div class="line row-a" onclick="jump(\'mfr\','+x.id+')">'
        +'<span class="tag t-na" style="min-width:96px;text-align:center">'
        +esc(KINDS[kindOf(x)].l.toLowerCase())+'</span>'
        +'<div class="line-m"><div>'+esc(x.name)+'</div>'
        +(x.country?('<div class="dim" style="font-size:12.5px;margin-top:2px">'
          +esc(x.country)+'</div>'):'')+'</div>'
        +'<span class="meta">'+matsOf(x).length+' material'+(matsOf(x).length===1?'':'s')
        +'</span></div>';}).join('')
    +'</div></div>';
  return out;
}

/* ================================================================
   REPORTS
   ----------------------------------------------------------------
   A tab of its own, because a report is a thing you go to rather
   than a button you happen upon. Excel only — every one of these ends
   up in somebody else's spreadsheet anyway.
   ================================================================ */

/* Documents are records of their own now, so the general log folds them
   back into the material they serve: one row per material, with the
   numbers of its documents in the columns the Main Log keeps them in.
   Exactly the shape the project already reads. */
var FOLD={
  MES:{n:'Method Statement Number',r:'MES Revision',s:'MES Status',d:null},
  ITP:{n:'ITP Number',r:'ITP Revision',s:'ITP Status',d:'ITP Submittal Date'},
  PID:{n:'PID Number',r:'PID Revision',s:'PID Status',d:'PID Submittal Date'},
  MIR:{n:'MIR Number',r:null,s:'MIR Status',d:'MIR Approval Date'},
  WIR:{n:'WIR Number',r:null,s:'WIR Status',d:'WIR Approval Date'},
  PQD:{n:'PQD Number',r:'PQD Revision',s:'PQD Status',d:'PQD Submittal Date'}
};
function foldDocs(m,raw){
  var by={};
  docsOf(m).forEach(function(d){
    var w=FOLD[d.doc];if(!w)return;
    (by[d.doc]=by[d.doc]||[]).push(d);
  });
  Object.keys(by).forEach(function(kind){
    var w=FOLD[kind],list=by[kind];
    /* several documents of one kind go into one cell separated by
       newlines, which is how the original file already holds them */
    function join(col,pick){
      if(!col)return;
      var vals=list.map(pick).filter(function(x){return x!==''&&x!=null;});
      if(vals.length)raw[col]=vals.join('\n');
    }
    join(w.n,function(d){return refOf(d);});
    join(w.s,function(d){return (d.raw||{})[w.s]||(d.raw||{})['MAT Status']||'';});
    join(w.r,function(d){return (d.raw||{})[w.r]||'';});
    join(w.d,function(d){return (d.raw||{})[w.d]||'';});
  });
  return raw;
}

function generalRows(){
  var mats=(DB.mats||[]).filter(function(m){return !isDoc(m);});
  var groups={},order=[];
  mats.forEach(function(m){
    var g=(m.disc||'Uncategorised').trim();
    if(!groups[g]){groups[g]=[];order.push(g);}
    groups[g].push(m);
  });
  var rows=[COLS.slice()];
  order.forEach(function(g,gi){
    if(gi>0){var div=new Array(COLS.length);div[0]=g;div.head=true;rows.push(div);}
    groups[g].forEach(function(m){
      var raw=foldDocs(m,rawOut(m));
      rows.push(COLS.map(function(c){
        var v=raw[c];
        if(v==null||v==='')return '';
        if(DATE_COLS[c]&&/^\d{4}-\d{2}-\d{2}$/.test(String(v)))return {date:String(v)};
        return v;
      }));
    });
  });
  return rows;
}

/* a document attached to nothing would vanish from a report shaped
   one-row-per-material, so it is written to a sheet of its own rather
   than quietly dropped */
function looseRows(){
  var loose=(DB.mats||[]).filter(function(m){
    return isDoc(m)&&servedBy(m).length===0;
  });
  var rows=[['Kind','Reference','Title','Discipline','Revision','Outcome','Note']];
  loose.forEach(function(d){
    var r=d.raw||{};
    rows.push([d.doc,refOf(d),d.name,d.disc||'',
      r['MAT Revision']||r['ITP Revision']||r['MES Revision']||r['PID Revision']||'',
      r['MAT Status']||r['ITP Status']||r['MES Status']||r['PID Status']||r['MIR Status']||'',
      'not linked to any material']);
  });
  return rows;
}

function vendorRows(){
  var rows=[['Vendor','Kind','Brought by','Country','Discipline','PQD Number','PQD Status',
             'PQD Date','Assessment','Materials']];
  (DB.mfrs||[]).forEach(function(v){
    var pq=(v.steps||{}).pqd||{};
    rows.push([v.name,KINDS[kindOf(v)].l,v.by||'',v.country||'',v.scope||v.disc||'',
      pq.ref||'',pq.status||'',pq.date?{date:pq.date}:'',
      (v.steps&&v.steps.pa&&v.steps.pa.status)||'',matsOf(v).length]);
  });
  return rows;
}

function aheadRows(){
  var t0=today(),end=addDays(t0,14),out=[];
  function add(date,what,who){if(date&&date>=t0&&date<=end)out.push([{date:date},what,who||'']);}
  (DB.mats||[]).forEach(function(m){
    if(isDoc(m))return;
    var s=m.steps||{};
    if(s.pfm&&s.pfm.date&&s.pfm.status!=='Approved')add(s.pfm.date,'Pre-fabrication meeting — '+m.name,'');
    if(s.fat&&s.fat.date)add(s.fat.date,'Final inspection or FAT — '+m.name,s.fat.by||'');
    matRoad(m).steps.forEach(function(x){
      var due=x.s.due?x.s.due(m,x.data):null;
      if(due)add(due.on,x.s.n+' — '+due.t+' — '+m.name,'');
    });
    (m.dels||[]).forEach(function(d){
      if(d.status==='Pending')add(d.date,'Delivery to inspect — '+m.name,'');});
  });
  out.sort(function(a,b){return String(a[0].date).localeCompare(String(b[0].date));});
  return [['Date','What','Who']].concat(out);
}

var REPORTS=[
 {k:'log',t:'The general log',
  d:'Every material on one row, with its documents folded back into the columns the project '
    +'already reads — the same seventy-seven columns in the same order. Documents attached to '
    +'nothing are written to a second sheet rather than dropped.',
  go:function(){
    var name=(DB.project||'Project Materials').replace(/[^\w \-]/g,'').trim();
    var sheets=[{name:'Main Log',rows:generalRows(),widths:widths()},
                {name:'Summary',rows:summaryRows()}];
    var loose=looseRows();
    if(loose.length>1)sheets.push({name:'Not linked',rows:loose});
    download(workbook(sheets),name+' — '+today()+'.xlsx');
  }},
 {k:'ven',t:'Vendors and who brought them',
  d:'Every company on the project, what kind it is, its pre-qualification, and the '
    +'subcontractor that first brought it on.',
  go:function(){
    download(workbook([{name:'Vendors',rows:vendorRows()}]),
      (DB.project||'Vendors')+' — vendors '+today()+'.xlsx');
  }},
 {k:'ahead',t:'Two-week look-ahead',
  d:'Everything falling due in the next fourteen days, drawn from the dates the steps '
    +'already carry. Submitted weekly under clause 2.2.6.',
  go:function(){
    download(workbook([{name:'Look-ahead',rows:aheadRows()}]),
      (DB.project||'Look-ahead')+' — look-ahead '+today()+'.xlsx');
  }}
];

function reportsPane(){
  var mats=(DB.mats||[]).filter(function(m){return !isDoc(m);}).length;
  var docs=(DB.mats||[]).length-mats;
  var loose=(DB.mats||[]).filter(function(m){
    return isDoc(m)&&servedBy(m).length===0;}).length;
  return '<div class="head"><div class="wrap"><div class="head-t">Reports</div>'
    +'<div class="head-m">'
    +'<span class="chip flat">'+mats+' materials</span>'
    +'<span class="chip flat">'+docs+' documents</span>'
    +'<span class="chip flat">'+(DB.mfrs||[]).length+' vendors</span>'
    +(loose?('<span class="chip warn">'+loose+' documents linked to nothing</span>'):'')
    +'</div></div></div>'
    +'<div class="body"><div class="wrap">'
    +REPORTS.map(function(r){
      return '<div class="panel"><div class="panel-b">'
        +'<div style="display:flex;gap:18px;align-items:flex-start;flex-wrap:wrap">'
        +'<div style="flex:1;min-width:260px">'
        +'<div class="panel-t">'+esc(r.t)+'</div>'
        +'<div class="swhy" style="margin-top:6px">'+esc(r.d)+'</div></div>'
        +'<button class="btn btn-p" onclick="runReport(\''+r.k+'\')">Download Excel</button>'
        +'</div></div></div>';
    }).join('')
    +'<div class="dim" style="font-size:13px;margin-top:18px;line-height:1.7">'
    +'More of these as you send the shapes you use. Each one comes out as a workbook — '
    +'nothing is printed from here.</div>'
    +'</div></div>';
}
window.runReport=function(k){
  var r=REPORTS.filter(function(x){return x.k===k;})[0];
  if(!r)return;
  try{r.go();toast(r.t+' — downloaded');}
  catch(e){toast('Could not build it — '+(e.message||e));}
};

/* ---------------------------------------------------------------
   10. WHERE THE BUTTONS LIVE
   The page's own menu is left as it is and added to, so this file
   can be removed again by deleting one line.
   --------------------------------------------------------------- */
function install(){
  if(!document.getElementById('xl-file')){
    var i=document.createElement('input');
    i.type='file';i.id='xl-file';i.accept='.xlsx';i.style.display='none';
    i.addEventListener('change',window.excelRead);
    document.body.appendChild(i);
  }
  if(!document.getElementById('xl-reg')){
    var g=document.createElement('input');
    g.type='file';g.id='xl-reg';g.accept='.xlsx';g.style.display='none';
    g.addEventListener('change',window.regRead);
    document.body.appendChild(g);
  }
  var orig=window.showMenu;
  if(!orig||orig.__xl)return;
  window.showMenu=function(){
    orig.apply(this,arguments);
    var b=document.getElementById('sheet-b');
    if(!b)return;
    b.insertAdjacentHTML('beforeend',
      '<div class="sec">The Main Log</div>'
     +'<div class="panel"><div class="panel-b">'
     +'<div class="dim" style="font-size:13.5px;line-height:1.7">'
     +'The whole log, written out with its '+COLS.length+' columns in their original order and read back the same way. '
     +'A column this page has no opinion about is carried through untouched.</div>'
     +'<div class="f-act" style="margin-top:14px">'
     +'<button class="btn btn-p" onclick="excelOut()">Download the workbook</button>'
     +'<button class="btn" onclick="excelPick()">Upload a workbook</button>'
     +'</div></div></div>'
     +'<div class="sec">The Aconex register</div>'
     +'<div class="panel"><div class="panel-b">'
     +'<div class="dim" style="font-size:13.5px;line-height:1.7">'
     +'The register export, read as it comes. One row per document and no history, so every '
     +'reference appears once. What is already in your file is brought up to date; what is not '
     +'is shown and left alone.</div>'
     +'<div class="f-act" style="margin-top:14px">'
     +'<button class="btn btn-p" onclick="regPick()">Upload the register</button>'
     +'<button class="btn" onclick="regReview()">Waiting to be reviewed</button>'
     +'</div></div></div>');
  };
  window.showMenu.__xl=true;
}
/* The records arrive from the database after the page has loaded, so
   labelling them at load time labels nothing. It runs when the workspace
   opens and again after every reload of it — and does nothing at all
   once everything already carries a label. */
function sortOut(){
  try{
    if(labelDocuments()){rList();rPane();}
    else paintTabs();
  }catch(e){}
}
function start(){
  install();
  install2();
  ['enter','refresh','loadAll'].forEach(function(fn){
    var orig=window[fn];
    if(typeof orig!=='function'||orig.__sorted)return;
    window[fn]=function(){
      var r=orig.apply(this,arguments);
      if(r&&typeof r.then==='function')return r.then(function(v){sortOut();return v;});
      sortOut();
      return r;
    };
    window[fn].__sorted=true;
  });
  sortOut();
}
if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',start);
else start();

/* handy from the console, and for anything built on top later */
window.EXCEL={cols:COLS,generalRows:generalRows,looseRows:looseRows,vendorRows:vendorRows,isDoc:isDoc,docsOf:docsOf,servedBy:servedBy,labelDocuments:labelDocuments,createFromRegister:createFromRegister,pending:pending,read:readMainLog,openBook:openBook,readRegister:readRegister,planRegister:planRegister,applyRegister:applyRegister,plan:planFrom,apply:applyPlan,rows:logRows,summary:summaryRows,book:workbook};
})();

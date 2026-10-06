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
/* the host page has these too, but a module that borrows a global it
   did not define breaks quietly the day the global moves. Quotes are
   escaped so a value can never close the attribute it sits in. */
function esc(s){return String(s==null?'':s).replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;')
  .replace(/"/g,'&quot;').replace(/'/g,'&#39;');}
function attr(s){return esc(s);}
/* a string handed to an inline handler: onclick="f('+jsq(x)+')" */
function jsq(s){return esc(JSON.stringify(String(s==null?'':s)));}
/* a cell a spreadsheet would read as a formula is written as text */
function csvCell(c){
  var s=String(c==null?'':c);
  if(/^[=+\-@\t\r]/.test(s)&&!/^[+-]?\d+(\.\d+)?$/.test(s))s="'"+s;
  return '"'+s.replace(/"/g,'""')+'"';
}
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
  if(v==null||v===''||v===0||v==='0')return '';
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
  /* a cell that also holds a later outcome is read by that outcome; a
     terminated line on its own is the weakest word in the cell */
  var best='',rank={'Terminated':0.5,'Rejected':1,'Resubmit':2,'Pending':3,'Approved with comments':4,'Approved':5};
  parts.forEach(function(p){
    var one=oneStatus(p);
    if(one&&(!best||(rank[one]||0)>(rank[best]||0)))best=one;
  });
  return best;
}
function oneStatus(p){
  var k=p.toLowerCase().replace(/[^a-z& ]/g,' ').replace(/\s+/g,' ').trim();
  if(!k)return '';
  /* read as Pending until now, which is how a dead document came to sit
     in the tracker waiting on a review that will never come */
  if(/terminat/.test(k))return 'Terminated';
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
/* red: the row holds something terminated, so every cell of it, empty
   ones too, is shaded — style 3 for text and numbers, 4 for dates */
function cellXml(ref,val,head,red){
  var st=red?' s="3"':'';
  if(val==null||val==='')return red?'<c r="'+ref+'" s="3"/>':'';
  if(head)return '<c r="'+ref+'" t="inlineStr" s="2"><is><t xml:space="preserve">'+xml(val)+'</t></is></c>';
  if(typeof val==='object'&&val.date!=null){
    var s=isoToSerial(val.date);
    if(s!=null)return '<c r="'+ref+'" s="'+(red?4:1)+'"><v>'+s+'</v></c>';
    val=showDate(val.date)||'';
    if(!val)return red?'<c r="'+ref+'" s="3"/>':'';
  }
  if(typeof val==='number'&&isFinite(val))return '<c r="'+ref+'"'+st+'><v>'+val+'</v></c>';
  var t=String(val);
  if(/^-?\d+(\.\d+)?$/.test(t)&&t.length<15&&!/^0\d/.test(t))
    return '<c r="'+ref+'"'+st+'><v>'+t+'</v></c>';
  /* several numbers in one cell, one to a line: the cell wraps, or Excel
     runs them together into one long string */
  if(t.indexOf('\n')>=0)st=' s="'+(red?6:5)+'"';
  return '<c r="'+ref+'" t="inlineStr"'+st+'><is><t xml:space="preserve">'+xml(t)+'</t></is></c>';
}
function isTerminated(v){return typeof v==='string'&&/terminat/i.test(v);}
/* tbl: the sheet is an Excel table. Its heading row is then written as
   shared strings and without a style of its own, so the table's style
   colours it, and the sheet points at its table part. */
/* bands: a row above the table naming each stage, merged across its
   columns and in its colour; the table's heading takes the same colours.
   The table itself starts under it, so a filter still covers every
   column at once. */
var BAND_FIRST=7;               /* the first band's style in styles.xml */
var BAND_COLOURS=['1F4E79','2E75B6','7030A0','C55A11','BF8F00','548235','0E7C7B','8B2252','4472C4','A5473D','5B6770','375623'];
function sheetXml(rows,widths,tbl,bands){
  var off=bands&&bands.length?1:0;
  var colStyle={};
  (bands||[]).forEach(function(b,i){for(var c=b.from;c<=b.to;c++)colStyle[c]=BAND_FIRST+(i%BAND_COLOURS.length);});
  var wide=rows.reduce(function(a,r){return Math.max(a,r.length);},0);
  var bandRow=off?('<row r="1" ht="24" customHeight="1">'+(bands||[]).map(function(b){
      var out='';
      for(var c=b.from;c<=b.to;c++)out+=c===b.from
        ?'<c r="'+colName(c)+'1" t="inlineStr" s="'+colStyle[c]+'"><is><t xml:space="preserve">'+xml(b.t)+'</t></is></c>'
        :'<c r="'+colName(c)+'1" s="'+colStyle[c]+'"/>';
      return out;
    }).join('')+'</row>'):'';
  var merges=(bands||[]).filter(function(b){return b.to>b.from;})
    .map(function(b){return '<mergeCell ref="'+colName(b.from)+'1:'+colName(b.to)+'1"/>';});
  var body=rows.map(function(row,r){
    /* a terminated submittal is shaded red across its whole row, so it is
       seen in a column of approvals without reading every cell */
    var red=r>0&&row.some(isTerminated);
    var line=red?row.concat(new Array(Math.max(0,wide-row.length)).fill('')):row;
    var cells=line.map(function(v,c){
      if(tbl&&r===0)return '<c r="'+colName(c)+(1+off)+'" t="s"'+(colStyle[c]?' s="'+colStyle[c]+'"':'')
        +'><v>'+tbl.sst(tbl.names[c])+'</v></c>';
      return cellXml(colName(c)+(r+1+off),v,r===0||(row.head&&c===0),red);
    }).join('');
    /* a row tall enough for its longest cell, since Excel does not grow
       one written this way by itself */
    var lines=r>0?row.reduce(function(a,v){return typeof v==='string'?Math.max(a,v.split('\n').length):a;},1):1;
    return cells?('<row r="'+(r+1+off)+'"'+(lines>1?' ht="'+(lines*15)+'" customHeight="1"':'')+'>'+cells+'</row>'):'';
  }).join('');
  var cols=widths?('<cols>'+widths.map(function(w,i){
    return '<col min="'+(i+1)+'" max="'+(i+1)+'" width="'+w+'" customWidth="1"/>';}).join('')+'</cols>'):'';
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
   +'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
   +'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
   +'<sheetViews><sheetView workbookViewId="0"><pane ySplit="'+(1+off)+'" topLeftCell="A'+(2+off)+'" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews>'
   +cols+'<sheetData>'+bandRow+body+'</sheetData>'
   +(merges.length?'<mergeCells count="'+merges.length+'">'+merges.join('')+'</mergeCells>':'')
   +(tbl?'<tableParts count="1"><tablePart r:id="rId1"/></tableParts>':'')
   +'</worksheet>';
}
/* An Excel table over a sheet's rows, in Excel's own Table Style
   Medium 2: dark teal heading, banded rows, a filter on every column.
   Its column names are the headings, each made unique as Excel needs. */
function tableXml(n,names,rows,off){
  off=off||0;
  var last=colName(names.length-1)+(Math.max(rows,2)+off);
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    +'<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="'+n+'" '
    +'name="Table'+n+'" displayName="Table'+n+'" ref="A'+(1+off)+':'+last+'" totalsRowShown="0">'
    +'<autoFilter ref="A'+(1+off)+':'+last+'"/>'
    +'<tableColumns count="'+names.length+'">'
    +names.map(function(t,i){return '<tableColumn id="'+(i+1)+'" name="'+xml(t)+'"/>';}).join('')
    +'</tableColumns>'
    +'<tableStyleInfo name="TableStyleMedium2" showFirstColumn="0" showLastColumn="0" '
    +'showRowStripes="1" showColumnStripes="0"/></table>';
}
function tableNames(head){
  var seen={};
  return head.map(function(h,i){
    var t=String(h==null||h===''?'Column '+(i+1):h).replace(/[\r\n]+/g,' ').trim(), b=t, k=2;
    while(seen[t.toLowerCase()])t=b+' '+(k++);
    seen[t.toLowerCase()]=1;return t;
  });
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
      +'<fonts count="4"><font><sz val="11"/><name val="Calibri"/></font>'
      +'<font><b/><sz val="11"/><name val="Calibri"/></font>'
      +'<font><sz val="11"/><color rgb="FF9C0006"/><name val="Calibri"/></font>'
      +'<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font></fonts>'
      +'<fills count="'+(3+BAND_COLOURS.length)+'"><fill><patternFill patternType="none"/></fill>'
      +'<fill><patternFill patternType="gray125"/></fill>'
      +'<fill><patternFill patternType="solid"><fgColor rgb="FFFFC7CE"/><bgColor indexed="64"/></patternFill></fill>'
      +BAND_COLOURS.map(function(c){return '<fill><patternFill patternType="solid"><fgColor rgb="FF'+c+'"/><bgColor indexed="64"/></patternFill></fill>';}).join('')
      +'</fills>'
      +'<borders count="1"><border/></borders>'
      +'<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
      +'<cellXfs count="'+(7+BAND_COLOURS.length)+'">'
      +'<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
      +'<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
      +'<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>'
      +'<xf numFmtId="0" fontId="2" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/>'
      +'<xf numFmtId="164" fontId="2" fillId="2" borderId="0" xfId="0" applyNumberFormat="1" applyFont="1" applyFill="1"/>'
      +'<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><alignment wrapText="1" vertical="top"/></xf>'
      +'<xf numFmtId="0" fontId="2" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment wrapText="1" vertical="top"/></xf>'
      +BAND_COLOURS.map(function(_,i){return '<xf numFmtId="0" fontId="3" fillId="'+(3+i)+'" borderId="0" xfId="0" '
        +'applyFont="1" applyFill="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>';}).join('')
      +'</cellXfs>'
      +'<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
      +'</styleSheet>'}
  ];
  /* the strings a table's heading needs, shared once for the workbook */
  var sst=[], at={};
  function sstIdx(t){if(!(t in at)){at[t]=sst.length;sst.push(t);}return at[t];}
  var nt=0, extraTypes='';
  sheets.forEach(function(s,i){
    var tbl=null;
    if(s.table&&s.rows.length&&s.rows[0].length){
      nt++;
      tbl={sst:sstIdx,names:tableNames(s.rows[0])};
      files.push({name:'xl/tables/table'+nt+'.xml',text:tableXml(nt,tbl.names,s.rows.length,s.bands&&s.bands.length?1:0)});
      files.push({name:'xl/worksheets/_rels/sheet'+(i+1)+'.xml.rels',
        text:'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        +'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        +'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" '
        +'Target="../tables/table'+nt+'.xml"/></Relationships>'});
      extraTypes+='<Override PartName="/xl/tables/table'+nt+'.xml" '
        +'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>';
    }
    files.push({name:'xl/worksheets/sheet'+(i+1)+'.xml',text:sheetXml(s.rows,s.widths,tbl,s.bands)});
  });
  if(sst.length){
    files.push({name:'xl/sharedStrings.xml',text:'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      +'<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="'+sst.length+'" uniqueCount="'+sst.length+'">'
      +sst.map(function(t){return '<si><t xml:space="preserve">'+xml(t)+'</t></si>';}).join('')+'</sst>'});
    extraTypes+='<Override PartName="/xl/sharedStrings.xml" '
      +'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>';
    files[3].text=files[3].text.replace('</Relationships>',
      '<Relationship Id="rId'+(sheets.length+2)+'" '
      +'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>'
      +'</Relationships>');
  }
  if(extraTypes)files[0].text=files[0].text.replace('</Types>',extraTypes+'</Types>');
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
    /* A spreadsheet fills a blank with zero more readily than anyone
       expects — a dragged formula, a cleared cell that kept its format.
       Zero is a date in Excel's reckoning (the thirtieth of December
       1899) and a name in nobody's. Ten cells of this file hold one.
       Everywhere here it means the cell is empty. */
    if(v===0||v==='0')return;
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

/* the vendor, other than this one, that already holds a pre-qualification
   number as its own */
function pqHolder(ref,not){
  var k=K(ref);
  return (DB.mfrs||[]).filter(function(v){
    return v!==not&&K(pqOf(v).ref||'')===k;})[0]||null;
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
  /* the pre-qualification lives on the vendor, so it is written there —
     unless the number already belongs to another company. The log puts
     the subcontractor's pre-qualification on every row of the makers it
     brought; copying it onto each maker gave one number to four vendors.
     The maker is marked as brought by the holder instead. */
  var pq=trim(raw['PQD Number']), st=normStatus(raw['PQD Status']);
  var holder=pq&&pq!=='-'?pqHolder(pq,found):null;
  if(holder){
    if(!found.by&&K(holder.name)!==K(found.name))found.by=holder.name;
    pq='';st='';
  }
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
/* Newest wins, cell by cell. The record's raw row is what the file said
   the last time it was read, so it is the common ancestor of both sides:
   a cell the file still holds as it was has not been edited there, and a
   cell the tracker would now write differently has been edited here. Only
   when both moved, to different values, does the clock decide — the
   file's saved time against the record's last save in the tracker. A
   record with no such stamp predates it, and the file wins as it did. */
function norm(c,v){
  var s=String(v==null?'':v).replace(/\s+/g,' ').trim();
  if(/Status$/.test(c)){var st=normStatus(s);if(st)return st;}
  return s.toLowerCase();
}
function appRow(m){return foldDocs(m,rawOut(m));}
function mergeRow(m,file,fileTime){
  var base=m.raw||{}, app=appRow(m), out={}, diff=[], kept=[], both=[];
  var appTime=Date.parse(m.edited||'')||0, appNewer=!!fileTime&&appTime>fileTime;
  COLS.forEach(function(c){
    var f=file[c], b=base[c], a=app[c];
    var fMoved=norm(c,f)!==norm(c,b), aMoved=norm(c,a)!==norm(c,b), pick=f;
    if(fMoved&&aMoved&&norm(c,f)!==norm(c,a)){
      both.push({col:c,file:f,app:a,won:appNewer?'tracker':'file'});
      if(appNewer)pick=a;
    }else if(!fMoved&&aMoved){pick=a;kept.push(c);}
    if(pick!=null&&pick!=='')out[c]=pick;
    /* a cell the file and the tracker already agree on is not news */
    if(fMoved&&pick===f&&norm(c,f)!==norm(c,a))diff.push(c);
  });
  return {raw:out,diff:diff,kept:kept,both:both};
}
/* A record carrying history the file has no column for — every delivery
   but the newest, a non-conformance, a factory visit — is not deleted by
   a row that merely went missing from a sheet. */
function history(m){
  return (m.dels||[]).length+(m.ncrs||[]).length+(m.visits||[]).length;
}
function planFrom(rows,fileTime){
  var head=rows[0]||[];
  var map={},shift=0;
  /* the header must be the one we know, or the columns land wrong */
  var seen=head.map(trim);
  var miss=COLS.filter(function(c,i){return K(seen[i]||'')!==K(c);});
  var out={add:[],change:[],same:[],gone:[],kept:[],both:[],skipped:0,dividers:0,header:miss.length};
  var byId={},byRef={},refN={};
  (DB.mats||[]).forEach(function(m){
    if(m.raw){
      byId[idOf(m.raw)]=m;
      REF_FIELDS.forEach(function(c){
        splitRefs(m.raw[c]).forEach(function(r){
          if(!byRef[K(r)])byRef[K(r)]=m;
          if(byRef[K(r)]!==m)refN[K(r)]=2;else refN[K(r)]=refN[K(r)]||1;
        });
      });
    }
  });
  var hit={}, hitRec=[];
  /* A row whose MAT number was edited is still the same material if one
     of its other references names exactly one record and nothing else
     has claimed it. A reference shared by several materials — one
     inspection plan covers twelve — proves nothing and is not used. */
  function byOtherRef(raw){
    for(var i=0;i<REF_FIELDS.length;i++){
      var refs=splitRefs(raw[REF_FIELDS[i]]);
      for(var j=0;j<refs.length;j++){
        var m=byRef[K(refs[j])];
        if(m&&refN[K(refs[j])]===1&&hitRec.indexOf(m)<0&&!isDoc(m))return m;
      }
    }
    return null;
  }

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
        return !m.raw&&K(m.name)===K(desc);})[0]
      ||byOtherRef(raw);
    if(!have){out.add.push({raw:raw,row:n+2});return;}
    hit[id]=1;hitRec.push(have);
    var mg=mergeRow(have,raw,fileTime);
    /* The tracker also writes cells it works out for itself — a delivery
       date, a running total — that the file never held. Those are kept
       quietly; only an edit made here after the file was saved is news. */
    if(mg.kept.length&&fileTime&&(Date.parse(have.edited||'')||0)>fileTime)
      out.kept.push({rec:have,row:n+2,cols:mg.kept});
    mg.both.forEach(function(b){out.both.push({rec:have,row:n+2,col:b.col,file:b.file,app:b.app,won:b.won});});
    if(mg.diff.length||mg.both.length)
      out.change.push({raw:mg.raw,row:n+2,rec:have,diff:mg.diff});
    else out.same.push({rec:have});
  });
  (DB.mats||[]).forEach(function(m){
    /* A document has no row in a log shaped one-per-material, so its
       absence from the file means nothing. Counting it as missing would
       put sixteen hundred records under a heading that invites deleting
       them. */
    if(isDoc(m))return;
    if(m.raw&&!hit[idOf(m.raw)]&&hitRec.indexOf(m)<0)out.gone.push(m);
  });
  out.goneSafe=out.gone.filter(function(m){return !history(m);});

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
      if(/^(-+|0+|n\/?a|avl|alv|tbd|tba|na)$/i.test(trim(v)))return;
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
  if(dropGone)DB.mats=DB.mats.filter(function(m){return p.goneSafe.indexOf(m)<0;});
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
  set('Item Description',m.name);
  set('Material Category',m.cat?('Category '+m.cat):'');
  set('Discipline',m.disc||'');
  set('Sub-contractor Name',m.sub||'');
  /* The log is written out only now, and the vendor linked on the
     material is the truth about it: its name, country and
     pre-qualification fill the row whether or not the row had them.
     A record from the register carries none of these columns, so the
     old rule (refresh only what was already there) left them blank. */
  if(v){
    if(v.kind==='sub'){if(!raw['Sub-contractor Name'])set('Sub-contractor Name',v.name);}
    else set('Manufacturer',v.name);
    /* a maker that also installs is the subcontractor too, unless another is named */
    if(v.kind==='makesub'&&!raw['Sub-contractor Name'])set('Sub-contractor Name',v.name);
    if(v.country)set('Country of Origin of Manufacture',v.country);
    var pq=pqOf(v);
    if(pq.ref)set('PQD Number',pq.ref);
    if(pq.rev)set('PQD Revision',pq.rev);
    if(pq.status)set('PQD Status',pq.status);
    if(pq.date)set('PQD Submittal Date',pq.date);
    /* the factory assessment, kept on the vendor */
    var lf=v.logf||{};
    VEN_LOG.forEach(function(f){if(lf[f[0]])set(f[0],lf[f[0]]);});
    /* the factory survey, from the vendor's own step */
    var pa=(v.steps||{}).pa||{};
    if(pa.status==='Scheduled'){if(pa.date&&!raw['PA Tentative Date'])set('PA Tentative Date',pa.date);}
    else if(pa.date)set('PA Date',pa.date);
    if(pa.ref)set('PA Document Number',pa.ref);
    if(pa.status&&pa.status!=='Pending'&&pa.status!=='Scheduled')set('Assessment Result',pa.status);
  }
  step('mts','MAT Number','MAT Submittal Date','MAT Status');
  step('itp','ITP Number','ITP Submittal Date','ITP Status');
  step('pid','PID Number','PID Submittal Date','PID Status');
  if(s.pfm&&s.pfm.date)set('Pre-Fabrication Meeting Date',s.pfm.date);
  if(s.fat){
    if(s.fat.ref)set(raw['FAT/TPI Results']!=null&&raw['FAT/TPI Results']!==''
      ?'FAT/TPI Results':'FAT Package/Procedure Number/ITP',s.fat.ref);
    if(s.fat.date)set('FAT Planned Date',s.fat.date);
    if(s.fat.status&&s.fat.status!=='Pending')
      setStatus('FAT Package Status',/Passed/.test(s.fat.status)?'Approved':s.fat.status);
    if(s.fat.by&&!raw['3rd Party Service Provider Name'])set('3rd Party Service Provider Name',s.fat.by);
  }
  /* deliveries: the newest one fills the inspection-request columns,
     and the running total fills the quantities */
  var dels=(m.dels||[]).slice().sort(function(a,b){return String(b.date||'').localeCompare(String(a.date||''));});
  if(dels.length&&!raw['1st Batch Delivery To Site Actual Date']){
    var first=dels.map(function(d){return d.date;}).filter(Boolean).sort()[0];
    if(first)set('1st Batch Delivery To Site Actual Date',first);
  }
  if(dels.length){
    set('MIR Number',dels[0].ref||'');
    set('MIR Approval Date',dels[0].date||'');
    setStatus('MIR Status',dels[0].status==='Received'?'Approved':(dels[0].status||''));
  }
  var ord=parseFloat(m.qty), got=(m.dels||[]).filter(function(d){
    return ['Received','Approved','Approved with comments'].indexOf(d.status)>=0;})
    .reduce(function(a,d){return a+(parseFloat(d.qty)||0);},0);
  got=Math.round(got*1000)/1000;      /* 113.29, not 113.28999999999999 */
  if(m.qty){set('Total Quantity',m.qty);set('Total Quantity Unit',m.unit||'');}
  if(m.dels&&m.dels.length){
    set('Delivered No',got);set('Delivered Unit',m.unit||'');
    if(isFinite(ord)&&ord>0){
      set('Remaining',Math.round(Math.max(0,ord-got)*1000)/1000);set('Remaining Unit',m.unit||'');
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
/* logRows once wrote the log with a heading row between disciplines,
   the way the file originally arrived. The headings are gone — they are
   a thing a spreadsheet does with sorting — and generalRows is the only
   way the log is written now. */

/* the same figures the dashboard sheet carried, recomputed */
function summaryRows(){
  /* the materials only, as the Main Log sheet beside it has them */
  var mats=(DB.mats||[]).filter(function(m){return !isDoc(m);});
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
  return COLS.map(function(c,i){return i===0?62:Math.min(30,Math.max(12,c.length*0.95));})
    .concat(LOG_EXTRA.map(function(x){return x.w;}));
}
window.excelOut=function(){
  try{
    var name=(DB.project||'Project Materials').replace(/[^\w \-]/g,'').trim();
    download(sevenLog(),name+' \u2014 Material Live Tracking Sheet '+today()+'.xlsx');
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
    PLAN=planFrom(got.rows,f.lastModified);
    PLAN.fileTime=f.lastModified;PLAN.file=f.name;PLAN.sheet=got.name;PLAN.guessed=!!got.guessed;
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

  var keptRecs=p.kept||[];
  body+=block('Kept from the tracker — edited here, untouched in the file',keptRecs,
    '<div class="dim" style="font-size:13.5px;margin-bottom:12px">The file still holds what these cells said '
    +'before, so the edit made in the tracker is the newer one and stays.</div>'
    +keptRecs.slice(0,40).map(function(k){
      return '<div class="line"><span class="tag t-ok">'+k.cols.length+' kept</span>'
        +'<div class="line-m"><div>'+esc(k.rec.name)+'</div>'
        +'<div class="dim" style="font-size:12.5px;margin-top:2px">'+esc(k.cols.slice(0,6).join(', '))
        +(k.cols.length>6?(' and '+(k.cols.length-6)+' more'):'')+'</div></div>'
        +'<span class="meta">row '+k.row+'</span></div>';
    }).join('')+more(keptRecs.length,40));

  var both=p.both||[];
  body+=block('Changed in both — the newer one wins',both,
    '<div class="dim" style="font-size:13.5px;margin-bottom:12px">These cells were changed in the file and in '
    +'the tracker, to different values. The file was saved '
    +esc(p.fileTime?new Date(p.fileTime).toLocaleString('en-GB',{day:'numeric',month:'short',hour:'2-digit',minute:'2-digit'}):'at an unknown time')
    +'; each record keeps whichever side was changed last.</div>'
    +both.slice(0,40).map(function(b){
      return '<div class="line"><span class="tag '+(b.won==='tracker'?'t-ok':'t-wait')+'">'
        +(b.won==='tracker'?'tracker':'file')+'</span>'
        +'<div class="line-m"><div>'+esc(b.rec.name)+' <span class="dim">· '+esc(b.col)+'</span></div>'
        +'<div class="dim" style="font-size:12.5px;margin-top:2px">file: <span class="mono">'+esc(b.file==null?'—':b.file)
        +'</span> · tracker: <span class="mono">'+esc(b.app==null?'—':b.app)+'</span></div></div>'
        +'<span class="meta">row '+b.row+'</span></div>';
    }).join('')+more(both.length,40));

  var guarded=p.gone.length-p.goneSafe.length;
  body+=block('In the tracker but not in the file',p.gone,
    '<div class="dim" style="font-size:13.5px;margin-bottom:12px">These are kept unless you say otherwise. '
    +'A row deleted from the sheet and a row never in it look the same from here, so nothing is removed by accident.'
    +(guarded?(' '+guarded+' of them '+(guarded===1?'carries':'carry')+' deliveries, non-conformances or visits the file has no column for, '
      +'so they are never deleted from here — delete them on the record if that is what you mean.'):'')+'</div>'
    +p.gone.slice(0,40).map(function(m){
      var h=history(m);
      return '<div class="line"><span class="tag '+(h?'t-ok':'t-na')+'">'+(h?'protected':'kept')+'</span>'
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
   +(p.goneSafe.length?('<button class="btn btn-d btn-s" onclick="excelApply(true)">Apply and delete the '
      +p.goneSafe.length+' missing</button>'):'')
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
  if(dropGone&&p.goneSafe.length)msg+=', '+p.goneSafe.length+' deleted';
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
    var st=v.steps||{};
    [st.pqd,st.appr].forEach(function(pq){
      if(pq&&pq.ref)(idx[K(pq.ref)]=idx[K(pq.ref)]||[]).push({v:v,col:'PQD Number',shared:1});
    });
    var pq={};
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
  /* An export with its history shows a document once per revision. Only
     the latest speaks for it — the highest revision, then the latest
     date — or an old revision's outcome could be written over the new. */
  var best={};
  function revKey(r){return /^\d+$/.test(r)?[0,+r]:[1,String(r).toUpperCase()];}
  function newer(a,b){
    var x=revKey(a.rev),y=revKey(b.rev);
    if(x[0]!==y[0])return x[0]>y[0];
    if(x[1]!==y[1])return x[1]>y[1];
    return String(a.date||'')>String(b.date||'');
  }
  out.forEach(function(d){var k=K(d.no);if(!best[k]||newer(d,best[k]))best[k]=d;});
  var kept=out.filter(function(d){return best[K(d.no)]===d;});
  return {rows:kept,sheet:name,about:about,cols:Object.keys(col),revisions:out.length-kept.length};
}

/* ---------------------------------------------------------------
   What the register says against what the tracker holds.
   --------------------------------------------------------------- */
/* Types the tracker does not keep. Work inspection requests are site
   work, and seventeen thousand of them made every page slow; the
   register's rows of that type are passed over, not brought in. */
var SKIP_TYPE={'work inspection request':1};
function planRegister(reg){
  var idx=refIndex();
  var p={moved:[],ended:[],locked:[],same:[],newC23:[],newPlain:[],newC01:[],unknown:0,dead:[]};
  reg.forEach(function(d){
    if(SKIP_TYPE[K(d.type)])return;
    var hits=idx[K(d.no)];
    /* Every document Aconex calls terminated, whatever the tracker holds
       and whichever column says it — listed on its own so it can be read
       as a whole. A document whose Status is Terminated while its Review
       Status still carries a verdict is the case worth a second look. */
    if(verdictOf(d.status)==='Terminated'||verdictOf(d.review)==='Terminated'){
      var h0=(hits||[])[0], w0=h0?STATUS_OF[h0.col]:null;
      p.dead.push({d:d,h:h0,
        now:!h0?'':h0.del?trim(h0.del.status):h0.m?trim((h0.m.raw||{})[w0?w0.s:'']):
          trim(((h0.v.steps||{})[(h0.v.kind==='agency')?'appr':'pqd']||{}).status)});
    }
    /* Review Status is the reviewer's word and Status the document's own;
       where they disagree the review is the later of the two — except
       when the document itself has been terminated. A dead document is
       dead whatever its review once said, and is recorded as such. */
    var want=verdictOf(d.status)==='Terminated'?'Terminated'
      :(verdictOf(d.review)||verdictOf(d.status));
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
      var vslot=h.v?((h.v.kind==='agency')?'appr':'pqd'):'';
      var now=h.del?trim(h.del.status)
             :h.m?trim((h.m.raw||{})[where?where.s:''])
                 :trim(h.second?(((h.v.pq2||[]).filter(function(x){
                     return K(x.ref)===K(d.no);})[0]||{}).status)
                   :(((h.v.steps||{})[vslot]||{}).status));
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
  /* Aconex says whose each pre-qualification is: its title names the
     company. Where one number sits on several vendors, the one the title
     names — or the one made from this very document — keeps it, and the
     others are listed to have it taken off. One number, one company. */
  p.wrongPq=[];
  reg.forEach(function(d){
    if(K(d.type)!=='pre-qualification')return;          /* by Type only */
    var holders=(DB.mfrs||[]).filter(function(v){return K(pqOf(v).ref||'')===K(d.no);});
    if(holders.length<2)return;
    var co=coName(companyOf(d.title));
    var owner=holders.filter(function(v){return K(v.scope||'')===K(d.title);})[0]
      ||holders.filter(function(v){var n=coName(v.name);return n&&co&&(n===co||co.indexOf(n)===0||n.indexOf(co)===0);})[0]
      ||pqGuess({list:holders});
    p.wrongPq.push({d:d,owner:owner,others:holders.filter(function(v){return v!==owner;})});
  });
  return p;
}

/* An upload is named by its file and the second it was applied; two in
   the same minute used to share a name and be taken back together. */
function regTag(p){
  return (p.file||'a register')+' · '+new Date().toLocaleString('en-GB',
    {day:'2-digit',month:'short',hour:'2-digit',minute:'2-digit',second:'2-digit'});
}
/* What an upload changed on a record that was already here is written
   down on that record — before and after — so taking the upload back
   can put it the way it was. A cell changed again since is left alone:
   that edit is newer than the upload and is not the upload's to undo.
   Only the last few uploads are remembered per record. */
var REG_KEEP=5;
function clone(x){return x==null?null:JSON.parse(JSON.stringify(x));}
function regState(owner,k){
  var s=k.split('|');
  if(s[0]==='v')return clone((owner.steps||{})[s[1]]);
  if(s[0]==='d'){
    var del=(owner.dels||[]).filter(function(x){return String(x.id)===s[1];})[0];
    return del?{status:del.status==null?null:del.status,date:del.date==null?null:del.date}:null;
  }
  var out={raw:{},step:s[2]?clone((owner.steps||{})[s[2]]):null};
  s[1].split(',').filter(Boolean).forEach(function(c){
    var v=(owner.raw||{})[c];out.raw[c]=v==null?null:v;});
  return out;
}
function regRestore(owner,k,was){
  var s=k.split('|');
  function put(obj,key,v){if(v==null)delete obj[key];else obj[key]=v;}
  if(s[0]==='v'){owner.steps=owner.steps||{};put(owner.steps,s[1],clone(was));return;}
  if(s[0]==='d'){
    var del=(owner.dels||[]).filter(function(x){return String(x.id)===s[1];})[0];
    if(del&&was){put(del,'status',was.status);put(del,'date',was.date);}
    return;
  }
  owner.raw=owner.raw||{};
  Object.keys(was.raw).forEach(function(c){put(owner.raw,c,was.raw[c]);});
  if(s[2]){owner.steps=owner.steps||{};put(owner.steps,s[2],clone(was.step));}
}
function regNote(owner,tag,k,before){
  var after=regState(owner,k);
  if(JSON.stringify(after)===JSON.stringify(before))return;
  owner.regWas=owner.regWas||{};
  (owner.regWas[tag]=owner.regWas[tag]||[]).push({k:k,was:before,set:after});
  var tags=Object.keys(owner.regWas);
  while(tags.length>REG_KEEP)delete owner.regWas[tags.shift()];
}
function regChanged(tag){
  var n=0;
  (DB.mats||[]).concat(DB.mfrs||[]).forEach(function(r){
    if(r.regWas&&r.regWas[tag])n+=r.regWas[tag].length;});
  return n;
}
function regUnchange(tag){
  var back=0,kept=0;
  (DB.mats||[]).concat(DB.mfrs||[]).forEach(function(r){
    var list=r.regWas&&r.regWas[tag];if(!list)return;
    list.slice().reverse().forEach(function(ch){
      if(JSON.stringify(regState(r,ch.k))===JSON.stringify(ch.set)){regRestore(r,ch.k,ch.was);back++;}
      else kept++;
    });
    delete r.regWas[tag];
    if(!Object.keys(r.regWas).length)delete r.regWas;
  });
  return {back:back,kept:kept};
}
var applyRegisterBare=function(){};
function applyRegister(p,alsoEnded,tag){
  var n=applyRegisterBare(p,alsoEnded,tag);
  liftDocSteps();
  return n;
}
applyRegisterBare=function(p,alsoEnded,tag){
  if(p.rows)stampDates(p.rows);
  var n=0;
  function write(it){
    var owner=it.h.v||it.h.m, k;
    if(it.h.v)k='v|'+((it.h.v.kind==='agency')?'appr':'pqd');
    else if(it.h.del)k='d|'+it.h.del.id;
    else k='m|'+[it.where.s,it.where.r,it.where.d].filter(Boolean).join(',')+'|'+(it.where.step||'');
    var before=owner&&tag?regState(owner,k):null;
    write1(it);
    if(owner&&tag)regNote(owner,tag,k,before);
  }
  function write1(it){
    var w=it.where,d=it.d;
    if(it.h.v){                                   /* a vendor's qualification */
      var v=it.h.v;
      v.steps=v.steps||{};
      var slot=(v.kind==='agency')?'appr':'pqd';
      var pq=v.steps[slot]||{};
      pq.status=it.want==='Approved as Noted'?'Approved with comments':it.want;
      if(d.date)pq.date=d.date;
      if(!pq.ref)pq.ref=d.no;
      v.steps[slot]=pq;n++;return;
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
  /* a number on a vendor that Aconex says is someone else's comes off it;
     the vendor is marked as brought by the holder if nothing else is said */
  (p.wrongPq||[]).forEach(function(w){
    w.others.forEach(function(v){
      var k='v|'+((v.kind==='agency')?'appr':'pqd');
      var before=tag?regState(v,k):null;
      if(v.steps)delete v.steps[(v.kind==='agency')?'appr':'pqd'];
      if(!v.by)v.by=w.owner.name;
      if(tag)regNote(v,tag,k,before);
      n++;
    });
  });
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
    REG.rows=reg.rows;
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
/* a terminated document: what Aconex says in both its columns, and
   what the tracker holds for it, if anything */
function deadRow(it){
  var d=it.d, h=it.h, rec=h?(h.m||h.v):null;
  var go=h?(h.m?("jump('mat',"+h.m.id+")"):("jump('mfr',"+h.v.id+")")):'';
  var rv=trim(d.review);
  return '<div class="line'+(go?' row-a':'')+'"'+(go?(' onclick="'+go+'"'):'')+'>'
    +'<span class="tag t-bad" style="min-width:118px;text-align:center;flex-shrink:0">'
    +esc(rec?(it.now||'nothing yet'):'not in the tracker')+'</span>'
    +'<div class="line-m"><div>'+esc(rec?rec.name:(d.title||d.no))+'</div>'
    +'<div class="dim" style="font-size:12.5px;margin-top:2px"><span class="mono">'+esc(d.no)+'</span>'
    +(d.rev!==''?(' · rev '+esc(d.rev)):'')
    +' · Status: '+esc(d.status||'—')+' · Review: '+esc(rv||'—')+'</div>'
    +'</div></div>';
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
   +(p.dead.length?('<div class="stat" style="cursor:pointer" onclick="document.getElementById(\'reg-dead\').scrollIntoView({behavior:\'smooth\'})">'
     +'<div class="stat-v">'+p.dead.length+'</div><div class="stat-l">Terminated in Aconex — see all</div></div>'):'')
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
  var wq=p.wrongPq||[];
  if(wq.length)body+='<div class="sec">One pre-qualification number on several vendors <span class="dim">'+wq.length+'</span></div>'
    +'<div class="panel"><div class="panel-b">'
    +'<div class="dim" style="font-size:12.5px;margin-bottom:10px">Aconex names the company each number belongs to. '
    +'On applying, the number stays with that company and comes off the others, which are marked as brought by it '
    +'where nothing else is recorded.</div>'
    +wq.slice(0,60).map(function(w){
      return '<div class="line"><span class="tag t-ok" style="min-width:118px;text-align:center;flex-shrink:0">keeps it</span>'
        +'<div class="line-m"><div><b>'+esc(w.owner.name)+'</b> <span class="mono dim">'+esc(w.d.no)+'</span></div>'
        +'<div class="dim" style="font-size:12.5px;margin-top:2px">comes off: '
        +esc(w.others.map(function(v){return v.name;}).join(', '))+'</div></div></div>';
    }).join('')+more(wq.length,60)+'</div></div>';
  if(p.dead.length)body+='<div id="reg-dead"></div>'+regBlock('Terminated in Aconex — all of them, whatever the tracker holds',
    p.dead,deadRow,p.dead.length);

  var fresh=p.newC23.length+p.newPlain.length+p.newC01.length;
  body+='<div class="f-act" style="margin-top:22px">'
   +((p.moved.length||fresh||wq.length)?('<button class="btn btn-p" onclick="regAll()">Do it all — '
      +p.moved.length+' updated, '+fresh+' brought in'+(wq.length?(', '+wq.length+' number'+(wq.length===1?'':'s')+' put right'):'')+'</button>'):'')
   +((p.moved.length||wq.length)?('<button class="btn" onclick="regApply(false)">'
      +(p.moved.length?('Only the '+p.moved.length+' that moved'):'')
      +(p.moved.length&&wq.length?' and ':'')
      +(wq.length?((p.moved.length?'the ':'Only put right the ')+wq.length+' shared number'+(wq.length===1?'':'s')):'')+'</button>'):'')
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
  var n=applyRegister(REG,alsoEnded,regTag(REG));
  closeSheet();
  toast(n+' document'+(n===1?'':'s')+' brought up to date — More → Waiting to be reviewed can take it back');
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
    +((p.wrongPq||[]).length?('<br><b>3.</b> Put right '+p.wrongPq.length+' pre-qualification number'
      +(p.wrongPq.length===1?'':'s')+' that sit on more than one vendor, keeping each on the company Aconex names.'):'')
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
  var tag=regTag(p);
  var n=applyRegister(p,true,tag);
  var made=createFromRegister(p,tag);
  closeSheet();
  var nl=looseMir().length;
  toast(n+' updated, '+made.mats+' materials and '+made.vendors+' vendors brought in'
    +(nl?(' · '+nl+' MIR not linked to a material'):''));
  REG=null;
  setTimeout(regReview,600);
};
window.regAddAll=function(){
  if(!REG)return;
  var p=REG;
  var tag=regTag(p);
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
        return csvCell(c);}).join(','));
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
 'material inspection request': {n:'MIR Number',d:'MIR Approval Date',r:null,s:'MIR Status'},
 /* the inspection of work done on site; its number keeps its own column,
    so the next upload finds it and brings its outcome up to date */
 'work inspection request':     {n:'WIR Number',d:'WIR Approval Date',r:null,s:'WIR Status'}
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
  /* the words that introduce a title rather than belong to it. The
     spelling of "prequalification" varies — Prequalificatoin is in this
     register — so the stem is matched rather than the word. */
  t=t.replace(/^\s*P4\s*-?\s*/i,' ')
     .replace(/\bmakkah\b/ig,' ')
     .replace(/\bpre[- ]?qualificat\w*\b/ig,' ')
     .replace(/\bprequalificat\w*\b/ig,' ')
     .replace(/\bPRQ\b/ig,' ')
     .replace(/\bsubmittal\b/ig,' ')
     .replace(/^[\s\-–:]*\bfor\b/i,' ')
     .replace(/\b(supplier|sub[- ]?contractor|subcontractor|manufacturer|vendor)\b\s*[-:]?/ig,' ')
     .replace(/^[\s\-:–,]+/,'');

  /* A name ends where the work begins. These are the words this project
     uses to start describing scope, and none of them has ever been part
     of a company's name here. */
  var scope=new RegExp('\\s(?:supply|supplies|supplying|subcontractor|installation|'
    +'install|works|work\\b|for\\s+[a-z]|category|all zones|provisional|lump sum|'
    +'prime cost)','i');

  var cut=t.split(/\s[-–]\s|\s*[-–]\s|\(|,|\u2013/)[0];
  var m=scope.exec(cut);
  if(m&&m.index>2)cut=cut.slice(0,m.index);
  cut=cut.split(/\b(?:category|all zones|rev\.?\s*\d)/i)[0];
  /* "Al Namlah factory. Supply for…" ends at the full stop, but "A. R."
     does not — a stop only ends it when a space and a capital follow and
     the word before it is longer than an initial. */
  cut=cut.replace(/([A-Za-z]{3,})\.\s+[A-Z].*$/,'$1');
  cut=trim(cut).replace(/^[\s\-–:.,]+/,'').replace(/[\s\-–:.,]+$/,'')
       .replace(/\s{2,}/g,' ');
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
/* What a register row is, by its Type column and nothing else — not the
   code in its number, not its title. A material submittal is a material,
   a pre-qualification a vendor, and every other type a document of that
   kind; a type not known here is kept as a document under its own name,
   so nothing can arrive as a material by mistake. */
var TYPE_KIND={'material submittal':'','pre-qualification':'PQD','method statement':'MES',
  'inspection & test plan':'ITP','material inspection request':'MIR','work inspection request':'WIR',
  'material sample':'MAS','personnel approval form':'PAA'};
function kindOfType(type){
  var t=K(type);
  if(!t)return 'Other';
  return (t in TYPE_KIND)?TYPE_KIND[t]:trim(type);
}
/* the old guess from a number's code — only for records made before
   the register's Type was read (see labelDocuments) */
function docKind(ref,type){
  var t=K(type);
  if(t==='material submittal')return '';
  if(t==='pre-qualification')return 'PQD';
  if(t==='method statement')return 'MES';
  if(t==='inspection & test plan')return 'ITP';
  if(t==='material inspection request')return 'MIR';
  /* in this project's register as well: material samples, and the forms
     that put the project's own staff up for approval (BIM, HSSE and the
     like) — documents, neither of them a material */
  if(t==='work inspection request')return 'WIR';
  if(t==='material sample')return 'MAS';
  if(t==='personnel approval form')return 'PAA';
  var r=String(ref||'').toUpperCase();
  if(/-MAS-/.test(r))return 'MAS';
  if(/-PAA-/.test(r))return 'PAA';
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
  var w=NEW_AS[K(d.type)];                    /* by its Type column only */
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

function createFromRegister(p,tag,opts){
  opts=opts||{};
  var made={n:1,vendors:[],id:idMaker()},mats=0,vends=0;
  var all=p.newC23.concat(p.newPlain,p.newC01);
  /* opts.fresh (a project built from the register alone, as it was once
     on 29 Sep 2026): every pre-qualification is a vendor of its own and
     nothing is marked for review. Terminated documents come in either
     way, carrying Terminated as their status. */
  all.forEach(function(d){
    if(K(d.type)==='pre-qualification'){
      var name=companyOf(d.title);
      var twin=opts.fresh?null:(DB.mfrs||[]).filter(function(v){return K(v.name)===K(name);})[0];
      var v=twin||{id:made.id(),name:name,kind:kindOfTitle(d.title),cat:d.cat||'',
        country:'',site:'',scope:'',steps:{},pq:{},added:today()};
      v.scope=v.scope||d.title;
      v.steps=v.steps||{};
      /* An agency is not pre-qualified, it is approved — clause 2.2.17 —
         and its road has no pre-qualification step at all. Writing the
         reference there left the record looking untouched, with the
         number sitting in a field nothing reads. */
      var slot=(v.kind==='agency')?'appr':'pqd';
      if(!(v.steps[slot]&&v.steps[slot].ref)){
        v.steps[slot]={ref:d.no,date:d.date||'',rev:d.rev||'',
          status:d.want==='Approved as Noted'?'Approved with comments':(d.want||'Pending')};
      }else if(K(v.steps[slot].ref)!==K(d.no)){
        /* the same company qualified twice. The page holds one
           qualification per vendor, so the second is kept beside it
           rather than dropped — losing it would make this document
           look new again on the next upload, for ever. */
        v.pq2=v.pq2||[];
        if(!v.pq2.some(function(x){return K(x.ref)===K(d.no);}))
          v.pq2.push({ref:d.no,date:d.date||'',status:d.want||'',title:d.title||''});
      }
      if(!twin){v.reg=tag;if(!opts.fresh)v.review=1;DB.mfrs.push(v);vends++;}
      return;
    }
    var m={id:made.id(),name:'',cat:'',ref:'',mfr:'',qty:'',unit:'',
      steps:{},dels:[],ncrs:[],added:today()};
    applyRaw(m,rawFromDoc(d),made);
    m.reg=tag;if(!opts.fresh)m.review=1;
    /* A record is placed by its Type alone: a material submittal is a
       material, and nothing in its title links it to a vendor — which
       vendor supplies it is set by hand. m.doc is always written, '' for
       a material, so nothing later re-guesses it from its number. */
    m.doc=kindOfType(d.type);
    if(d.date)m.acxDate=d.date;
    DB.mats.push(m);mats++;
  });
  touch();rList();rPane();
  return {mats:mats,vendors:vends};
}
/* Each record keeps the date Aconex gives its latest revision, so the
   calendar can put everything from the register on its day — method
   statements included, which have no date column of their own. A
   material's row also holds the numbers of documents folded into it;
   only its own number dates it. */
function stampDates(rows){
  var idx=refIndex(),n=0;
  rows.forEach(function(d){
    if(!d.date&&!d.rev)return;
    (idx[K(d.no)]||[]).forEach(function(h){
      /* a vendor's pre-qualification keeps its revision and the date of
         that revision, so a resubmission can be told from a first one */
      if(h.v&&!h.second){
        var pq=(h.v.steps||{})[(h.v.kind==='agency')?'appr':'pqd'];
        if(pq&&K(pq.ref)===K(d.no)){
          if(d.rev&&pq.rev!==d.rev){pq.rev=d.rev;n++;}
          if(d.date&&pq.date!==d.date){pq.date=d.date;n++;}
        }
        return;
      }
      if(!d.date)return;
      if(!h.m||h.del||h.ncr)return;
      if(!isDoc(h.m)&&h.col!=='MAT Number')return;
      if(h.m.acxDate!==d.date){h.m.acxDate=d.date;n++;}
    });
  });
  return n;
}
/* the date a record went in to Aconex: stamped by the register, or the
   kind's own date column for a record made before the stamp */
function acxDateOf(m){
  if(m.acxDate)return m.acxDate;
  var raw=m.raw||{};
  var v=raw[{MIR:'MIR Approval Date',ITP:'ITP Submittal Date'}[m.doc]||'MAT Submittal Date'];
  return /^\d{4}-\d{2}-\d{2}$/.test(String(v||''))?String(v):'';
}
/* everything the register brought in, on the calendar */
window.acxCal=function(put){
  /* a vendor's tentative physical assessment, on its day */
  (DB.mfrs||[]).forEach(function(v){
    var t=(v.logf||{})['PA Tentative Date'];
    if(t)put(t,'PA (tentative) \u00b7 '+v.name,'wait',"jump('mfr',"+v.id+")",{c:'visit',kind:'PA',name:v.name,st:'Tentative'});
  });
  (DB.mats||[]).forEach(function(m){
    var on=acxDateOf(m);if(!on)return;
    /* a request linked to a material shows there, as its consignment */
    if(m.doc==='MIR'&&isLinked(m))return;
    var mat=!isDoc(m);
    var st=mat?(stepOf(m,'mts','status')||(m.raw||{})['MAT Status']||''):rawEnd(m,'Status');
    /* the number from its kind code on (MAT-00010), which says the kind too */
    var no=refOf(m).replace(/^.*?-(?=[A-Z]{3}-)/,'');
    put(on,(no||(mat?'MAT':m.doc))+' \u00b7 '+m.name
      +(st?(' \u00b7 '+st):''),statusTone(st)||'na',"jump('mat',"+m.id+")",
      {c:mat?'amat':m.doc==='MIR'?'amir':'adoc',kind:mat?'MAT':m.doc,name:m.name,no:refOf(m),st:st,
       disc:m.disc||'',cat:m.cat||'',docId:mat?null:m.id,
       unlinked:!mat&&!NOT_FOR_MAT[m.doc]&&!SITE_KIND[m.doc]&&!m.site&&!isLinked(m),
       linkedTo:(!mat&&isLinked(m))?servedBy(m).map(function(x){return x.name;}).join(', '):''});
  });
  (DB.mfrs||[]).forEach(function(v){
    var pq=pqOf(v);if(!pq.date)return;
    put(pq.date,'PQD \u00b7 '+v.name+(pq.status?(' \u00b7 '+pq.status):''),statusTone(pq.status)||'na',
      "jump('mfr',"+v.id+")",{c:'apqd',kind:'PQD',name:v.name,no:pq.ref||'',st:pq.status||''});
  });
};

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
/* every upload that can still be taken back: what it brought in and
   has not been reviewed, and what it changed on records already here */
function batches(){
  var b={};
  function at(t){return b[t]=b[t]||{n:0,s:0};}
  pending().forEach(function(x){at(x.r.reg||'(unknown upload)').n++;});
  (DB.mats||[]).concat(DB.mfrs||[]).forEach(function(r){
    Object.keys(r.regWas||{}).forEach(function(t){at(t).s+=r.regWas[t].length;});
  });
  return b;
}
function batchLabel(x){
  return [x.n?(x.n+' new'):'',x.s?(x.s+' status'+(x.s===1?'':'es')):''].filter(Boolean).join(' · ');
}
window.regReview=function(){
  var list=pending(), b=batches();
  if(!list.length&&!Object.keys(b).length)return sheet('Nothing waiting',
    '<div class="dim" style="padding:26px 0;text-align:center">'
    +'Nothing is waiting to be reviewed. Records brought in from a register appear here '
    +'until you have been through them.</div>');
  var cap=60;
  sheet('Waiting to be reviewed',
    '<div class="dim" style="font-size:13.5px;margin-bottom:16px">'
    +list.length+' record'+(list.length===1?'':'s')+' came in from a register and have not '
    +'been looked at. They sit in Materials and Vendors like any other, and this list is '
    +'only a way of finding them again.</div>'
    +'<div class="panel"><div class="panel-b">'
    +Object.keys(b).map(function(t){
      return '<div class="line"><span class="tag t-na">'+esc(batchLabel(b[t]))+'</span>'
        +'<div class="line-m">'+esc(t)+'</div>'
        +'<button class="btn btn-s btn-d" onclick="regUndo('+jsq(t)+')">Take this upload back</button>'
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
  var s=regChanged(tag);
  sheet('Take back this upload?',
    '<div style="font-size:14px;line-height:1.75"><b>'+esc(tag)+'</b><br><br>'
    +(n?(n+' record'+(n===1?'':'s')+' that came in from it and '+(n===1?'has':'have')
      +' not been reviewed will be deleted; anything already marked reviewed stays. '):'')
    +(s?(s+' status'+(s===1?'':'es')+' it changed on records that were already here '
      +(s===1?'goes':'go')+' back to what '+(s===1?'it was':'they were')
      +' — except any changed again since, which are newer than the upload and stay.'):'')
    +'</div>'
    +'<div class="f-act" style="margin-top:20px">'
    +'<button class="btn btn-d" onclick="regUndoYes('+jsq(tag)+')">Take it back</button>'
    +'<button class="btn-q" onclick="regReview()">Keep them</button></div>');
};
window.regUndoYes=function(tag){
  function drop(r){return r.review&&(r.reg||'(unknown upload)')===tag;}
  var before=(DB.mats||[]).length+(DB.mfrs||[]).length;
  DB.mats=(DB.mats||[]).filter(function(m){return !drop(m);});
  DB.mfrs=(DB.mfrs||[]).filter(function(v){return !drop(v);});
  var gone=before-(DB.mats.length+DB.mfrs.length);
  var st=regUnchange(tag);
  if(SEL.mat&&!mat(SEL.mat))SEL.mat=DB.mats.length?DB.mats[0].id:null;
  if(SEL.mfr&&!mfr(SEL.mfr))SEL.mfr=DB.mfrs.length?DB.mfrs[0].id:null;
  touch();rList();rPane();closeSheet();
  toast([gone?(gone+' record'+(gone===1?'':'s')+' taken back'):'',
    st.back?(st.back+' status'+(st.back===1?'':'es')+' restored'):'',
    st.kept?(st.kept+' changed since, left as they are'):''].filter(Boolean).join(', ')||'Nothing to take back');
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
var DOC_KINDS=['MES','ITP','MAS','PAA','PID','PQD','Report','Procedure','WIR'];

function isDoc(m){return !!(m&&m.doc);}
/* Documents is the work still to do: what serves a material and has not
   been linked to one. Pre-qualification belongs to a vendor and a PAA to
   an inspector, so neither ever waits there. Every kind has its own page
   under Other, linked or not, to look things up in. */
var NOT_FOR_MAT={PQD:1,PAA:1};
/* work on site by default: a WIR that belongs to a material is linked
   from the material, and the rest never wait in Documents */
var SITE_KIND={WIR:1};
var OTHER_PAGES=[
  ['itp','ITP','Inspection & Test Plan'],
  ['mes','MES','Method Statement'],
  ['wir','WIR','Work Inspection Request'],
  ['mas','MAS','Material Sample'],
  ['paa','PAA','Personnel Approval (PAA)'],
  ['pqd','PQD','Pre-qualification (PQD)']];
var OTHER_KIND={};OTHER_PAGES.forEach(function(p){OTHER_KIND[p[0]]=p[1];});
function isOtherView(v){return !!OTHER_KIND[v]||v==='odoc'||v==='alldoc'||v==='allmir';}
/* the kinds without a page of their own */
function oddKind(m){
  return isDoc(m)&&m.doc!=='MIR'&&!OTHER_PAGES.some(function(p){return p[1]===m.doc;});
}
/* Which materials link each document, worked out once for the whole
   record set and again only after an edit. With twenty thousand
   documents, asking every material about every row was the slow part
   of a page. While a page draws from a narrowed set, the full one is
   still what is read. */
var MATS_ALL=null, LINKIDX=null;
function linkIndex(){
  var all=MATS_ALL||DB.mats||[];
  var key=(typeof EDITS!=='undefined'?EDITS:0)+'|'+all.length;
  if(LINKIDX&&LINKIDX.arr===all&&LINKIDX.key===key)return LINKIDX;
  var by={};
  all.forEach(function(m){
    (m.docs||[]).forEach(function(id){var k=String(id);(by[k]=by[k]||[]).push(m);});
  });
  LINKIDX={arr:all,key:key,by:by};
  return LINKIDX;
}
function isLinked(d){return !!linkIndex().by[String(d.id)];}
function waitsForLink(m){
  return isDoc(m)&&m.doc!=='MIR'&&!NOT_FOR_MAT[m.doc]&&!SITE_KIND[m.doc]&&!m.site&&!isLinked(m);
}
/* the user's own word that a document is work on site, for no material:
   it leaves Documents for its kind's page, and can be taken back */
window.markSite=function(id,on){
  var d=mat(id);if(!d)return;
  if(on)d.site=true;else delete d.site;
  touch();rList();rPane();
  toast(on?('Marked as site work \u2014 it is on the '+d.doc+' page now'):'Back in Documents');
};
function viewHas(v,m){
  if(v==='mat'||v==='ipi'||v==='fat'||v==='irn')return !isDoc(m);
  if(v==='mir')return m.doc==='MIR'&&!isLinked(m);
  if(v==='allmir')return m.doc==='MIR';
  if(v==='doc')return waitsForLink(m)&&(!DOCKIND||m.doc===DOCKIND);
  if(OTHER_KIND[v])return m.doc===OTHER_KIND[v];
  if(v==='odoc')return oddKind(m);
  if(v==='alldoc')return isDoc(m)&&m.doc!=='MIR';
  return false;
}
function inView(m){return viewHas(VIEW,m);}
/* the page a record opened from elsewhere belongs on */
function viewFor(m){
  if(viewHas(VIEW,m))return VIEW;
  if(!isDoc(m))return 'mat';
  if(m.doc==='MIR')return isLinked(m)?'allmir':'mir';
  if(waitsForLink(m))return 'doc';
  var p=OTHER_PAGES.filter(function(x){return x[1]===m.doc;})[0];
  return p?p[0]:'odoc';
}
/* A material's page is drawn from a set narrowed to materials, so its
   documents are looked up in the whole record set — or a link just
   made saved, and the page still said nothing was linked. */
function docsOf(m){
  var ids=(m&&m.docs)||[];
  var all=MATS_ALL||DB.mats||[];
  return ids.map(function(id){
    return all.filter(function(x){return String(x.id)===String(id);})[0];
  }).filter(Boolean);
}
/* inspection requests that no material has linked yet */
function looseMir(){
  return (DB.mats||[]).filter(function(m){return m.doc==='MIR'&&!isLinked(m);});
}
window.showLooseMir=function(){setTab('mir');};
function servedBy(doc){return (linkIndex().by[String(doc.id)]||[]).slice();}

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
  [['mir','Material Inspection Request'],['doc','Documents'],
   ['ipi','In-Process Inspection'],['fat','FAT/Final Inspection'],['irn','Inspection Release Note']].forEach(function(pair){
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
  /* Other: the pages looked things up in rather than worked through,
     behind one tab so the row stays short */
  var ob=document.createElement('button');
  ob.className='tab';ob.id='tab-other';ob.setAttribute('role','tab');
  ob.setAttribute('aria-selected','false');ob.setAttribute('aria-haspopup','menu');
  ob.textContent='Other \u25BE';
  ob.onclick=function(e){e.stopPropagation();otherMenu(ob);};
  tabs.appendChild(ob);
  /* No Table tab: each list page is its own table now — vendors on
     Vendors, materials on Materials — so a second copy of them is gone. */
  [['avl','AVL'],['rep','Reports']].forEach(function(p){
    var b=document.createElement('button');
    b.className='tab';b.id='tab-'+p[0];b.setAttribute('role','tab');
    b.setAttribute('aria-selected','false');
    b.appendChild(document.createTextNode(p[1]));
    b.onclick=function(){setTab(p[0]);};
    tabs.appendChild(b);
  });
  tableCSS();barCSS();
}

/* does a material's road include this step (by its category)? */
function stepApplies(m,k){
  var s=MAT_ROAD.filter(function(x){return x.k===k;})[0];
  return !!(s&&applies(s,m.cat));
}
function paintTabs(){
  TCACHE=null;
  var n={mat:0,mir:0,doc:0,ipi:0,fat:0,irn:0};
  (DB.mats||[]).forEach(function(m){
    if(!isDoc(m)){
      n.mat++;
      /* each inspection page counts the reports recorded on it */
      (m.visits||[]).forEach(function(v){if(n[v.step]!=null&&stepApplies(m,v.step))n[v.step]++;});
    }
    else if(m.doc==='MIR'){if(!isLinked(m))n.mir++;}
    else if(waitsForLink(m))n.doc++;
  });
  var ot=document.getElementById('tab-other');
  if(ot)ot.setAttribute('aria-selected',String(TAB==='mat'&&isOtherView(VIEW)&&VIEW!=='paa'&&VIEW!=='pqd'&&VIEW!=='allmir'));
  var mt=document.getElementById('tab-mir');
  if(mt&&TAB==='mat'&&VIEW==='allmir')mt.setAttribute('aria-selected','true');
  /* the PQD page is the Vendors tab's second page */
  var vt=document.getElementById('tab-mfr');
  if(vt)vt.setAttribute('aria-selected',String(TAB==='mfr'||(TAB==='mat'&&VIEW==='pqd')));
  /* the PAA page is the Inspectors tab's second page */
  var it=document.getElementById('tab-insp');
  if(it)it.setAttribute('aria-selected',String(TAB==='insp'||(TAB==='mat'&&VIEW==='paa')));
  ['mat','mir','doc','ipi','fat','irn'].forEach(function(k){
    var c=document.getElementById('n-'+k);if(c)c.textContent=n[k];
    var t=document.getElementById('tab-'+k);
    if(t)t.setAttribute('aria-selected',String(TAB==='mat'&&(VIEW===k||(k==='mir'&&VIEW==='allmir'))));
  });
  ['rep','tbl','avl'].forEach(function(k){
    var t=document.getElementById('tab-'+k);
    if(t)t.setAttribute('aria-selected',String(TAB===k));
  });
  var td=document.getElementById('tab-today');
  if(td)td.setAttribute('aria-selected',String(TAB==='home'));
}

/* ---------------------------------------------------------------
   The two views share the Materials machinery and differ only in
   which records they are given, so the list is drawn by the page's
   own code over a filtered set rather than reimplemented here.
   --------------------------------------------------------------- */
function withSubset(fn){
  var all=DB.mats,was=MATS_ALL;
  MATS_ALL=all;DB.mats=all.filter(inView);
  try{return fn();}finally{DB.mats=all;MATS_ALL=was;}
}

/* The board asks what to do today, and the answer is about materials.
   A method statement has a status, not a road: twelve hundred inspection
   requests each sitting at "step 1 of 2, waiting on you" turned a real
   number into four and a half thousand, which is no number at all. */
function withMaterials(fn){
  var all=DB.mats,was=MATS_ALL;
  MATS_ALL=all;DB.mats=all.filter(function(m){return !isDoc(m);});
  try{return fn();}finally{DB.mats=all;MATS_ALL=was;}
}

/* Which pre-qualification a vendor is actually at — the chips already
   there read the whole road, so a vendor approved but missing an ISO
   date reads "in progress". This is the other question. */
var VENSTAT='';
var VEN_STATES=['Approved','Approved as Noted','Under Review','Revise & Resubmit',
  'Rejected','Terminated'];
function pqStatus(v){
  var st=pqOf(v).status||'';
  if(!st)return '';
  var k=K(st);
  if(k==='approved with comments')return 'Approved as Noted';
  if(k==='pending')return 'Under Review';
  if(k==='resubmit')return 'Revise & Resubmit';
  for(var i=0;i<VEN_STATES.length;i++)if(K(VEN_STATES[i])===k)return VEN_STATES[i];
  return st;
}
function withVendors(fn){
  if(!VENSTAT)return fn();
  var all=DB.mfrs;
  DB.mfrs=all.filter(function(v){
    return VENSTAT==='(none)'?!pqStatus(v):pqStatus(v)===VENSTAT;});
  try{return fn();}finally{DB.mfrs=all;}
}
function venChips(){
  if(TAB!=='mfr')return;
  var box=document.getElementById('filters');
  if(!box)return;
  var n={},none=0;
  (DB.mfrs||[]).forEach(function(v){
    var st=pqStatus(v);
    if(!st)none++;else n[st]=(n[st]||0)+1;
  });
  var html='<div style="flex-basis:100%;height:0"></div>'
    +'<span class="fchip" style="border:none;background:none;color:var(--rail-t3);'
    +'cursor:default;padding-left:0">pre-qualification</span>';
  function chip(key,label,count){
    html+='<button class="fchip" aria-pressed="'+(VENSTAT===key)+'" '
      +'onclick="setVenStat(\''+key+'\')">'+esc(label)
      +'<span class="fn">'+count+'</span></button>';
  }
  VEN_STATES.forEach(function(st){if(n[st])chip(st,st,n[st]);});
  if(none)chip('(none)','none recorded',none);
  if(VENSTAT)html+='<button class="fchip" onclick="setVenStat(\'\')">clear</button>';
  box.insertAdjacentHTML('beforeend',html);
}
window.setVenStat=function(k){VENSTAT=(VENSTAT===k?'':k);rList();rPane();};

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
function otherMenu(btn){
  var old=document.getElementById('other-menu');
  if(old){old.remove();return;}
  var n={},all=0;
  (DB.mats||[]).forEach(function(m){
    if(!isDoc(m)||m.doc==='MIR')return;
    all++;
    var p=OTHER_PAGES.filter(function(x){return x[1]===m.doc;})[0];
    var k=p?p[0]:'odoc';n[k]=(n[k]||0)+1;
  });
  var items=OTHER_PAGES.filter(function(p){return p[0]!=='paa'&&p[0]!=='pqd'&&n[p[0]];}).map(function(p){return [p[0],p[2]];})
    .concat(n.odoc?[['odoc','Other kinds']]:[]).concat([['alldoc','All documents']]);
  var box=document.createElement('div');
  box.id='other-menu';box.className='other-menu';box.setAttribute('role','menu');
  box.innerHTML=items.map(function(it){
    var c=it[0]==='alldoc'?all:(n[it[0]]||0);
    return '<button role="menuitem"'+(TAB==='mat'&&VIEW===it[0]?' aria-current="true"':'')
      +' onclick="otherGo(\''+it[0]+'\')"><span>'+esc(it[1])+'</span>'
      +'<span class="n">'+c+'</span></button>';
  }).join('')
    +(n.wir?('<button role="menuitem" style="color:var(--bad,#b3261e);border-top:1px solid var(--line);'
      +'border-radius:0 0 7px 7px;margin-top:4px" onclick="dropWir()"><span>Remove all WIR</span>'
      +'<span class="n">'+n.wir+'</span></button>'):'');
  document.body.appendChild(box);
  var r=btn.getBoundingClientRect();
  box.style.top=(r.bottom+4)+'px';
  box.style.left=Math.max(8,Math.min(r.left,window.innerWidth-box.offsetWidth-8))+'px';
  setTimeout(function(){
    function shut(e){
      if(e.type==='keydown'&&e.key!=='Escape')return;
      if(e.type==='click'&&box.contains(e.target))return;
      box.remove();
      document.removeEventListener('click',shut);document.removeEventListener('keydown',shut);
    }
    document.addEventListener('click',shut);document.addEventListener('keydown',shut);
  },0);
}
window.dropWir=function(){
  var m0=document.getElementById('other-menu');if(m0)m0.remove();
  var n=(DB.mats||[]).filter(function(m){return m.doc==='WIR';}).length;
  if(!n)return toast('No WIR left');
  sheet('Remove all WIR',
     '<div style="font-size:14px;line-height:1.8">This deletes the <b>'+n+'</b> work inspection '
    +'requests from the tracker, and takes them off any material they were linked to. '
    +'Nothing else is touched, and Aconex keeps them all.'
    +'<br><br>Uploading the register afterwards will not bring them back \u2014 it passes over '
    +'the WIR type. Saving the deletion takes a moment.</div>'
    +'<div class="f-act" style="margin-top:22px">'
    +'<button class="btn btn-d" onclick="dropWirYes()">Remove '+n+' WIR</button>'
    +'<button class="btn-q" onclick="closeSheet()">Cancel</button></div>');
};
window.dropWirYes=function(){
  var gone={};
  (DB.mats||[]).forEach(function(m){if(m.doc==='WIR')gone[String(m.id)]=1;});
  var n=Object.keys(gone).length;
  DB.mats=(DB.mats||[]).filter(function(m){return !gone[String(m.id)];});
  DB.mats.forEach(function(m){
    if(m.docs&&m.docs.some(function(id){return gone[String(id)];}))
      m.docs=m.docs.filter(function(id){return !gone[String(id)];});
  });
  closeSheet();
  if(TAB==='mat'&&VIEW==='wir')setTab('doc');
  touch();rList();rPane();
  toast(n+' WIR removed \u2014 saving');
};
window.otherGo=function(k){
  var m=document.getElementById('other-menu');if(m)m.remove();
  setTab(k);
};
window.setKind=function(k){DOCKIND=(DOCKIND===k?'':k);rList();rPane();};

/* RECORD is true while one record is open over its list page */
var RECORD=false;
function listKey(){
  if(TAB==='mat')return VIEW;                /* mat | mir | doc */
  if(TAB==='mfr'||TAB==='insp')return TAB;
  return null;
}
function install2(){
  if(window.__docs)return;
  window.__docs=true;
  liftTabs();

  /* The way to Today was the project name on the rail. A list page puts
     the rail away, so Today also gets a tab of its own, first in the row. */
  var bar=document.querySelector('.topbar');
  if(bar&&!document.getElementById('tab-today')){
    var tb=document.createElement('button');
    tb.className='tab';tb.id='tab-today';tb.setAttribute('role','tab');
    tb.textContent='Today';tb.onclick=function(){setTab('home');};
    bar.insertBefore(tb,bar.firstChild);
    /* Every page has the same frame: the tab row, and at its end what
       used to sit on the rail — the save state, What is due, More. */
    var tools=document.createElement('div');
    tools.className='topbar-tools';
    tools.innerHTML='<button class="btn-q" onclick="navBack()" title="Back to where you were (Alt+\u2190)">\u2190 Back</button>'
      +'<span id="saved2"></span>'
      +'<button class="btn btn-s" onclick="showDue()">What is due</button>'
      +'<button class="btn-q" onclick="showMenu()">More</button>';
    bar.appendChild(tools);
    /* The order of the row: Today, then the work in the order it happens —
       materials, vendors, the inspections at the factory, the inspection
       on arrival, the release note — and the rest after. */
    ['tab-today','tab-mat','tab-mfr','tab-ipi','tab-fat','tab-mir','tab-irn',
     'tab-doc','tab-other','tab-insp','tab-cal','tab-avl','tab-rep'].forEach(function(id){
      var b=document.getElementById(id);if(b)bar.insertBefore(b,tools);
    });
  }

  var origSetTab=window.setTab;
  window.setTab=function(t){
    RECORD=false;                            /* a tab opens on its list */
    if(t==='mir'||t==='doc'||t==='ipi'||t==='fat'||t==='irn'||isOtherView(t)){
      VIEW=t;if(t!=='doc')DOCKIND='';origSetTab('mat');}
    else if(t==='rep'||t==='tbl'||t==='avl'){
      VIEW='mat';DOCKIND='';
      origSetTab('home');          /* borrows the shape of a page with no list */
      window.TAB=t;
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
    /* Reports has no list beside it, and the page's filter row reads
       BUCKETS[TAB] — which for a tab the page has never heard of is
       undefined. Applying a file while standing on Reports threw there,
       and the throw came before the save, so the work sat in memory
       looking as though it were being written. */
    if(TAB==='rep'||TAB==='tbl'||TAB==='avl'){paintTabs();return;}
    if(TAB==='mfr'){withVendors(origList);venChips();paintTabs();return;}
    if(TAB!=='mat'){origList();paintTabs();return;}
    withSubset(origList);
    kindChips();
    paintTabs();
  };

  var origPane=window.rPane;
  window.rPane=function(){
    /* A list tab is its table, full width, with the rail put away; a
       record opened from it takes the page, with a way back. */
    var lk=listKey();
    document.body.classList.toggle('listmode',!!lk);
    if(lk&&!RECORD){
      TBL=lk;
      var el0=document.getElementById('pane');
      if(el0)el0.innerHTML=tablePane(true);
      return;
    }
    if(TAB==='rep'||TAB==='tbl'||TAB==='avl'){
      var el=document.getElementById('pane');
      if(el)el.innerHTML=(TAB==='tbl'?tablePane():TAB==='avl'?avlPane():reportsPane());
      return;
    }
    if(TAB==='home')return withMaterials(origPane);
    if(TAB==='mfr')withVendors(origPane);
    else if(TAB!=='mat')origPane();
    else withSubset(origPane);
    if(lk)backBar(lk);
  };
  function backBar(lk){
    var el=document.getElementById('pane');if(!el)return;
    /* first in the band's one row, so a record's band is no taller than any other */
    var h=el.querySelector('.head .head-m')||el.querySelector('.head .wrap')||el;
    h.insertAdjacentHTML('afterbegin','<button class="btn btn-s backbtn no-print" onclick="listBack()">← '
      +esc(/^All /.test(TABLES_DEF[lk].label)?TABLES_DEF[lk].label
        :'All '+(/^[a-z][a-z ]*$/i.test(TABLES_DEF[lk].label)&&!/ [A-Z]/.test(TABLES_DEF[lk].label)
        ?TABLES_DEF[lk].label.toLowerCase():TABLES_DEF[lk].label))+'</button>');
  }
  window.listBack=function(){RECORD=false;rPane();};
  /* One band, one height, on every page: its title and one row. Any
     further line a page puts in it — a clause note, a row count — is
     moved to a plain strip just beneath, where it is read, not framed. */
  /* The band's row, in two: what the record is (chips, free to wrap) on
     the left, and what can be done to it on the right, kept together — a
     lone "Delete" used to fall to a line of its own. */
  function splitRow(head){
    var hm=head.querySelector('.head-m');
    if(!hm||hm.querySelector('.head-acts'))return;
    var kids=[].slice.call(hm.childNodes);
    var at=kids.findIndex(function(n){return n.nodeType===1&&n.tagName==='SPAN'&&/flex:\s*1/.test(n.getAttribute('style')||'');});
    if(at<0)return;
    var chips=document.createElement('div');chips.className='head-chips';
    var acts=document.createElement('div');acts.className='head-acts';
    kids.forEach(function(n,i){if(i<at)chips.appendChild(n);else if(i>at)acts.appendChild(n);else hm.removeChild(n);});
    hm.appendChild(chips);hm.appendChild(acts);
    hm.classList.add('split');
  }
  function tidyHead(){
    var el=document.getElementById('pane');if(!el)return;
    var head=el.querySelector('.head');if(!head)return;
    splitRow(head);
    var wrap=head.querySelector('.wrap')||head;
    var extra=[].slice.call(wrap.children).filter(function(c){
      return !c.classList.contains('head-t')&&!c.classList.contains('head-m');});
    if(!extra.length)return;
    var note=document.createElement('div');
    note.className='head-note';
    var inner=document.createElement('div');
    inner.className=wrap===head?'':'wrap';
    if(wrap!==head&&wrap.getAttribute('style'))inner.setAttribute('style',wrap.getAttribute('style'));
    extra.forEach(function(c){inner.appendChild(c);});
    note.appendChild(inner);
    head.parentNode.insertBefore(note,head.nextSibling);
  }
  var rpInner=window.rPane;
  window.rPane=function(){var r=rpInner.apply(this,arguments);try{tidyHead();}catch(e){}return r;};
  /* "Add" on the board used to point at the rail's search box, which a
     list page no longer shows; it opens the list page's own Add instead */
  var origAdd=window.addNew;
  window.addNew=function(){
    if(TAB!=='home')return origAdd();
    setTab('mat');
    setTimeout(function(){
      var b=document.querySelector('#pane [onclick^="listAdd"]');if(b)b.click();},0);
  };
  /* the save state lives on the rail, which a list page puts away, so it
     is repeated in the page's own bar */
  var origStamp=window.stamp;
  window.stamp=function(){
    origStamp();
    var a=document.getElementById('saved'), b=document.getElementById('saved2');
    if(a&&b){b.textContent=a.textContent;
      /* it sits on the navy bar: pale when saved, amber while not */
      b.style.color=/dirty/.test(a.className)?'#ffcf7a':'rgba(255,255,255,.7)';
      b.style.fontSize='12.5px';}
  };
  window.stamp();
  /* adding from a list page: a name, then the new record opens */
  window.listAdd=function(ev){
    var noun={mat:'material',mfr:'vendor',insp:'inspector'}[TAB];
    popText(ev.currentTarget,'New '+noun,'','Its name as it should appear.',function(n){
      if(!n)return;
      var q=document.getElementById('q');if(!q)return;
      q.value=n;addFromSearch();
      RECORD=true;rPane();
    });
  };

  var origDue=window.showDue;
  if(typeof origDue==='function')
    window.showDue=function(){return withMaterials(origDue);};
  /* the board's breakdowns count what the board counts: materials, not
     the documents that sit beside them */
  /* Opening a vendor from the board or a list must show it, even when
     the vendor chips are narrowed to something it is not. */
  var origJump=window.jump;
  if(typeof origJump==='function')window.jump=function(tab,id){
    if(id)RECORD=true;                       /* the record, over its list */
    if(tab==='mat'&&id){
      var rec=(DB.mats||[]).filter(function(x){return String(x.id)===String(id);})[0];
      if(rec){var v2=viewFor(rec);if(v2!==VIEW){VIEW=v2;if(v2!=='doc')DOCKIND='';}}
    }
    if(tab==='mfr'&&id&&VENSTAT){
      var v=(DB.mfrs||[]).filter(function(x){return String(x.id)===String(id);})[0];
      if(v&&(VENSTAT==='(none)'?!!pqStatus(v):pqStatus(v)!==VENSTAT))VENSTAT='';
    }
    return origJump(tab,id);
  };
  ['showBucket','showSnoozed','vendorWork'].forEach(function(n){
    var o=window[n];
    if(typeof o==='function')window[n]=function(){
      var a=arguments;return withMaterials(function(){return o.apply(null,a);});};
  });

  /* Redrawing must never be able to stop a save. Everything above is
     display; the writing to the database happens after it, and a broken
     screen is a small thing beside work that never leaves the tab. */
  ['rList','rPane'].forEach(function(fn){
    var draw=window[fn];
    window[fn]=function(){
      try{return draw.apply(this,arguments);}
      catch(e){
        if(window.console)console.error(fn+' failed',e);
        /* Swallowing this left the last screen sitting there looking
           like the right one. A save must survive a broken draw, but
           the person must not be left reading a stale page and guessing.
           The pane says what happened; the list keeps quiet, since a
           list that cannot draw is obvious on its own. */
        if(fn!=='rPane')return;
        var el=document.getElementById('pane');
        if(!el)return;
        el.innerHTML='<div class="head"><div class="wrap">'
          +'<div class="head-t">This page could not be drawn</div></div></div>'
          +'<div class="body"><div class="wrap"><div class="panel"><div class="panel-b">'
          +'<div style="font-size:14px;line-height:1.8">'
          +'Your work is safe and saved — this is the drawing of one screen, '
          +'nothing else. Move to another tab and back, or reload the page.</div>'
          +'<div class="swhy" style="margin-top:14px"><span class="mono">'
          +esc(String((e&&e.message)||e))+'</span></div>'
          +'<div class="swhy" style="margin-top:6px;white-space:pre-wrap">'
          +esc(String((e&&e.stack)||'').split('\n').slice(0,4).join('\n'))
          +'</div></div></div></div></div>';
      }
    };
  });

  /* A redraw replaces the whole pane, which throws away where you were
     reading. Saving once a minute therefore threw you to the top of a
     long record once a minute. The place is kept across a redraw of the
     same thing, and let go when you move to another. */
  var WAS='';
  function place(){
    return [TAB,VIEW,DOCKIND,TBL,SEL.mat,SEL.mfr,SEL.insp].join('|');
  }
  var drawPane=window.rPane;
  window.rPane=function(){
    var here=place();
    var body=document.querySelector('.main .body')||document.getElementById('tbl-body');
    var was=(here===WAS&&body)?body.scrollTop:0;
    var r=drawPane.apply(this,arguments);
    var now=document.querySelector('.main .body')||document.getElementById('tbl-body');
    /* the same thing redrawn keeps its place; a different thing starts
       at the top, said out loud rather than left to the browser */
    if(now)now.scrollTop=(here===WAS)?was:0;
    WAS=here;
    return r;
  };
  var drawList=window.rList;
  window.rList=function(){
    var el=document.getElementById('list');
    var was=el?el.scrollTop:0;
    var r=drawList.apply(this,arguments);
    var now=document.getElementById('list');
    if(now&&was)now.scrollTop=was;
    return r;
  };

  var origInsp=window.inspPane;
  if(typeof origInsp==='function'&&!origInsp.__bar){
    window.inspPane=function(p){
      var html=readingBar(origInsp(p),'insp',p.id);
      var i=html.lastIndexOf('</div></div>');
      if(i<0)return html+paaPanel(p);
      return html.slice(0,i)+paaPanel(p)+html.slice(i);
    };
    window.inspPane.__bar=true;
  }

  /* the vendor's page gains the one about who brought them */
  var origMfr=window.mfrPane;
  window.mfrPane=function(v){
    var html=readingBar(refChip(nameable(origMfr(v),'mfr',v.id),v),'mfr',v.id);
    var i=html.lastIndexOf('</div></div>');
    if(i<0)return html+broughtPanel(v);
    return html.slice(0,i)+broughtPanel(v)+html.slice(i);
  };

  /* the two inspection steps grow a list under them */
  var origCard=window.stepCard;
  if(typeof origCard==='function'&&!origCard.__visits){
    window.stepCard=function(rec,x,i,kind){
      var html=origCard(rec,x,i,kind);
      if(kind!=='mat'||!VISIT_STEPS[x.s.k])return html;
      /* These two steps happen more than once, so the single-entry form
         underneath them says the opposite of what the list above it
         says. Two ways to record the same thing, one of which quietly
         replaces the other's work, is worse than either alone — so the
         buttons that open it go, and openStep below refuses to. */
      html=html.replace(/<div class="sacts no-print">[\s\S]*?<\/div>/,'');
      var j=html.lastIndexOf('</div></div>');
      if(j<0)return html+visitPanel(rec,x.s.k);
      return html.slice(0,j)+visitPanel(rec,x.s.k)+html.slice(j);
    };
    window.stepCard.__visits=true;
  }

  /* A step that keeps a list is never edited through the single form.
     Anything that asks for it — an old button, a keyboard shortcut —
     gets the visit sheet instead. */
  var origOpen=window.openStep;
  if(typeof origOpen==='function'&&!origOpen.__visits){
    window.openStep=function(k){
      var m=/^mat:(\w+)$/.exec(k||'');
      if(m&&VISIT_STEPS[m[1]]&&SEL.mat)return editVisit(SEL.mat,m[1]);
      return origOpen(k);
    };
    window.openStep.__visits=true;
  }

  /* the material's page gains the panel where the linking happens */
  var origMat=window.matPane;
  window.matPane=function(m){
    var html=readingBar(nameable(origMat(m),'mat',m.id),'mat',m.id);
    var add=linkPanel(m)+(isDoc(m)?'':logPanel(m));
    var i=html.lastIndexOf('</div></div>');
    html=(i<0)?html+add:html.slice(0,i)+add+html.slice(i);
    return isDoc(m)?html:withIndex(html,m);
  };
  /* A long page gets an index down its side: every section, in its
     colour, a press away — and the Main Log's share filled at its head.
     It stays put while the page scrolls; a narrow screen goes without. */
  function withIndex(html,m){
    var items=[],re=/data-nav="([^"]*)" data-tone="([^"]*)" id="([^"]+)"/g,x;
    while((x=re.exec(html)))items.push({t:x[1],tone:x[2],id:x[3]});
    if(items.length<4)return html;
    var f=logFill(m,true), pct=Math.round(f.n/f.total*100);
    var nav='<nav class="mat-nav no-print">'
      +'<button class="mn-log" onclick="navTo(\'pg-mainlog\')"><span>Main Log</span><b>'+pct+'%</b>'
      +'<i><u style="width:'+pct+'%;background:'+(pct>=80?'var(--ok)':pct>=50?'var(--wait)':'var(--now)')+'"></u></i></button>'
      +items.map(function(it){
        return '<button onclick="navTo(\''+it.id+'\')"><span class="mn-dot mn-'+it.tone+'"></span>'+it.t+'</button>';
      }).join('')+'</nav>';
    var start='<div class="body"><div class="wrap">';
    var a=html.indexOf(start);if(a<0)return html;
    html=html.slice(0,a)+'<div class="body"><div class="wrap mat-wrap">'+nav+'<div class="mat-main">'+html.slice(a+start.length);
    var b=html.lastIndexOf('</div></div>');
    return html.slice(0,b)+'</div>'+html.slice(b);
  }
  /* scrolled by hand: the page body is its own scroller, which
     scrollIntoView does not always move */
  window.navTo=function(id){
    var el=document.getElementById(id);if(!el)return;
    var box=el.parentNode;
    while(box&&box!==document.body){
      var oy=getComputedStyle(box).overflowY;
      if((oy==='auto'||oy==='scroll')&&box.scrollHeight>box.clientHeight)break;
      box=box.parentNode;
    }
    if(!box||box===document.body)box=document.scrollingElement;
    var top=el.getBoundingClientRect().top-(box===document.scrollingElement?0:box.getBoundingClientRect().top)+box.scrollTop-12;
    var was=box.style.scrollBehavior;box.style.scrollBehavior='auto';
    box.scrollTop=top;box.style.scrollBehavior=was;
  };
}

/* Every chip on these pages opens an editor except the one that matters
   most. A name pulled out of an Aconex title comes through as
   "- - -Sodamco-Concrete Admixtures & Mortar Based Solutions", and until
   now there was nowhere to fix it. */
/* The name sits in the middle of the page and was a button, which made
   it the easiest thing of all to open by accident. It is read here and
   changed in the sheet, along with everything else. */
function nameable(html){return html;}
window.renameMat=function(ev,id){
  var m=mat(id);if(!m)return;
  popText(ev.currentTarget,'Name of this record',m.name,
    'The reference is what ties this row to the register, so renaming it breaks nothing.',
    function(v){if(!v)return;m.name=v;if(m.raw)m.raw['Item Description']=v;
      touch();rList();rPane();});
};
window.renameInsp=function(ev,id){
  var p=(DB.people||[]).filter(function(x){return String(x.id)===String(id);})[0];
  if(!p)return;
  var old=p.name;
  popText(ev.currentTarget,'Name of this inspector',p.name,
    'The name as it appears on the reports. Anywhere it has been used follows the change.',
    function(n){
      if(!n||K(n)===K(old))return;
      p.name=n;
      /* a name typed onto a step is loose text, so it is carried over
         rather than left pointing at somebody who no longer exists */
      (DB.mats||[]).concat(DB.mfrs||[]).forEach(function(r){
        Object.keys(r.steps||{}).forEach(function(k){
          if(K(r.steps[k].by)===K(old))r.steps[k].by=n;
        });
        (r.visits||[]).forEach(function(v){if(K(v.by)===K(old))v.by=n;});
      });
      touch();rList();rPane();
    });
};
window.renameVen=function(ev,id){
  var v=mfr(id);if(!v)return;
  var old=v.name;
  popText(ev.currentTarget,'Name of this company',v.name,
    'Anything that named this company — a maker that it brought on — follows the change.',
    function(n){
      if(!n||K(n)===K(old))return;
      v.name=n;
      (DB.mfrs||[]).forEach(function(x){if(x.by&&K(x.by)===K(old))x.by=n;});
      touch();rList();rPane();
    });
};

function linkPanel(m){
  if(m.doc==='PQD'){
    var ven=vendorByPq(refOf(m));
    return '<div class="sec">The vendor this PQD qualifies</div>'
      +'<div class="panel"><div class="panel-b">'
      +(ven?('<div class="line row-a" onclick="jump(\'mfr\','+ven.id+')">'
          +'<span class="tag t-na">'+esc(pqStatus(ven)||'—')+'</span>'
          +'<div class="line-m">'+esc(ven.name)+'</div></div>')
        :'<span class="dim">No vendor carries this PQD number yet.</span>')
      +'</div></div>';
  }
  if(m.doc==='PAA'){
    var who=paaOwners(m);
    return '<div class="sec">The inspector this PAA approves</div>'
      +'<div class="panel"><div class="panel-b">'
      +(who.length?who.map(function(x){
          return '<div class="line row-a" onclick="jump(\'insp\','+x.id+')">'
            +'<span class="tag t-na">'+esc(x.status||'Pending')+'</span>'
            +'<div class="line-m">'+esc(x.name)+'</div></div>';}).join('')
        :'<span class="dim">Not linked to an inspector yet. Open the inspector and link it '
         +'from there.</span>')
      +'</div></div>';
  }
  if(isDoc(m)){
    var on=servedBy(m);
    return '<div class="sec">The materials this '+esc(m.doc)+' serves</div>'
      +'<div class="panel"><div class="panel-b">'
      +(on.length?on.map(function(x){
          return '<div class="line row-a" onclick="jump(\'mat\','+x.id+')">'
            +'<span class="tag t-na">'+esc(x.cat||'—')+'</span>'
            +'<div class="line-m">'+esc(x.name)+'</div></div>';}).join('')
        :(m.site||SITE_KIND[m.doc]
          ?(m.site?'<span class="dim">Marked as site work \u2014 it serves no material.</span>'
            :'<span class="dim">Site work \u2014 no material links it. If it belongs to one, link it from the material.</span>')
          :'<span class="dim">Not linked to any material yet. Open the material and link it '
           +'from there — the material is where a link is made and unmade.</span>'))
      +(on.length||NOT_FOR_MAT[m.doc]||SITE_KIND[m.doc]?''
        :('<div class="no-print" style="margin-top:12px">'
          +(m.site
            ?'<button class="btn btn-s" onclick="markSite('+m.id+',false)">Not site work \u2014 back to Documents</button>'
            :'<button class="btn btn-s" onclick="markSite('+m.id+',true)">Site work, no material</button>'
             +'<span class="dim" style="font-size:12.5px;margin-left:10px">It leaves Documents and stays on the '
             +esc(m.doc)+' page.</span>')
          +'</div>'))
      +'</div></div>';
  }
  /* What has a place of its own on the page is not listed again here:
     inspection requests are consignments, and an ITP or PID shows in its
     step — when the material's road has that step at all. */
  var inStep=function(d){return DOC_STEP[d.doc]&&stepApplies(m,DOC_STEP[d.doc]);};
  var all=docsOf(m), list=all.filter(function(d){return d.doc!=='MIR'&&!inStep(d);});
  var nmir=all.filter(function(d){return d.doc==='MIR';}).length;
  var stepped={};all.filter(inStep).forEach(function(d){stepped[d.doc]=(stepped[d.doc]||0)+1;});
  return '<div class="sec"'+navAttr('Documents','na','pg-docs')+'>Documents</div><div class="panel">'
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
      :'<span class="dim">Nothing else linked. A method statement that belongs to this material is '
      +'attached here, and one document can serve many materials.</span>')
    +Object.keys(stepped).map(function(k){
      return '<div class="dim" style="font-size:12.5px;margin-top:10px">'+stepped[k]+' '+k+(stepped[k]===1?' is':'s are')+' shown in '+(k==='ITP'?'the Inspection and test plan':'the Pre-inspection dossier')+' step.</div>';}).join('')
    +(nmir?'<div class="dim" style="font-size:12.5px;margin-top:10px">'+nmir+' inspection request'
      +(nmir===1?' is':'s are')+' under Deliveries, one consignment each.</div>':'')
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

/* ---------------------------------------------------------------
   The Main Log's columns that nothing else on the page fills — the
   package, samples and mock-ups, purchasing, fabrication, delivery and
   installation. They are kept on the material under the log's own
   headings, so what is typed here is what the log writes out.
   --------------------------------------------------------------- */
var YN=['Yes','No'];
var LOG_FORM=[
  ['Package and supply',[
    ['Common Package (Yes/No)','sel',YN],
    ['Finishing (Internal/External)','sel',['Internal','External']],
    ['Package (Lump Sum / Provisional Sum / Prime Cost)','sel',['Lump Sum','Provisional Sum','Prime Cost']],
    ['Supply & Install / Supply / Install Only','sel',['Supply & Install','Supply','Install Only']]]],
  ['Sample and mock-up',[
    ['Sample','text'],
    ['Mock-up Delivered Date','date'],
    ['Mock-up First in Place','text'],
    ['Mock-up Approved by PMC','text'],
    ['Mock-up per Client DLA','text']]],
  ['Purchasing',[
    ['Purchase Order Issued (Yes/No)','sel',YN],
    ['PO Number','text'],
    ['PO Date','date']]],
  ['Fabrication',[
    ['MES Incl. ITP (Yes/No)','sel',YN],
    ['3rd Party Assigned (Yes/No)','sel',YN],
    ['3rd Party Service Provider Name','text'],
    ['Design Verification/Calculation Status','text'],
    ['Fabrication Planned Date','date'],
    ['Fabrication 1st Batch Started Date','date'],
    ['Fabrication Percentage Completion (%)','text'],
    ['FAT Location / City','text']]],
  ['Delivery',[
    ['1st Batch Delivery To Site Planned Date','date'],
    ['1st Batch Delivery To Site Actual Date','date'],
    ['Storage (Site/Offsite)','sel',['Site','Offsite']]]],
  /* installed, then inspected as work done: one group. Work inspection
     requests are not brought in from Aconex; their columns are typed */
  ['Installation & WIR',[
    ['Installation Planned Date','date'],
    ['Installation Actual Date','date'],
    ['Installer Name','text'],
    ['WIR Number','text'],
    ['WIR Approval Date','date'],
    ['WIR Status','sel',['Approved','Approved as Noted','Under Review','Revise and Resubmit','Rejected','Terminated']]]]
];
/* The factory assessment belongs to the company, not to one material:
   it is kept on the vendor, beside its physical assessment, and every
   material from that vendor carries it into the Main Log. */
var VEN_LOG=[
  ['PA Tentative Date','date'],
  ['3rd Party Assessment Done','sel',YN],
  ['Client Assessment Done','sel',YN],
  ['PMC/LDC Assessment Done','sel',YN],
  ['Contractor Assessment Done','sel',YN]];
window.venLogFor=function(v,keys){
  var o=v.logf||{}, at=(keys||[]).indexOf('pa')>=0?'pa':(keys||[])[(keys||[]).length-1]||'';
  var filled=VEN_LOG.filter(function(f){return logShow(f,o[f[0]]);}).length;
  var h='<div class="pgroup"><div class="pg-h"><span class="pg-t">Assessment</span>'
    +'<span class="tag t-'+(filled===VEN_LOG.length?'ok':filled?'wait':'na')+'">'+filled+' of '+VEN_LOG.length+'</span>'
    +'<span style="flex:1"></span>'
    +'<button class="btn btn-s no-print" onclick="editVenLog('+v.id+')">Edit</button></div>'
    +'<div class="props">'+VEN_LOG.map(function(f){
      return '<div class="pr"><div class="pr-l">'+esc(f[0])+'</div><div class="pr-v">'+esc(logShow(f,o[f[0]]))+'</div></div>';
    }).join('')+'</div>'
    +'<div class="pg-note" style="color:var(--ink-3)">Carried into the Main Log of every material from this vendor.</div></div>';
  /* it comes before the physical assessment itself */
  var out={};out[at==='pa'?'^pa':at]=h;return out;
};
window.editVenLog=function(id){
  var v=mfr(id);if(!v)return;
  var o=v.logf||{};
  sheet('Assessment \u2014 '+v.name,
     '<div class="form" style="margin:0;padding:0;border:none">'
    +VEN_LOG.map(function(f,i){
      var fid='vl-'+i, val=o[f[0]];
      var h='<div class="f"><label for="'+fid+'">'+esc(f[0])+'</label>';
      if(f[2]===YN)return '<div class="f"><label style="display:flex;align-items:center;gap:8px;cursor:pointer">'
        +'<input type="checkbox" id="'+fid+'"'+(val==='Yes'?' checked':'')+' style="width:18px;height:18px;margin:0;flex:none">'
        +esc(f[0])+'</label></div>';
      if(f[1]==='sel')h+='<select id="'+fid+'"><option value="">\u2014</option>'
        +f[2].map(function(x){return '<option'+(x===val?' selected':'')+'>'+esc(x)+'</option>';}).join('')+'</select>';
      else h+='<input id="'+fid+'" class="mono" value="'+attr(logShow(f,val))+'" placeholder="dd/mm/yyyy" autocomplete="off">'
        +'<span class="err" id="e-'+fid+'"></span>';
      return h+'</div>';
    }).join('')
    +'<div class="f-act"><button class="btn btn-p" onclick="saveVenLog('+id+')">Save</button>'
    +'<button class="btn-q" onclick="closeSheet()">Cancel</button></div></div>');
};
window.saveVenLog=function(id){
  var v=mfr(id);if(!v)return;
  var o=Object.assign({},v.logf||{}),bad=false;
  VEN_LOG.forEach(function(f,i){
    var el=document.getElementById('vl-'+i);if(!el)return;
    var x=el.type==='checkbox'?(el.checked?'Yes':'No'):trim(el.value);
    if(f[1]==='date'&&x){var iso=parseDate(x);
      if(!iso){bad=true;document.getElementById('e-vl-'+i).textContent='Use dd/mm/yyyy';return;}x=iso;}
    if(x)o[f[0]]=x;else delete o[f[0]];
  });
  if(bad)return;
  v.logf=o;touch();closeSheet();rPane();
  toast('Saved \u2014 every material from '+v.name+' carries it');
};
/* Where each group sits on the material's page: after the first of its
   steps the material's road has (a C0 material has no vendor step, a C2
   no pre-fabrication meeting). '_mat' is the Material group itself. */
var LOG_ANCHOR={'Package and supply':['_mat'],'Sample and mock-up':['mts'],
  'Purchasing':['mfr','mts'],
  'Fabrication':['pfm','itp','pid','mfr','mts'],
  'Delivery':['del'],'Installation & WIR':['del']};
/* a group drawn inside a step's own group instead of beside it */
var LOG_INLINE={'Delivery':'del'};
window.logInline=function(m,name){
  var gi=LOG_FORM.map(function(g){return g[0];}).indexOf(name);if(gi<0)return null;
  var raw=m.raw||{}, row=appRow(m);
  return {gi:gi,props:LOG_FORM[gi][1].map(function(f){
    /* the first actual delivery, when not typed, is the first consignment */
    var v=logShow(f,raw[f[0]])||(f[0]==='1st Batch Delivery To Site Actual Date'&&row[f[0]]?logShow(f,row[f[0]]):'');
    return '<div class="pr"><div class="pr-l">'+esc(f[0])+'</div><div class="pr-v">'+esc(v)+'</div></div>';
  }).join('')};
};
window.logGroupsFor=function(m,keys){
  var raw=m.raw||{}, out={}, have={_mat:1};
  (keys||[]).forEach(function(k){have[k]=1;});
  LOG_FORM.forEach(function(g,gi){
    if(LOG_INLINE[g[0]]&&have[LOG_INLINE[g[0]]])return;
    var at=(LOG_ANCHOR[g[0]]||['del']).filter(function(k){return have[k];})[0]||'_mat';
    var filled=g[1].filter(function(f){return logShow(f,raw[f[0]]);}).length;
    out[at]=(out[at]||'')
      +'<div class="pgroup"'+navAttr(g[0],filled===g[1].length?'ok':filled?'wait':'na','pg-log-'+gi)+'><div class="pg-h"><span class="pg-t">'+esc(g[0])+'</span>'
      +'<span class="tag t-'+(filled===g[1].length?'ok':filled?'wait':'na')+'">'+filled+' of '+g[1].length+'</span>'
      +'<span style="flex:1"></span>'
      +'<button class="btn btn-s no-print" onclick="editLog('+m.id+','+gi+')">Edit</button></div>'
      +'<div class="props">'+g[1].map(function(f){
        var v=logShow(f,raw[f[0]]);
        return '<div class="pr"><div class="pr-l">'+esc(f[0])+'</div><div class="pr-v">'+esc(v)+'</div></div>';
      }).join('')+'</div></div>';
  });
  return out;
};
function logShow(f,v){
  if(v==null||v==='')return '';
  return f[1]==='date'?(show(v)||String(v)):String(v);
}
/* How much of a material's Main Log row the system fills today, and
   which of the 77 columns are still empty — worked out from the very
   row the export writes, so the two never disagree. */
function logFill(m,fresh){
  var c=tc();c.fill=c.fill||{};
  if(c.fill[m.id]&&!fresh)return c.fill[m.id];
  var row=appRow(m),missing=[];
  COLS.forEach(function(k){var v=row[k];if(v==null||String(v).trim()==='')missing.push(k);});
  return (c.fill[m.id]={n:COLS.length-missing.length,total:COLS.length,missing:missing});
}
/* where in the system an empty column is filled */
var LOG_WHERE=[
  [/^(Item Description|Material Category|Discipline|Sub-contractor Name)$/,'the Material group'],
  [/^(Manufacturer|Country of Origin|PQD |PA |Assessment Result|3rd Party Assessment|Client Assessment|PMC\/LDC Assessment|Contractor Assessment)/,'the vendor’s page'],
  [/^MAT /,'Technical submittal, or Aconex'],
  [/^(ITP )/,'link the ITP, or the ITP step'],
  [/^(Method Statement Number|MES Revision|MES Status)$/,'link the method statement'],
  [/^PID /,'Pre-inspection dossier'],
  [/^Pre-Fabrication/,'Pre-fabrication meeting'],
  [/^(FAT Package|FAT Planned|FAT\/TPI)/,'FAT / Final inspection'],
  [/^MIR /,'link the MIR, or record a delivery'],
  [/^Total Quantity/,'the quantity at the top'],
  [/^(Delivered|Remaining)/,'record deliveries']
];
function logWhere(c){
  var g=LOG_FORM.filter(function(g){return g[1].some(function(f){return f[0]===c;});})[0];
  if(g)return 'the '+g[0]+' group';
  var w=LOG_WHERE.filter(function(x){return x[0].test(c);})[0];
  return w?w[1]:'';
}
function logPanel(m){
  var raw=m.raw||{},filled=0,total=0;
  var fill=logFill(m,true), pct=Math.round(fill.n/fill.total*100);
  var body='';
  /* the empty columns, each with where it is filled */
  var miss=fill.missing.map(function(c){
    var w=logWhere(c);
    return '<div style="display:flex;gap:12px;padding:3px 0;font-size:13px;border-top:1px solid var(--line)">'
      +'<span style="flex:0 0 46%">'+esc(c)+'</span><span class="dim">'+esc(w)+'</span></div>';
  }).join('');
  return '<div class="sec"'+navAttr('Main Log',pct>=80?'ok':pct>=50?'wait':'now','pg-mainlog')+'>Main Log</div><div class="panel">'
    +'<div class="panel-h"><div style="flex:1;min-width:220px">'
    +'<div class="panel-t">'+fill.n+' of '+fill.total+' columns filled <span class="dim" style="font-weight:400">· '+pct+'%</span></div>'
    +'<div style="height:6px;border-radius:3px;background:var(--sunk);margin-top:8px;overflow:hidden">'
    +'<div style="height:100%;width:'+pct+'%;background:'+(pct>=80?'var(--ok)':pct>=50?'var(--wait)':'var(--now)')+'"></div></div></div>'
    +'<button class="btn btn-s no-print" onclick="editLog('+m.id+')">Edit all Main Log fields</button></div>'
    +'<div class="panel-b">'
    +(fill.missing.length?('<details><summary style="cursor:pointer;font-weight:600;font-size:13.5px">'
      +fill.missing.length+' empty column'+(fill.missing.length===1?'':'s')+' — and where each is filled</summary>'
      +'<div style="margin-top:8px">'+miss+'</div></details>'):'<span class="dim">Every column is filled.</span>')
    +(body?('<div style="margin-top:14px;font-weight:600;font-size:13.5px">Main Log details '
      +'<span class="dim" style="font-weight:400">'+filled+' of '+total+'</span></div>'+body):'')
    +'</div></div>';
}
window.editLog=function(id,only){
  var m=mat(id);if(!m)return;
  var raw=m.raw||{};
  var one=(only!=null&&LOG_FORM[only])?only:null;
  sheet((one!=null?LOG_FORM[one][0]:'Main Log fields')+' \u2014 '+m.name,
     '<div class="form" style="margin:0;padding:0;border:none">'
    +LOG_FORM.map(function(g,gi){
      if(one!=null&&gi!==one)return '';
      return (one!=null?'':'<div class="f wide" style="grid-column:1/-1;margin:14px 0 0;font-weight:650;color:var(--ink)">'+esc(g[0])+'</div>')
        +g[1].map(function(f,fi){
          var fid='lg-'+gi+'-'+fi, v=raw[f[0]];
          var h='<div class="f'+(f[1]==='text'?' wide':'')+'"><label for="'+fid+'">'+esc(f[0])+'</label>';
          if(f[1]==='sel'){
            var opts=f[2].slice();
            if(v&&opts.indexOf(v)<0)opts.push(v);       /* a value typed some other way is kept */
            h+='<select id="'+fid+'"><option value="">\u2014</option>'
              +opts.map(function(o){return '<option'+(o===v?' selected':'')+'>'+esc(o)+'</option>';}).join('')
              +'</select>';
          }else if(f[1]==='date'){
            h+='<input id="'+fid+'" class="mono" value="'+attr(logShow(f,v))+'" placeholder="dd/mm/yyyy" autocomplete="off">'
              +'<span class="err" id="e-'+fid+'"></span>';
          }else h+='<input id="'+fid+'" value="'+attr(v==null?'':String(v))+'" autocomplete="off">';
          return h+'</div>';
        }).join('');
    }).join('')
    +'<div class="f-act"><button class="btn btn-p" onclick="saveLog('+id+')">Save</button>'
    +'<button class="btn-q" onclick="closeSheet()">Cancel</button></div></div>');
};
window.saveLog=function(id){
  var m=mat(id);if(!m)return;
  var out={},bad=false;
  LOG_FORM.forEach(function(g,gi){g[1].forEach(function(f,fi){
    var fid='lg-'+gi+'-'+fi, el=document.getElementById(fid);if(!el)return;
    var v=trim(el.value);
    if(f[1]==='date'&&v){
      var iso=parseDate(v);
      if(!iso){bad=true;var e=document.getElementById('e-'+fid);if(e)e.textContent='Use dd/mm/yyyy';return;}
      v=iso;
    }
    out[f[0]]=v;
  });});
  if(bad)return;
  m.raw=m.raw||{};
  Object.keys(out).forEach(function(c){if(out[c])m.raw[c]=out[c];else delete m.raw[c];});
  touch();closeSheet();rPane();
  toast('Saved \u2014 the Main Log carries it');
};
/* ---------------------------------------------------------------
   A PAA approves a person, so it is linked to an inspector — by hand,
   from the inspector's page, the way a material takes its documents.
   --------------------------------------------------------------- */
function paaOwners(d){
  return (DB.people||[]).filter(function(p){
    return (p.docs||[]).some(function(id){return String(id)===String(d.id);});});
}
function paaPanel(p){
  var list=(p.docs||[]).map(function(id){return mat(id);}).filter(Boolean);
  return '<div class="sec">Personnel approval (PAA)</div><div class="panel">'
    +'<div class="panel-h"><div class="panel-t">From Aconex</div>'
    +'<button class="btn btn-s no-print" onclick="paaPick('+p.id+')">Link a PAA</button></div>'
    +'<div class="panel-b">'
    +(list.length?list.map(function(d){
        var st=rawEnd(d,'Status');
        return '<div class="line">'
          +'<span class="tag t-na" style="min-width:54px;text-align:center">PAA</span>'
          +'<div class="line-m"><div>'+esc(d.name)+'</div>'
          +'<div class="dim mono" style="font-size:12.5px;margin-top:2px">'
          +esc(refOf(d))+(st?(' \u00b7 '+esc(st)):'')+'</div></div>'
          +'<button class="btn-q" onclick="jump(\'mat\','+d.id+')">Open</button>'
          +'<button class="btn-q" onclick="paaUnlink('+p.id+','+d.id+')">Unlink</button></div>';
      }).join('')
      :'<span class="dim">No PAA linked yet. Link the approval Aconex holds for this person.</span>')
    +'</div></div>';
}
window.paaPick=function(pid,q,go){
  var p=insp(pid);if(!p)return;
  var need=K(q||'');
  var all=(DB.mats||[]).filter(function(d){return d.doc==='PAA';});
  var rows=all.filter(function(d){
    if((p.docs||[]).some(function(id){return String(id)===String(d.id);}))return false;
    if(!need)return !paaOwners(d).length;            /* start with the ones nobody has */
    return K(d.name+' '+refOf(d)).indexOf(need)>=0;
  });
  if(go&&need&&rows.length===1)return paaLink(pid,rows[0].id);
  sheet('Link a PAA to '+p.name,
     '<div class="dim" style="font-size:13.5px;margin-bottom:14px">'
    +(need?('Searching all '+all.length+' PAA files.')
          :('Showing the '+rows.length+' not linked to anyone \u2014 search to see all '+all.length+'.'))
    +'</div>'
    +'<div class="f" style="margin-bottom:14px"><label for="lk">Search by number or title</label>'
    +'<input id="lk" value="'+attr(q||'')+'" autocomplete="off" '
    +'onkeydown="if(event.key===\'Enter\'){event.preventDefault();paaPick('+pid+',this.value,true);}" '
    +'oninput="searchSoon(function(v){paaPick('+pid+',v);},this.value)">'
    +'<span class="dim" style="font-size:12px">The list follows as you type</span></div>'
    +'<div class="panel"><div class="panel-b">'
    +(rows.length?rows.slice(0,60).map(function(d){
        var o=paaOwners(d);
        return '<div class="line row-a" onclick="paaLink('+pid+','+d.id+')">'
          +'<span class="tag t-na" style="min-width:54px;text-align:center">PAA</span>'
          +'<div class="line-m"><div>'+esc(d.name)+'</div>'
          +'<div class="dim mono" style="font-size:12.5px;margin-top:2px">'+esc(refOf(d))
          +(o.length?(' \u00b7 linked to '+esc(o.map(function(x){return x.name;}).join(', '))):'')
          +'</div></div></div>';}).join('')
        +(rows.length>60?('<div class="dim" style="font-size:13px;padding-top:10px">and '
          +(rows.length-60)+' more \u2014 narrow the search</div>'):'')
      :'<span class="dim">Nothing matches.</span>')
    +'</div></div>');
  var f=document.getElementById('lk');
  if(f){f.focus();f.setSelectionRange(f.value.length,f.value.length);}
};
window.paaLink=function(pid,docId){
  var p=insp(pid);if(!p)return;
  p.docs=p.docs||[];
  if(!p.docs.some(function(x){return String(x)===String(docId);}))p.docs.push(docId);
  touch();closeSheet();rList();rPane();
  toast('PAA linked to '+p.name);
};
window.paaUnlink=function(pid,docId){
  var p=insp(pid);if(!p)return;
  p.docs=(p.docs||[]).filter(function(x){return String(x)!==String(docId);});
  touch();rList();rPane();
};
window.showLoosePaa=function(){
  setTab('paa');
  TBLQ[TBL]=TBLQ[TBL]||{};TBLQ[TBL].lnk='Not linked';
  TBLN[TBL]=PAGE_ROWS;rPane();
};
window.unlink=function(matId,docId){
  var m=mat(matId);if(!m)return;
  m.docs=(m.docs||[]).filter(function(x){return String(x)!==String(docId);});
  /* the consignment the link made goes with it, unless something was
     written on it by hand — then it stays, no longer tied to the link */
  m.dels=(m.dels||[]).filter(function(c){
    if(String(c.fromDoc)!==String(docId))return true;
    if(c.qty||c.note){delete c.fromDoc;return true;}
    return false;
  });
  syncDocSteps(m);
  touch();rPane();paintTabs();
};
/* An inspection request is a consignment arriving: linked to a
   material, it becomes one under Deliveries — its number, date and
   outcome from Aconex, the quantity left for the person to add. The
   register keeps its outcome up to date from then on, as it does for
   any consignment. */
function delWordOf(s){
  var w=normStatus(s)||trim(s||'');
  if(w==='Approved')return 'Received';
  if(w==='Approved as Noted'||w==='Approved with comments')return 'Approved with comments';
  return w||'Pending';
}
var DELID=null;
function mirToDel(m,d){
  if(!m||isDoc(m)||!d||d.doc!=='MIR')return false;
  var no=refOf(d), raw=d.raw||{};
  m.dels=m.dels||[];
  var have=m.dels.filter(function(c){return (c.fromDoc&&String(c.fromDoc)===String(d.id))||(no&&K(c.ref)===K(no));})[0];
  if(have){if(!have.fromDoc){have.fromDoc=d.id;return true;}return false;}
  m.dels.push({id:(DELID||(DELID=idMaker()))(),qty:'',date:raw['MIR Approval Date']||d.acxDate||'',ref:no,
    status:delWordOf(raw['MIR Status']),note:'',fromDoc:d.id});
  return true;
}
/* the requests linked before this, made into their consignments once */
/* An inspection and test plan or a pre-inspection dossier linked to a
   material is that material's step: its number, outcome and date come
   from the documents themselves, read again whenever a link changes or
   the register brings a new outcome. Several linked (fabrication and
   erection, say) give every number, the latest date, and the outcome
   that still holds the step back — all must be approved for it to be. */
var DOC_STEP={ITP:'itp',PID:'pid'};
var STEP_RANK={'Rejected':0,'Resubmit':1,'Pending':2,'Approved with comments':3,'Approved':4};
function syncDocSteps(m){
  if(!m||isDoc(m))return false;
  var changed=false, docs=docsOf(m);
  m.steps=m.steps||{};
  Object.keys(DOC_STEP).forEach(function(kind){
    var k=DOC_STEP[kind], list=docs.filter(function(d){return d.doc===kind;});
    var cur=m.steps[k];
    if(!list.length){
      if(cur&&cur.fromDocs){delete m.steps[k];changed=true;}
      return;
    }
    var refs=[],date='',sts=[];
    list.forEach(function(d){
      var raw=d.raw||{};
      var r=refOf(d);if(r)refs.push(r);
      var dt=raw[kind+' Submittal Date']||d.acxDate||'';
      if(/^\d{4}-\d{2}-\d{2}$/.test(String(dt))&&dt>date)date=dt;
      sts.push(normStatus(raw[kind+' Status'])||'Pending');
    });
    var live=sts.filter(function(x){return x!=='Terminated';});
    var status=live.length?live.reduce(function(a,b){return (STEP_RANK[b]<STEP_RANK[a])?b:a;}):'Terminated';
    var next={ref:refs.join(', '),date:date,status:status,fromDocs:true};
    if(!cur||cur.ref!==next.ref||cur.date!==next.date||cur.status!==next.status||!cur.fromDocs){
      m.steps[k]=Object.assign({},cur||{},next);changed=true;
    }
  });
  return changed;
}
function liftDocSteps(){
  var n=0;
  (DB.mats||[]).forEach(function(m){if(syncDocSteps(m))n++;});
  if(n)touch();
  return n;
}
function liftMirLinks(){
  var n=0;
  (DB.mats||[]).forEach(function(m){
    if(isDoc(m))return;
    docsOf(m).forEach(function(d){if(mirToDel(m,d))n++;});
  });
  if(n)touch();
  return n;
}
/* a search that redraws a short pause after the last key, not on every one */
var SEARCH_T=null;
window.searchSoon=function(fn,v){clearTimeout(SEARCH_T);SEARCH_T=setTimeout(function(){fn(v);},250);};
window.linkPick=function(matId,q,go){
  var m=mat(matId);if(!m)return;
  var has={};(m.docs||[]).forEach(function(id){has[String(id)]=1;});
  var need=K(q||'');
  var all=(DB.mats||[]).filter(isDoc);
  var rows=all.filter(function(d){
    if(has[String(d.id)])return false;
    if(!need)return K(d.disc||'')===K(m.disc||'');   /* start with its own trade */
    return K(d.name+' '+refOf(d)+' '+(d.doc||'')).indexOf(need)>=0;
  });
  /* Enter on a search that leaves one document links it there and then */
  if(go&&need&&rows.length===1)return linkAdd(matId,rows[0].id);
  var already=need?all.filter(function(d){
    return has[String(d.id)]&&K(d.name+' '+refOf(d)+' '+(d.doc||'')).indexOf(need)>=0;}):[];
  sheet('Link a document to '+m.name,
     '<div class="dim" style="font-size:13.5px;margin-bottom:14px">'
    +(need?('Searching all '+all.length+' documents.')
          :('Showing the '+rows.length+' in '+esc(m.disc||'no discipline')
            +' — search to see the other '+(all.length-rows.length)+'.'))
    +' A document can serve many materials; linking it here does not take it from anywhere else.'
    +'</div>'
    +'<div class="f" style="margin-bottom:14px"><label for="lk">Search by number, title or kind</label>'
    +'<input id="lk" value="'+attr(q||'')+'" autocomplete="off" '
    +'onkeydown="if(event.key===\'Enter\'){event.preventDefault();linkPick('+matId+',this.value,true);}" '
    +'oninput="searchSoon(function(v){linkPick('+matId+',v);},this.value)">'
    +'<span class="dim" style="font-size:12px">The list follows as you type. Press a document to link it — '
    +'or Enter, when the search leaves only one.</span></div>'
    +'<div class="panel"><div class="panel-b">'
    +(rows.length?rows.slice(0,60).map(function(d){
        return '<div class="line row-a" onclick="linkAdd('+matId+','+d.id+')">'
          +'<span class="tag t-na" style="min-width:54px;text-align:center">'+esc(d.doc)+'</span>'
          +'<div class="line-m"><div>'+esc(d.name)+'</div>'
          +'<div class="dim mono" style="font-size:12.5px;margin-top:2px">'+esc(refOf(d))+'</div></div>'
          +'</div>';}).join('')
        +(rows.length>60?('<div class="dim" style="font-size:13px;padding-top:10px">and '
          +(rows.length-60)+' more — narrow the search</div>'):'')
      :(already.length
        ?'<span class="dim">Already linked to this material: <b>'+esc(already.map(function(d){return refOf(d)||d.name;}).join(', '))+'</b>.</span>'
        :'<span class="dim">Not in the tracker. If it is in Aconex, upload the latest register '
         +'(More → Upload the register) and it comes in; then link it here.</span>'))
    +'</div></div>');
  var f=document.getElementById('lk');
  if(f){f.focus();f.setSelectionRange(f.value.length,f.value.length);}
};
window.linkAdd=function(matId,docId){
  var m=mat(matId);if(!m)return;
  m.docs=m.docs||[];
  if(m.docs.indexOf(docId)<0&&!m.docs.some(function(x){return String(x)===String(docId);}))
    m.docs.push(docId);
  var d=mat(docId);
  var made=mirToDel(m,d);
  syncDocSteps(m);
  touch();closeSheet();rPane();paintTabs();
  if(made)return toast('Linked '+refOf(d)+' — it is a consignment under Deliveries now; add its quantity there');
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
    var A=(a.kind==='sub'||a.kind==='makesub')?0:1,B=(b.kind==='sub'||b.kind==='makesub')?0:1;
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

/* The chip that carries the scope showed the whole Aconex title, which
   for a pre-qualification is a paragraph — and sitting where a chip
   normally holds a reference, it read as one. The reference goes where
   the eye looks for it, and the title is cut to a chip's worth with the
   rest on hover. */
function refChip(html,v){
  var st=v.steps||{};
  var ref=pqOf(v).ref||(st.pqd&&st.pqd.ref)||(st.appr&&st.appr.ref)||'';
  var i=html.indexOf('<div class="head-m">');
  if(i>=0&&v.locality){
    html=html.slice(0,i+20)+'<span class="chip flat">'+esc(v.locality)+'</span>'+html.slice(i+20);
  }
  if(i>=0&&ref){
    var chip='<span class="chip flat" title="'+attr(ref)+'">'
      +'<span class="mono">'+esc(ref.length>44?(ref.slice(0,42)+'…'):ref)+'</span></span>';
    html=html.slice(0,i+20)+chip+html.slice(i+20);
  }
  /* the scope chip, shortened in place */
  html=html.replace(/(onclick="editText\(event,'scope','Scope of work'\)">)([^<]{46,})(<)/,
    function(all,open,text,close){
      return open+esc(trim(text).slice(0,44))+'…'+close;
    });
  return html;
}

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

/* Grouped by discipline, but without the heading rows the old file used
   to separate them. A row that is a heading and not a material has to be
   recognised again on the way back in, and sorting is what a spreadsheet
   is for. */
/* Columns the log carries after its seventy-seven. The reader checks
   the first seventy-seven by position and stops there, so anything past
   them travels out without disturbing the way back in. These are read
   off the vendor, and they are not read back onto it: one vendor serves
   many rows, and a correction belongs on the vendor sheet where there
   is one row per company. */
/* The company behind a row. Many rows here name no manufacturer at all
   and only a subcontractor — the UAAC curtain wall is one — so reading
   the manufacturer alone left them blank however the company was set.
   The maker first; failing that, the subcontractor by name. */
function companyOfMat(m){
  /* The company behind a row, found the way the project links them: by
     the pre-qualification the row carries. A name is spelt one way on
     the material ("UAAC") and another on the vendor ("United Arab
     Aluminum Company"); PRQ-00011 is the same on both. The linked vendor
     comes first if it says anything, then the reference, then the
     subcontractor's name as a last resort. */
  var v=m.mfr?mfr(m.mfr):null;
  if(v&&v.locality)return v;
  var refs=splitRefs((m.raw||{})['PQD Number']);
  for(var i=0;i<refs.length;i++){
    var want=K(refs[i]);
    var hit=(DB.mfrs||[]).filter(function(x){
      if(K(pqOf(x).ref||'')===want)return true;
      return (x.pq2||[]).some(function(p){return K(p.ref)===want;});
    })[0];
    if(hit)return hit;
  }
  var sub=trim(m.sub||((m.raw||{})['Sub-contractor Name'])||'');
  if(sub){
    var s2=(DB.mfrs||[]).filter(function(x){return K(x.name)===K(sub);})[0];
    if(s2)return s2;
  }
  return v;
}
var LOG_EXTRA=[
  {t:'Local / Foreign',w:14,read:function(m){var v=companyOfMat(m);return v?(v.locality||''):'';}},
  /* every report, oldest first, one to a line: reference, date, result */
  {t:'In-Process Inspection Reports',w:40,read:function(m){return visitLines(m,'ipi');}},
  {t:'Inspection Release Notes',w:34,read:function(m){return visitLines(m,'irn');}}
];
function visitLines(m,k){
  return visitsOf(m,k).slice().reverse().map(function(v){
    return [v.ref||'(no reference)',v.date?showDate(v.date):'',v.result||''].filter(Boolean).join(' \u00b7 ');
  }).join('\n');
}
/* the Main Log's columns by stage, for the coloured row above the table */
var LOG_STAGES=[
  ['Material','Item Description'],['Vendor & PQD','Manufacturer'],['MAT submittal','MAT Number'],
  ['Assessment','PA Tentative Date'],['Sample & Mock-up','Sample'],['Purchase','Purchase Order Issued (Yes/No)'],
  ['MES / ITP / PID','Method Statement Number'],['Fabrication & FAT','Pre-Fabrication Meeting Date'],
  ['Delivery & MIR','1st Batch Delivery To Site Planned Date'],['Installation & WIR','Installation Planned Date'],
  ['Quantities','Total Quantity'],['From the tracker','Local / Foreign']];
function logBands(head){
  var starts=LOG_STAGES.map(function(s){return {t:s[0],from:head.indexOf(s[1])};})
    .filter(function(b){return b.from>=0;}).sort(function(a,b){return a.from-b.from;});
  starts.forEach(function(b,i){b.to=(i+1<starts.length?starts[i+1].from:head.length)-1;});
  return starts;
}
function generalRows(){
  var mats=(DB.mats||[]).filter(function(m){return !isDoc(m);});
  mats.sort(function(a,b){
    var d=String(a.disc||'~').localeCompare(String(b.disc||'~'));
    return d||String(a.name).localeCompare(String(b.name));
  });
  var rows=[COLS.concat(LOG_EXTRA.map(function(x){return x.t;}))];
  mats.forEach(function(m){
    var raw=foldDocs(m,rawOut(m));
    rows.push(COLS.map(function(c){
      var v=raw[c];
      if(v==null||v==='')return '';
      if(DATE_COLS[c]&&/^\d{4}-\d{2}-\d{2}$/.test(String(v)))return {date:String(v)};
      if(DATE_COLS[c]&&String(v).indexOf('\n')>=0)
        return String(v).split('\n').map(function(x){return /^\d{4}-\d{2}-\d{2}$/.test(x)?showDate(x):x;}).join('\n');
      return v;
    }).concat(LOG_EXTRA.map(function(x){return x.read(m)||'';})));
  });
  return rows;
}

/* a document attached to nothing would vanish from a report shaped
   one-row-per-material, so it is written to a sheet of its own rather
   than quietly dropped */
function looseRows(){
  var loose=(DB.mats||[]).filter(function(m){
    return waitsForLink(m)||(m.doc==='MIR'&&!isLinked(m));
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

/* Every column here can be edited and sent back, including the name and
   the pre-qualification number — which is why the first column exists.
   It is the only thing on the row that must not be touched: it is how a
   row finds its way home after somebody has corrected everything else
   about it. */
var VEN_COLS=['ID','Vendor','Kind','Category','Brought by','Country','Local / Foreign','Production site','Scope',
  'PQD Number','PQD Status','PQD Date','ISO Number','ISO Expires','ISO Status',
  'Visit result','Visit date','QMS auditor','Technical expert','Visit report','Technical expert report',
  'Materials'];
/* headings an older copy of the sheet used, still read as the new ones */
var VEN_ALIAS={'Visit result':'Assessment','Visit report':'Assessment Reference','Visit date':'Assessment Date'};
/* Written as typed, a status the road does not recognise stops it:
   "approved" is not "Approved". Each is read back into the words the
   page uses; anything it cannot place is kept as typed. */
var VISIT_WORDS=['Pending','Scheduled','Passed','Failed','Waived'];
function pqWord(s){return normStatus(s)||s;}
function isoWordIn(s){var k=K(s);return /valid/.test(k)?'Valid':/expir/.test(k)?'Expired':s;}
function visitWordIn(s){
  var k=K(s);
  var hit=VISIT_WORDS.filter(function(w){return k.indexOf(w.toLowerCase().slice(0,4))===0;})[0];
  return hit||s;
}
var VEN_KINDS={'manufacturer':'maker','maker':'maker','subcontractor':'sub','sub':'sub',
  'manufacturer & subcontractor':'makesub','manufacturer and subcontractor':'makesub','makesub':'makesub',
  'supplier':'supplier','inspection agency':'agency','agency':'agency'};

/* ---------------------------------------------------------------
   One pre-qualification number, one company. Where several vendors hold
   the same number, one of them is its real holder and the rest were
   brought in under it. The holder is guessed — the vendor whose scope
   carries the pre-qualification's own title (it came from the register),
   then a subcontractor — and can be changed before anything is written.
   --------------------------------------------------------------- */
function pqGroups(){
  var by={};
  (DB.mfrs||[]).forEach(function(v){
    var ref=trim(pqOf(v).ref||'');if(!ref||ref==='-')return;
    (by[K(ref)]=by[K(ref)]||{ref:ref,list:[]}).list.push(v);
  });
  return Object.keys(by).map(function(k){return by[k];}).filter(function(g){return g.list.length>1;})
    .sort(function(a,b){return b.list.length-a.list.length||String(a.ref).localeCompare(String(b.ref));});
}
function pqGuess(g){
  var titled=g.list.filter(function(v){return /\b(PRQ|pre-?qualification)\b/i.test(v.scope||'')||v.reg;});
  if(titled.length===1)return titled[0];
  var subs=g.list.filter(function(v){return v.kind==='sub';});
  if(subs.length===1)return subs[0];
  return titled[0]||subs[0]||g.list[0];
}
/* ---------------------------------------------------------------
   The materials sheet that goes out to be edited and comes back. Aconex
   owns a material's number and its outcome, so those travel for reading
   only; what can be changed is what Aconex does not carry — the name as
   you want it, category, discipline, vendor, subcontractor, quantity
   and purchase order. A row finds its material by ID, then by MAT
   number. A sheet cannot add a material: they come from Aconex.
   --------------------------------------------------------------- */
var MAT_COLS=['ID','MAT Number','Item Description','Category','Discipline','Vendor','Vendor PQD','Sub-contractor',
  'Quantity','Unit','PO Number','MAT Status'];
/* the vendor that holds a pre-qualification number — the whole number,
   or its tail ("PRQ-00024", "00024") when only one vendor's ends so */
function vendorByPq(x){
  var k=K(x);if(!k)return null;
  var refs=function(v){return [pqOf(v).ref].concat((v.pq2||[]).map(function(p){return p.ref;})).filter(Boolean).map(K);};
  var exact=(DB.mfrs||[]).filter(function(v){return refs(v).indexOf(k)>=0;});
  if(exact.length)return exact[0];
  var tail=k.replace(/^0+/,'');
  if(tail.length<2)return null;
  var ends=(DB.mfrs||[]).filter(function(v){return refs(v).some(function(r){
    return r.slice(-k.length)===k||r.replace(/.*-/,'').replace(/^0+/,'')===tail;});});
  return ends.length===1?ends[0]:null;
}
function matNo(m){return trim(m.ref||(m.raw||{})['MAT Number']||'');}
function vendorByName(name){
  var k=K(name), c=coName(name);
  var hits=(DB.mfrs||[]).filter(function(v){return K(v.name)===k;});
  if(!hits.length)hits=(DB.mfrs||[]).filter(function(v){return c&&coName(v.name)===c;});
  /* two records of one company (two pre-qualifications): the live one first */
  hits.sort(function(a,b){return (pqOf(a).status==='Terminated')-(pqOf(b).status==='Terminated');});
  return hits[0]||null;
}
var MAT_FIELDS=[
  {c:'Item Description',get:function(m){return m.name;},set:function(m,x){if(x){m.name=x;m.raw=m.raw||{};m.raw['Item Description']=x;}}},
  {c:'Category',get:function(m){return m.cat;},
    read:function(x){return normCat(x)||(/^[0-3]$/.test(trim(x))?'C'+trim(x):'');},
    set:function(m,x){m.cat=x;m.catSure=true;m.catFrom=x?'set in Excel':'';}},
  {c:'Discipline',get:function(m){return m.disc;},set:function(m,x){m.disc=x;}},
  {c:'Vendor',get:function(m){var v=m.mfr?mfr(m.mfr):null;return v?v.name:'';},
    set:function(m,x){if(!x){m.mfr='';return;}var v=vendorByName(x);if(v)m.mfr=String(v.id);}},
  /* the vendor by its pre-qualification number: exact where the name is
     not, and applied after the name, so it wins where the two disagree */
  {c:'Vendor PQD',get:function(m){var v=m.mfr?mfr(m.mfr):null;return v?(pqOf(v).ref||''):'';},
    set:function(m,x){if(!x){m.mfr='';return;}var v=vendorByPq(x);if(v)m.mfr=String(v.id);}},
  {c:'Sub-contractor',get:function(m){return m.sub;},set:function(m,x){m.sub=x;}},
  {c:'Quantity',get:function(m){return m.qty;},set:function(m,x){m.qty=x;}},
  {c:'Unit',get:function(m){return m.unit;},set:function(m,x){m.unit=x;}},
  {c:'PO Number',get:function(m){return (m.raw||{})['PO Number'];},
    set:function(m,x){m.raw=m.raw||{};if(x)m.raw['PO Number']=x;else delete m.raw['PO Number'];}}
];
/* The Main Log's own fields go out and come back too, after the twelve:
   a blank cell leaves a value alone, "-" clears it, a date is read in any
   of the ways a sheet writes one, and a choice is matched to its list
   whatever its capitals. PO Number is already among the twelve. */
/* Documents linked by number from the sheet. A cell may hold several,
   written any way — on lines of their own, with commas, or run together
   with no space at all: each document of that kind whose full number is
   found inside the cell is linked. Nothing is unlinked except by "-". */
function docsInText(t,kind){
  var flat=String(t||'').toUpperCase().replace(/[\s,;\/|]+/g,'');
  if(!flat)return [];
  return (MATS_ALL||DB.mats||[]).filter(function(d){
    if(d.doc!==kind)return false;
    var no=String(refOf(d)||'').toUpperCase().replace(/\s+/g,'');
    return no.length>6&&flat.indexOf(no)>=0;
  });
}
[['ITP Number','ITP'],['MES Number','MES']].forEach(function(p){
  MAT_COLS.push(p[0]);
  MAT_FIELDS.push({c:p[0],docKind:p[1],
    get:function(m){return linkedVals(m,p[1],'no');},
    set:function(m,x){
      m.docs=m.docs||[];
      if(!x){                                      /* "-" : unlink them all */
        var gone={};docsOf(m).forEach(function(d){if(d.doc===p[1])gone[String(d.id)]=1;});
        m.docs=m.docs.filter(function(id){return !gone[String(id)];});
      }else docsInText(x,p[1]).forEach(function(d){
        if(!m.docs.some(function(id){return String(id)===String(d.id);}))m.docs.push(d.id);
      });
      syncDocSteps(m);
    }});
});
LOG_FORM.forEach(function(g){g[1].forEach(function(f){
  if(MAT_COLS.indexOf(f[0])>=0)return;
  MAT_COLS.push(f[0]);
  MAT_FIELDS.push({c:f[0],log:f,
    get:function(m){return (m.raw||{})[f[0]]||'';},
    read:function(x){
      if(f[1]==='date')return anyDate(x)||trim(x);
      if(f[1]==='sel'){var k=K(x),hit=f[2].filter(function(o){return K(o)===k;})[0];return hit||trim(x);}
      return trim(x);
    },
    set:function(m,x){m.raw=m.raw||{};if(x)m.raw[f[0]]=x;else delete m.raw[f[0]];}});
});});
function matWant(r,fd){
  var raw=trim(r[fd.c]);
  if(raw==='')return '';
  if(raw==='-')return null;
  return fd.read?fd.read(raw):raw;
}
function materialRows(list){
  var rows=[MAT_COLS.slice()];
  list.forEach(function(m){
    var v=m.mfr?mfr(m.mfr):null;
    rows.push([String(m.id),matNo(m),m.name||'',m.cat||'',m.disc||'',v?v.name:'',v?(pqOf(v).ref||''):'',m.sub||'',
      m.qty||'',m.unit||'',(m.raw||{})['PO Number']||'',stepOf(m,'mts','status')||(m.raw||{})['MAT Status']||'']
      .concat(MAT_FIELDS.filter(function(fd){return fd.log||fd.docKind;}).map(function(fd){
        var x=fd.get(m);
        if(fd.docKind)return x;
        return (fd.log[1]==='date'&&/^\d{4}-\d{2}-\d{2}$/.test(String(x)))?{date:String(x)}:x;
      })));
  });
  return rows;
}
async function readMaterialSheet(file){
  var book=await openBook(file);
  for(var i=0;i<book.sheets.length;i++){
    var rows=await book.rows(book.sheets[i]);
    for(var h=0;h<Math.min(rows.length,10);h++){
      var head=(rows[h]||[]).map(function(x){return K(x);});
      if(head.indexOf('item description')>=0&&(head.indexOf('id')>=0||head.indexOf('mat number')>=0)){
        var col={};head.forEach(function(n,j){if(n)col[n]=j;});
        var out=[];
        rows.slice(h+1).forEach(function(line){
          var o={};MAT_COLS.forEach(function(c){var j=col[K(c)];o[c]=(j==null)?'':line[j];});
          if(trim(o['ID'])||trim(o['MAT Number']))out.push(o);
        });
        if(out.length)return {rows:out};
      }
    }
  }
  throw new Error('No sheet in that file has the materials columns. Download it with Edit in Excel first.');
}
function planMaterials(rows){
  var byId={},byNo={};
  (DB.mats||[]).forEach(function(m){if(isDoc(m))return;byId[String(m.id)]=m;var n=K(matNo(m));if(n&&!byNo[n])byNo[n]=m;});
  var p={change:[],same:0,lost:[],noVendor:[],noDoc:[]};
  rows.forEach(function(r){
    var m=byId[trim(r['ID'])]||byNo[K(trim(r['MAT Number']))];
    if(!m){p.lost.push(r);return;}
    var diff=[];
    MAT_FIELDS.forEach(function(fd){
      var want=matWant(r,fd), now=trim(fd.get(m));
      if(want===null){if(now!=='')diff.push(fd.c+' cleared');return;}
      if(want===''||K(now)===K(want))return;
      if(fd.c==='Vendor'&&!vendorByName(want)){p.noVendor.push({m:m,name:want});return;}
      if(fd.c==='Vendor PQD'&&!vendorByPq(want)){p.noVendor.push({m:m,name:want,pq:true});return;}
      if(fd.docKind){
        var found=docsInText(want,fd.docKind);
        if(!found.length){p.noDoc.push({m:m,name:want,kind:fd.docKind});return;}
        var have={};docsOf(m).forEach(function(d){have[String(d.id)]=1;});
        var fresh=found.filter(function(d){return !have[String(d.id)];});
        if(fresh.length)diff.push(fd.docKind+' +'+fresh.length);
        return;
      }
      diff.push(fd.c);
    });
    if(diff.length)p.change.push({m:m,r:r,diff:diff});else p.same++;
  });
  return p;
}
var MATP=null;
function showMaterials(){
  var p=MATP;
  sheet('From '+p.file,
    '<div class="dim" style="font-size:13.5px;margin-bottom:16px">'+p.count+' rows read. MAT numbers and '
    +'statuses come from Aconex and are not changed from here. Nothing has been changed yet.</div>'
    +'<div class="grid" style="margin-bottom:18px">'+stat(p.change.length,'Changed')+stat(p.same,'Unchanged')
    +stat(p.noVendor.length,'Vendor not found')+stat(p.lost.length,'No such material')+'</div>'
    +(p.change.length?('<div class="sec">Changed</div><div class="panel"><div class="panel-b">'
      +p.change.slice(0,60).map(function(c){
        return '<div class="line"><span class="tag t-wait">'+c.diff.length+'</span><div class="line-m"><div>'+esc(c.m.name)
          +'</div><div class="dim" style="font-size:12.5px;margin-top:2px">'+esc(c.diff.join(', '))+'</div></div></div>';}).join('')
      +more(p.change.length,60)+'</div></div>'):'')
    +(p.noVendor.length?('<div class="sec">Vendor not found — left as it is</div><div class="panel"><div class="panel-b">'
      +p.noVendor.slice(0,40).map(function(x){
        return '<div class="line"><span class="tag t-now">no match</span><div class="line-m"><div>'+esc(x.m.name)
          +'</div><div class="dim" style="font-size:12.5px;margin-top:2px">"'+esc(x.name)+'" '
          +(x.pq?'is no vendor’s PQD number here':'is not a vendor here — write it as it appears on the Vendors page')+'</div></div></div>';}).join('')
      +more(p.noVendor.length,40)+'</div></div>'):'')
    +(p.noDoc.length?('<div class="sec">Document number not found — left as it is</div><div class="panel"><div class="panel-b">'
      +p.noDoc.slice(0,40).map(function(x){
        return '<div class="line"><span class="tag t-now">'+esc(x.kind)+'</span><div class="line-m"><div>'+esc(x.m.name)
          +'</div><div class="dim" style="font-size:12.5px;margin-top:2px">"'+esc(x.name)+'" matches no '+esc(x.kind)
          +' in the tracker \u2014 write the full number as in Aconex</div></div></div>';}).join('')
      +more(p.noDoc.length,40)+'</div></div>'):'')
    +(p.lost.length?('<div class="sec">No such material — ignored</div><div class="panel"><div class="panel-b">'
      +'<div class="dim" style="font-size:12.5px;margin-bottom:8px">Materials come from Aconex; a row that matches none is not added.</div>'
      +p.lost.slice(0,20).map(function(r){return '<div class="line"><span class="tag t-na">'+esc(trim(r['MAT Number'])||'no number')
        +'</span><div class="line-m">'+esc(trim(r['Item Description']))+'</div></div>';}).join('')
      +more(p.lost.length,20)+'</div></div>'):'')
    +'<div class="f-act" style="margin-top:22px">'
    +(p.change.length?('<button class="btn btn-p" onclick="matApply()">Apply — '+p.change.length+'</button>'):'')
    +'<button class="btn-q" onclick="closeSheet()">Cancel</button></div>');
}
window.matApply=function(){
  if(!MATP)return;
  MATP.change.forEach(function(c){
    MAT_FIELDS.forEach(function(fd){
      var want=matWant(c.r,fd);
      if(want==='')return;
      if(fd.c==='Vendor'&&want!==null&&!vendorByName(want))return;
      if(fd.c==='Vendor PQD'&&want!==null&&!vendorByPq(want))return;
      fd.set(c.m,want===null?'':want);
    });
  });
  var n=MATP.change.length;MATP=null;
  closeSheet();touch();rList();rPane();
  toast(n+' material'+(n===1?'':'s')+' updated');
};
/* the edit sheet's columns by group, coloured like the Main Log's */
function matBands(head){
  var b=[{t:'Material',from:0,to:11}];
  var di=head.indexOf('ITP Number'), dm=head.indexOf('MES Number');
  if(di>=0&&dm>=0)b.push({t:'ITP / MES',from:Math.min(di,dm),to:Math.max(di,dm)});
  LOG_FORM.forEach(function(g){
    var cols=g[1].map(function(f){return head.indexOf(f[0]);}).filter(function(i){return i>11;});
    if(cols.length)b.push({t:g[0],from:Math.min.apply(null,cols),to:Math.max.apply(null,cols)});
  });
  return b;
}
window.matOut=function(){
  var prev=TBL;TBL='mat';
  var list=filtered();TBL=prev;
  var mr=materialRows(list);
  download(workbook([{name:'Materials',rows:mr,table:true,bands:matBands(mr[0])}]),
    (DB.project||'Materials')+' — materials to edit '+today()+'.xlsx');
  toast(list.length+' materials written — edit, keep the ID column, then Upload edited');
};
window.venOut=function(){
  var prev=TBL;TBL='mfr';
  var list=filtered();TBL=prev;
  download(workbook([{name:'Vendors',rows:vendorRows(list)}]),
    (DB.project||'Vendors')+' — vendors to edit '+today()+'.xlsx');
  toast(list.length+' vendors written — edit, keep the ID column, then Upload edited');
};
window.pqShared=function(){
  var gs=pqGroups();
  if(!gs.length){closeSheet();return toast('Every pre-qualification number belongs to one vendor');}
  sheet('Shared pre-qualification numbers — '+gs.length,
    '<div class="dim" style="font-size:13.5px;line-height:1.7;margin-bottom:16px">Each of these numbers is held by '
    +'more than one vendor. A pre-qualification belongs to one company; the others usually came onto the '
    +'project through it. Pick the holder of each — the likely one is already picked. Applying keeps the '
    +'number on the holder only, and marks each of the others as <b>brought by</b> the holder where nothing '
    +'else is recorded. Nothing changes until you apply.</div>'
    +gs.map(function(g,i){
      var guess=pqGuess(g);
      return '<div class="sec"><span class="mono">'+esc(g.ref)+'</span> <span class="dim">'+g.list.length+' vendors</span></div>'
        +'<div class="panel"><div class="panel-b">'
        +g.list.map(function(v){
          return '<label class="line" style="cursor:pointer"><input type="radio" name="pqh'+i+'" value="'+v.id+'"'
            +(v===guess?' checked':'')+' style="margin-right:10px">'
            +'<div class="line-m"><div>'+esc(v.name)+' <span class="dim">· '+esc(KINDS[kindOf(v)].l)+'</span></div>'
            +(v.scope?'<div class="dim" style="font-size:12.5px;margin-top:2px">'+esc(v.scope)+'</div>':'')+'</div>'
            +'<span class="meta">'+esc(pqOf(v).status||'')+'</span></label>';
        }).join('')+'</div></div>';
    }).join('')
    +'<div class="f-act" style="margin-top:20px"><button class="btn btn-p" onclick="pqSharedApply()">Apply to all '+gs.length+'</button>'
    +'<button class="btn-q" onclick="closeSheet()">Cancel</button></div>');
};
window.pqSharedApply=function(){
  var gs=pqGroups(), moved=0;
  gs.forEach(function(g,i){
    var pick=document.querySelector('input[name="pqh'+i+'"]:checked');
    var holder=pick?g.list.filter(function(v){return String(v.id)===pick.value;})[0]:pqGuess(g);
    if(!holder)return;
    g.list.forEach(function(v){
      if(v===holder)return;
      var slot=(v.kind==='agency')?'appr':'pqd';
      if(v.steps&&v.steps[slot])delete v.steps[slot];
      if(!v.by)v.by=holder.name;
      moved++;
    });
  });
  closeSheet();touch();rList();rPane();
  toast(moved+' vendor'+(moved===1?'':'s')+' no longer carry a number that is not theirs');
};
function pqOf(v){
  var st=v.steps||{};
  return ((v.kind==='agency')?st.appr:st.pqd)||{};
}
/* Local or Foreign, and nothing else — "local", "KSA", "Saudi", "L" all
   mean the first; "foreign", "import", "F" the second. A column people
   filter on has to hold two values, not twelve spellings of two. */
function localityOf(x){
  var k=K(x);
  if(!k)return '';
  if(/^(l|local|ksa|saudi|saudi arabia|domestic|in-kingdom)$/.test(k))return 'Local';
  if(/^(f|foreign|import|imported|international|overseas|outside)$/.test(k))return 'Foreign';
  return '';
}
function vendorRows(list){
  var rows=[VEN_COLS.slice()];
  (list||DB.mfrs||[]).forEach(function(v){
    var st=v.steps||{}, pq=pqOf(v), iso=st.iso||{};
    var pa=st.pa||{};
    /* The ID goes out as text. It has sixteen digits and Excel keeps
       fifteen, so as a number it came back rounded — and found nothing. */
    rows.push([String(v.id),v.name,KINDS[kindOf(v)].l,v.cat||'',v.by||'',v.country||'',v.locality||'',v.site||'',v.scope||'',
      pq.ref||'',pq.status||'',pq.date?{date:pq.date}:'',
      iso.ref||'',iso.date?{date:iso.date}:'',iso.status||'',
      pa.status||'',pa.date?{date:pa.date}:'',pa.by||'',pa.by2||'',pa.ref||'',
      (pa.ref2&&pa.ref2!==pa.ref)?pa.ref2:'',matsOf(v).length]);
  });
  return rows;
}

/* reading it back: the identifier first, then the pre-qualification
   number, then the name — each a weaker claim than the one before */
async function readVendorSheet(file){
  var book=await openBook(file);
  for(var i=0;i<book.sheets.length;i++){
    var rows=await book.rows(book.sheets[i]);
    for(var h=0;h<Math.min(rows.length,10);h++){
      var head=(rows[h]||[]).map(function(x){return K(x);});
      if(head.indexOf('vendor')>=0&&head.indexOf('kind')>=0){
        var col={};head.forEach(function(n,j){if(n)col[n]=j;});
        var out=[];
        rows.slice(h+1).forEach(function(line){
          var name=trim(line[col['vendor']]);
          if(!name)return;
          var o={};
          VEN_COLS.forEach(function(c){
            var j=col[K(c)];
            if(j==null&&VEN_ALIAS[c])j=col[K(VEN_ALIAS[c])];
            o[c]=(j==null)?'':line[j];
          });
          out.push(o);
        });
        if(out.length)return {rows:out,sheet:book.sheets[i].name};
      }
    }
  }
  throw new Error('No sheet in that file has a Vendor and a Kind column. '
    +'This reads the vendor report back — download it from Reports first.');
}

function planVendors(rows){
  var byId={},byPq={},byName={};
  (DB.mfrs||[]).forEach(function(v){
    byId[String(v.id)]=v;
    var pq=pqOf(v).ref;
    if(pq&&!byPq[K(pq)])byPq[K(pq)]=v;
    byName[K(v.name)]=v;
  });
  var p={add:[],change:[],same:[],gone:[]};
  var hit={};
  rows.forEach(function(r){
    var v=byId[String(trim(r['ID']))]||byPq[K(trim(r['PQD Number']))]||byName[K(trim(r['Vendor']))];
    if(!v){p.add.push(r);return;}
    hit[String(v.id)]=1;
    var diff=vendorDiff(v,r);
    if(diff.length)p.change.push({v:v,r:r,diff:diff});
    else p.same.push(v);
  });
  (DB.mfrs||[]).forEach(function(v){if(!hit[String(v.id)])p.gone.push(v);});
  return p;
}
/* Every editable column, once: which heading, where it lives on the
   vendor, and how a typed value is read. The comparison and the writing
   both walk this one list, so they cannot disagree about a column.
   A blank cell says nothing and leaves the value alone; a cell holding
   only "-" clears it. */
var VEN_FIELDS=[
  {c:'Vendor',get:function(v){return v.name;},set:function(v,x){if(x)v.name=x;}},
  {c:'Kind',get:function(v){return KINDS[kindOf(v)].l;},
    read:function(x){var k=VEN_KINDS[K(x)];return k?KINDS[k].l:'';},
    set:function(v,x){var k=VEN_KINDS[K(x)];if(k)v.kind=k;}},
  {c:'Category',get:function(v){return v.cat;},
    read:function(x){return normCat(x)||(/^[0-3]$/.test(trim(x))?'C'+trim(x):'');},
    set:function(v,x){v.cat=x;}},
  {c:'Brought by',get:function(v){return v.by;},set:function(v,x){v.by=x;}},
  {c:'Country',get:function(v){return v.country;},set:function(v,x){v.country=x;}},
  {c:'Local / Foreign',get:function(v){return v.locality;},read:localityOf,set:function(v,x){v.locality=x;}},
  {c:'Production site',get:function(v){return v.site;},set:function(v,x){v.site=x;}},
  {c:'Scope',get:function(v){return v.scope;},set:function(v,x){v.scope=x;}},
  {c:'PQD Number',step:'pq',f:'ref'},
  {c:'PQD Status',step:'pq',f:'status',read:pqWord},
  {c:'PQD Date',step:'pq',f:'date',date:true},
  {c:'ISO Number',step:'iso',f:'ref'},
  {c:'ISO Expires',step:'iso',f:'date',date:true},
  {c:'ISO Status',step:'iso',f:'status',read:isoWordIn},
  {c:'Visit result',step:'pa',f:'status',read:visitWordIn},
  {c:'Visit date',step:'pa',f:'date',date:true},
  {c:'QMS auditor',step:'pa',f:'by'},
  {c:'Technical expert',step:'pa',f:'by2'},
  {c:'Visit report',step:'pa',f:'ref'},
  {c:'Technical expert report',step:'pa',f:'ref2'}
];
function venSlot(v,step){return step==='pq'?((v.kind==='agency')?'appr':'pqd'):step;}
function venNow(v,fd){
  if(fd.get)return trim(fd.get(v));
  var d=((v.steps||{})[venSlot(v,fd.step)])||{};
  return trim(d[fd.f]);
}
/* what a cell asks for: '' says nothing, null clears, anything else is
   the value in the page's own words */
function venWant(r,fd){
  var raw=trim(r[fd.c]);
  if(raw==='')return '';
  if(raw==='-')return null;
  if(fd.date)return anyDate(r[fd.c])||'';
  return fd.read?fd.read(raw):raw;
}
function vendorDiff(v,r){
  var out=[];
  VEN_FIELDS.forEach(function(fd){
    var want=venWant(r,fd), now=venNow(v,fd);
    if(want===null){if(now!=='')out.push(fd.c+' cleared');return;}
    if(want!==''&&K(now)!==K(want))out.push(fd.c);
  });
  return out;
}
function applyVendors(p){
  var made={n:1,vendors:[],id:idMaker()};
  function write(v,r){
    v.steps=v.steps||{};
    /* Kind first: it decides which slot the pre-qualification lives in */
    VEN_FIELDS.slice().sort(function(a,b){return (b.c==='Kind')-(a.c==='Kind');}).forEach(function(fd){
      var want=venWant(r,fd);
      if(want==='')return;
      if(fd.get){if(want===null){if(fd.c!=='Vendor'&&fd.c!=='Kind')fd.set(v,'');}else fd.set(v,want);return;}
      var slot=venSlot(v,fd.step), d=v.steps[slot]||{};
      if(want===null)delete d[fd.f];else d[fd.f]=want;
      if(Object.keys(d).length)v.steps[slot]=d;else delete v.steps[slot];
    });
    /* a certificate with a date and no word said about it is valid until
       that date, which is what a certificate means */
    var iso=v.steps.iso;
    if(iso&&(iso.ref||iso.date)&&!iso.status)iso.status='Valid';
    /* the same visit report typed twice is one report */
    var pa=v.steps.pa;
    if(pa&&pa.ref2&&pa.ref2===pa.ref)delete pa.ref2;
  }
  p.change.forEach(function(c){write(c.v,c.r);});
  p.add.forEach(function(r){
    var v={id:made.id(),name:trim(r['Vendor']),kind:VEN_KINDS[K(r['Kind'])]||'maker',
      cat:'',country:'',site:'',scope:'',steps:{},pq:{},added:today()};
    write(v,r);
    DB.mfrs.push(v);
  });
  touch();rList();rPane();
  return {changed:p.change.length,added:p.add.length};
}

/* ================================================================
   TWO WEEK LOOK-AHEAD SCHEDULE (TWLAS)
   ----------------------------------------------------------------
   The project's own layout: this week and next, Saturday to Friday, a
   column a day; each section a grey band, each item a row with its
   discipline, and a document mark on every day something falls due or
   is planned. A section with nothing in the fortnight says "No update".
   ================================================================ */
var TW_SECTIONS=[['1','Physical Assessment'],['2','Pre-Qualification (PQD)'],['6','Pre-inspection Dossier'],
  ['7','Pre-inspection Meeting'],['8','Post-inspection Dossier'],['9','FAT/Final Inspection'],['10','Materials Release']];
var TW_MARK='🗎';            /* the document mark */
/* the discipline as the project writes it: the code in the document
   number (…-MBL-ST-MAT-…), or the first letters of the discipline */
function twDisc(no,disc){
  var m=/-([A-Z]{2,3})-(?:MAT|PRQ|ITP|PID|MIR|MES|MAS|IRN|WIR)-/.exec(String(no||'').toUpperCase());
  if(m)return m[1];
  var d=trim(disc);
  var c=/^([A-Z]{2,3})\s*-/.exec(d);
  return c?c[1]:d.slice(0,2).toUpperCase();
}
function twStart(){
  var d=new Date(today()+'T00:00:00');
  d.setDate(d.getDate()-((d.getDay()+1)%7));              /* back to Saturday */
  return d.getFullYear()+'-'+('0'+(d.getMonth()+1)).slice(-2)+'-'+('0'+d.getDate()).slice(-2);
}
function twItems(t0,end){
  var sec={};TW_SECTIONS.forEach(function(x){sec[x[0]]=[];});
  function inWin(d){return d&&d>=t0&&d<=end;}
  function put(k,label,disc,dates){
    dates=dates.filter(inWin);
    if(dates.length)sec[k].push({label:label,disc:disc,dates:dates});
  }
  /* the day of the thing itself, and a real deadline after it — not a
     day something merely becomes allowed (a purchase order may follow,
     fabrication may start), which is no activity on the schedule */
  /* the day it went in, and nothing after it */
  function stepDates(rec,st,d){
    return [d.date];
  }
  (DB.mfrs||[]).forEach(function(v){
    var stp=v.steps||{}, pq=pqOf(v), code=twDisc(pq.ref,'');
    var pa=stp.pa||{};
    /* the tentative date — the visit as planned, before the official one
       and its report; failing that, a survey recorded as Scheduled */
    var tent=(v.logf||{})['PA Tentative Date']||(pa.status==='Scheduled'?pa.date:'');
    if(tent)put('1',v.name+(tent>=today()?' ( upcoming )':''),code,[tent]);
    var pqs=MFR_ROAD.filter(function(x){return x.k==='pqd';})[0];
    /* a first submission only: a resubmission (revision 1 and on) is not new */
    var first=pq.rev==null||pq.rev===''||/^0*$/.test(String(pq.rev).trim());
    if(first&&(pq.date||pq.ref))put('2',v.name+(pq.status?' — '+pq.status:''),code,stepDates(v,pqs,pq));
  });
  (DB.mats||[]).forEach(function(m){
    if(isDoc(m))return;
    var code=twDisc(matNo(m),m.disc), road=matRoad(m).steps;
    var step=function(k){return road.filter(function(x){return x.s.k===k;})[0];};
    [['pid','6'],['pfm','7'],['post','8']].forEach(function(p){
      var x=step(p[0]);if(!x)return;
      put(p[1],m.name,code,stepDates(m,x.s,x.data));
    });
    var vs=function(k){return (m.visits||[]).filter(function(y){return y.step===k;}).map(function(y){return y.date;});};
    if(step('fat'))put('9',m.name,code,vs('fat').concat(stepDates(m,step('fat').s,step('fat').data)));
    /* materials released on site: each inspection request (consignment)
       on its day, its number beside the material */
    if(m.cat==='C2'||m.cat==='C3')(m.dels||[]).forEach(function(c){
      if(c.date)put('10',m.name+(c.ref?' — '+c.ref:''),code,[c.date]);
    });
  });
  /* and every inspection request from Aconex not linked to a material
     yet, by its own title and number, on the day it went in */
  (DB.mats||[]).forEach(function(d){
    if(d.doc!=='MIR'||isLinked(d)||(d.cat!=='C2'&&d.cat!=='C3'))return;          /* C2 and C3 only */
    var no=refOf(d);
    put('10',d.name+(no?' — '+no:''),twDisc(no,d.disc),[acxDateOf(d)]);
  });
  return sec;
}
function twlasBook(){
  var t0=twStart(), days=[];
  for(var i=0;i<14;i++)days.push(addDays(t0,i));
  var end=days[13], sec=twItems(t0,end);
  var DAY=['SUN','MON','TUE','WED','THU','FRI','SAT'];
  var styles='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    +'<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
    +'<fonts count="4"><font><sz val="11"/><name val="Calibri"/></font>'
    +'<font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="Calibri"/></font>'
    +'<font><b/><sz val="11"/><name val="Calibri"/></font>'
    +'<font><sz val="14"/><name val="Segoe UI Symbol"/></font></fonts>'
    +'<fills count="8"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>'
    +['FF808080','FFC6E0B4','FFDDEBF7','FFFFC000','FFA6A6A6','FFBFBFBF'].map(function(c){
      return '<fill><patternFill patternType="solid"><fgColor rgb="'+c+'"/><bgColor indexed="64"/></patternFill></fill>';}).join('')
    +'</fills>'
    +'<borders count="2"><border/><border><left style="thin"><color rgb="FF595959"/></left><right style="thin"><color rgb="FF595959"/></right>'
    +'<top style="thin"><color rgb="FF595959"/></top><bottom style="thin"><color rgb="FF595959"/></bottom></border></borders>'
    +'<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
    +'<cellXfs count="13">'
    +'<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
    /* 1 heading, dark */      +'<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>'
    /* 2 week 1 */             +'<xf numFmtId="0" fontId="2" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>'
    /* 3 week 2 */             +'<xf numFmtId="0" fontId="2" fillId="4" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>'
    /* 4 day, yellow */        +'<xf numFmtId="0" fontId="0" fillId="5" borderId="1" xfId="0" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>'
    /* 5 Friday, grey */       +'<xf numFmtId="0" fontId="0" fillId="2" borderId="1" xfId="0" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>'
    /* 6 section number */     +'<xf numFmtId="0" fontId="2" fillId="6" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>'
    /* 7 section title */      +'<xf numFmtId="0" fontId="2" fillId="6" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>'
    /* 8 item number */        +'<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>'
    /* 9 item text */          +'<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center" wrapText="1"/></xf>'
    /* 10 discipline */        +'<xf numFmtId="0" fontId="0" fillId="7" borderId="1" xfId="0" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="left" vertical="center"/></xf>'
    /* 11 a day */             +'<xf numFmtId="0" fontId="3" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>'
    /* 12 a Friday */          +'<xf numFmtId="0" fontId="3" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>'
    +'</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>';
  var rows=[], merges=[], r=0;
  function txt(ref,v,st){return v===''||v==null?'<c r="'+ref+'" s="'+st+'"/>'
    :'<c r="'+ref+'" s="'+st+'" t="inlineStr"><is><t xml:space="preserve">'+xml(String(v))+'</t></is></c>';}
  function dayCol(i){return colName(3+i);}
  /* the three heading rows */
  r=1;
  rows.push('<row r="1" ht="20" customHeight="1">'+txt('A1','S No',1)+txt('B1','DOCUMENT DESCRIPTION',1)+txt('C1','Discipline',1)
    +days.map(function(_,i){return i===0?txt(dayCol(0)+'1','Week 1',2):i===7?txt(dayCol(7)+'1','Week 2',3):txt(dayCol(i)+'1','',i<7?2:3);}).join('')+'</row>');
  rows.push('<row r="2" ht="18" customHeight="1">'+txt('A2','',1)+txt('B2','',1)+txt('C2','',1)
    +days.map(function(d,i){var fri=new Date(d+'T00:00:00').getDay()===5;
      return txt(dayCol(i)+'2',DAY[new Date(d+'T00:00:00').getDay()],fri?5:4);}).join('')+'</row>');
  rows.push('<row r="3" ht="18" customHeight="1">'+txt('A3','',1)+txt('B3','',1)+txt('C3','',1)
    +days.map(function(d,i){var fri=new Date(d+'T00:00:00').getDay()===5;
      return '<c r="'+dayCol(i)+'3" s="'+(fri?5:4)+'"><v>'+(+d.slice(8,10))+'</v></c>';}).join('')+'</row>');
  merges.push('A1:A3','B1:B3','C1:C3',dayCol(0)+'1:'+dayCol(6)+'1',dayCol(7)+'1:'+dayCol(13)+'1');
  r=3;
  TW_SECTIONS.forEach(function(x){
    r++;
    rows.push('<row r="'+r+'" ht="18" customHeight="1">'+txt('A'+r,x[0],6)+txt('B'+r,x[1],7)+txt('C'+r,'',7)
      +days.map(function(_,i){return txt(dayCol(i)+r,'',7);}).join('')+'</row>');
    merges.push('B'+r+':'+dayCol(13)+r);
    var items=sec[x[0]];
    if(!items.length)items=[{label:'No update',disc:'',dates:[]}];
    items.forEach(function(it,j){
      r++;
      var no=x[0]+'.'+('0'+(j+1)).slice(-2);
      rows.push('<row r="'+r+'" ht="20" customHeight="1">'+txt('A'+r,no,8)+txt('B'+r,it.label,9)+txt('C'+r,it.disc,10)
        +days.map(function(d,i){var fri=new Date(d+'T00:00:00').getDay()===5;
          return txt(dayCol(i)+r,it.dates.indexOf(d)>=0?TW_MARK:'',fri?12:11);}).join('')+'</row>');
    });
  });
  /* the period, at the foot, as the project's sheet has it */
  r++;
  var fmt=function(d){var x=new Date(d+'T00:00:00');
    return ('0'+x.getDate()).slice(-2)+'-'+['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'][x.getMonth()]+'-'+x.getFullYear();};
  rows.push('<row r="'+r+'" ht="22" customHeight="1">'+txt('A'+r,'',1)+txt('B'+r,'',1)+txt('C'+r,'',1)
    +days.map(function(_,i){return txt(dayCol(i)+r,i===0?fmt(t0)+' to '+fmt(end):'',1);}).join('')+'</row>');
  merges.push('A'+r+':C'+r,dayCol(0)+r+':'+dayCol(13)+r);
  var sheet='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    +'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
    +'<sheetViews><sheetView workbookViewId="0" showGridLines="0"><pane xSplit="3" ySplit="3" topLeftCell="D4" activePane="bottomRight" state="frozen"/></sheetView></sheetViews>'
    +'<cols><col min="1" max="1" width="8" customWidth="1"/><col min="2" max="2" width="70" customWidth="1"/>'
    +'<col min="3" max="3" width="11" customWidth="1"/><col min="4" max="17" width="9" customWidth="1"/></cols>'
    +'<sheetData>'+rows.join('')+'</sheetData>'
    +'<mergeCells count="'+merges.length+'">'+merges.map(function(m){return '<mergeCell ref="'+m+'"/>';}).join('')+'</mergeCells>'
    +'<pageMargins left="0.4" right="0.4" top="0.5" bottom="0.5" header="0.3" footer="0.3"/>'
    +'<pageSetup orientation="landscape" fitToHeight="0"/></worksheet>';
  return zipUp([
    {name:'[Content_Types].xml',text:'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      +'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      +'<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      +'<Default Extension="xml" ContentType="application/xml"/>'
      +'<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
      +'<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
      +'<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
      +'</Types>'},
    {name:'_rels/.rels',text:'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      +'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      +'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
      +'</Relationships>'},
    {name:'xl/workbook.xml',text:'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      +'<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
      +'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>'
      +'<sheet name="TWLAS" sheetId="1" r:id="rId1"/></sheets></workbook>'},
    {name:'xl/_rels/workbook.xml.rels',text:'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      +'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      +'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
      +'<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
      +'</Relationships>'},
    {name:'xl/styles.xml',text:styles},
    {name:'xl/worksheets/sheet1.xml',text:sheet}
  ]);
}
/* ================================================================
   A WORKBOOK OF ONE STYLED SHEET
   ----------------------------------------------------------------
   The project's own sheets carry their own colours. A style is named
   once ({fill, color, b, sz, u, i, align, wrap, border}) and the
   fonts, fills and formats Excel needs are made from the names used.
   ================================================================ */
function styledBook(spec){
  var fonts=['<font><sz val="10"/><name val="Calibri"/></font>'], fills=['<fill><patternFill patternType="none"/></fill>','<fill><patternFill patternType="gray125"/></fill>'];
  var xfs=['<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'], idx={};
  Object.keys(spec.styles).forEach(function(k){
    var s=spec.styles[k];
    var f='<font>'+(s.b?'<b/>':'')+(s.i?'<i/>':'')+(s.u?'<u/>':'')+'<sz val="'+(s.sz||10)+'"/>'
      +(s.color?'<color rgb="FF'+s.color+'"/>':'')+'<name val="Calibri"/></font>';
    var fi=fonts.indexOf(f);if(fi<0){fonts.push(f);fi=fonts.length-1;}
    var fl=0;
    if(s.fill){var x='<fill><patternFill patternType="solid"><fgColor rgb="FF'+s.fill+'"/><bgColor indexed="64"/></patternFill></fill>';
      fl=fills.indexOf(x);if(fl<0){fills.push(x);fl=fills.length-1;}}
    xfs.push('<xf numFmtId="0" fontId="'+fi+'" fillId="'+fl+'" borderId="'+(s.border===false?0:1)+'" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">'
      +'<alignment horizontal="'+(s.align||'center')+'" vertical="center"'+(s.wrap===false?'':' wrapText="1"')+'/></xf>');
    idx[k]=xfs.length-1;
  });
  var styles='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    +'<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
    +'<fonts count="'+fonts.length+'">'+fonts.join('')+'</fonts><fills count="'+fills.length+'">'+fills.join('')+'</fills>'
    +'<borders count="2"><border/><border><left style="thin"><color rgb="FF8EA9C1"/></left><right style="thin"><color rgb="FF8EA9C1"/></right>'
    +'<top style="thin"><color rgb="FF8EA9C1"/></top><bottom style="thin"><color rgb="FF8EA9C1"/></bottom></border></borders>'
    +'<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
    +'<cellXfs count="'+xfs.length+'">'+xfs.join('')+'</cellXfs>'
    +'<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>';
  var rows=spec.rows.map(function(row,r){
    var ht=(spec.heights||{})[r+1];
    return '<row r="'+(r+1)+'"'+(ht?' ht="'+ht+'" customHeight="1"':'')+'>'+row.map(function(c,j){
      if(!c)return '';
      var ref=colName(j)+(r+1), s=idx[c.st]||0, v=c.v;
      if(v==null||v==='')return '<c r="'+ref+'" s="'+s+'"/>';
      if(typeof v==='number')return '<c r="'+ref+'" s="'+s+'"><v>'+v+'</v></c>';
      return '<c r="'+ref+'" s="'+s+'" t="inlineStr"><is><t xml:space="preserve">'+xml(String(v))+'</t></is></c>';
    }).join('')+'</row>';
  }).join('');
  var sheet='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    +'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
    +'<sheetViews><sheetView workbookViewId="0" showGridLines="0">'
    +(spec.freeze?'<pane xSplit="'+(spec.freeze.x||0)+'" ySplit="'+spec.freeze.y+'" topLeftCell="'+colName(spec.freeze.x||0)+(spec.freeze.y+1)+'" activePane="bottomRight" state="frozen"/>':'')
    +'</sheetView></sheetViews>'
    +'<cols>'+spec.cols.map(function(w,i){return '<col min="'+(i+1)+'" max="'+(i+1)+'" width="'+w+'" customWidth="1"/>';}).join('')+'</cols>'
    +'<sheetData>'+rows+'</sheetData>'
    +(spec.filter?'<autoFilter ref="'+spec.filter+'"/>':'')
    +(spec.merges&&spec.merges.length?'<mergeCells count="'+spec.merges.length+'">'+spec.merges.map(function(m){return '<mergeCell ref="'+m+'"/>';}).join('')+'</mergeCells>':'')
    +'<pageMargins left="0.4" right="0.4" top="0.5" bottom="0.5" header="0.3" footer="0.3"/>'
    +'<pageSetup orientation="landscape"/></worksheet>';
  return zipUp([
    {name:'[Content_Types].xml',text:'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      +'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      +'<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      +'<Default Extension="xml" ContentType="application/xml"/>'
      +'<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
      +'<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
      +'<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
      +'</Types>'},
    {name:'_rels/.rels',text:'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      +'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      +'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
      +'</Relationships>'},
    {name:'xl/workbook.xml',text:'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      +'<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
      +'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>'
      +'<sheet name="'+xml(spec.name)+'" sheetId="1" r:id="rId1"/></sheets>'
      +(spec.filter?'<definedNames><definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">\''
        +xml(spec.name)+'\'!'+spec.filter.replace(/([A-Z]+)(\d+)/g,'$$$1$$$2')+'</definedName></definedNames>':'')
      +'</workbook>'},
    {name:'xl/_rels/workbook.xml.rels',text:'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      +'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      +'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
      +'<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
      +'</Relationships>'},
    {name:'xl/styles.xml',text:styles},
    {name:'xl/worksheets/sheet1.xml',text:sheet}
  ]);
}
/* 12-Jan-26, as the project's trackers write a date */
function dShort(iso){
  var m=/^(\d{4})-(\d{2})-(\d{2})$/.exec(String(iso||''));if(!m)return iso||'';
  return (+m[3])+'-'+['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'][+m[2]-1]+'-'+m[1].slice(2);
}

/* ---------------- Final Inspection / FAT Schedule Tracker ---------------- */
function fatTrackerBook(){
  var list=(DB.mats||[]).filter(function(m){
    if(isDoc(m))return false;
    var has=(m.visits||[]).some(function(v){return v.step==='fat';})||((m.steps||{}).fat&&((m.steps.fat.date)||(m.steps.fat.status)));
    return has;
  });
  var rowsData=list.map(function(m){
    var v=m.mfr?mfr(m.mfr):null, raw=appRow(m);
    var fats=(m.visits||[]).filter(function(x){return x.step==='fat';}).sort(function(a,b){return String(b.date||'').localeCompare(String(a.date||''));});
    var last=fats[0]||(m.steps||{}).fat||{};
    var res=last.result||last.status||'';
    var by=last.by||'', who=by?(DB.people||[]).filter(function(p){return K(p.name)===K(by);})[0]:null;
    var remark=/Failed/.test(res)?'Re-FAT is required'
      :/Passed/.test(res)?(m.cat==='C2'?'Final inspection done':'FAT done'):'upcoming';
    var first=raw['1st Batch Delivery To Site Actual Date']||raw['1st Batch Delivery To Site Planned Date']||'';
    first=String(first).split('\n')[0];
    var fab=raw['Fabrication 1st Batch Started Date']||raw['Fabrication Planned Date']||'';
    var pfm=((m.steps||{}).pfm||{}).date||'';
    return {m:m,date:last.date||'',remark:remark,cells:[
      m.name,v?v.name:'',m.cat||'',v?[v.site,v.country].filter(Boolean).join(' / '):'',v?v.name:'',
      who?(who.agency||'TBA'):(remark==='upcoming'?'TBA':'N/A'),by||(remark==='upcoming'?'TBA':'N/A'),
      who?(who.status||'Pending'):(remark==='upcoming'?'TBA':'N/A'),
      pfm?dShort(pfm):'',fab?dShort(fab):'',last.date?dShort(last.date):'',first?dShort(first):'',remark]};
  }).sort(function(a,b){
    var A=a.remark==='upcoming'?1:0,B=b.remark==='upcoming'?1:0;
    return A-B||String(a.date).localeCompare(String(b.date));
  });
  var S={
    title:{fill:'9BC2E6',b:true,sz:13,color:'1F1F1F'},sub:{fill:'BDD7EE',b:true,u:true,sz:11,color:'1F1F1F'},
    head:{fill:'1F4E78',b:true,color:'FFFFFF',sz:9},cell:{fill:'F2F2F2'},name:{fill:'F2F2F2'},
    done:{fill:'92D050'},refat:{fill:'F8CBAD'},up:{fill:'FFFF00'},
    note:{color:'FF0000',align:'left',border:false,wrap:false,sz:11},note2:{align:'left',border:false,wrap:false,sz:8}};
  var head=['S#','Material','Manufacturer / Supplier Name','Material Category','Country / Location','Manufacturer / Supplier Assessment',
    'Third Party Inspection Agency','Third Party Inspection Name','TPIA Approval','Pre-fabrication Dates (Kick-off meeting)',
    'Fabrication Dates (Start of Fabrication)','Factory Acceptance Test (FAT) / Final Inspection Date','First Delivery Date','Remark/Comments'];
  var rows=[
    [{v:'Project:   '+(DB.project||''),st:'title'}].concat(head.slice(1).map(function(){return {v:'',st:'title'};})),
    [{v:'Final Inspection  /  FAT Schedule Tracker',st:'sub'}].concat(head.slice(1).map(function(){return {v:'',st:'sub'};})),
    head.map(function(h){return {v:h,st:'head'};})
  ];
  rowsData.forEach(function(x,i){
    rows.push([{v:i+1,st:'cell'}].concat(x.cells.map(function(c,j){
      return {v:c,st:j===x.cells.length-1?(x.remark==='upcoming'?'up':x.remark==='Re-FAT is required'?'refat':'done'):'cell'};})));
  });
  rows.push([{v:'* Proposed dates are tentative dates, actual dates will be updated regularly based on manufacturing progress.',st:'note'}]);
  rows.push([{v:'TBA- To Be Assigned; TBD- To Be Determined',st:'note2'}]);
  var heights={1:22,2:18,3:52};
  for(var i=4;i<4+rowsData.length;i++)heights[i]=30;
  return styledBook({name:'FAT Schedule',styles:S,rows:rows,heights:heights,
    cols:[5,26,16,10,16,20,18,18,11,13,13,18,12,18],
    merges:['A1:N1','A2:N2'],freeze:{x:2,y:3}});
}

/* ---------------- Physical Assessment Summary ---------------- */
function paSummaryBook(){
  var list=(DB.mfrs||[]).filter(function(v){
    var k=kindOf(v);if(k!=='maker'&&k!=='makesub')return false;
    /* only a survey that has taken place: passed or failed */
    var pa=(v.steps||{}).pa||{};
    return /Passed|Failed/.test(pa.status||'');
  }).sort(function(a,b){return String(a.name).localeCompare(String(b.name));});
  var S={
    title:{fill:'F8CBAD',sz:14,color:'843C0C',border:false},grp:{fill:'B4C6E7',b:true,sz:9},
    head:{fill:'B4C6E7',b:true,sz:9},cell:{},left:{align:'left'},bad:{align:'left',color:'FF0000'},
    scope:{fill:'DDEBF7',color:'C00000'},done:{fill:'548235',color:'FFFFFF'},fail:{fill:'F4B6C2'},none:{fill:'FF0000'}};
  var head=['S#','Suppliers proposed for participation','Material Category','Status','Ref Document No','Scope','Location',
    'Assessment Performed','Technical Assessor','Technical report','QMS Assessor','QMS Report','Status','Ref Document No'];
  var rows=[
    head.map(function(_,i){return {v:i===0?'Physical Assessment Summary':'',st:'title'};}),
    head.map(function(_,i){return {v:i===3?'PQD':i===7?'Physical Assessment':'',st:'grp'};}),
    head.map(function(h){return {v:h,st:'head'};})];
  var word=function(s){return /Pass/.test(s||'')?'Pass':/Fail/.test(s||'')?'Failed':(s||'');};
  list.forEach(function(v,i){
    var pq=pqOf(v), pa=(v.steps||{}).pa||{};
    var done=/Passed|Failed|Waived/.test(pa.status||'');
    var overall=word(pa.status);
    var tech=pa.tres||overall, qms=pa.qres||overall;
    var rejected=/Rejected|Terminated/.test(pq.status||'');
    /* the vendor's category, or the highest among its materials */
    var cat=v.cat||matsOf(v).map(function(m){return m.cat||'';}).filter(Boolean).sort().pop()||'';
    rows.push([
      {v:i+1,st:'cell'},{v:v.name,st:rejected?'bad':'left'},{v:cat,st:'cell'},
      {v:pq.status||'',st:pq.ref?'cell':'none'},{v:pq.ref||'',st:pq.ref?'cell':'none'},
      {v:v.scope||'',st:'scope'},{v:[v.site,v.country].filter(Boolean).join(' - '),st:'cell'},
      {v:done?'Done':(pa.status==='Scheduled'?'Scheduled':''),st:done?'done':'cell'},
      {v:pa.by2||'',st:'cell'},{v:tech,st:/Failed/.test(tech)?'fail':'cell'},
      {v:pa.by||'',st:'cell'},{v:qms,st:/Failed/.test(qms)?'fail':'cell'},
      {v:done?overall:'',st:/Failed/.test(overall)?'fail':'cell'},{v:pa.ref||'',st:'cell'}]);
  });
  var heights={1:24,2:16,3:30};
  return styledBook({name:'Physical Assessment',styles:S,rows:rows,heights:heights,
    cols:[4,30,10,18,34,16,16,12,16,13,16,13,9,32],
    merges:['A1:N1','D2:E2','H2:N2'],freeze:{x:2,y:3},filter:'A3:N'+(3+list.length)});
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

/* ---------------------------------------------------------------
   Inspectors, out and back.
   Clause 2.2.17 approves an inspector by reference, and that reference
   is what the sheet is keyed on — a name is spelt three ways across
   four files in this project, but PAA-00086 is one person.
   --------------------------------------------------------------- */
var INS_COLS=['ID','Name','Agency','Discipline','Approval','Approval Reference','Assignments'];

function inspectorRows(){
  var rows=[INS_COLS.slice()];
  (DB.people||[]).forEach(function(p){
    rows.push([p.id,p.name||'',p.agency||'',p.disc||'',p.status||'Pending',p.ref||'',
      assignments(p).length]);
  });
  return rows;
}
async function readInspectorSheet(file){
  var book=await openBook(file);
  for(var i=0;i<book.sheets.length;i++){
    var rows=await book.rows(book.sheets[i]);
    for(var h=0;h<Math.min(rows.length,10);h++){
      var head=(rows[h]||[]).map(function(x){return K(x);});
      if(head.indexOf('name')>=0&&head.indexOf('approval')>=0){
        var col={};head.forEach(function(n,j){if(n)col[n]=j;});
        var out=[];
        rows.slice(h+1).forEach(function(line){
          if(!trim(line[col['name']]))return;
          var o={};
          INS_COLS.forEach(function(c){var j=col[K(c)];o[c]=(j==null)?'':line[j];});
          out.push(o);
        });
        if(out.length)return {rows:out,sheet:book.sheets[i].name};
      }
    }
  }
  throw new Error('No sheet in that file has a Name and an Approval column. '
    +'This reads the inspector report back — download it from Reports first.');
}
function planInspectors(rows){
  var byId={},byRef={},byName={};
  (DB.people||[]).forEach(function(p){
    byId[String(p.id)]=p;
    if(p.ref)byRef[K(p.ref)]=p;
    byName[K(p.name)]=p;
  });
  var p={add:[],change:[],same:[]};
  rows.forEach(function(r){
    var f=byId[String(trim(r['ID']))]||byRef[K(trim(r['Approval Reference']))]
      ||byName[K(trim(r['Name']))];
    if(!f){p.add.push(r);return;}
    var diff=[];
    [['Name','name'],['Agency','agency'],['Discipline','disc'],
     ['Approval','status'],['Approval Reference','ref']].forEach(function(pair){
      var want=trim(r[pair[0]]);
      if(want!==''&&K(f[pair[1]]||'')!==K(want))diff.push(pair[0]);
    });
    if(diff.length)p.change.push({p:f,r:r,diff:diff});else p.same.push(f);
  });
  return p;
}
function applyInspectors(p){
  var made={id:idMaker()};
  function write(x,r){
    function set(k,v){v=trim(v);if(v!=='')x[k]=v;}
    var old=x.name;
    set('name',r['Name']);
    if(old&&K(old)!==K(x.name)){
      /* a name typed onto a step is loose text and would be left
         pointing at somebody who no longer exists under that spelling */
      (DB.mats||[]).concat(DB.mfrs||[]).forEach(function(rec){
        Object.keys(rec.steps||{}).forEach(function(k){
          if(K(rec.steps[k].by)===K(old))rec.steps[k].by=x.name;});
        (rec.visits||[]).forEach(function(v){if(K(v.by)===K(old))v.by=x.name;});
      });
    }
    set('agency',r['Agency']);set('disc',r['Discipline']);
    set('status',r['Approval']);set('ref',r['Approval Reference']);
  }
  p.change.forEach(function(c){write(c.p,c.r);});
  p.add.forEach(function(r){
    var x={id:made.id(),name:'',agency:'',disc:'',status:'Pending',ref:'',added:today()};
    write(x,r);
    DB.people.push(x);
  });
  touch();rList();rPane();
  return {changed:p.change.length,added:p.add.length};
}

/* ================================================================
   SEVEN'S LIVE TRACKING SHEET
   ----------------------------------------------------------------
   The client's own sheet, written from the tracker: its styles, widths,
   merged headings and the colour of each column are SEVEN's template
   (seven-tpl.js); every material is a row under the six heading rows.
   Its 76 columns are filled from the same row the Main Log writes, with
   the few the template has and the log does not — a running number, the
   factory's city, the sample's outcome, whether everything has arrived.
   ================================================================ */
function sevenRow(m,i){
  var r=appRow(m), v=m.mfr?mfr(m.mfr):null;
  var g=function(c){var x=r[c];return x==null?'':x;};
  var unit=function(q,u){q=g(q);return q===''?'':(q+(g(u)?' '+g(u):''));};
  var city=v?(v.site||v.country||''):'';
  var tot=parseFloat(g('Total Quantity')), rem=parseFloat(g('Remaining'));
  return [i+1,'',g('Item Description'),m.cat||'',g('Common Package (Yes/No)'),g('Discipline'),
    g('Finishing (Internal/External)'),g('Package (Lump Sum / Provisional Sum / Prime Cost)'),
    g('Sub-contractor Name'),g('Supply & Install / Supply / Install Only'),g('Manufacturer'),city,
    g('PQD Number'),g('PQD Revision'),g('PQD Status'),g('PQD Submittal Date'),
    g('MAT Number'),g('MAT Submittal Date'),g('MAT Revision'),g('MAT Status'),
    g('PA Tentative Date'),city,g('PA Document Number'),g('PA Date'),
    g('3rd Party Assessment Done'),g('Client Assessment Done'),g('PMC/LDC Assessment Done'),
    g('Contractor Assessment Done'),g('Assessment Result'),
    g('Sample'),g('Mock-up Delivered Date'),g('Mock-up First in Place'),g('Mock-up Approved by PMC'),g('Mock-up per Client DLA'),
    g('Purchase Order Issued (Yes/No)'),g('PO Date'),
    g('Method Statement Number'),g('MES Incl. ITP (Yes/No)'),g('MES Revision'),g('MES Status'),
    g('ITP Number'),g('ITP Submittal Date'),g('ITP Revision'),g('ITP Status'),
    g('PID Number'),g('PID Submittal Date'),g('PID Revision'),g('PID Status'),
    g('Pre-Fabrication Meeting Date'),g('3rd Party Assigned (Yes/No)'),g('3rd Party Service Provider Name'),
    g('Design Verification/Calculation Status'),
    g('Fabrication Planned Date'),g('Fabrication 1st Batch Started Date'),g('Fabrication Percentage Completion (%)'),
    g('FAT Package/Procedure Number/ITP'),g('FAT Package Status'),g('FAT Planned Date'),g('FAT Location / City'),g('FAT/TPI Results'),
    g('1st Batch Delivery To Site Planned Date'),g('1st Batch Delivery To Site Actual Date'),g('Storage (Site/Offsite)'),
    g('MIR Number'),g('MIR Approval Date'),g('MIR Status'),
    g('Installation Planned Date'),g('Installation Actual Date'),g('Installer Name'),
    g('WIR Number'),g('WIR Approval Date'),g('WIR Status'),
    unit('Total Quantity','Total Quantity Unit'),unit('Delivered No','Delivered Unit'),unit('Remaining','Remaining Unit'),
    (isFinite(tot)&&tot>0&&isFinite(rem))?(rem<=0?'Yes':'No'):'']
    .concat(LOG_EXTRA.map(function(x){return x.read(m)||'';}));
}
function sevenCell(ref,v,s){
  if(v==null||v==='')return '<c r="'+ref+'" s="'+s+'"/>';
  if(typeof v==='number'&&isFinite(v))return '<c r="'+ref+'" s="'+s+'"><v>'+v+'</v></c>';
  var t=String(v);
  /* a date reads as dd/mm/yyyy, one to a line where a cell holds several */
  t=t.split('\n').map(function(x){return /^\d{4}-\d{2}-\d{2}$/.test(x)?showDate(x):x;}).join('\n');
  return '<c r="'+ref+'" s="'+s+'" t="inlineStr"><is><t xml:space="preserve">'+xml(t)+'</t></is></c>';
}
/* the Main Log, in SEVEN's format: every material, by discipline */
function sevenMaterials(){
  return (DB.mats||[]).filter(function(m){return !isDoc(m);}).sort(function(a,b){
    var d=String(a.disc||'~').localeCompare(String(b.disc||'~'));
    return d||String(a.name).localeCompare(String(b.name));
  });
}
function sevenLog(){
  var loose=looseRows(), more=[{name:'Summary',rows:summaryRows()}];
  if(loose.length>1)more.push({name:'Not linked',rows:loose});
  return sevenBook(sevenMaterials(),more);
}
/* a plain sheet in the same workbook: grey bold heading, bordered cells */
function sevenPlain(rows){
  var wide=rows.reduce(function(a,r){return Math.max(a,r.length);},0);
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    +'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
    +'<cols><col min="1" max="1" width="44" customWidth="1"/><col min="2" max="'+Math.max(2,wide)+'" width="22" customWidth="1"/></cols>'
    +'<sheetData>'+rows.map(function(row,r){
      return '<row r="'+(r+1)+'">'+row.map(function(v,c){return sevenCell(colName(c)+(r+1),v,r===0?51:2);}).join('')+'</row>';
    }).join('')+'</sheetData></worksheet>';
}
/* A data cell's style: the template's own, except the green and red its
   sample rows were filled with, which were somebody's data and not the
   format — those cells are written plain (bordered, centred, wrapped). */
var SEVEN_PLAIN={3:2,30:2,31:2};
function sevenStyle(T,c){
  var st=c<76?(T.data[c]||0):2;
  return SEVEN_PLAIN[st]!=null?SEVEN_PLAIN[st]:st;
}
function sevenBook(list,more){
  var T=window.SEVEN_TPL;
  if(!T)throw new Error('seven-tpl.js was not found beside the page');
  more=more||[];
  /* SEVEN's 76 columns, then the tracker's own, under a heading of their
     own in the template's styles so the sheet still reads as one */
  var X=LOG_EXTRA.map(function(x){return x.t;}), x0=76, xl=x0+X.length-1;
  var rowsHead=T.head.split('</row>');
  var add=[
    X.map(function(_,i){return '<c r="'+colName(x0+i)+'1" s="5"/>';}).join(''),
    X.map(function(_,i){return '<c r="'+colName(x0+i)+'2" s="7"/>';}).join(''),
    X.map(function(_,i){return i?'<c r="'+colName(x0+i)+'3" s="51"/>'
      :'<c r="'+colName(x0)+'3" s="51" t="inlineStr"><is><t>From the tracker</t></is></c>';}).join(''),
    X.map(function(t,i){return '<c r="'+colName(x0+i)+'4" s="14" t="inlineStr"><is><t>'+xml(t)+'</t></is></c>';}).join(''),
    X.map(function(_,i){return '<c r="'+colName(x0+i)+'5" s="14"/>';}).join(''),
    X.map(function(_,i){return '<c r="'+colName(x0+i)+'6" s="25"/>';}).join('')];
  var head0=rowsHead.slice(0,6).map(function(r,i){return r+add[i];}).join('</row>')+'</row>';
  var merges=T.merges.replace(/<mergeCells count="(\d+)">/,function(_,n){return '<mergeCells count="'+(+n+1+X.length)+'">';})
    .replace('</mergeCells>','<mergeCell ref="'+colName(x0)+'3:'+colName(xl)+'3"/>'
      +X.map(function(_,i){return '<mergeCell ref="'+colName(x0+i)+'4:'+colName(x0+i)+'5"/>';}).join('')+'</mergeCells>');
  var cols=T.cols.replace('</cols>','<col min="'+(x0+1)+'" max="'+(xl+1)+'" width="26" customWidth="1"/></cols>');
  /* one extra style: a terminated material's row, shaded red */
  var st=T.styles;
  var nf=+(/<fills count="(\d+)"/.exec(st)[1]), nx=+(/<cellXfs count="(\d+)"/.exec(st)[1]);
  st=st.replace(/<fills count="\d+">/,'<fills count="'+(nf+1)+'">')
    .replace('</fills>','<fill><patternFill patternType="solid"><fgColor rgb="FFFFC7CE"/><bgColor indexed="64"/></patternFill></fill></fills>')
    .replace(/<cellXfs count="\d+">/,'<cellXfs count="'+(nx+1)+'">')
    .replace('</cellXfs>','<xf numFmtId="0" fontId="5" fillId="'+nf+'" borderId="6" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1">'
      +'<alignment horizontal="center" vertical="center" wrapText="1"/></xf></cellXfs>');
  var RED=nx;
  var rows=list.map(function(m,i){
    var vals=sevenRow(m,i), r=7+i;
    var dead=vals.some(function(x){return /terminat/i.test(String(x));});
    var lines=vals.reduce(function(a,x){return Math.max(a,String(x).split('\n').length);},1);
    return '<row r="'+r+'"'+(lines>1?' ht="'+(lines*15)+'" customHeight="1"':'')+'>'
      +vals.map(function(x,c){return sevenCell(colName(c)+r,x,dead?RED:sevenStyle(T,c));}).join('')+'</row>';
  }).join('');
  var last=6+Math.max(list.length,1);
  var name=(DB.project||'Materials').replace(/[\\\/\?\*\[\]:]/g,' ').slice(0,31)||'Materials';
  var head=head0.replace('%%PROJECT%%',xml((DB.project||'').toUpperCase()));
  var sheetXml='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    +'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
    +'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
    +'<sheetViews><sheetView zoomScale="70" zoomScaleNormal="70" workbookViewId="0">'
    +'<pane xSplit="6" ySplit="6" topLeftCell="G7" activePane="bottomRight" state="frozen"/></sheetView></sheetViews>'
    +'<sheetFormatPr defaultColWidth="8.796875" defaultRowHeight="14.25"/>'
    +cols+'<sheetData>'+head+rows+'</sheetData>'
    +'<autoFilter ref="A6:'+colName(xl)+last+'"/>'+merges
    +'<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/></worksheet>';
  var files=[
    {name:'[Content_Types].xml',text:'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      +'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
      +'<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
      +'<Default Extension="xml" ContentType="application/xml"/>'
      +'<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
      +'<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
      +more.map(function(_,i){return '<Override PartName="/xl/worksheets/sheet'+(i+2)+'.xml" '
        +'ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>';}).join('')
      +'<Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>'
      +'<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
      +'</Types>'},
    {name:'_rels/.rels',text:'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      +'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      +'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
      +'</Relationships>'},
    {name:'xl/workbook.xml',text:'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      +'<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
      +'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>'
      +'<sheet name="'+xml(name)+'" sheetId="1" r:id="rId1"/>'
      +more.map(function(m,i){return '<sheet name="'+xml(m.name)+'" sheetId="'+(i+2)+'" r:id="rId'+(i+10)+'"/>';}).join('')
      +'</sheets>'
      +'<definedNames><definedName name="_xlnm._FilterDatabase" localSheetId="0" hidden="1">\''
      +xml(name.replace(/'/g,"''"))+'\'!$A$6:$'+colName(xl)+'$'+last+'</definedName></definedNames></workbook>'},
    {name:'xl/_rels/workbook.xml.rels',text:'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
      +'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
      +'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
      +'<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>'
      +'<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
      +more.map(function(m,i){return '<Relationship Id="rId'+(i+10)+'" '
        +'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet'+(i+2)+'.xml"/>';}).join('')
      +'</Relationships>'},
    {name:'xl/theme/theme1.xml',text:T.theme},
    {name:'xl/styles.xml',text:st},
    {name:'xl/worksheets/sheet1.xml',text:sheetXml}
  ].concat(more.map(function(m,i){return {name:'xl/worksheets/sheet'+(i+2)+'.xml',text:sevenPlain(m.rows)};}));
  return zipUp(files);
}
var REPORTS=[
 {k:'log',t:'The general log',back:false,   /* written out only: Aconex is the source */
  d:'SEVEN\u2019s Material Live Tracking Sheet, in the client\u2019s own format \u2014 its headings, colours '
    +'and 76 columns \u2014 with every material on a row, by discipline, and the tracker\u2019s own columns '
    +'after them. A Summary sheet, and the documents linked to nothing, go in the same workbook.',
  go:function(){
    var name=(DB.project||'Project Materials').replace(/[^\w \-]/g,'').trim();
    download(sevenLog(),name+' \u2014 Material Live Tracking Sheet '+today()+'.xlsx');
  }},
 {k:'ven',t:'Vendors and who brought them',back:true,
  d:'Every company on the project: its kind and category, its pre-qualification, its ISO '
    +'certificate, the factory visit and who made it, where it is and who brought it on. '
    +'Correct any of it — including the name and the pre-qualification number — and send it '
    +'back. A blank cell leaves a value as it is; a cell with only "-" in it clears it. Leave '
    +'the first column alone; it is how a row finds its way home.',
  go:function(){
    download(workbook([{name:'Vendors',rows:vendorRows()}]),
      (DB.project||'Vendors')+' — vendors '+today()+'.xlsx');
  }},
 {k:'fatx',t:'Final Inspection / FAT Schedule Tracker',
  d:'Every material with a final inspection or FAT recorded: its maker, category and place, the '
    +'third-party agency and inspector and their approval, the pre-fabrication, fabrication, FAT and '
    +'first delivery dates, and whether the FAT is done, upcoming or to be repeated.',
  go:function(){download(fatTrackerBook(),(DB.project||'Project')+' — FAT Schedule Tracker '+today()+'.xlsx');}},
 {k:'pasum',t:'Physical Assessment Summary',
  d:'Every vendor whose factory is surveyed: its pre-qualification and number, scope and location, '
    +'then the physical assessment — done or not, the technical and QMS assessors and their '
    +'reports, the result and the report’s number.',
  go:function(){download(paSummaryBook(),(DB.project||'Project')+' — Physical Assessment Summary '+today()+'.xlsx');}},
 {k:'ahead',t:'Two-week look-ahead (TWLAS)',
  d:'This week and next, Saturday to Friday, in the project\u2019s TWLAS layout: physical assessments, '
    +'pre-qualifications, pre-inspection dossiers and meetings, post-inspection dossiers, FAT and '
    +'materials release, each marked on the day it is planned or falls due. Submitted weekly under clause 2.2.6.',
  go:function(){
    download(twlasBook(),(DB.project||'Project')+' \u2014 TWLAS '+twStart()+'.xlsx');
  }},
 {k:'insp',t:'Inspectors',back:true,
  d:'Everyone approved to inspect on this project, their agency and the reference SEVEN '
    +'approved them under — clause 2.2.17. Correct any of it and send it back. A row is '
    +'found by that reference before it is found by the name, because a name is spelt '
    +'three ways across four files and a reference is one person.',
  go:function(){
    download(workbook([{name:'Inspectors',rows:inspectorRows()}]),
      (DB.project||'Inspectors')+' — inspectors '+today()+'.xlsx');
  }}
];

window.repPick=function(k){
  REP_WANT=k;
  document.getElementById('xl-rep').click();
};
var REP_WANT='';
window.repRead=async function(ev){
  var f=ev.target.files[0];ev.target.value='';
  if(!f)return;
  if(REP_WANT==='log')return window.excelRead({target:{files:[f],value:''}});
  if(REP_WANT==='mat'){
    if(typeof busy==='function')busy(true,'Reading the materials sheet');
    try{
      var gm=await readMaterialSheet(f);
      MATP=planMaterials(gm.rows);MATP.file=f.name;MATP.count=gm.rows.length;
      if(typeof busy==='function')busy(false);
      showMaterials();
    }catch(e){
      if(typeof busy==='function')busy(false);
      sheet('That sheet could not be read','<div style="font-size:14px;line-height:1.75">'+esc(e.message||String(e))+'</div>');
    }
    return;
  }
  if(REP_WANT==='insp'){
    if(typeof busy==='function')busy(true,'Reading the inspector sheet');
    try{
      var g=await readInspectorSheet(f);
      var pl=planInspectors(g.rows);
      if(typeof busy==='function')busy(false);
      var n=applyInspectors(pl);
      toast(n.added+' added, '+n.changed+' updated, '+pl.same.length+' unchanged');
    }catch(e){
      if(typeof busy==='function')busy(false);
      sheet('That sheet could not be read',
        '<div style="font-size:14px;line-height:1.75">'+esc(e.message||String(e))+'</div>');
    }
    return;
  }
  if(typeof busy==='function')busy(true,'Reading the vendor sheet');
  try{
    var got=await readVendorSheet(f);
    VENP=planVendors(got.rows);
    VENP.file=f.name;VENP.count=got.rows.length;
    if(typeof busy==='function')busy(false);
    showVendors();
  }catch(e){
    if(typeof busy==='function')busy(false);
    sheet('That sheet could not be read',
      '<div style="font-size:14px;line-height:1.75">'+esc(e.message||String(e))+'</div>');
  }
};
var VENP=null;
function showVendors(){
  var p=VENP;
  sheet('From '+p.file,
    '<div class="dim" style="font-size:13.5px;margin-bottom:16px">'
    +p.count+' rows read. A row is found by its ID first, then by its pre-qualification '
    +'number, then by its name — so everything else on it can be corrected freely. '
    +'Nothing has been changed yet, and a vendor missing from the sheet is kept.</div>'
    +'<div class="grid" style="margin-bottom:18px">'
    +stat(p.add.length,'New')+stat(p.change.length,'Changed')
    +stat(p.same.length,'Unchanged')+stat(p.gone.length,'Not in the sheet')
    +'</div>'
    +(p.change.length?('<div class="sec">Changed</div><div class="panel"><div class="panel-b">'
      +p.change.slice(0,40).map(function(c){
        return '<div class="line"><span class="tag t-wait">'+c.diff.length+'</span>'
          +'<div class="line-m"><div>'+esc(c.v.name)
          +(K(c.v.name)!==K(trim(c.r['Vendor']))?(' → <b>'+esc(trim(c.r['Vendor']))+'</b>'):'')
          +'</div><div class="dim" style="font-size:12.5px;margin-top:2px">'
          +esc(c.diff.join(', '))+'</div></div></div>';}).join('')
      +more(p.change.length,40)+'</div></div>'):'')
    +(p.add.length?('<div class="sec">New</div><div class="panel"><div class="panel-b">'
      +p.add.slice(0,30).map(function(r){
        return '<div class="line"><span class="tag t-ok">new</span>'
          +'<div class="line-m"><div>'+esc(trim(r['Vendor']))+'</div>'
          +'<div class="dim" style="font-size:12.5px;margin-top:2px">'
          +esc([trim(r['Kind']),trim(r['Country']),trim(r['PQD Number'])].filter(Boolean).join(' · '))
          +'</div></div></div>';}).join('')
      +more(p.add.length,30)+'</div></div>'):'')
    +'<div class="f-act" style="margin-top:22px">'
    +((p.add.length||p.change.length)
      ?('<button class="btn btn-p" onclick="repApplyVendors()">Apply — '
        +(p.add.length+p.change.length)+'</button>'):'')
    +'<button class="btn-q" onclick="closeSheet()">Cancel</button></div>');
}
window.repApplyVendors=function(){
  if(!VENP)return;
  var n=applyVendors(VENP);
  closeSheet();
  toast(n.changed+' updated, '+n.added+' added');
  VENP=null;
};

function reportsPane(){
  /* counted the way the tabs count, so the same thing is not given two
     different numbers on two screens */
  var n={mat:0,mir:0,doc:0,loose:0};
  (DB.mats||[]).forEach(function(m){
    if(!isDoc(m)){n.mat++;return;}
    if(m.doc==='MIR')n.mir++;else n.doc++;
    if(waitsForLink(m)||(m.doc==='MIR'&&!isLinked(m)))n.loose++;
  });
  return '<div class="head"><div class="wrap"><div class="head-t">Reports</div>'
    +'<div class="head-m">'
    +'<span class="chip flat">'+n.mat+' materials</span>'
    +'<span class="chip flat">'+n.mir+' inspections</span>'
    +'<span class="chip flat">'+n.doc+' documents</span>'
    +'<span class="chip flat">'+(DB.mfrs||[]).length+' vendors</span>'
    +(n.loose?('<span class="chip warn">'+n.loose+' linked to no material</span>'):'')
    +'</div></div></div>'
    +'<div class="body"><div class="wrap">'
    +REPORTS.map(function(r){
      return '<div class="panel"><div class="panel-b">'
        +'<div style="display:flex;gap:18px;align-items:flex-start;flex-wrap:wrap">'
        +'<div style="flex:1;min-width:260px">'
        +'<div class="panel-t">'+esc(r.t)+'</div>'
        +'<div class="swhy" style="margin-top:6px">'+esc(r.d)+'</div></div>'
        +'<div style="display:flex;gap:8px;flex-wrap:wrap">'
        +'<button class="btn btn-p" onclick="runReport(\''+r.k+'\')">Download Excel</button>'
        +(r.back?('<button class="btn" onclick="repPick(\''+r.k+'\')">Upload edited</button>'):'')
        +'</div></div></div></div>';
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

/* ================================================================
   THE TABLE
   ----------------------------------------------------------------
   Rows and columns, because that is how this project already thinks.
   Five of them — materials, inspection requests, documents, vendors,
   inspectors — kept apart rather than folded into one, since a method
   statement and a manufacturer have almost no columns in common and a
   table that holds both is mostly empty.

   The filters matter more than the rest. Every column has one, they
   add together, and a column of a few repeated values offers them as
   a list rather than asking anyone to remember how a status is spelt.

   Nine hundred rows would take a while to build all at once and most
   of them are below the fold, so a screenful is drawn and the rest
   follows as it is scrolled to.
   ================================================================ */

var TBL='mat';                 /* mat | mir | doc | mfr | insp */
var TBLQ={};                   /* the filters, per table */
var TBLSORT={};                /* which column, which way */
var TBLSHOW={};                /* which columns are visible */
var TBLN={};                   /* how many rows have been drawn */
var PAGE_ROWS=120;

/* A column is a name, where to read it from, and what kind of filter
   it deserves. "pick" means the values repeat and are worth listing;
   "text" means they do not. */
function col(t,k,read,kind,w){return {t:t,k:k,read:read,kind:kind||'text',w:w||160};}
/* how a cell is drawn: a status as a coloured tag, a count to the right */
function asTag(c,f){c.tone=f||statusTone;return c;}
function asNum(c){c.num=true;return c;}
function statusTone(v){
  var s=normStatus(v)||trim(v);
  if(/^(Approved|Approved with comments|Passed|Passed with comments|Received|Valid|Waived|Closed)$/.test(s))return 'ok';
  if(/^(Rejected|Resubmit|Terminated|Expired|Failed|Suspended)$/.test(s))return 'bad';
  if(/^(Pending|Under Review)$/.test(s))return 'wait';
  return '';
}
/* A table draws hundreds of rows and reads each one several times — for
   the cell, the filter list, the sort. A material's road and a vendor's
   waiting materials are worked out once per drawing and kept here. */
var TCACHE=null;
function tc(){return TCACHE||(TCACHE={road:{},hold:null});}
function roadOfRow(r){var c=tc();return c.road[r.id]||(c.road[r.id]=matRoad(r));}
function holdsUp(v){
  var c=tc();
  if(!c.hold){
    c.hold={};
    (DB.mats||[]).forEach(function(m){
      if(isDoc(m)||!m.mfr)return;
      if(roadOfRow(m).at>=0)c.hold[String(m.mfr)]=(c.hold[String(m.mfr)]||0)+1;
    });
  }
  return c.hold[String(v.id)]||0;
}
function stageOf(r){var rd=roadOfRow(r);return rd.at<0?'Cleared':rd.steps[rd.at].s.n;}
function whoOf(r){
  var rd=roadOfRow(r);if(rd.at<0)return 'Cleared';
  var x=rd.steps[rd.at];
  if(x.s.linked)return 'Vendor';
  return x.state==='bad'?'Blocked':x.state==='wait'?'Reviewer':'You';
}
var WHO_TONE={'You':'now','Vendor':'now','Reviewer':'wait','Blocked':'bad','Cleared':'ok'};
/* where one requirement of a vendor stands, in one word; blank when that
   kind of company does not need it. Anything failed shows its own word. */
function stepWord(v,k){
  var x=mfrRoad(v).filter(function(y){return y.s.k===k;})[0];
  if(!x)return '';
  var iso=k==='iso';
  if(x.state==='done')return iso?'Valid':'Done';
  if(x.state==='wait')return iso?'Not confirmed':'With reviewer';
  if(x.state==='open')return iso?'Not recorded':'Not started';
  return iso?'Expired':(trim(x.data.status)||'Rejected');
}
var STEP_TONE={'Done':'ok','Valid':'ok','With reviewer':'wait','Not confirmed':'wait',
  'Not started':'now','Not recorded':'now'};
/* a vendor's ISO certificate in one word, for every kind of vendor */
function isoWord(r){
  var d=(r.steps||{}).iso||{};
  if(d.date&&daysTo(d.date)<0)return 'Expired';
  if(d.status==='Expired')return 'Expired';
  if(d.status==='Valid')return 'Valid';
  return (d.ref||d.date)?'Not confirmed':'Not recorded';
}
var MS_TONE={ok:'ok',wait:'wait',bad:'bad',now:'now'};

function stepOf(r,k,f){var d=((r.steps||{})[k])||{};return d[f]||'';}
function rawOf(r,c){return ((r.raw||{})[c])||'';}

/* An inspection page: every material whose road has that step, and where
   the step stands on it — read from the material itself, as recorded. */
/* A release note is issued, not visited: no inspector, its own words */
var VISIT_TEXT={
  ipi:{one:'Report',date:'Visit date',by:true,ref:'Report reference',res:'Result',none:'No report yet',
       add:'Record a visit',edit:'Edit the visit',what:'visit'},
  fat:{one:'Report',date:'Inspection date',by:true,ref:'Report reference',res:'Result',none:'No report yet',
       add:'Record a visit',edit:'Edit the visit',what:'visit'},
  irn:{one:'Note',date:'Issued',by:false,ref:'Release note number',res:'Employer copy',none:'No note yet',
       add:'Add a release note',edit:'Edit the release note',what:'release note'}
};
/* A page of reports rather than materials: one row for each visit
   recorded on a material's step — an in-process inspection is repeated
   while the thing is being made, and every report counts. A material
   with none yet has one row saying so, so what is still owed shows. */
function visitTable(k,label){
  var w=VISIT_TEXT[k];
  var num=function(r){return r.v?(r.n+' of '+r.of):'';};
  var cols=[
    (function(c){c.sub=function(r){var v=r.m.mfr?mfr(r.m.mfr):null;return v?v.name:'';};return c;})(
      col('Material','name',function(r){return r.m.name;},'text',380)),
    col('Category','cat',function(r){return r.m.cat;},'pick',90),
    col(w.one,'n',num,'text',80),
    col(w.date,'dt',function(r){return r.v?show(r.v.date):'';},'text',110)];
  if(w.by)cols.push(col('Inspector','by',function(r){return r.v?r.v.by:'';},'pick',180));
  cols.push(
    col(w.ref,'ref',function(r){return r.v?r.v.ref:'';},'text',280),
    asTag(col(w.res,'st',function(r){return r.v?(trim(r.v.result)||'Pending'):w.none;},'pick',160),
      function(v){return statusTone(v)||(v===w.none?'now':v==='Scheduled'||v==='Sent'?'wait':'');}),
    col('Discipline','disc',function(r){return r.m.disc;},'pick',150));
  return {label:label,
    rows:function(){
      var out=[];
      (DB.mats||[]).forEach(function(m){
        if(isDoc(m)||!stepApplies(m,k))return;
        var list=visitsOf(m,k);
        if(!list.length){out.push({id:-Number(m.id),m:m,v:null});return;}
        /* numbered oldest first, so report 1 is the first visit */
        list.slice().reverse().forEach(function(v,i){out.push({id:Number(v.id),m:m,v:v,n:i+1,of:list.length});});
      });
      return out;
    },
    open:function(r){jump('mat',r.m.id);},
    cols:cols,show:cols.map(function(c){return c.k;})};
}
/* A page of documents. Status and revision are read from whichever of
   the register's columns the kind fills. */
function rawEnd(r,end){
  var raw=r.raw||{},pre=r.doc?(r.doc+' '+end):'';
  if(pre&&raw[pre])return raw[pre];
  for(var k in raw)if(raw[k]&&k.slice(-end.length-1)===' '+end)return raw[k];
  return '';
}
function docTable(label,keep,withKind,withLink){
  var cols=[];
  if(withKind)cols.push(col('Kind','kind',function(r){return r.doc;},'pick',110));
  cols.push(
    col('Title','name',function(r){return r.name;},'text',440),
    col('Number','no',function(r){return refOf(r);},'text',300),
    col('Discipline','disc',function(r){return r.disc;},'pick',150),
    asTag(col('Status','st',function(r){return rawEnd(r,'Status');},'pick',170)),
    col('Revision','rev',function(r){return rawEnd(r,'Revision');},'pick',90));
  if(withLink)cols.push(
    asTag(col('Linked','lnk',function(r){
        return NOT_FOR_MAT[r.doc]?'':isLinked(r)?'Linked':(r.site||SITE_KIND[r.doc])?'Site work':'Not linked';},'pick',120),
      function(v){return v==='Not linked'?'bad':v==='Linked'?'ok':'';}),
    col('Materials served','n',function(r){var n=NOT_FOR_MAT[r.doc]?0:servedBy(r).length;return n?String(n):'';},'text',130));
  return {label:label,
    rows:function(){return (DB.mats||[]).filter(keep);},
    open:function(r){jump('mat',r.id);},
    cols:cols,show:cols.map(function(c){return c.k;})};
}
function mirTable(label,keep,withLink,extra){
  var cols=[
    col('Title','name',function(r){return r.name;},'text',460),
    col('Number','no',function(r){return refOf(r);},'text',300),
    col('Category','cat',function(r){return r.cat;},'pick',90),
    col('Discipline','disc',function(r){return r.disc;},'pick',150),
    asTag(col('Status','st',function(r){return rawOf(r,'MIR Status')||rawOf(r,'MAT Status');},'pick',170)),
    col('Date','dt',function(r){return show(rawOf(r,'MIR Approval Date')||rawOf(r,'MAT Submittal Date'));},'text',110)];
  if(withLink)cols.push(
    asTag(col('Linked','lnk',function(r){return isLinked(r)?'Linked':'Not linked';},'pick',120),
      function(v){return v==='Not linked'?'bad':'ok';}),
    col('Linked to','on',function(r){var x=servedBy(r);return x.length?x[0].name:'';},'text',360));
  return {label:label,rows:function(){return (DB.mats||[]).filter(keep);},
    open:function(r){jump('mat',r.id);},extra:extra,
    cols:cols,show:cols.map(function(c){return c.k;})};
}
var TABLES_DEF={
 ipi:visitTable('ipi','In-Process Inspection'),
 itp:docTable('Inspection & Test Plan',function(m){return m.doc==='ITP';},false,true),
 mes:docTable('Method Statement',function(m){return m.doc==='MES';},false,true),
 wir:docTable('Work Inspection Request',function(m){return m.doc==='WIR';},false,true),
 mas:docTable('Material Sample',function(m){return m.doc==='MAS';},false,true),
 paa:(function(d){
   d.cols=d.cols.concat([
     asTag(col('Linked','lnk',function(r){return paaOwners(r).length?'Linked':'Not linked';},'pick',120),
       function(v){return v==='Not linked'?'bad':'ok';}),
     col('Inspector','who',function(r){return paaOwners(r).map(function(p){return p.name;}).join(', ');},'text',240)]);
   d.show=d.cols.map(function(c){return c.k;});
   d.extra=function(){return '<button class="btn btn-s" onclick="setTab(\'insp\')">\u2190 Personnel Approval</button>';};
   return d;
 })(docTable('PAA files',function(m){return m.doc==='PAA';},false,false)),
 /* a PQD finds its vendor by its number, which Aconex gives each company alone */
 pqd:(function(d){
   d.cols=d.cols.concat([
     asTag(col('Vendor found','lnk',function(r){return vendorByPq(refOf(r))?'Found':'Not found';},'pick',130),
       function(v){return v==='Not found'?'bad':'ok';}),
     col('Vendor','who',function(r){var v=vendorByPq(refOf(r));return v?v.name:'';},'text',280)]);
   d.show=d.cols.map(function(c){return c.k;});
   d.extra=function(){return '<button class="btn btn-s" onclick="setTab(\'mfr\')">\u2190 Vendors</button>';};
   return d;
 })(docTable('Pre-qualification (PQD)',function(m){return m.doc==='PQD';},false,false)),
 fat:visitTable('fat','FAT/Final Inspection'),
 irn:visitTable('irn','Inspection Release Note'),
 mat:{label:'Materials',rows:function(){return (DB.mats||[]).filter(function(m){return !isDoc(m);});},
   open:function(r){jump('mat',r.id);},
   /* the materials the filters leave, out to Excel to be edited, and back */
   extra:function(){return '<button class="btn btn-s" onclick="matOut()">Edit in Excel</button>'
     +'<button class="btn btn-s" onclick="repPick(\'mat\')">Upload edited</button>';},
   cols:[
    col('Item Description','name',function(r){return r.name;},'text',420),
    col('Category','cat',function(r){return r.cat;},'pick',90),
    col('Discipline','disc',function(r){return r.disc;},'pick',150),
    col('Vendor','ven',function(r){var v=r.mfr?mfr(r.mfr):null;return v?v.name:'';},'pick',180),
    asTag(col('Vendor status','vst',function(r){var v=r.mfr?mfr(r.mfr):null;return v?mfrState(v).word:'no vendor';},'pick',150),
      function(v,r){var x=r.mfr?mfr(r.mfr):null;return x?MS_TONE[mfrState(x).tone]:'now';}),
    col('Stage','stage',stageOf,'pick',190),
    asTag(col('Waiting on','who',whoOf,'pick',110),function(v){return WHO_TONE[v]||'';}),
    (function(c){c.sort=function(r){var rd=roadOfRow(r);return rd.steps.length?rd.done/rd.steps.length:0;};return c;})(
      asNum(col('Progress','prog',function(r){var rd=roadOfRow(r);return rd.done+' / '+rd.steps.length;},'text',90))),
    /* how much of its Main Log row is filled: sort on it to find the gaps */
    (function(c){c.sort=function(r){var f=logFill(r);return f.n/f.total;};return c;})(
      asNum(col('Main Log','logp',function(r){var f=logFill(r);return Math.round(f.n/f.total*100)+'%';},'text',100))),
    col('Local / Foreign','loc',function(r){var v=companyOfMat(r);return v?(v.locality||''):'';},'pick',120),
    col('Sub-contractor','sub',function(r){return r.sub;},'pick',150),
    col('MAT Number','matno',function(r){return rawOf(r,'MAT Number')||r.ref;},'text',300),
    asTag(col('MAT Status','matst',function(r){return stepOf(r,'mts','status')||rawOf(r,'MAT Status');},'pick',150)),
    col('MAT Date','matdt',function(r){return show(stepOf(r,'mts','date')||rawOf(r,'MAT Submittal Date'));},'text',110),
    /* read from the documents linked to the material, as the Main Log
       is, and only failing those from what the material once carried */
    col('ITP Number','itpno',function(r){return linkedVals(r,'ITP','no')||rawOf(r,'ITP Number');},'text',300),
    asTag(col('ITP Status','itpst',function(r){return stepOf(r,'itp','status')||rawOf(r,'ITP Status');},'pick',150)),
    col('MES Number','mesno',function(r){return linkedVals(r,'MES','no')||rawOf(r,'Method Statement Number');},'text',300),
    asTag(col('MES Status','messt',function(r){return linkedVals(r,'MES','st')||rawOf(r,'MES Status');},'pick',150)),
    col('PO Number','pono',function(r){return rawOf(r,'PO Number');},'text',150),
    col('Package','pkg',function(r){return rawOf(r,'Package (Lump Sum / Provisional Sum / Prime Cost)');},'pick',150),
    col('Quantity','qty',function(r){return r.qty?(r.qty+' '+(r.unit||'')):'';},'text',110),
    asNum(col('Delivered','got',function(r){return (r.dels||[]).length?String(received(r)):'';},'text',100)),
    asNum(col('Documents','ndoc',function(r){return String((r.docs||[]).length||'');},'text',100)),
    asNum(col('Consignments','ndel',function(r){return String((r.dels||[]).length||'');},'text',110))],
   show:['name','cat','disc','ven','vst','stage','who','prog','logp','matst','itpst','mesno','messt']},

 /* The MIR tab is the work still to do, as Documents is: a request
    leaves it once a material links it, and every request, linked or
    not, is on "All MIR" to look things up in. */
 mir:mirTable('Material Inspection Request',function(m){return m.doc==='MIR'&&!isLinked(m);},false,
   function(){var n=(DB.mats||[]).filter(function(m){return m.doc==='MIR';}).length;
     return '<span class="chip flat">Not linked to a material yet</span>'
       +'<button class="btn btn-s" onclick="setTab(\'allmir\')">All MIR '+n+'</button>';}),
 allmir:mirTable('All MIR',function(m){return m.doc==='MIR';},true,
   function(){return '<button class="btn btn-s" onclick="setTab(\'mir\')">\u2190 Not linked yet</button>';}),

 /* the work still to do; once linked a document is found under Other */
 doc:(function(d){d.extra=function(){return '<span class="chip flat">Not linked to a material yet</span>';};return d;})(
   docTable('Documents',waitsForLink,true,false)),
 alldoc:docTable('All documents',function(m){return isDoc(m)&&m.doc!=='MIR';},true,true),
 odoc:docTable('Other kinds',oddKind,true,true),

 mfr:{label:'Vendors',rows:function(){return (DB.mfrs||[]).slice();},
   open:function(r){jump('mfr',r.id);},
   /* The sheet that goes out to be edited and comes back: the vendors the
      filters leave, in the vendor report's shape (its ID column finds each
      row's vendor again). The shared-number button shows only while some
      pre-qualification number sits on more than one vendor. */
   extra:function(){var n=pqGroups().length;
     var nq=(DB.mats||[]).filter(function(m){return m.doc==='PQD';}).length;
     return '<button class="btn btn-s" onclick="setTab(\'pqd\')">PQD files '+nq+'</button>'
       +'<button class="btn btn-s" onclick="venOut()">Edit in Excel</button>'
       +'<button class="btn btn-s" onclick="repPick(\'ven\')">Upload edited</button>'
       +(n?'<button class="btn btn-s" style="border-color:#ffcc3e" onclick="pqShared()">'+n
       +' shared PQD number'+(n===1?'':'s')+'</button>':'');},
   cols:[
    /* the name, with what the company does written small beneath it */
    (function(c){c.sub=function(r){return r.scope||'';};return c;})(
      col('Vendor','name',function(r){return r.name;},'text',300)),
    col('Kind','kind',function(r){return KINDS[kindOf(r)].l;},'pick',150),
    /* The same few questions for every vendor, whatever kind it is: its
       category, where its pre-qualification stands, its ISO certificate,
       the factory visit, and whether it is local and where. */
    col('Category','vcat',function(r){return r.cat||'';},'pick',100),
    asTag(col('PQD','pqd',function(r){return trim(pqOf(r).status)||'Not started';},'pick',170),
      function(v){return statusTone(v)||(v==='Not started'?'now':'');}),
    asTag(col('ISO 9001','isoall',isoWord,'pick',140),function(v){return STEP_TONE[v]||'bad';}),
    asTag(col('Visit','visit',function(r){return trim(stepOf(r,'pa','status'))||'Not done';},'pick',150),
      function(v){return statusTone(v)||(v==='Not done'?'now':'');}),
    col('Local / Foreign','loc',function(r){return r.locality||'';},'pick',120),
    col('Location','where',function(r){return [r.site,r.country].filter(Boolean).join(' · ');},'text',220),
    /* a few more, for when they are wanted */
    asTag(col('Qualification','qual',function(r){return mfrState(r).word;},'pick',160),
      function(v,r){return MS_TONE[mfrState(r).tone]||'';}),
    asNum(col('Holds up','hold',function(r){var n=holdsUp(r);return n?String(n):'';},'text',90)),
    asNum(col('Materials','n',function(r){return String(matsOf(r).length);},'text',100)),
    asTag(asNum(col('ISO days left','isodays',function(r){
      var d=stepOf(r,'iso','date');return d?String(daysTo(d)):'';},'text',110)),
      function(v){if(v==='')return '';var n=+v;return n<0?'bad':n<=60?'now':'ok';}),
    col('ISO Expires','isodt',function(r){return show(stepOf(r,'iso','date'));},'text',110),
    col('PQD Number','pq',function(r){return pqOf(r).ref||'';},'text',300),
    col('Brought by','by',function(r){return r.by||'';},'pick',180)],
   show:['name','kind','vcat','pqd','isoall','visit','loc','where']},

 avl:{label:'SEVEN vendor list',rows:function(){return AVL?AVL.rows:[];},
   open:function(){},                      /* reference: nothing to open */
   extra:function(){
     return '<span class="chip flat">'+esc(AVL&&AVL.rev?('Revision '+AVL.rev):(AVL?AVL.file:''))+'</span>'
       +'<button class="btn btn-s" onclick="avlPick()">Upload a new revision</button>';},
   cols:[
    col('Material','mat',function(r){return r.mat;},'text',320),
    col('Manufacturer','mfr',function(r){return r.mfr;},'text',260),
    col('Category','cat',function(r){return r.cat;},'pick',90),
    asTag(col('Status','status',function(r){return r.status;},'pick',130),
      function(v){return {'Approved':'ok','Re-assessed':'ok','Disapproved':'bad','On hold':'now',
        'Pending':'wait','Under monitoring':'now'}[v]||'';}),
    /* which of SEVEN's lists the row came from: its letter and its title */
    col('List','sh',function(r){return avlListName(r.sh);},'pick',230),
    col('Discipline','disc',function(r){return r.disc;},'pick',150),
    col('Segment','seg',function(r){return r.seg;},'pick',160),
    col('Country','country',function(r){return r.country;},'pick',140),
    /* already one of this project's vendors, by name */
    asTag(col('On project','onp',function(r){return onProject(r.mfr)?'Yes':'';},'pick',100),function(v){return v?'wait':'';}),
    col('Project','project',function(r){return r.project;},'pick',140),
    col('Limitations','limit',function(r){return r.limit;},'text',260),
    col('Supplier','supplier',function(r){return r.supplier;},'text',200),
    col('Remarks','remarks',function(r){return r.remarks;},'text',220),
    col('Assessed','date',function(r){return r.date?show(r.date):'';},'text',110)],
   show:['mat','mfr','cat','status','sh','disc','country','onp','limit']},

 insp:{label:'Personnel Approval',rows:function(){return (DB.people||[]).slice();},
   extra:function(){
     var all=(DB.mats||[]).filter(function(m){return m.doc==='PAA';});
     var loose=all.filter(function(d){return !paaOwners(d).length;}).length;
     return '<button class="btn btn-s" onclick="setTab(\'paa\')">PAA files '+all.length+'</button>'
       +(loose?('<button class="btn btn-s" style="border-color:#ffcc3e" onclick="showLoosePaa()">'
         +loose+' not linked to an inspector</button>'):'');},
   open:function(r){jump('insp',r.id);},
   cols:[
    col('Name','name',function(r){return r.name;},'text',240),
    col('Agency','agency',function(r){return r.agency||'';},'pick',200),
    col('Discipline','disc',function(r){return r.disc||'';},'pick',160),
    asTag(col('Approval','status',function(r){return r.status||'Pending';},'pick',130)),
    col('Reference','ref',function(r){return r.ref||'';},'text',240),
    asNum(col('PAA','npaa',function(r){var n=(r.docs||[]).length;return n?String(n):'';},'text',80))],
   show:['name','agency','disc','status','ref','npaa']}
};

/* ================================================================
   SEVEN's vendor list (the AVL). The client's own list of which makers
   are approved for which material — reference, not the project's own
   vendors, so it sits in a tab of its own and is searched, not edited.
   It is uploaded as SEVEN issues it and kept in one settings row, read
   only when the tab is opened; a new revision replaces the old.
   ================================================================ */
var AVL=null, AVL_STATE='idle';               /* idle | loading | none | ready */
/* The workbook read with the fill of each cell, because sheet A says
   "re-assessed" and "disapproved" in colour, not in words. */
async function avlBook(file){
  var zip=await unzip(await file.arrayBuffer());
  var wbDoc=parse(await textOf(zip['xl/workbook.xml']));
  var relDoc=parse(await textOf(zip['xl/_rels/workbook.xml.rels']));
  var rels={};
  [].forEach.call(relDoc.getElementsByTagName('Relationship'),function(r){
    rels[r.getAttribute('Id')]=r.getAttribute('Target').replace(/^\/?xl\//,'').replace(/^\//,'');});
  var shared=[], ss=await textOf(zip['xl/sharedStrings.xml']);
  if(ss)[].forEach.call(parse(ss).getElementsByTagName('si'),function(si){
    var t='';[].forEach.call(si.getElementsByTagName('t'),function(n){t+=n.textContent;});shared.push(t);});
  var xfFill=[], st=await textOf(zip['xl/styles.xml']);
  if(st){
    var sd=parse(st), fills=[];
    var fe=sd.getElementsByTagName('fills')[0];
    if(fe)[].forEach.call(fe.getElementsByTagName('fill'),function(f){
      var pf=f.getElementsByTagName('patternFill')[0], fg=pf&&pf.getElementsByTagName('fgColor')[0];
      fills.push(pf&&pf.getAttribute('patternType')==='solid'&&fg&&fg.getAttribute('rgb')?fg.getAttribute('rgb').toUpperCase():'');
    });
    var cx=sd.getElementsByTagName('cellXfs')[0];
    if(cx)[].forEach.call(cx.getElementsByTagName('xf'),function(x){xfFill.push(fills[+x.getAttribute('fillId')||0]||'');});
  }
  var out=[];
  [].forEach.call(wbDoc.getElementsByTagName('sheet'),function(s){
    var id=s.getAttribute('r:id')||s.getAttributeNS('http://schemas.openxmlformats.org/officeDocument/2006/relationships','id');
    out.push({name:s.getAttribute('name'),path:'xl/'+(rels[id]||'')});
  });
  for(var i=0;i<out.length;i++){
    var doc=parse(await textOf(zip[out[i].path])), rows=[];
    [].forEach.call(doc.getElementsByTagName('row'),function(r){
      var line=[];
      [].forEach.call(r.getElementsByTagName('c'),function(c){
        var j=colNum(c.getAttribute('r')), t=c.getAttribute('t'), v='';
        if(t==='inlineStr'){[].forEach.call(c.getElementsByTagName('t'),function(n){v+=n.textContent;});}
        else{var vn=c.getElementsByTagName('v')[0];v=vn?vn.textContent:'';if(t==='s')v=shared[+v]||'';}
        line[j]={v:v,fill:xfFill[+c.getAttribute('s')||0]||''};
      });
      rows[(+r.getAttribute('r')||rows.length+1)-1]=line;
    });
    out[i].rows=rows;
  }
  return out;
}
function avlStatus(s){
  var k=K(s);
  if(!k)return '';
  if(/disapprov/.test(k))return 'Disapproved';
  if(/hold/.test(k))return 'On hold';
  if(/re-?assess/.test(k))return 'Re-assessed';
  if(/pending/.test(k))return 'Pending';
  if(/monitor/.test(k))return 'Under monitoring';
  if(/approv/.test(k))return 'Approved';
  if(/closed|award/.test(k))return 'Approved';
  return trim(s);
}
/* Every sheet, read by its own headings. A segment or discipline written
   once over a block of rows is carried down it. The ICT sheet lists up
   to five vendors across a row and is read as a row each. */
/* Which of SEVEN's lists a row came from, by the letter its sheet ends
   in. The file's own title for each list is read when it is uploaded;
   these stand in for a list saved before titles were kept. */
var AVL_LISTS={A:'General materials',B:'Bulk, framework agreements',C:'Special materials and finishes',
  D:'ICT',E:'Cinema bulk materials',F:'Controlled materials',G:'CxA and third parties'};
function avlListName(tag){
  var t=(AVL&&AVL.lists&&AVL.lists[tag])||AVL_LISTS[tag]||'';
  return t?(tag+' · '+t):tag;
}
/* "SEVEN APPROVED VENDOR LIST FOR SPECIAL MATERIALS" -> "Special materials" */
function avlTitle(s){
  var t=trim(s).replace(/^.*?\bLIST\s+FOR\s+/i,'').replace(/^.*?\bAPPROVED\s+/i,'').replace(/\s+/g,' ')
    .replace(/\s*-?\s*list\s*\([a-z]\)\s*$/i,'').trim();   /* "... - LIST (A)": the letter is shown already */
  if(!t||t.length>60)return '';
  t=t.toLowerCase().replace(/\bcxa\b/,'CxA').replace(/\bict\b/,'ICT');
  return t.charAt(0).toUpperCase()+t.slice(1);
}
function avlRows(sheets){
  var all=[], rev='', lists={};
  sheets.forEach(function(sh){
    var R=sh.rows, cell=function(r,j){return trim(((R[r]||[])[j]||{}).v);};
    var tag=(/\(([A-Z])\)\s*$/.exec(sh.name)||[])[1]||sh.name;
    /* the list's title sits above its table, below the document block */
    for(var tr=4;tr<Math.min(R.length,14);tr++)
      (R[tr]||[]).forEach(function(c){
        if(c&&!lists[tag]&&/SEVEN\b.*\b(LIST|APPROVED)\b/i.test(c.v)){var t=avlTitle(c.v);if(t)lists[tag]=t;}});
    for(var r=0;r<Math.min(R.length,8);r++)
      (R[r]||[]).forEach(function(c,j){if(c&&K(c.v)==='revision date'&&!rev)rev=cell(r,j+2)||cell(r,j+1);});
    /* the legend: a filled cell with its meaning beside it */
    var legend={};
    for(var lr=0;lr<Math.min(R.length,12);lr++){
      var a=(R[lr]||[])[0];
      if(a&&a.fill&&cell(lr,1))legend[a.fill]=avlStatus(cell(lr,1));
    }
    var head=-1,ict=false;
    for(var h=0;h<Math.min(R.length,30)&&head<0;h++){
      (R[h]||[]).forEach(function(c){
        if(!c)return;
        if(/^sr ?#?$/.test(K(c.v))&&head<0)head=h;
        if(K(c.v)==='system name'&&head<0){head=h;ict=true;}
      });
    }
    if(head<0)return;
    var H={};(R[head]||[]).forEach(function(c,j){if(c&&trim(c.v))H[j]=K(c.v);});
    function find(test){for(var j in H)if(test(H[j]))return +j;return -1;}
    if(ict){
      var sysJ=find(function(x){return x==='system name';});
      var vJ=Object.keys(H).filter(function(j){return /^vendor \d/.test(H[j]);}).map(Number);
      var seg='';
      for(var r2=head+1;r2<R.length;r2++){
        if(cell(r2,0))seg=cell(r2,0);
        var sys=cell(r2,sysJ);if(!sys)continue;
        vJ.forEach(function(j){var m=cell(r2,j);
          if(m)all.push({sh:tag,seg:seg,disc:'ICT',mat:sys,cat:'',mfr:m,country:'',status:'Approved'});});
      }
      return;
    }
    var J={mat:find(function(x){return /material description/.test(x);}),
      seg:find(function(x){return /segment/.test(x);}),
      disc:find(function(x){return /discipline/.test(x);}),
      cat:find(function(x){return /inspection category/.test(x);}),
      mfr:find(function(x){return /^(mfr name|manufacturer name|company name)$/.test(x);}),
      country:find(function(x){return x==='country';}),
      status:find(function(x){return /approved/.test(x)&&!/commercial|pc approved/.test(x);}),
      project:find(function(x){return x==='project';}),
      limit:find(function(x){return /limitation/.test(x);}),
      supplier:find(function(x){return /supplier name/.test(x);}),
      remarks:find(function(x){return /remarks/.test(x);}),
      date:find(function(x){return /assessment date/.test(x);})};
    if(J.mfr<0)return;
    var carry={seg:'',disc:'',mat:'',cat:''};
    for(var r3=head+1;r3<R.length;r3++){
      var get=function(k){return J[k]<0?'':cell(r3,J[k]);};
      ['seg','disc','mat','cat'].forEach(function(k){var v=get(k);if(v)carry[k]=v;});
      var mfr=get('mfr');if(!mfr)continue;
      var fill=((R[r3]||[])[J.mfr]||{}).fill;
      var status=avlStatus(get('status'))||(fill&&legend[fill])||'Approved';
      var d=get('date');
      all.push({sh:tag,seg:carry.seg,disc:carry.disc,mat:J.mat<0?carry.seg:carry.mat,cat:normCat(get('cat'))||get('cat'),
        mfr:mfr,country:get('country'),status:status,project:get('project'),
        limit:get('limit'),supplier:get('supplier'),remarks:get('remarks'),
        date:d?(anyDate(isNaN(d)?d:Number(d))||d):''});
    }
  });
  all.forEach(function(x,i){x.id=i+1;});
  return {rev:rev,rows:all,lists:lists};
}
async function avlLoad(){
  if(AVL_STATE==='loading'||AVL_STATE==='ready')return;
  AVL_STATE='loading';
  try{
    var r=await client().from('settings').select('value').eq('key','avl').maybeSingle();
    if(r.error)throw r.error;
    AVL=(r.data&&r.data.value)||null;
    AVL_STATE=AVL?'ready':'none';
    if(AVL)AVL.rows.forEach(function(x,i){x.id=i+1;});
  }catch(e){AVL_STATE='idle';toast('Could not read the vendor list — '+niceError(e));}
  if(TAB==='avl')rPane();
}
window.avlPick=function(){
  var f=document.getElementById('avl-file');
  if(!f){f=document.createElement('input');f.type='file';f.id='avl-file';f.accept='.xlsx';
    f.style.display='none';f.onchange=avlRead;document.body.appendChild(f);}
  f.value='';f.click();
};
var AVL_NEW=null;
async function avlRead(ev){
  var file=ev.target.files[0];if(!file)return;
  busy(true,'Reading the vendor list');
  try{
    var got=avlRows(await avlBook(file));
    busy(false);
    if(!got.rows.length)throw new Error('No manufacturer rows were found in that file.');
    AVL_NEW={rev:got.rev,file:file.name,at:new Date().toISOString(),rows:got.rows,lists:got.lists};
    var bySh={},bySt={};
    got.rows.forEach(function(x){bySh[x.sh]=(bySh[x.sh]||0)+1;bySt[x.status]=(bySt[x.status]||0)+1;});
    sheet('SEVEN vendor list — '+file.name,
      '<div class="dim" style="font-size:13.5px;margin-bottom:16px">'
      +got.rows.length+' manufacturer rows read'+(got.rev?(', revision '+esc(got.rev)):'')
      +'. Nothing is saved yet. Saving replaces the list the project holds now'
      +(AVL?(' ('+esc(AVL.rev||AVL.file)+')'):'')+'.</div>'
      +'<div class="grid" style="margin-bottom:18px">'
      +Object.keys(bySt).map(function(k){return stat(bySt[k],k);}).join('')+'</div>'
      +'<div class="sec">By sheet</div><div class="panel"><div class="panel-b">'
      +Object.keys(bySh).map(function(k){return '<div class="line"><span class="tag t-na">'+esc(k)+'</span>'
        +'<div class="line-m">'+esc(got.lists[k]||AVL_LISTS[k]||'')+' <span class="dim">· '+bySh[k]+' rows</span></div></div>';}).join('')+'</div></div>'
      +'<div class="f-act" style="margin-top:20px"><button class="btn btn-p" onclick="avlSave()">Save to the project</button>'
      +'<button class="btn-q" onclick="closeSheet()">Cancel</button></div>');
  }catch(e){busy(false);sheet('That file could not be read','<div style="font-size:14px;line-height:1.75">'
    +esc(e.message||String(e))+'<br><br>This reads SEVEN’s vendor list as it is issued '
    +'(0DMQL00-DLVR-00-SEV-QM-TEM-00007), with its sheets A to G.</div>');}
}
window.avlSave=async function(){
  if(!AVL_NEW)return;
  busy(true,'Saving the vendor list');
  var keep=AVL_NEW;
  var slim={rev:keep.rev,file:keep.file,at:keep.at,lists:keep.lists||{},rows:keep.rows.map(function(x){
    var o={};Object.keys(x).forEach(function(k){if(k!=='id'&&x[k]!=='')o[k]=x[k];});return o;})};
  var r=await client().from('settings').upsert({key:'avl',value:slim},{onConflict:'key'});
  busy(false);
  if(r.error)return toast('Could not save the vendor list — '+niceError(r.error));
  AVL=keep;AVL_STATE='ready';AVL_NEW=null;closeSheet();
  toast(keep.rows.length+' rows saved'+(keep.rev?(' — revision '+keep.rev):''));
  rPane();
};
/* is a maker on SEVEN's list already one of this project's vendors? By
   name, loosely: case, punctuation and the usual company words aside */
var ONP=null, ONP_N=-1;
function coName(s){return K(s).replace(/\b(co|company|ltd|limited|llc|est|factory|industries|industrial|group|trading|the)\b/g,' ')
  .replace(/[^a-z0-9]+/g,' ').trim();}
function onProject(name){
  var n=(DB.mfrs||[]).length;
  if(!ONP||ONP_N!==n){ONP={};ONP_N=n;(DB.mfrs||[]).forEach(function(v){var k=coName(v.name);if(k)ONP[k]=1;});}
  var k=coName(name);return !!(k&&ONP[k]);
}
function avlPane(){
  if(AVL_STATE==='idle'){avlLoad();}
  if(AVL_STATE!=='ready'){
    var loading=AVL_STATE!=='none';
    return '<div class="head"><div class="wrap"><div class="head-t">SEVEN vendor list</div>'
      +'<div class="head-m">'+(loading?'<span class="chip flat">Loading…</span>'
        :'<button class="btn btn-s" onclick="avlPick()">Upload the vendor list</button>')+'</div></div></div>'
      +(loading?'':'<div class="body"><div class="wrap"><div class="empty" style="padding:40px 0;text-align:center">'
        +'No vendor list yet. Upload SEVEN’s workbook as it is issued — every sheet in it is read, '
        +'and the colours in sheet A (re-assessed, disapproved) are read as statuses.</div></div></div>');
  }
  TBL='avl';
  return tablePane(true);
}

function tdef(){return TABLES_DEF[TBL];}
/* the columns a table offers today: one that says nothing about anything
   on the project — a requirement no vendor has — is left out */
function colsOf(d){return d.cols.filter(function(c){return !c.need||c.need();});}
function anyVendorNeeds(k){
  return (DB.mfrs||[]).some(function(v){
    return needsQual(v)&&KINDS[kindOf(v)].steps.indexOf(k)>=0;});
}
function shownCols(){
  var d=tdef();
  if(!TBLSHOW[TBL])TBLSHOW[TBL]=d.show.slice();
  var want={};TBLSHOW[TBL].forEach(function(k){want[k]=1;});
  return colsOf(d).filter(function(c){return want[c.k];});
}
function cellOf(r,c){
  try{return trim(c.read(r));}catch(e){return '';}
}

/* the values a "pick" column actually holds, so the list offers what is
   there rather than what might be */
function choices(c,rows){
  var n={};
  rows.forEach(function(r){
    var v=cellOf(r,c);
    if(v)n[v]=(n[v]||0)+1;
  });
  return Object.keys(n).sort(function(a,b){return n[b]-n[a];})
    .slice(0,40).map(function(v){return {v:v,n:n[v]};});
}

function filtered(){
  TCACHE=null;                       /* a fresh reading for every drawing */
  var d=tdef(),rows=d.rows(),q=TBLQ[TBL]||{};
  var active=Object.keys(q).filter(function(k){return q[k]!=='';});
  if(active.length){
    var byKey={};d.cols.forEach(function(c){byKey[c.k]=c;});
    rows=rows.filter(function(r){
      for(var i=0;i<active.length;i++){
        var c=byKey[active[i]];if(!c)continue;
        var want=q[active[i]],have=cellOf(r,c);
        if(c.kind==='pick'){
          if(want==='(blank)'){if(have!=='')return false;}
          else if(have!==want)return false;
        }else if(K(have).indexOf(K(want))<0)return false;
      }
      return true;
    });
  }
  var s=TBLSORT[TBL];
  if(s){
    var byKey2={};d.cols.forEach(function(c){byKey2[c.k]=c;});
    var c2=byKey2[s.k];
    if(c2)rows.sort(function(a,b){
      var A=c2.sort?c2.sort(a):cellOf(a,c2),B=c2.sort?c2.sort(b):cellOf(b,c2);
      var num=(A!==''&&B!==''&&!isNaN(A)&&!isNaN(B));
      var d2=num?(parseFloat(A)-parseFloat(B)):String(A).localeCompare(String(B));
      return s.dir<0?-d2:d2;
    });
  }
  return rows;
}

function tablePane(list){
  var d=tdef(),rows=filtered(),all=d.rows().length;
  var cols=shownCols();
  var q=TBLQ[TBL]||{};
  var nActive=Object.keys(q).filter(function(k){return q[k]!=='';}).length;
  if(!TBLN[TBL])TBLN[TBL]=PAGE_ROWS;
  var upto=Math.min(TBLN[TBL],rows.length);
  var canAdd={mat:'material',mfr:'vendor',insp:'inspector'}[TBL];

  var head='<div class="head no-print"><div class="wrap" style="max-width:none">'
    +'<div class="head-t">'+(list?esc(d.label):'Table')+'</div><div class="head-m">'
    /* a list page is already one table; its tab says which */
    +(list?((canAdd?'<button class="btn btn-s btn-p" data-pop onclick="listAdd(event)">Add a '+canAdd+'</button>':'')
      +(d.extra?d.extra():'')):
      Object.keys(TABLES_DEF).filter(function(k){return k!=='avl';}).map(function(k){
      return '<button class="chip'+(TBL===k?' set':'')+'" onclick="setTable(\''+k+'\')">'
        +esc(TABLES_DEF[k].label)+' <span class="meta">'+TABLES_DEF[k].rows().length+'</span></button>';
    }).join(''))
    +(list?'':'<span style="flex:1"></span>')
    +'<button class="btn btn-s" onclick="tblCols()">Columns</button>'
    +(nActive?('<button class="btn btn-s btn-d" onclick="tblClear()">Clear '+nActive
      +' filter'+(nActive===1?'':'s')+'</button>'):'')
    +'<button class="btn btn-s" onclick="window.print()">Print</button>'
    +'<button class="btn btn-s'+(list?'':' btn-p')+'" onclick="tblExport()">Excel</button>'
    +'</div>'
    +'<div class="swhy" style="margin-top:8px">'
    +(nActive?(rows.length+' of '+all+' rows'):(all+' rows'))
    +' · type under a heading to filter, click a heading to sort'
    +'</div></div></div>';

  var body='<div class="body" id="tbl-body" onscroll="tblScroll(this)">'
    +'<table class="tbl"><thead><tr>'
    +cols.map(function(c){
      var s=TBLSORT[TBL];
      var mark=(s&&s.k===c.k)?(s.dir<0?' ▾':' ▴'):'';
      return '<th style="min-width:'+c.w+'px">'
        +'<button class="th-b" onclick="tblSort(\''+c.k+'\')">'+esc(c.t)+mark+'</button>'
        +tblFilter(c,rows)+'</th>';
    }).join('')
    +'</tr></thead><tbody id="tbl-rows">'
    +rows.slice(0,upto).map(function(r){return tblRow(r,cols);}).join('')
    +'</tbody></table>'
    +(upto<rows.length?('<div class="dim" style="padding:16px;text-align:center" id="tbl-more">'
      +'showing '+upto+' of '+rows.length+' — scroll for more</div>'):'')
    +(rows.length?'':'<div class="empty">Nothing matches these filters.</div>')
    +'</div>';
  return head+body;
}

function tblFilter(c,rows){
  var q=(TBLQ[TBL]||{})[c.k]||'';
  if(c.kind==='pick'){
    var opts=choices(c,tdef().rows());
    return '<select class="th-f" onchange="tblSet(\''+c.k+'\',this.value)">'
      +'<option value=""'+(q===''?' selected':'')+'>all</option>'
      +opts.map(function(o){
        return '<option value="'+attr(o.v)+'"'+(q===o.v?' selected':'')+'>'
          +esc(o.v)+' ('+o.n+')</option>';}).join('')
      +'<option value="(blank)"'+(q==='(blank)'?' selected':'')+'>— blank —</option>'
      +'</select>';
  }
  return '<input class="th-f" value="'+attr(q)+'" placeholder="filter" '
    +'oninput="tblSetLater(\''+c.k+'\',this.value)">';
}

window.setTable=function(k){TBL=k;TBLN[k]=PAGE_ROWS;rPane();};
window.tblSort=function(k){
  var s=TBLSORT[TBL];
  TBLSORT[TBL]=(s&&s.k===k)?{k:k,dir:-s.dir}:{k:k,dir:1};
  TBLN[TBL]=PAGE_ROWS;rPane();
};
window.tblSet=function(k,v){
  TBLQ[TBL]=TBLQ[TBL]||{};TBLQ[TBL][k]=v;
  TBLN[TBL]=PAGE_ROWS;rPane();
};
/* typing redraws nine hundred rows on every keystroke otherwise */
var tblTimer=null;
window.tblSetLater=function(k,v){
  TBLQ[TBL]=TBLQ[TBL]||{};TBLQ[TBL][k]=v;
  clearTimeout(tblTimer);
  tblTimer=setTimeout(function(){
    TBLN[TBL]=PAGE_ROWS;rPane();
    /* the table is rebuilt, so the box being typed into is a new one */
    var inputs=document.querySelectorAll('.th-f');
    for(var i=0;i<inputs.length;i++)if(inputs[i].value===v){inputs[i].focus();
      inputs[i].setSelectionRange(v.length,v.length);break;}
  },260);
};
window.tblClear=function(){TBLQ[TBL]={};TBLN[TBL]=PAGE_ROWS;rPane();};
window.tblOpen=function(id){
  var d=tdef(),rows=d.rows();
  var r=rows.filter(function(x){return String(x.id)===String(id);})[0];
  if(r)d.open(r);
};
/* A terminated record is shaded red across its row, as in the exported
   sheets: a material by its MAT status, a vendor by its pre-qualification,
   a document by its own status. */
/* the numbers or outcomes of a material's linked documents of one kind,
   one to a line */
function linkedVals(m,kind,what){
  return docsOf(m).filter(function(d){return d.doc===kind;}).map(function(d){
    return what==='no'?refOf(d):(normStatus(rawEnd(d,'Status'))||rawEnd(d,'Status'));
  }).filter(Boolean).join('\n');
}
function rowDead(r){
  if(!r||r.m)return false;                         /* a row of a visit page */
  var st;
  if(TBL==='mfr')st=pqOf(r).status;
  else if(r.steps&&!isDoc(r))st=stepOf(r,'mts','status')||(r.raw||{})['MAT Status'];
  else if(isDoc(r))st=rawEnd(r,'Status');
  return /terminat/i.test(String(st||''));
}
function tblRow(r,cols){
  return '<tr'+(rowDead(r)?' class="row-dead"':'')+' onclick="tblOpen('+r.id+')">'
    +cols.map(function(c){
      var v=cellOf(r,c), t=(v!==''&&c.tone)?c.tone(v,r):'';
      var cls=[c.kind==='pick'?'nowrap':'',c.num?'num':''].filter(Boolean).join(' ');
      var sub=c.sub?trim(c.sub(r)):'';
      return '<td'+(cls?' class="'+cls+'"':'')+(v.length>40?' title="'+attr(v)+'"':'')+'>'
        +(t?('<span class="tag t-'+t+'">'+esc(v)+'</span>'):esc(v))
        +(sub?'<div class="cell-sub" title="'+attr(sub)+'">'+esc(sub)+'</div>':'')+'</td>';
    }).join('')+'</tr>';
}
/* Near the bottom, the next rows are added under the ones already drawn.
   Redrawing the whole table instead, then setting the scroll back, ran
   through the page's smooth scrolling from the top — so reaching the end
   of a long list seemed to throw you back to its start. */
var TBL_ADDING=false;
window.tblScroll=function(el){
  if(TBL_ADDING)return;
  if(el.scrollTop+el.clientHeight<el.scrollHeight-400)return;
  var body=document.getElementById('tbl-rows');
  var rows=filtered();
  var from=TBLN[TBL]||PAGE_ROWS;
  if(from>=rows.length)return;
  TBL_ADDING=true;
  try{
    var upto=Math.min(from+PAGE_ROWS,rows.length);
    TBLN[TBL]=upto;
    if(!body){rPane();return;}
    var cols=shownCols();
    body.insertAdjacentHTML('beforeend',rows.slice(from,upto).map(function(r){return tblRow(r,cols);}).join(''));
    var more=document.getElementById('tbl-more');
    if(more){
      if(upto>=rows.length)more.parentNode.removeChild(more);
      else more.textContent='showing '+upto+' of '+rows.length+' — scroll for more';
    }
  }finally{TBL_ADDING=false;}
};
window.tblCols=function(){
  var d=tdef();
  if(!TBLSHOW[TBL])TBLSHOW[TBL]=d.show.slice();
  var on={};TBLSHOW[TBL].forEach(function(k){on[k]=1;});
  sheet('Columns — '+d.label,
    '<div class="dim" style="font-size:13.5px;margin-bottom:14px">'
    +'Tick what you want to see. The order is fixed; the choice is not.</div>'
    +'<div class="panel"><div class="panel-b">'
    +colsOf(d).map(function(c){
      return '<label class="cl-item"><input type="checkbox" '+(on[c.k]?'checked ':'')
        +'onchange="tblToggle(\''+c.k+'\',this.checked)"><span>'+esc(c.t)+'</span></label>';
    }).join('')
    +'</div></div>'
    +'<div class="f-act" style="margin-top:16px">'
    +'<button class="btn" onclick="tblColsAll(1)">All</button>'
    +'<button class="btn" onclick="tblColsAll(0)">Back to the usual</button>'
    +'<button class="btn btn-p" onclick="closeSheet()">Done</button></div>');
};
window.tblToggle=function(k,on){
  var d=tdef();
  if(!TBLSHOW[TBL])TBLSHOW[TBL]=d.show.slice();
  var list=TBLSHOW[TBL].filter(function(x){return x!==k;});
  if(on){
    list=[];
    d.cols.forEach(function(c){
      if(c.k===k||TBLSHOW[TBL].indexOf(c.k)>=0)list.push(c.k);
    });
  }
  TBLSHOW[TBL]=list;rPane();
};
window.tblColsAll=function(all){
  var d=tdef();
  TBLSHOW[TBL]=all?colsOf(d).map(function(c){return c.k;}):d.show.slice();
  closeSheet();rPane();
};
window.tblExport=function(){
  var d=tdef(),cols=shownCols(),rows=filtered();
  var out=[cols.map(function(c){return c.t;})];
  rows.forEach(function(r){out.push(cols.map(function(c){return cellOf(r,c);}));});
  download(workbook([{name:d.label.slice(0,28),rows:out,
    widths:cols.map(function(c){return Math.min(60,Math.max(12,c.w/7));})}]),
    (DB.project||'Table')+' — '+d.label+' '+today()+'.xlsx');
  toast(rows.length+' rows exported');
};

/* the heading band's pattern: short rounded marks in a deeper shade of
   the same purple, drawn once and tiled */
var HEAD_PATTERN='url("data:image/svg+xml,%3Csvg%20xmlns%3D%22http%3A%2F%2Fwww.w3.org%2F2000%2Fsvg%22%20width%3D%22160%22%20height%3D%22160%22%3E%3Cg%20stroke%3D%22%237e2aa6%22%20stroke-linecap%3D%22round%22%3E%3Cline%20x1%3D%22-1.0%22%20y1%3D%22144.0%22%20x2%3D%221.1%22%20y2%3D%22150.6%22%20stroke-opacity%3D%220.65%22%20stroke-width%3D%223.2%22%2F%3E%3Cline%20x1%3D%22159.0%22%20y1%3D%22144.0%22%20x2%3D%22161.1%22%20y2%3D%22150.6%22%20stroke-opacity%3D%220.65%22%20stroke-width%3D%223.2%22%2F%3E%3Cline%20x1%3D%2246.7%22%20y1%3D%2227.2%22%20x2%3D%2245.9%22%20y2%3D%2240.2%22%20stroke-opacity%3D%220.79%22%20stroke-width%3D%223.6%22%2F%3E%3Cline%20x1%3D%22118.7%22%20y1%3D%2279.2%22%20x2%3D%22114.7%22%20y2%3D%2291.2%22%20stroke-opacity%3D%220.85%22%20stroke-width%3D%222.8%22%2F%3E%3Cline%20x1%3D%2254.4%22%20y1%3D%22117.7%22%20x2%3D%2251.0%22%20y2%3D%22123.0%22%20stroke-opacity%3D%220.76%22%20stroke-width%3D%223.5%22%2F%3E%3Cline%20x1%3D%22111.6%22%20y1%3D%2296.2%22%20x2%3D%22107.5%22%20y2%3D%22107.3%22%20stroke-opacity%3D%220.79%22%20stroke-width%3D%223.6%22%2F%3E%3Cline%20x1%3D%22129.1%22%20y1%3D%2235.3%22%20x2%3D%22132.3%22%20y2%3D%2242.1%22%20stroke-opacity%3D%220.81%22%20stroke-width%3D%223.9%22%2F%3E%3Cline%20x1%3D%22114.1%22%20y1%3D%227.9%22%20x2%3D%22109.3%22%20y2%3D%229.3%22%20stroke-opacity%3D%220.55%22%20stroke-width%3D%223.8%22%2F%3E%3Cline%20x1%3D%22114.1%22%20y1%3D%22167.9%22%20x2%3D%22109.3%22%20y2%3D%22169.3%22%20stroke-opacity%3D%220.55%22%20stroke-width%3D%223.8%22%2F%3E%3Cline%20x1%3D%22151.2%22%20y1%3D%22126.9%22%20x2%3D%22144.0%22%20y2%3D%22131.4%22%20stroke-opacity%3D%220.72%22%20stroke-width%3D%223.1%22%2F%3E%3Cline%20x1%3D%2216.4%22%20y1%3D%2276.6%22%20x2%3D%2222.9%22%20y2%3D%2287.6%22%20stroke-opacity%3D%220.51%22%20stroke-width%3D%223.4%22%2F%3E%3Cline%20x1%3D%22152.1%22%20y1%3D%22109.6%22%20x2%3D%22143.4%22%20y2%3D%22110.7%22%20stroke-opacity%3D%220.81%22%20stroke-width%3D%223.4%22%2F%3E%3Cline%20x1%3D%2260.7%22%20y1%3D%2247.2%22%20x2%3D%2258.4%22%20y2%3D%2260.2%22%20stroke-opacity%3D%220.79%22%20stroke-width%3D%222.7%22%2F%3E%3Cline%20x1%3D%2229.7%22%20y1%3D%2243.1%22%20x2%3D%2219.2%22%20y2%3D%2252.2%22%20stroke-opacity%3D%220.68%22%20stroke-width%3D%222.6%22%2F%3E%3Cline%20x1%3D%22154.6%22%20y1%3D%2284.8%22%20x2%3D%22145.4%22%20y2%3D%2293.6%22%20stroke-opacity%3D%220.59%22%20stroke-width%3D%223.7%22%2F%3E%3Cline%20x1%3D%22117.5%22%20y1%3D%22137.3%22%20x2%3D%22115.9%22%20y2%3D%22143.6%22%20stroke-opacity%3D%220.71%22%20stroke-width%3D%222.8%22%2F%3E%3Cline%20x1%3D%2213.4%22%20y1%3D%2242.4%22%20x2%3D%225.4%22%20y2%3D%2245.2%22%20stroke-opacity%3D%220.76%22%20stroke-width%3D%223.4%22%2F%3E%3Cline%20x1%3D%22173.4%22%20y1%3D%2242.4%22%20x2%3D%22165.4%22%20y2%3D%2245.2%22%20stroke-opacity%3D%220.76%22%20stroke-width%3D%223.4%22%2F%3E%3Cline%20x1%3D%2293.2%22%20y1%3D%2280.2%22%20x2%3D%2286.2%22%20y2%3D%2282.7%22%20stroke-opacity%3D%220.57%22%20stroke-width%3D%223.5%22%2F%3E%3Cline%20x1%3D%2260.7%22%20y1%3D%2230.8%22%20x2%3D%2274.1%22%20y2%3D%2231.0%22%20stroke-opacity%3D%220.46%22%20stroke-width%3D%223.4%22%2F%3E%3Cline%20x1%3D%22144.3%22%20y1%3D%224.9%22%20x2%3D%22146.6%22%20y2%3D%2212.7%22%20stroke-opacity%3D%220.71%22%20stroke-width%3D%223.8%22%2F%3E%3Cline%20x1%3D%22144.3%22%20y1%3D%22164.9%22%20x2%3D%22146.6%22%20y2%3D%22172.7%22%20stroke-opacity%3D%220.71%22%20stroke-width%3D%223.8%22%2F%3E%3Cline%20x1%3D%2280.6%22%20y1%3D%2223.1%22%20x2%3D%2284.2%22%20y2%3D%2229.9%22%20stroke-opacity%3D%220.69%22%20stroke-width%3D%222.7%22%2F%3E%3Cline%20x1%3D%2254.7%22%20y1%3D%2282.1%22%20x2%3D%2251.7%22%20y2%3D%2288.0%22%20stroke-opacity%3D%220.45%22%20stroke-width%3D%223.0%22%2F%3E%3Cline%20x1%3D%2294.2%22%20y1%3D%2235.4%22%20x2%3D%22102.3%22%20y2%3D%2239.2%22%20stroke-opacity%3D%220.64%22%20stroke-width%3D%223.1%22%2F%3E%3Cline%20x1%3D%22146.3%22%20y1%3D%2266.4%22%20x2%3D%22143.8%22%20y2%3D%2280.1%22%20stroke-opacity%3D%220.66%22%20stroke-width%3D%222.9%22%2F%3E%3Cline%20x1%3D%2277.4%22%20y1%3D%22106.8%22%20x2%3D%2284.2%22%20y2%3D%22113.7%22%20stroke-opacity%3D%220.76%22%20stroke-width%3D%222.9%22%2F%3E%3Cline%20x1%3D%22108.0%22%20y1%3D%2237.7%22%20x2%3D%22118.7%22%20y2%3D%2242.7%22%20stroke-opacity%3D%220.47%22%20stroke-width%3D%223.9%22%2F%3E%3Cline%20x1%3D%2294.0%22%20y1%3D%2253.0%22%20x2%3D%2293.8%22%20y2%3D%2259.2%22%20stroke-opacity%3D%220.70%22%20stroke-width%3D%223.4%22%2F%3E%3Cline%20x1%3D%2216.9%22%20y1%3D%22126.4%22%20x2%3D%2218.6%22%20y2%3D%22135.8%22%20stroke-opacity%3D%220.59%22%20stroke-width%3D%222.7%22%2F%3E%3Cline%20x1%3D%2216.9%22%20y1%3D%2263.9%22%20x2%3D%227.9%22%20y2%3D%2266.2%22%20stroke-opacity%3D%220.64%22%20stroke-width%3D%223.1%22%2F%3E%3Cline%20x1%3D%2225.7%22%20y1%3D%227.0%22%20x2%3D%2223.5%22%20y2%3D%2217.9%22%20stroke-opacity%3D%220.61%22%20stroke-width%3D%222.7%22%2F%3E%3Cline%20x1%3D%2271.8%22%20y1%3D%2290.1%22%20x2%3D%2274.5%22%20y2%3D%2299.6%22%20stroke-opacity%3D%220.45%22%20stroke-width%3D%222.7%22%2F%3E%3Cline%20x1%3D%2226.9%22%20y1%3D%22146.6%22%20x2%3D%2217.2%22%20y2%3D%22148.4%22%20stroke-opacity%3D%220.75%22%20stroke-width%3D%223.4%22%2F%3E%3Cline%20x1%3D%22139.2%22%20y1%3D%22139.8%22%20x2%3D%22144.0%22%20y2%3D%22150.7%22%20stroke-opacity%3D%220.72%22%20stroke-width%3D%223.0%22%2F%3E%3Cline%20x1%3D%2210.4%22%20y1%3D%22-4.5%22%20x2%3D%2213.9%22%20y2%3D%225.0%22%20stroke-opacity%3D%220.59%22%20stroke-width%3D%223.8%22%2F%3E%3Cline%20x1%3D%2210.4%22%20y1%3D%22155.5%22%20x2%3D%2213.9%22%20y2%3D%22165.0%22%20stroke-opacity%3D%220.59%22%20stroke-width%3D%223.8%22%2F%3E%3Cline%20x1%3D%2261.4%22%20y1%3D%22138.9%22%20x2%3D%2258.2%22%20y2%3D%22147.3%22%20stroke-opacity%3D%220.70%22%20stroke-width%3D%223.9%22%2F%3E%3Cline%20x1%3D%2298.1%22%20y1%3D%22146.0%22%20x2%3D%22105.7%22%20y2%3D%22149.0%22%20stroke-opacity%3D%220.62%22%20stroke-width%3D%224.0%22%2F%3E%3Cline%20x1%3D%2229.5%22%20y1%3D%2289.8%22%20x2%3D%2237.3%22%20y2%3D%2290.8%22%20stroke-opacity%3D%220.48%22%20stroke-width%3D%223.0%22%2F%3E%3Cline%20x1%3D%2279.1%22%20y1%3D%221.2%22%20x2%3D%2280.3%22%20y2%3D%229.3%22%20stroke-opacity%3D%220.56%22%20stroke-width%3D%223.8%22%2F%3E%3Cline%20x1%3D%2279.1%22%20y1%3D%22161.2%22%20x2%3D%2280.3%22%20y2%3D%22169.3%22%20stroke-opacity%3D%220.56%22%20stroke-width%3D%223.8%22%2F%3E%3Cline%20x1%3D%2225.4%22%20y1%3D%22109.7%22%20x2%3D%2216.4%22%20y2%3D%22110.5%22%20stroke-opacity%3D%220.60%22%20stroke-width%3D%222.7%22%2F%3E%3Cline%20x1%3D%22118.7%22%20y1%3D%2264.9%22%20x2%3D%22126.1%22%20y2%3D%2269.4%22%20stroke-opacity%3D%220.51%22%20stroke-width%3D%223.0%22%2F%3E%3Cline%20x1%3D%2243.1%22%20y1%3D%2263.0%22%20x2%3D%2232.0%22%20y2%3D%2266.6%22%20stroke-opacity%3D%220.48%22%20stroke-width%3D%223.3%22%2F%3E%3Cline%20x1%3D%22151.4%22%20y1%3D%2229.6%22%20x2%3D%22142.5%22%20y2%3D%2234.3%22%20stroke-opacity%3D%220.51%22%20stroke-width%3D%223.0%22%2F%3E%3Cline%20x1%3D%22105.0%22%20y1%3D%2219.6%22%20x2%3D%22101.5%22%20y2%3D%2223.8%22%20stroke-opacity%3D%220.60%22%20stroke-width%3D%223.0%22%2F%3E%3Cline%20x1%3D%22108.7%22%20y1%3D%2265.7%22%20x2%3D%22102.4%22%20y2%3D%2271.6%22%20stroke-opacity%3D%220.47%22%20stroke-width%3D%223.0%22%2F%3E%3Cline%20x1%3D%224.8%22%20y1%3D%2219.1%22%20x2%3D%227.0%22%20y2%3D%2224.4%22%20stroke-opacity%3D%220.49%22%20stroke-width%3D%223.4%22%2F%3E%3Cline%20x1%3D%22164.8%22%20y1%3D%2219.1%22%20x2%3D%22167.0%22%20y2%3D%2224.4%22%20stroke-opacity%3D%220.49%22%20stroke-width%3D%223.4%22%2F%3E%3Cline%20x1%3D%22104.0%22%20y1%3D%22123.2%22%20x2%3D%22111.6%22%20y2%3D%22124.6%22%20stroke-opacity%3D%220.75%22%20stroke-width%3D%223.4%22%2F%3E%3Cline%20x1%3D%22129.7%22%20y1%3D%22149.4%22%20x2%3D%22128.8%22%20y2%3D%22162.4%22%20stroke-opacity%3D%220.57%22%20stroke-width%3D%223.2%22%2F%3E%3Cline%20x1%3D%22129.7%22%20y1%3D%22-10.6%22%20x2%3D%22128.8%22%20y2%3D%222.4%22%20stroke-opacity%3D%220.57%22%20stroke-width%3D%223.2%22%2F%3E%3Cline%20x1%3D%2246.1%22%20y1%3D%22102.9%22%20x2%3D%2241.1%22%20y2%3D%22104.0%22%20stroke-opacity%3D%220.85%22%20stroke-width%3D%223.1%22%2F%3E%3Cline%20x1%3D%2264.5%22%20y1%3D%2211.8%22%20x2%3D%2267.8%22%20y2%3D%2217.4%22%20stroke-opacity%3D%220.61%22%20stroke-width%3D%223.0%22%2F%3E%3Cline%20x1%3D%2239.5%22%20y1%3D%22145.8%22%20x2%3D%2235.5%22%20y2%3D%22149.9%22%20stroke-opacity%3D%220.68%22%20stroke-width%3D%222.8%22%2F%3E%3Cline%20x1%3D%22124.9%22%20y1%3D%22119.6%22%20x2%3D%22123.8%22%20y2%3D%22124.9%22%20stroke-opacity%3D%220.71%22%20stroke-width%3D%223.6%22%2F%3E%3Cline%20x1%3D%2271.2%22%20y1%3D%2275.1%22%20x2%3D%2269.0%22%20y2%3D%2280.8%22%20stroke-opacity%3D%220.50%22%20stroke-width%3D%222.7%22%2F%3E%3Cline%20x1%3D%22129.7%22%20y1%3D%2211.7%22%20x2%3D%22125.8%22%20y2%3D%2221.7%22%20stroke-opacity%3D%220.66%22%20stroke-width%3D%223.4%22%2F%3E%3Cline%20x1%3D%22134.5%22%20y1%3D%2290.6%22%20x2%3D%22125.3%22%20y2%3D%2298.7%22%20stroke-opacity%3D%220.83%22%20stroke-width%3D%223.0%22%2F%3E%3Cline%20x1%3D%2210.4%22%20y1%3D%2293.8%22%20x2%3D%224.1%22%20y2%3D%22102.4%22%20stroke-opacity%3D%220.73%22%20stroke-width%3D%222.7%22%2F%3E%3Cline%20x1%3D%22170.4%22%20y1%3D%2293.8%22%20x2%3D%22164.1%22%20y2%3D%22102.4%22%20stroke-opacity%3D%220.73%22%20stroke-width%3D%222.7%22%2F%3E%3Cline%20x1%3D%2274.7%22%20y1%3D%22123.2%22%20x2%3D%2275.0%22%20y2%3D%22135.4%22%20stroke-opacity%3D%220.75%22%20stroke-width%3D%223.6%22%2F%3E%3Cline%20x1%3D%2272.6%22%20y1%3D%2243.5%22%20x2%3D%2274.3%22%20y2%3D%2252.1%22%20stroke-opacity%3D%220.72%22%20stroke-width%3D%222.8%22%2F%3E%3Cline%20x1%3D%2288.2%22%20y1%3D%22126.7%22%20x2%3D%2296.6%22%20y2%3D%22133.1%22%20stroke-opacity%3D%220.68%22%20stroke-width%3D%223.6%22%2F%3E%3Cline%20x1%3D%2238.7%22%20y1%3D%2245.3%22%20x2%3D%2249.1%22%20y2%3D%2251.8%22%20stroke-opacity%3D%220.53%22%20stroke-width%3D%223.3%22%2F%3E%3Cline%20x1%3D%2218.4%22%20y1%3D%2227.8%22%20x2%3D%2220.5%22%20y2%3D%2237.1%22%20stroke-opacity%3D%220.47%22%20stroke-width%3D%223.4%22%2F%3E%3Cline%20x1%3D%2244.6%22%20y1%3D%220.6%22%20x2%3D%2253.6%22%20y2%3D%228.5%22%20stroke-opacity%3D%220.73%22%20stroke-width%3D%223.8%22%2F%3E%3Cline%20x1%3D%2244.6%22%20y1%3D%22160.6%22%20x2%3D%2253.6%22%20y2%3D%22168.5%22%20stroke-opacity%3D%220.73%22%20stroke-width%3D%223.8%22%2F%3E%3Cline%20x1%3D%2295.3%22%20y1%3D%2298.9%22%20x2%3D%2287.4%22%20y2%3D%22100.2%22%20stroke-opacity%3D%220.59%22%20stroke-width%3D%223.1%22%2F%3E%3Cline%20x1%3D%22140.4%22%20y1%3D%2251.7%22%20x2%3D%22127.7%22%20y2%3D%2256.0%22%20stroke-opacity%3D%220.78%22%20stroke-width%3D%223.6%22%2F%3E%3Cline%20x1%3D%2292.0%22%20y1%3D%225.3%22%20x2%3D%22100.7%22%20y2%3D%225.6%22%20stroke-opacity%3D%220.61%22%20stroke-width%3D%223.3%22%2F%3E%3Cline%20x1%3D%2292.0%22%20y1%3D%22165.3%22%20x2%3D%22100.7%22%20y2%3D%22165.6%22%20stroke-opacity%3D%220.61%22%20stroke-width%3D%223.3%22%2F%3E%3Cline%20x1%3D%2259.0%22%20y1%3D%22157.7%22%20x2%3D%2272.3%22%20y2%3D%22157.7%22%20stroke-opacity%3D%220.74%22%20stroke-width%3D%223.7%22%2F%3E%3Cline%20x1%3D%2259.0%22%20y1%3D%22-2.3%22%20x2%3D%2272.3%22%20y2%3D%22-2.3%22%20stroke-opacity%3D%220.74%22%20stroke-width%3D%223.7%22%2F%3E%3Cline%20x1%3D%22-0.9%22%20y1%3D%2275.5%22%20x2%3D%2210.9%22%20y2%3D%2282.7%22%20stroke-opacity%3D%220.61%22%20stroke-width%3D%222.9%22%2F%3E%3Cline%20x1%3D%22159.1%22%20y1%3D%2275.5%22%20x2%3D%22170.9%22%20y2%3D%2282.7%22%20stroke-opacity%3D%220.61%22%20stroke-width%3D%222.9%22%2F%3E%3Cline%20x1%3D%2239.1%22%20y1%3D%22125.0%22%20x2%3D%2236.4%22%20y2%3D%22131.5%22%20stroke-opacity%3D%220.56%22%20stroke-width%3D%223.4%22%2F%3E%3Cline%20x1%3D%2285.1%22%20y1%3D%22145.0%22%20x2%3D%2273.7%22%20y2%3D%22150.1%22%20stroke-opacity%3D%220.57%22%20stroke-width%3D%223.2%22%2F%3E%3Cline%20x1%3D%227.9%22%20y1%3D%22111.9%22%20x2%3D%227.2%22%20y2%3D%22122.8%22%20stroke-opacity%3D%220.81%22%20stroke-width%3D%222.8%22%2F%3E%3Cline%20x1%3D%22167.9%22%20y1%3D%22111.9%22%20x2%3D%22167.2%22%20y2%3D%22122.8%22%20stroke-opacity%3D%220.81%22%20stroke-width%3D%222.8%22%2F%3E%3Cline%20x1%3D%22151.5%22%20y1%3D%2253.5%22%20x2%3D%22162.7%22%20y2%3D%2253.7%22%20stroke-opacity%3D%220.79%22%20stroke-width%3D%223.3%22%2F%3E%3Cline%20x1%3D%22-8.5%22%20y1%3D%2253.5%22%20x2%3D%222.7%22%20y2%3D%2253.7%22%20stroke-opacity%3D%220.79%22%20stroke-width%3D%223.3%22%2F%3E%3Cline%20x1%3D%2241.4%22%20y1%3D%2213.7%22%20x2%3D%2234.6%22%20y2%3D%2224.8%22%20stroke-opacity%3D%220.51%22%20stroke-width%3D%223.3%22%2F%3E%3Cline%20x1%3D%2267.9%22%20y1%3D%22108.0%22%20x2%3D%2259.8%22%20y2%3D%22109.7%22%20stroke-opacity%3D%220.77%22%20stroke-width%3D%223.4%22%2F%3E%3C%2Fg%3E%3C%2Fsvg%3E")';
function tableCSS(){
  if(document.getElementById('tbl-css'))return;
  var css=document.createElement('style');
  css.id='tbl-css';
  css.textContent=
   '.tbl{border-collapse:separate;border-spacing:0;font-size:13px;width:max-content;min-width:100%}'
  +'.tbl th{position:sticky;top:0;z-index:2;background:var(--card);text-align:left;'
  +'padding:8px 10px 10px;border-bottom:1px solid var(--line-2);vertical-align:top}'
  +'.tbl td{padding:9px 10px;border-bottom:1px solid var(--line);vertical-align:top;'
  +'max-width:520px;overflow:hidden;text-overflow:ellipsis}'
  +'.tbl td.nowrap{white-space:nowrap}'
  +'.tbl td.num{text-align:right;font-variant-numeric:tabular-nums;white-space:nowrap}'
  +'.tbl td .tag{white-space:nowrap}'
  +'.cell-sub{font-size:12px;font-weight:400;color:var(--ink-3);margin-top:3px;line-height:1.4;'
  +'white-space:nowrap;overflow:hidden;text-overflow:ellipsis}'
  +'.tbl tbody tr{cursor:pointer}'
  /* a quiet band on every other row, so a long row can be followed across */
  +'.tbl tbody td{background:var(--card)}'
  +'.tbl tbody tr:nth-child(even) td{background:var(--hover)}'
  +'.tbl tbody tr:hover td{background:var(--sunk)}'
  +'.tbl tbody tr.row-dead td{background:#fde7e9;color:#8a1c24}'
  +'.tbl tbody tr.row-dead:hover td{background:#f9d3d7}'
  /* the first column stays put while the rest scroll sideways */
  +'.tbl th:first-child,.tbl td:first-child{position:sticky;left:0;z-index:1;'
  +'box-shadow:1px 0 0 var(--line-2);font-weight:500}'
  +'.tbl th:first-child{z-index:3}'
  /* the gutter lives inside the pinned column, or scrolled text shows in it */
  +'.tbl th:first-child,.tbl td:first-child{padding-left:18px}'
  /* pinned, it must leave room for the rest on a narrow screen */
  +'.tbl th:first-child{min-width:min(360px,42vw)!important}'
  +'.tbl td:first-child{max-width:min(360px,42vw)}'
  +'.th-b{display:block;width:100%;text-align:left;background:none;border:none;padding:0 0 6px;'
  +'font-weight:600;font-size:12.5px;color:var(--ink);cursor:pointer;white-space:nowrap}'
  +'.th-b:hover{color:var(--wait-t)}'
  +'.th-f{width:100%;font:inherit;font-size:12px;padding:5px 7px;border:1px solid var(--line-2);'
  +'border-radius:6px;background:var(--card);color:var(--ink);outline:none}'
  +'.th-f:focus{border-color:var(--wait);box-shadow:0 0 0 2px var(--wait-b)}'
  +'#tbl-body{padding:0 18px 40px 0;overflow:auto}'
  /* every page is the tab row and the page; the rail is put away */
  +'.side{display:none}'
  /* and every page uses the whole width, its words starting on the left
     as the list pages' do, rather than a centred column on some pages */
  +'.head .wrap,.head-note .wrap,.body .wrap{max-width:none!important;margin:0!important}'
  +'.other-menu{position:fixed;z-index:60;min-width:260px;background:var(--card);border:1px solid var(--line);'
  +'border-radius:10px;box-shadow:0 10px 30px rgba(0,22,58,.18);padding:6px;display:flex;flex-direction:column}'
  +'.other-menu button{display:flex;justify-content:space-between;gap:18px;align-items:center;text-align:left;'
  +'background:none;border:0;border-radius:7px;padding:9px 12px;font:inherit;font-size:14px;color:var(--ink);cursor:pointer}'
  +'.other-menu button:hover,.other-menu button:focus-visible{background:var(--hover)}'
  +'.other-menu button[aria-current]{font-weight:600;box-shadow:inset 3px 0 0 #ffcc3e}'
  +'.other-menu .n{color:var(--ink-4);font-family:var(--mono);font-size:11.5px}'
  +'.topbar-tools{margin-left:auto;display:flex;align-items:center;gap:8px;padding:0 4px 6px;flex-shrink:0}'
  /* the tab row in the project navy, with light words on it */
  +'.topbar{background:#00163a;border-bottom-color:#00163a}'
  +'.topbar .tab{color:rgba(255,255,255,.72)}'
  +'.topbar .tab:hover{background:rgba(255,255,255,.08);color:#fff}'
  +'.topbar .tab[aria-selected=true]{color:#fff;border-bottom-color:#fff}'
  +'.topbar .tab .n{color:rgba(255,255,255,.5)}'
  +'.topbar-tools .btn{background:rgba(255,255,255,.1);border-color:rgba(255,255,255,.28);color:#fff;box-shadow:none}'
  +'.topbar-tools .btn:hover{background:rgba(255,255,255,.18);border-color:rgba(255,255,255,.45)}'
  +'.topbar-tools .btn-q{color:rgba(255,255,255,.75)}'
  +'.topbar-tools .btn-q:hover{background:rgba(255,255,255,.08);color:#fff}'
  +'.topbar :focus-visible{outline-color:#9cc3ff}'
  /* Today in a full yellow block, the colour of the rule under the band,
     from the top of the bar to its foot */
  +'#tab-today{background:#ffcc3e;color:#00163a;font-weight:650;border-radius:0;'
  +'margin:-8px 10px 0 -14px;padding:18px 22px 12px;border-bottom:2px solid transparent}'
  +'#tab-today:hover{background:#ffd866;color:#00163a}'
  +'.topbar #tab-today[aria-selected=true]{color:#00163a;border-bottom-color:#00163a}'
  /* each page's heading band in the project purple, white words on it;
     the quiet grey chips turn to a light glass, the coloured ones keep
     their colour so a warning still reads as one */
  +'.head{background:#963cbd '+HEAD_PATTERN+' repeat;background-size:160px 160px;border-bottom:4px solid #ffcc3e;color:#fff}'
  +'.head .head-t{color:#fff}'
  +'.head .swhy,.head .dim,.head .meta{color:rgba(255,255,255,.82)}'
  +'.head .swhy b{color:#fff}'
  +'.head .chip.flat{background:rgba(255,255,255,.16);color:#fff;border-color:transparent}'
  +'.head .chip.flat:hover{background:rgba(255,255,255,.16)}'
  +'.head .btn-q{color:rgba(255,255,255,.85)}'
  +'.head .btn-q:hover{background:rgba(255,255,255,.12);color:#fff}'
  +'.head :focus-visible{outline-color:#fff}'
  +'@media print{.head{background:none;color:inherit;border-bottom-color:var(--line)}'
  +'.head .head-t,.head .swhy,.head .dim{color:inherit}}'
  /* the band's row split: chips wrap on the left, actions stay together on the right */
  +'.head-m.split{flex-wrap:nowrap;align-items:flex-start}'
  +'.head-chips{display:flex;gap:8px;flex-wrap:wrap;align-items:center;flex:1;min-width:0}'
  +'.head-acts{display:flex;gap:6px;align-items:center;flex-shrink:0;margin-left:14px}'
  /* every action on the band looks the same: a light outlined button */
  +'.head .head-acts .btn-q{border:1px solid rgba(255,255,255,.35);background:rgba(255,255,255,.1);'
  +'color:#fff;border-radius:8px;padding:6px 12px;font-size:13px;min-height:32px}'
  +'.head .head-acts .btn-q:hover{background:rgba(255,255,255,.2);border-color:rgba(255,255,255,.55)}'
  +'.head .head-acts [onclick^="remove"]:hover{background:#c62828;border-color:#c62828}'
  +'@media(max-width:760px){.head-m.split{flex-wrap:wrap}.head-acts{margin-left:0}}'
  /* the way back sits in the band's row, first */
  +'.head .backbtn{background:rgba(255,255,255,.14);border-color:rgba(255,255,255,.35);color:#fff;box-shadow:none}'
  +'.head .backbtn:hover{background:rgba(255,255,255,.24);border-color:rgba(255,255,255,.55)}'
  /* every band the same height: title and one row, centred */
  +'.head{box-sizing:border-box;min-height:128px;display:flex;flex-direction:column;justify-content:center}'
  /* its content keeps its full width, so the words stay where they were */
  +'.head>.wrap{width:100%;box-sizing:border-box}'
  +'.head-note{background:var(--card);border-bottom:1px solid var(--line);padding:10px 40px;flex-shrink:0}'
  +'.head-note .swhy{margin-top:0}'
  +'@media(max-width:880px){.head-note{padding:10px 18px}}'
  +'#saved2{white-space:nowrap}'
  +'@media print{.tbl th{position:static}#tbl-body{padding:0;overflow:visible}'
  +'.th-f{display:none}.tbl{font-size:9px}}';
  document.head.appendChild(css);
}

/* ================================================================
   VISITS
   ----------------------------------------------------------------
   "Visits to the factory while the material is being made" — the words
   were already plural and the form held one. An in-process inspection
   happens every fortnight and a final test can be repeated; keeping
   only the last one makes a handover file that cannot be defended.

   So these two steps keep a list, the way deliveries already do. The
   newest visit is also written onto the step itself, which is what the
   road reads and what the log exports — one column cannot hold four
   dates, and the newest is the one that column should carry.
   ================================================================ */

var VISIT_STEPS={ipi:'In-process inspection',fat:'Final inspection or FAT',irn:'Inspection release note'};
window.VISIT_KINDS=VISIT_STEPS;window.visitPanel=function(m,k){return visitPanel(m,k);};
var VISIT_RESULTS={
  ipi:['Pending','Passed','Passed with comments','Failed'],
  fat:['Pending','Scheduled','Passed','Passed with comments','Failed'],
  irn:['Pending','Sent']
};

function visitsOf(m,k){
  return (m.visits||[]).filter(function(v){return v.step===k;})
    .sort(function(a,b){return String(b.date||'').localeCompare(String(a.date||''));});
}
/* the newest visit is what the step, the road and the log all read */
function restep(m,k){
  var list=visitsOf(m,k);
  m.steps=m.steps||{};
  if(!list.length){delete m.steps[k];return;}
  var n=list[0];
  m.steps[k]={date:n.date||'',by:n.by||'',ref:n.ref||'',status:n.result||'Pending'};
}

function visitPanel(m,k){
  var list=visitsOf(m,k),w=VISIT_TEXT[k];
  return '<div class="cl" style="margin-top:14px"><div class="cl-body" style="padding:4px 15px 6px">'
    +(list.length?list.map(function(v){
        var t=/Passed/.test(v.result)?'ok':/Failed/.test(v.result)?'bad'
             :v.result==='Scheduled'?'wait':'na';
        return '<div class="cl-item" style="cursor:default">'
          +'<span class="tag t-'+t+'" style="min-width:118px;text-align:center;flex-shrink:0">'
          +esc(v.result||'Pending')+'</span>'
          +'<span style="flex:1;min-width:0">'
          +'<span class="mono">'+esc(v.ref||'—')+'</span>'
          +(v.by?('<span class="dim"> · '+esc(v.by)+'</span>'):'')
          +(v.note?('<div class="dim" style="font-size:12.5px;margin-top:2px">'+esc(v.note)+'</div>'):'')
          +'</span>'
          +'<span class="meta" style="flex-shrink:0">'+esc(show(v.date))+'</span>'
          +'<button class="btn-q no-print" onclick="editVisit('+m.id+',\''+k+'\','+v.id+')">Edit</button>'
          +'</div>';
      }).join('')
      :'<div class="dim" style="padding:10px 0;font-size:13.5px">No '+w.what+' recorded yet. '
       +'This step happens more than once — each '+w.what+' is kept, and the newest one is what '
       +'the road and the log read.</div>')
    +'<div class="cl-foot"><button class="btn btn-s" onclick="editVisit('+m.id+',\''+k+'\')">'
    +w.add+'</button>'
    +(list.length>1?('<span class="dim" style="font-size:12.5px">'+list.length+' '+w.what+'s</span>'):'')
    +'</div></div></div>';
}

/* a visit added from elsewhere (the calendar): planned, the step brought up to date */
window.addVisit=function(m,k,date,result){
  m.visits=m.visits||[];
  m.visits.push({id:idMaker()(),step:k,date:date,by:'',ref:'',result:result||'Pending',note:''});
  restep(m,k);
};
window.editVisit=function(matId,k,id){
  var m=mat(matId);if(!m)return;
  var w=VISIT_TEXT[k];
  var v=(m.visits||[]).filter(function(x){return String(x.id)===String(id);})[0]||{};
  sheet((id?w.edit:w.add)+' — '+VISIT_STEPS[k],
     '<div class="form" style="margin:0;padding:0;border:none">'
    +'<div class="f"><label for="v-date">'+w.date+'</label>'
    +'<input id="v-date" class="mono" value="'+attr(show(v.date||today()))+'" '
    +'placeholder="dd/mm/yyyy" autocomplete="off"><span class="err" id="e-v-date"></span></div>'
    +'<div class="f"'+(w.by?'':' hidden')+'><label for="v-by">Inspector</label>'
    +personInput('v-by',v.by||'')+'</div>'
    +'<div class="f wide"><label for="v-ref">'+w.ref+'</label>'
    +'<input id="v-ref" class="mono" value="'+attr(v.ref||'')+'" autocomplete="off"></div>'
    +'<div class="f"><label for="v-res">'+w.res+'</label><select id="v-res">'
    +VISIT_RESULTS[k].map(function(o){
        return '<option'+((v.result||'Pending')===o?' selected':'')+'>'+esc(o)+'</option>';}).join('')
    +'</select></div>'
    +'<div class="f wide"><label for="v-note">Note</label>'
    +'<input id="v-note" value="'+attr(v.note||'')+'" '
    +'placeholder="what was seen, what was left open…" autocomplete="off"></div>'
    +'<div class="f-act"><button class="btn btn-p" onclick="saveVisit('+matId+',\''+k+'\','
    +(id||'null')+')">Save</button>'
    +'<button class="btn-q" onclick="closeSheet()">Cancel</button>'
    +(id?('<span style="flex:1"></span><button class="btn btn-d btn-s" data-pop '
      +'onclick="dropVisit(event,'+matId+',\''+k+'\','+id+')">Delete</button>'):'')
    +'</div></div>');
};
window.saveVisit=function(matId,k,id){
  var m=mat(matId);if(!m)return;
  var d=parseDate(document.getElementById('v-date').value);
  if(d===null||!d){document.getElementById('e-v-date').textContent='Use dd/mm/yyyy';return;}
  var o={id:id||idMaker()(),step:k,date:d,
    by:trim(document.getElementById('v-by').value),
    ref:trim(document.getElementById('v-ref').value),
    result:document.getElementById('v-res').value,
    note:trim(document.getElementById('v-note').value)};
  m.visits=m.visits||[];
  var i=m.visits.map(function(x){return String(x.id);}).indexOf(String(id));
  if(id&&i>-1)m.visits[i]=o;else m.visits.push(o);
  restep(m,k);
  touch();closeSheet();rList();rPane();
};
window.dropVisit=function(ev,matId,k,id){
  var m=mat(matId);if(!m)return;
  popConfirm(ev.currentTarget,'Delete this '+VISIT_TEXT[k].what+'?',
    'The others stay. If this was the newest, the step goes back to the one before it.',
    'Delete',function(){
      m.visits=(m.visits||[]).filter(function(x){return String(x.id)!==String(id);});
      restep(m,k);touch();rList();rPane();
    });
};

/* Records that already carry a single visit on the step keep it — it
   becomes the first entry in the list rather than being left behind. */
/* The Main Log said Terminated and the reader, not knowing the word, made
   the step Pending — so a dead submittal sat waiting on a reviewer, and
   the register, comparing against the cell, thought it already matched.
   A step still Pending under a cell that says Terminated is put right
   once; anything already moved on is left alone. */
var TERM_STEPS={mts:'MAT Status',itp:'ITP Status',pid:'PID Status'};
function liftTerminated(){
  var n=0;
  (DB.mats||[]).forEach(function(m){
    if(!m.raw)return;
    Object.keys(TERM_STEPS).forEach(function(k){
      if(normStatus(m.raw[TERM_STEPS[k]])!=='Terminated')return;
      m.steps=m.steps||{};
      var d=m.steps[k]||{};
      if(d.status&&d.status!=='Pending')return;
      d.status='Terminated';m.steps[k]=d;n++;
    });
  });
  if(n)touch();
  return n;
}
function liftVisits(){
  var n=0;
  (DB.mats||[]).forEach(function(m){
    Object.keys(VISIT_STEPS).forEach(function(k){
      var d=(m.steps||{})[k];
      if(!d||!(d.date||d.ref||d.by))return;
      m.visits=m.visits||[];
      if(m.visits.some(function(v){return v.step===k;}))return;
      m.visits.push({id:idMaker()(),step:k,date:d.date||'',by:d.by||'',ref:d.ref||'',
        result:d.status||'Pending',note:''});
      n++;
    });
  });
  if(n)touch();
  return n;
}

/* ================================================================
   THE HEADER READS BEFORE IT WRITES
   ----------------------------------------------------------------
   Every chip along the top was a button. Reaching for a reference to
   copy it opened an editor; a stray click changed a category. A header
   is the part of a record people look at most and change least, so it
   is text until somebody asks for it to be otherwise.

   Edit turns the row live and turns it back. The state belongs to the
   moment, not to the record: move to another and it is reading again.
   ================================================================ */

/* A chip stops being a button. The markup is the page's own, so this
   turns the tag around rather than rebuilding it — anything the page
   adds to a chip later keeps working. */
function chipsToText(s){
  return s.replace(/<button\s+class="chip([^"]*)"([^>]*)>([\s\S]*?)<\/button>/g,
    function(all,cls,attrs,inner){
      var title=/title="([^"]*)"/.exec(attrs);
      return '<span class="chip'+cls+' still"'+(title?(' title="'+title[1]+'"'):'')+'>'
        +inner+'</span>';
    });
}

function readingBar(html,kind,id){
  var at=html.indexOf('<div class="head-m">');
  if(at<0)return html;
  var from=at+20;
  var flex=html.indexOf('<span style="flex:1"></span>',from);
  var end=(flex<0)?html.indexOf('</div>',from):flex;
  if(end<0)return html;

  var chips=chipsToText(html.slice(from,end));

  var btn='<button class="btn btn-s no-print" onclick="editRecord()" '
    +'title="Change any of this">Edit</button>';

  /* the button sits with the other actions on the right, and when there
     is no right-hand group it makes one */
  var out=html.slice(0,from)+chips+(flex<0?'<span style="flex:1"></span>':'')
    +html.slice(end);
  var mark=out.indexOf('<span style="flex:1"></span>');
  out=out.slice(0,mark+28)+btn+out.slice(mark+28);

  /* "Category C3 — from the Main Log." says nothing a person needs while
     reading, and the category is already on a chip two lines above. */
  out=out.replace(/<div class="swhy" style="margin-top:10px">Category [\s\S]*?<\/div>/,'');
  return out;
}
window.__bar=readingBar;

function barCSS(){
  if(document.getElementById('bar-css'))return;
  var css=document.createElement('style');
  css.id='bar-css';
  css.textContent=
   /* the name, while it can be changed, should look like it can */
   '.namebtn{background:none;border:none;padding:2px 8px;margin-left:-8px;text-align:left;'
  +'cursor:pointer;font:inherit;color:inherit;border-radius:8px;'
  +'box-shadow:inset 0 0 0 1px var(--line-2)}'
  +'.namebtn:hover{background:var(--hover);box-shadow:inset 0 0 0 1px var(--line-3)}'
   /* a chip that is not a button should not look like one, and the text
      inside it has to be selectable or copying a reference is no easier
      than it was */
  +'.chip.still{cursor:text;user-select:text;-webkit-user-select:text}'
  +'.chip.still:hover{background:var(--card);border-color:var(--line-2)}'
  +'.chip.still.set:hover{background:var(--wait-b);border-color:var(--wait)}'
  +'.chip.still.warn:hover{background:var(--now-b);border-color:var(--now)}'
  +'.chip.still.flat:hover{background:var(--sunk);border-color:transparent}';
  document.head.appendChild(css);
}

/* ================================================================
   ONE PLACE TO CHANGE THINGS
   ----------------------------------------------------------------
   Edit opens a sheet with everything the header carries on it, and a
   single Save. The header itself never becomes editable — it is a
   thing to read and copy from, and a reference copied is the commonest
   reason anyone touches it.

   The fields are the same ones the chips used to open one at a time,
   so nothing new is being asked for; they are simply all visible at
   once, and a change is not committed until it is asked for.
   ================================================================ */

function optList(list,cur){
  return list.map(function(o){
    var v=(typeof o==='string')?o:o.v, l=(typeof o==='string')?o:o.l;
    return '<option value="'+attr(v)+'"'+(K(v)===K(cur||'')?' selected':'')+'>'+esc(l)+'</option>';
  }).join('');
}

window.editRecord=function(){
  if(TAB==='mfr')return editVendorSheet();
  if(TAB==='insp')return editInspSheet();
  return editMatSheet();
};

/* ---------------------------------------------------------------- material */
function editMatSheet(){
  var m=mat(SEL.mat);if(!m)return;
  var vendors=(DB.mfrs||[]).filter(function(v){return kindOf(v)!=='agency';})
    .sort(function(a,b){return String(a.name).localeCompare(String(b.name));});
  var cur=m.mfr?mfr(m.mfr):null;
  var discs={};(DB.mats||[]).forEach(function(x){if(x.disc)discs[x.disc]=1;});
  sheet('Edit — '+(m.doc?m.doc:'material'),
     '<div class="form" style="margin:0;padding:0;border:none">'
    +'<div class="f wide"><label for="ed-name">Name</label>'
      +'<input id="ed-name" value="'+attr(m.name||'')+'" autocomplete="off"></div>'
    +'<div class="f"><label for="ed-cat">Category</label><select id="ed-cat">'
      +'<option value="">—</option>'+optList(['C0','C1','C2','C3'],m.cat)+'</select></div>'
    +'<div class="f"><label for="ed-disc">Discipline</label>'
      +'<input id="ed-disc" list="disc-list" value="'+attr(m.disc||'')+'" autocomplete="off">'
      +'<datalist id="disc-list">'
      +Object.keys(discs).sort().map(function(d){return '<option value="'+attr(d)+'">';}).join('')
      +'</datalist></div>'
    +'<div class="f wide"><label for="ed-ven">Vendor '
      +'<span class="dim">— type to narrow the list</span></label>'
      +'<input id="ed-ven" list="ven-list" value="'+attr(cur?cur.name:'')+'" '
      +'placeholder="none" autocomplete="off">'
      +'<datalist id="ven-list">'
      +vendors.map(function(v){return '<option value="'+attr(v.name)+'">';}).join('')
      +'</datalist></div>'
    +'<div class="f wide"><label for="ed-sub">Sub-contractor '
      +'<span class="dim">— type to narrow the list</span></label>'
      +'<input id="ed-sub" list="sub-list" value="'+attr(m.sub||'')+'" placeholder="none" autocomplete="off">'
      /* the subcontractors on the vendor list, and any name already
         written on a material */
      +'<datalist id="sub-list">'
      +(function(){
        var seen={},out=[];
        (DB.mfrs||[]).forEach(function(x){if((kindOf(x)==='sub'||kindOf(x)==='makesub')&&x.name&&!seen[K(x.name)]){seen[K(x.name)]=1;out.push(x.name);}});
        (DB.mats||[]).forEach(function(x){if(x.sub&&!seen[K(x.sub)]){seen[K(x.sub)]=1;out.push(x.sub);}});
        return out.sort().map(function(n){return '<option value="'+attr(n)+'">';}).join('');
      })()
      +'</datalist></div>'
    +'<div class="f wide"><label for="ed-ref">Aconex reference</label>'
      +'<input id="ed-ref" class="mono" value="'+attr(m.ref||'')+'" autocomplete="off"></div>'
    +'<div class="f"><label for="ed-qty">Quantity</label>'
      +'<input id="ed-qty" value="'+attr(m.qty||'')+'" autocomplete="off"></div>'
    +'<div class="f"><label for="ed-unit">Unit</label>'
      +'<input id="ed-unit" value="'+attr(m.unit||'')+'" autocomplete="off"></div>'
    +'<div class="f-act"><button class="btn btn-p" onclick="saveMatSheet()">Save</button>'
    +'<button class="btn-q" onclick="closeSheet()">Cancel</button>'
    +'<span style="flex:1"></span>'
    +'<span class="dim" style="font-size:12.5px">Nothing changes until you save.</span>'
    +'</div></div>');
}
window.saveMatSheet=function(){
  var m=mat(SEL.mat);if(!m)return;
  function v(id){return trim((document.getElementById(id)||{}).value);}
  var name=v('ed-name');
  if(name){m.name=name;if(m.raw)m.raw['Item Description']=name;}
  m.cat=v('ed-cat');
  m.catFrom=m.cat?'set by you':'';m.catSure=true;
  m.disc=v('ed-disc');
  m.sub=v('ed-sub');
  m.ref=v('ed-ref');
  m.qty=v('ed-qty');
  m.unit=v('ed-unit');
  /* the vendor is named, not numbered — an unknown name is made rather
     than dropped, because typing one is how a person says it exists */
  var want=v('ed-ven');
  if(!want)m.mfr='';
  else{
    var found=(DB.mfrs||[]).filter(function(x){return K(x.name)===K(want);})[0];
    if(!found){
      found={id:idMaker()(),name:want,kind:'maker',cat:m.cat||'',country:'',site:'',
        scope:'',steps:{},pq:{},added:today()};
      DB.mfrs.push(found);
      toast('Added '+want+' to the vendors');
    }
    m.mfr=String(found.id);
  }
  touch();closeSheet();rList();rPane();
};

/* ---------------------------------------------------------------- vendor */
function editVendorSheet(){
  var v=mfr(SEL.mfr);if(!v)return;
  var st=v.steps||{}, pq=pqOf(v), iso=st.iso||{};
  var subs=(DB.mfrs||[]).filter(function(x){return String(x.id)!==String(v.id);})
    .sort(function(a,b){
      var A=(a.kind==='sub'||a.kind==='makesub')?0:1,B=(b.kind==='sub'||b.kind==='makesub')?0:1;
      return A-B||String(a.name).localeCompare(String(b.name));});
  sheet('Edit — vendor',
     '<div class="form" style="margin:0;padding:0;border:none">'
    +'<div class="f wide"><label for="ev-name">Name</label>'
      +'<input id="ev-name" value="'+attr(v.name||'')+'" autocomplete="off"></div>'
    +'<div class="f"><label for="ev-kind">Kind</label><select id="ev-kind">'
      +optList(Object.keys(KINDS).map(function(k){return {v:k,l:KINDS[k].l};}),kindOf(v))
      +'</select></div>'
    +'<div class="f"><label for="ev-cat">Highest category supplied</label><select id="ev-cat">'
      +'<option value="">—</option>'+optList(['C0','C1','C2','C3'],v.cat)+'</select></div>'
    +'<div class="f"><label for="ev-country">Country</label>'
      +'<input id="ev-country" value="'+attr(v.country||'')+'" autocomplete="off"></div>'
    +'<div class="f"><label for="ev-loc">Local / Foreign</label><select id="ev-loc">'
      +'<option value="">—</option>'+optList(['Local','Foreign'],v.locality)+'</select></div>'
    +'<div class="f"><label for="ev-site">Production site</label>'
      +'<input id="ev-site" value="'+attr(v.site||'')+'" autocomplete="off"></div>'
    +'<div class="f wide"><label for="ev-by">Brought onto the project by</label>'
      +'<input id="ev-by" list="by-list" value="'+attr(v.by||'')+'" placeholder="nobody recorded" '
      +'autocomplete="off"><datalist id="by-list">'
      +subs.map(function(x){return '<option value="'+attr(x.name)+'">';}).join('')
      +'</datalist></div>'
    +'<div class="f wide"><label for="ev-scope">Scope of work</label>'
      +'<input id="ev-scope" value="'+attr(v.scope||'')+'" autocomplete="off"></div>'
    +'<div class="f wide"><label for="ev-pq">'
      +((v.kind==='agency')?'Approval reference':'Pre-qualification reference')+'</label>'
      +'<input id="ev-pq" class="mono" value="'+attr(pq.ref||'')+'" autocomplete="off"></div>'
    +'<div class="f"><label for="ev-pqst">Outcome</label><select id="ev-pqst">'
      +'<option value="">—</option>'
      +optList(['Approved','Approved with comments','Under Review','Revise & Resubmit',
                'Rejected','Terminated'],pq.status)+'</select></div>'
    +'<div class="f"><label for="ev-pqdt">Submitted</label>'
      +'<input id="ev-pqdt" class="mono" value="'+attr(show(pq.date||''))+'" '
      +'placeholder="dd/mm/yyyy" autocomplete="off"><span class="err" id="e-ev-pqdt"></span></div>'
    +'<div class="f"><label for="ev-iso">ISO 9001 certificate</label>'
      +'<input id="ev-iso" class="mono" value="'+attr(iso.ref||'')+'" autocomplete="off"></div>'
    +'<div class="f"><label for="ev-isodt">Expires</label>'
      +'<input id="ev-isodt" class="mono" value="'+attr(show(iso.date||''))+'" '
      +'placeholder="dd/mm/yyyy" autocomplete="off"><span class="err" id="e-ev-isodt"></span></div>'
    +'<div class="f-act"><button class="btn btn-p" onclick="saveVendorSheet()">Save</button>'
    +'<button class="btn-q" onclick="closeSheet()">Cancel</button>'
    +'<span style="flex:1"></span>'
    +'<span class="dim" style="font-size:12.5px">Nothing changes until you save.</span>'
    +'</div></div>');
}
window.saveVendorSheet=function(){
  var v=mfr(SEL.mfr);if(!v)return;
  function g(id){return trim((document.getElementById(id)||{}).value);}
  function date(id){
    var raw=g(id);
    if(!raw)return '';
    var d=parseDate(raw);
    if(d===null||!d){
      var e=document.getElementById('e-'+id);
      if(e)e.textContent='Use dd/mm/yyyy';
      return null;
    }
    return d;
  }
  var pqdt=date('ev-pqdt'), isodt=date('ev-isodt');
  if(pqdt===null||isodt===null)return;
  var old=v.name;
  var name=g('ev-name');
  if(name&&K(name)!==K(old)){
    v.name=name;
    (DB.mfrs||[]).forEach(function(x){if(x.by&&K(x.by)===K(old))x.by=name;});
  }
  v.kind=g('ev-kind')||v.kind;
  v.cat=g('ev-cat');
  v.country=g('ev-country');
  v.locality=g('ev-loc');
  v.site=g('ev-site');
  v.by=g('ev-by');
  v.scope=g('ev-scope');
  v.steps=v.steps||{};
  var slot=(v.kind==='agency')?'appr':'pqd';
  var pq=v.steps[slot]||{};
  pq.ref=g('ev-pq');pq.status=g('ev-pqst');
  if(pqdt)pq.date=pqdt;else delete pq.date;
  if(pq.ref||pq.status||pq.date)v.steps[slot]=pq;else delete v.steps[slot];
  var iso=v.steps.iso||{};
  iso.ref=g('ev-iso');
  if(isodt)iso.date=isodt;else delete iso.date;
  /* a certificate with a date and nothing said about it is valid until
     that date, which is what a certificate means */
  if((iso.ref||iso.date)&&!iso.status)iso.status='Valid';
  if(iso.ref||iso.date)v.steps.iso=iso;else delete v.steps.iso;
  touch();closeSheet();rList();rPane();
};

/* ---------------------------------------------------------------- inspector */
function editInspSheet(){
  var p=(DB.people||[]).filter(function(x){return String(x.id)===String(SEL.insp);})[0];
  if(!p)return;
  var agencies={};(DB.mfrs||[]).forEach(function(v){if(kindOf(v)==='agency')agencies[v.name]=1;});
  (DB.people||[]).forEach(function(x){if(x.agency)agencies[x.agency]=1;});
  sheet('Edit — inspector',
     '<div class="form" style="margin:0;padding:0;border:none">'
    +'<div class="f wide"><label for="ei-name">Name</label>'
      +'<input id="ei-name" value="'+attr(p.name||'')+'" autocomplete="off"></div>'
    +'<div class="f wide"><label for="ei-agency">Agency</label>'
      +'<input id="ei-agency" list="ag-list" value="'+attr(p.agency||'')+'" autocomplete="off">'
      +'<datalist id="ag-list">'
      +Object.keys(agencies).sort().map(function(a){return '<option value="'+attr(a)+'">';}).join('')
      +'</datalist></div>'
    +'<div class="f"><label for="ei-disc">Discipline</label>'
      +'<input id="ei-disc" value="'+attr(p.disc||'')+'" autocomplete="off"></div>'
    +'<div class="f"><label for="ei-status">Approval '
      +'<span class="dim">— clause 2.2.17</span></label><select id="ei-status">'
      +optList(['Pending','Approved','Suspended'],p.status||'Pending')+'</select></div>'
    +'<div class="f wide"><label for="ei-ref">Approval reference</label>'
      +'<input id="ei-ref" class="mono" value="'+attr(p.ref||'')+'" autocomplete="off"></div>'
    +'<div class="f-act"><button class="btn btn-p" onclick="saveInspSheet()">Save</button>'
    +'<button class="btn-q" onclick="closeSheet()">Cancel</button>'
    +'<span style="flex:1"></span>'
    +'<span class="dim" style="font-size:12.5px">Nothing changes until you save.</span>'
    +'</div></div>');
}
window.saveInspSheet=function(){
  var p=(DB.people||[]).filter(function(x){return String(x.id)===String(SEL.insp);})[0];
  if(!p)return;
  function g(id){return trim((document.getElementById(id)||{}).value);}
  var old=p.name, name=g('ei-name');
  if(name&&K(name)!==K(old)){
    p.name=name;
    /* a name typed onto a step is loose text and would otherwise point
       at somebody who no longer exists under that spelling */
    (DB.mats||[]).concat(DB.mfrs||[]).forEach(function(r){
      Object.keys(r.steps||{}).forEach(function(k){
        if(K(r.steps[k].by)===K(old))r.steps[k].by=name;});
      (r.visits||[]).forEach(function(vv){if(K(vv.by)===K(old))vv.by=name;});
    });
  }
  p.agency=g('ei-agency');
  p.disc=g('ei-disc');
  p.status=g('ei-status')||'Pending';
  p.ref=g('ei-ref');
  touch();closeSheet();rList();rPane();
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
  if(!document.getElementById('xl-rep')){
    var p=document.createElement('input');
    p.type='file';p.id='xl-rep';p.accept='.xlsx';p.style.display='none';
    p.addEventListener('change',window.repRead);
    document.body.appendChild(p);
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
      /* Aconex is the source now. The Main Log is only written out, as a
         report; it is no longer read back in. */
      '<div class="sec">The Aconex register</div>'
     +'<div class="panel"><div class="panel-b">'
     +'<div class="dim" style="font-size:13.5px;line-height:1.7">'
     +'The register export, read as it comes. What is already here is brought up to date from it; '
     +'what is not is shown, and can be brought in.</div>'
     +'<div class="f-act" style="margin-top:14px">'
     +'<button class="btn btn-p" onclick="regPick()">Upload the register</button>'
     +'<button class="btn" onclick="regReview()">Waiting to be reviewed</button>'
     +'</div></div></div>'
     +'<div class="sec">The Main Log</div>'
     +'<div class="panel"><div class="panel-b">'
     +'<div class="dim" style="font-size:13.5px;line-height:1.7">'
     +'Written out in its '+COLS.length+' columns, for whoever still reads it. It is not read back: '
     +'the records come from Aconex.</div>'
     +'<div class="f-act" style="margin-top:14px">'
     +'<button class="btn" onclick="excelOut()">Download the workbook</button>'
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
    var n=labelDocuments()+liftVisits()+liftTerminated()+liftMirLinks()+liftDocSteps();
    if(n){rList();rPane();}
    else paintTabs();
  }catch(e){}
}
/* ================================================================
   GOING BACK TO WHERE YOU WERE
   ----------------------------------------------------------------
   Every move between pages — a tab, a record, back to a list — is a
   step in the browser's history, carrying the page, the record, the
   view and how far down it was scrolled. Back (the button, the
   browser's, the mouse's) returns to the same list, the same filters
   and the same row. The filters and the sort already stay with each
   table; the scroll and the page are what this keeps.
   ================================================================ */
var NAVING=false, NAVDEPTH=0, LISTPOS={}, TABLAST={};
/* the top-row tab a place belongs to: Materials, MIR, Documents… are
   views of one tab underneath, but each is its own tab on the row */
function navTabKey(x){return x.tab==='mat'?x.view:x.tab;}
function navStillThere(x){
  var id=x.sel&&x.sel[x.tab];
  if(x.tab==='mat')return !!mat(id);
  if(x.tab==='mfr')return !!mfr(id);
  if(x.tab==='insp')return !!insp(id);
  return false;
}
/* what actually scrolls: the page's body on a wide screen, the whole
   document on a narrow one */
function navScroller(){
  var b=document.querySelector('#pane .body');
  if(b&&b.scrollHeight>b.clientHeight+2)return b;
  return document.scrollingElement||document.documentElement;
}
function navScrollTo(y){
  if(!y)return;
  /* twice: once the page is drawn, and again once its rows are laid out */
  [0,60].forEach(function(t){setTimeout(function(){
    var b=navScroller();
    var was=b.style.scrollBehavior;b.style.scrollBehavior='auto';b.scrollTop=y;b.style.scrollBehavior=was;
  },t);});
}
function navScroll(){
  var l=document.querySelector('.side-l');
  return {b:navScroller().scrollTop,l:l?l.scrollTop:0};
}
function navSnap(){
  return {nav:1,tab:TAB,view:VIEW,record:!!RECORD,dockind:DOCKIND,
    sel:{mat:SEL.mat,mfr:SEL.mfr,insp:SEL.insp},
    tbl:listKey()||'',tbln:listKey()?(TBLN[listKey()]||0):0,scroll:navScroll()};
}
function navKey(x){
  return [x.tab,x.view,x.record,x.record?x.sel[x.tab]:'',x.dockind].join('|');
}
function navRestore(x){
  NAVING=true;
  try{
    closeSheet();
    DOCKIND=x.dockind||'';
    if(x.tbl&&x.tbln)TBLN[x.tbl]=Math.max(TBLN[x.tbl]||0,x.tbln);
    if(x.record&&x.sel[x.tab]){VIEW=x.view;window.jump(x.tab,x.sel[x.tab]);}
    else window.setTab(x.tab==='mat'?x.view:x.tab);
    DOCKIND=x.dockind||'';
  }finally{NAVING=false;}
  /* after the page is drawn: the same distance down, without the
     smooth scroll the page otherwise uses */
  navScrollTo(x.scroll&&x.scroll.b);
  setTimeout(function(){
    var l=document.querySelector('.side-l');
    if(l&&x.scroll&&x.scroll.l)l.scrollTop=x.scroll.l;
  },0);
}
function installNav(){
  if(window.__nav)return;window.__nav=true;
  /* Choosing a record anywhere opens that record, over its list, and is
     a step Back returns from — never the list it sits in. */
  window.pick=function(id){return window.jump(TAB,id);};
  /* Each tab of the top row remembers where it was left: a material
     open on Materials is still open when Materials is pressed again
     from another tab. Pressed while already on it, a tab shows its list,
     as "All …" does. */
  window.tabGo=function(k){
    var here=navTabKey(navSnap()), was=TABLAST[k];
    if(k!==here&&was&&was.record&&navStillThere(was)){
      var before=navSnap();
      TABLAST[here]=before;
      try{window.history.replaceState(before,'');}catch(e){}
      navRestore(was);
      try{window.history.pushState(navSnap(),'');}catch(e){}
      return;
    }
    window.setTab(k);
  };
  /* the top row's tabs go through it */
  setTimeout(function(){
    [].forEach.call(document.querySelectorAll('.topbar .tab'),function(b){
      var k=b.id==='tab-today'?'home':b.id.replace(/^tab-/,'');
      if(!k||b.id==='tab-other')return;
      b.removeAttribute('onclick');
      b.onclick=function(){window.tabGo(k);};
    });
  },0);
  ['setTab','jump','listBack'].forEach(function(fn){
    var orig=window[fn];if(typeof orig!=='function')return;
    window[fn]=function(){
      if(NAVING||NAVDEPTH||!DB)return orig.apply(this,arguments);
      var before=navSnap();
      TABLAST[navTabKey(before)]=before;
      if(!before.record&&before.tbl)LISTPOS[before.tbl]=before.scroll.b;
      try{window.history.replaceState(before,'');}catch(e){}
      NAVDEPTH++;
      try{return orig.apply(this,arguments);}
      finally{
        NAVDEPTH--;
        var after=navSnap();
        if(navKey(after)!==navKey(before)){try{window.history.pushState(after,'');}catch(e){}}
        /* a list comes back where it was left, however it is reached */
        if(!after.record&&after.tbl&&LISTPOS[after.tbl])navScrollTo(LISTPOS[after.tbl]);
      }
    };
  });
  window.addEventListener('popstate',function(e){
    if(e.state&&e.state.nav&&DB)navRestore(e.state);
  });
  /* Alt+Left, as in a browser, for when the hand is on the keyboard */
  document.addEventListener('keydown',function(e){
    if(e.altKey&&e.key==='ArrowLeft'){e.preventDefault();window.history.back();}
  });
}
window.navBack=function(){window.history.back();};

/* ================================================================
   THE PAGE'S OWN DROPDOWNS
   ----------------------------------------------------------------
   A list opened from a field is drawn by the page, not by the browser:
   the same card, type and colours as everything else. The field
   underneath stays a plain <select> or <input>, so everything that
   reads its value or listens for its change still does — picking an
   item sets the value and raises the same events the browser would.
   A text field with suggestions (an inspector, who brought a vendor)
   shows those that contain what is typed, the ones that start with it
   first.
   ================================================================ */
var DD=null;                     /* {box, field, items, at, kind} */
function ddClose(){
  if(!DD)return;
  DD.box.remove();
  if(DD.field)DD.field.classList.remove('dd-open');
  DD=null;
}
function ddPlace(box,field){
  var r=field.getBoundingClientRect();
  box.style.minWidth=Math.max(r.width,180)+'px';
  var below=window.innerHeight-r.bottom, h=Math.min(box.scrollHeight,320);
  box.style.left=Math.max(8,Math.min(r.left,window.innerWidth-box.offsetWidth-8))+'px';
  if(below<h+12&&r.top>below){box.style.top='';box.style.bottom=(window.innerHeight-r.top+4)+'px';}
  else{box.style.bottom='';box.style.top=(r.bottom+4)+'px';}
}
function ddFire(el){
  el.dispatchEvent(new Event('input',{bubbles:true}));
  el.dispatchEvent(new Event('change',{bubbles:true}));
}
function ddDraw(){
  if(!DD)return;
  var list=DD.box.querySelector('.dd-list');
  list.innerHTML=DD.items.length?DD.items.map(function(it,i){
    return '<div class="dd-i'+(i===DD.at?' at':'')+(it.sel?' sel':'')+'" data-i="'+i+'" role="option"'
      +(it.sel?' aria-selected="true"':'')+'>'
      +'<span>'+esc(it.text)+'</span>'+(it.note?'<span class="dd-n">'+esc(it.note)+'</span>':'')+'</div>';
  }).join(''):'<div class="dd-none">Nothing matches</div>';
  var at=list.querySelector('.dd-i.at');
  if(at){var t=at.offsetTop,b=t+at.offsetHeight;
    if(t<list.scrollTop)list.scrollTop=t;else if(b>list.scrollTop+list.clientHeight)list.scrollTop=b-list.clientHeight;}
}
function ddPick(i){
  if(!DD||!DD.items[i])return;
  var it=DD.items[i], f=DD.field;
  ddClose();
  if(f.tagName==='SELECT'){f.selectedIndex=it.idx;}
  else f.value=it.value;
  ddFire(f);
  f.focus();
}
/* a <select>: every option, the chosen one marked; a search line
   when the list is long */
function ddSelect(sel){
  ddClose();
  var opts=[].slice.call(sel.options);
  var box=document.createElement('div');
  box.className='dd-menu';box.setAttribute('role','listbox');
  box.innerHTML=(opts.length>10?'<input class="dd-q" placeholder="Search" autocomplete="off">':'')+'<div class="dd-list"></div>';
  document.body.appendChild(box);
  DD={box:box,field:sel,kind:'sel',all:opts.map(function(o,i){
      return {text:o.text,value:o.value,idx:i,sel:i===sel.selectedIndex,dis:o.disabled};
    }).filter(function(x){return !x.dis;})};
  DD.items=DD.all;DD.at=Math.max(0,DD.items.findIndex(function(x){return x.sel;}));
  sel.classList.add('dd-open');
  ddDraw();ddPlace(box,sel);
  var q=box.querySelector('.dd-q');
  if(q){
    q.addEventListener('input',function(){
      var k=q.value.trim().toLowerCase();
      DD.items=DD.all.filter(function(x){return !k||x.text.toLowerCase().indexOf(k)>=0;});
      DD.at=0;ddDraw();
    });
    q.addEventListener('keydown',ddKeys);
    setTimeout(function(){q.focus();},0);
  }
}
/* a text field with a list of suggestions */
function ddSuggest(inp){
  var id=inp.getAttribute('data-list');
  var dl=id&&document.getElementById(id);if(!dl)return;
  /* compared without spaces or punctuation: "alra" finds "Al Ramasat",
     "binladen" finds "Bin Laden" */
  var flat=function(t){return String(t||'').toLowerCase().replace(/[^a-z0-9؀-ۿ]+/g,'');};
  var k=inp.value.trim().toLowerCase(), kf=flat(k);
  var all=[].slice.call(dl.options).map(function(o){return {text:o.value,value:o.value,note:o.label&&o.label!==o.value?o.label:''};});
  var hits=all.filter(function(x){return !kf||flat(x.text).indexOf(kf)>=0;});
  hits.sort(function(a,b){
    var A=flat(a.text).indexOf(kf)===0?0:1, B=flat(b.text).indexOf(kf)===0?0:1;
    return A-B||a.text.localeCompare(b.text);
  });
  hits=hits.slice(0,60).map(function(x){x.sel=x.text.toLowerCase()===k;return x;});
  if(DD&&DD.field===inp){DD.items=hits;DD.at=hits.length?0:-1;ddDraw();ddPlace(DD.box,inp);return;}
  ddClose();
  var box=document.createElement('div');
  box.className='dd-menu';box.setAttribute('role','listbox');
  box.innerHTML='<div class="dd-list"></div>';
  document.body.appendChild(box);
  DD={box:box,field:inp,kind:'text',items:hits,at:-1};
  inp.classList.add('dd-open');
  ddDraw();ddPlace(box,inp);
}
function ddKeys(e){
  if(!DD)return;
  var n=DD.items.length;
  if(e.key==='ArrowDown'){e.preventDefault();DD.at=n?(DD.at+1)%n:-1;ddDraw();}
  else if(e.key==='ArrowUp'){e.preventDefault();DD.at=n?(DD.at-1+n)%n:-1;ddDraw();}
  else if(e.key==='Enter'){if(DD.at>=0){e.preventDefault();ddPick(DD.at);}else ddClose();}
  else if(e.key==='Escape'){e.preventDefault();var f=DD.field;ddClose();f.focus();}
  else if(e.key==='Tab')ddClose();
}
function installDropdowns(){
  if(window.__dd)return;window.__dd=true;
  var css=document.createElement('style');
  css.textContent=
   '.dd-menu{position:fixed;z-index:950;background:var(--card);border:1px solid var(--line-2);border-radius:10px;'
  +'box-shadow:0 12px 32px rgba(0,22,58,.18);padding:5px;display:flex;flex-direction:column;max-width:min(520px,94vw)}'
  +'.dd-list{max-height:300px;overflow-y:auto;overscroll-behavior:contain}'
  +'.dd-i{display:flex;justify-content:space-between;align-items:center;gap:14px;padding:8px 11px;border-radius:7px;'
  +'font-size:14px;color:var(--ink);cursor:pointer;line-height:1.35}'
  +'.dd-i:hover,.dd-i.at{background:var(--hover)}'
  +'.dd-i.at{box-shadow:inset 0 0 0 1px var(--line-2)}'
  +'.dd-i.sel{font-weight:600;box-shadow:inset 3px 0 0 #ffcc3e}'
  +'.dd-n{font-size:12px;color:var(--ink-4);flex-shrink:0}'
  +'.dd-none{padding:9px 11px;font-size:13px;color:var(--ink-4)}'
  +'.dd-q{margin:2px 2px 6px;padding:8px 10px;border:1px solid var(--line-2);border-radius:7px;font:inherit;font-size:13.5px;'
  +'outline:none;background:var(--card);color:var(--ink)}'
  +'.dd-q:focus{border-color:var(--wait)}'
  +'select.dd-open,input.dd-open{border-color:var(--wait)!important;box-shadow:0 0 0 3px var(--wait-b)!important}';
  document.head.appendChild(css);

  /* a press on a select opens ours instead of the browser's */
  document.addEventListener('mousedown',function(e){
    if(DD&&DD.box.contains(e.target)){
      e.preventDefault();
      var it=e.target.closest('.dd-i');if(it)ddPick(Number(it.getAttribute('data-i')));
      return;
    }
    var sel=e.target.closest&&e.target.closest('select');
    if(sel&&!sel.disabled&&!sel.multiple){
      e.preventDefault();
      if(DD&&DD.field===sel){ddClose();return;}
      sel.focus();ddSelect(sel);return;
    }
    if(DD&&e.target!==DD.field)ddClose();
  },true);
  /* the keyboard opens it too, as the browser's would */
  document.addEventListener('keydown',function(e){
    var t=e.target;
    if(DD&&(t===DD.field))return ddKeys(e);
    if(t&&t.tagName==='SELECT'&&!t.multiple&&(e.key===' '||e.key==='Enter'||(e.altKey&&e.key==='ArrowDown'))){
      e.preventDefault();ddSelect(t);
    }
  },true);
  /* a field with suggestions loses the browser's list and gets ours */
  function takeList(t){
    if(t&&t.tagName==='INPUT'&&t.hasAttribute('list')){
      t.setAttribute('data-list',t.getAttribute('list'));t.removeAttribute('list');
    }
    return t&&t.tagName==='INPUT'&&t.hasAttribute('data-list');
  }
  document.addEventListener('focusin',function(e){if(takeList(e.target))ddSuggest(e.target);});
  document.addEventListener('input',function(e){if(e.isTrusted&&takeList(e.target))ddSuggest(e.target);});
  document.addEventListener('focusout',function(e){
    setTimeout(function(){
      if(DD&&DD.field===e.target&&!DD.box.contains(document.activeElement))ddClose();
    },120);
  });
  /* the page moving under an open list closes it */
  window.addEventListener('resize',ddClose);
  document.addEventListener('scroll',function(e){if(DD&&!DD.box.contains(e.target))ddClose();},true);
}

function start(){
  install();
  install2();
  installNav();
  installDropdowns();
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
window.__tablePane=tablePane;
window.__tbl={label:function(){return tdef().label;},count:function(){return tdef().rows().length;},
  cols:function(){return shownCols();},filtered:function(){return filtered().length;},
  first:function(){var r=filtered()[0];return r?r.name:'';},
  choices:function(k){var c=tdef().cols.filter(function(x){return x.k===k;})[0];
    return c?choices(c,tdef().rows()):[];}};
window.__v={lift:liftVisits,list:visitsOf};
/* from a day of the calendar: link a document to the material it
   serves, then come back to that day for the next one */
window.docLinkPick=function(docId,q,ds){
  var d=mat(docId);if(!d)return;
  var need=K(q||'');
  var all=(DB.mats||[]).filter(function(m){return !isDoc(m);});
  var rows=all.filter(function(m){
    if(!need)return d.disc?K(m.disc||'')===K(d.disc):true;
    return K(m.name+' '+matNo(m)+' '+(m.disc||'')).indexOf(need)>=0;
  });
  sheet('Link '+(d.doc||'document')+' to a material',
     '<div style="font-size:13.5px;margin-bottom:12px"><b>'+esc(d.name)+'</b>'
    +'<div class="dim mono" style="font-size:12.5px;margin-top:2px">'+esc(refOf(d))+(d.disc?' · '+esc(d.disc):'')+'</div></div>'
    +'<div class="f" style="margin-bottom:12px"><label for="lk">Search the materials by name, MAT number or discipline</label>'
    +'<input id="lk" value="'+attr(q||'')+'" autocomplete="off" '
    +'oninput="searchSoon(function(v){docLinkPick('+docId+',v,\''+ds+'\');},this.value)">'
    +'<span class="dim" style="font-size:12px">'+(need?('Searching all '+all.length+' materials.')
      :('Showing the '+rows.length+(d.disc?' in '+esc(d.disc):'')+' — type to search them all.'))+'</span></div>'
    +'<div class="panel"><div class="panel-b">'
    +(rows.length?rows.slice(0,60).map(function(m){
        return '<div class="line row-a" onclick="docLinkDo('+m.id+','+docId+',\''+ds+'\')">'
          +'<span class="tag t-na" style="min-width:40px;text-align:center">'+esc(m.cat||'—')+'</span>'
          +'<div class="line-m"><div>'+esc(m.name)+'</div>'
          +'<div class="dim mono" style="font-size:12.5px;margin-top:2px">'+esc(matNo(m))+'</div></div></div>';}).join('')
        +(rows.length>60?('<div class="dim" style="font-size:13px;padding-top:10px">and '+(rows.length-60)+' more — narrow the search</div>'):'')
      :'<span class="dim">No material matches.</span>')
    +'</div></div>'
    +'<div class="f-act" style="margin-top:14px"><button class="btn-q" onclick="daySheet(\''+ds+'\')">← Back to the day</button></div>');
  var f=document.getElementById('lk');if(f){f.focus();f.setSelectionRange(f.value.length,f.value.length);}
};
window.docLinkDo=function(matId,docId,ds){
  window.linkAdd(matId,docId);
  if(ds)setTimeout(function(){daySheet(ds);},60);
};
window.docRef=refOf;
/* the documents behind a step filled from them, one by one */
window.stepDocs=function(m,k){
  var kind=Object.keys(DOC_STEP).filter(function(x){return DOC_STEP[x]===k;})[0];
  if(!kind)return [];
  return docsOf(m).filter(function(d){return d.doc===kind;}).map(function(d){
    var raw=d.raw||{};
    return {id:d.id,no:refOf(d),title:d.name,status:normStatus(raw[kind+' Status'])||trim(raw[kind+' Status']||'')||'Pending',
      date:raw[kind+' Submittal Date']||d.acxDate||''};
  });
};
window.EXCEL={cols:COLS,inspectorRows:inspectorRows,readInspectorSheet:readInspectorSheet,planInspectors:planInspectors,applyInspectors:applyInspectors,vendorRows:vendorRows,readVendorSheet:readVendorSheet,planVendors:planVendors,applyVendors:applyVendors,generalRows:generalRows,looseRows:looseRows,vendorRows:vendorRows,isDoc:isDoc,docsOf:docsOf,servedBy:servedBy,labelDocuments:labelDocuments,createFromRegister:createFromRegister,pending:pending,read:readMainLog,openBook:openBook,readRegister:readRegister,planRegister:planRegister,applyRegister:applyRegister,plan:planFrom,apply:applyPlan,rows:generalRows,summary:summaryRows,book:workbook};
})();

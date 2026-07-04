/* ═══════════════════════════════════════════════════════════════════════════
   MOTOR COMPARTIDO — Matriz Consolidada de Consumo · MIDES · Comedores Sociales
   -----------------------------------------------------------------------------
   Toma uno o varios "Cuadros Consolidados" (plantilla normal o FEGUA) y produce
   el "CUADRO CONSOLIDADO DE CONSUMO POR LICITACIÓN/CONTRATO ABIERTO" (la Matriz):
   un bloque por comedor, agrupando corridas de días con la MISMA ración diaria,
   con SUBTOTAL por comedor y TOTAL general. Genera XLSX (SheetJS/xlsx-js-style)
   y PDF (pdfmake), replicando el formato de las matrices oficiales.

   Regla de agrupación (confirmada con los ejemplos oficiales):
     días consecutivos con (desayuno, almuerzo) idénticos → una sola línea
     "del X al Y", días = cuenta. Cuando cambia el número, se corta la línea.
     Los días 100% en cero (sin servicio) se omiten salvo que se pida incluirlos.

   Todo el costo y metadatos salen del CUADRO INTEGRACIÓN del propio cuadro:
     PRECIO desayuno / PRECIO almuerzo, LICITACIÓN/CONTRATO ABIERTO, Lote, fecha.
   ═══════════════════════════════════════════════════════════════════════════ */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.MatrizConsolidada = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  const MESES = ['enero','febrero','marzo','abril','mayo','junio','julio','agosto',
    'septiembre','octubre','noviembre','diciembre'];

  /* ─────────────── utilidades ─────────────── */
  function toInt(v){
    if(v===null||v===undefined||v==='') return 0;
    const n=parseInt(String(v).replace(/\.0$/,'').replace(/,/g,''),10);
    return isNaN(n)?0:n;
  }
  function toFloat(v){
    if(v===null||v===undefined||v==='') return 0;
    const n=parseFloat(String(v).replace(/,/g,''));
    return isNaN(n)?0:n;
  }
  function norm(v){ return String(v==null?'':v).replace(/\s+/g,' ').trim().toUpperCase(); }
  function tit(t){
    if(!t) return '';
    const exc=['de','la','el','los','las','y','en','del','a'];
    return String(t).toLowerCase().split(/\s+/).filter(Boolean)
      .map((p,i)=>(i===0||!exc.includes(p))?p.charAt(0).toUpperCase()+p.slice(1):p).join(' ')
      .split('S.a.').join('S.A.').split('Fegua').join('FEGUA').split('Car').join('CAR');
  }
  function aFecha(v){
    if(v instanceof Date && !isNaN(v)) return new Date(v.getFullYear(),v.getMonth(),v.getDate());
    if(typeof v==='number' && v>20000 && v<80000){ // serial Excel
      const d=new Date(Date.UTC(1899,11,30)+v*86400000);
      return new Date(d.getUTCFullYear(),d.getUTCMonth(),d.getUTCDate());
    }
    const s=String(v==null?'':v).trim();
    let m=s.match(/^(\d{4})-(\d{2})-(\d{2})/); if(m) return new Date(+m[1],+m[2]-1,+m[3]);
    m=s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/); if(m) return new Date(+m[3],+m[2]-1,+m[1]);
    return null;
  }
  function fmtDMY(d){ return d? String(d.getDate()).padStart(2,'0')+'/'+String(d.getMonth()+1).padStart(2,'0')+'/'+d.getFullYear() : ''; }
  function fechaLarga(d){ return d? d.getDate()+' de '+MESES[d.getMonth()]+' del '+d.getFullYear() : ''; }

  /* ─────────────── 1) Extracción desde el CUADRO INTEGRACIÓN ─────────────── */
  // Recibe un workbook ya leído con SheetJS (XLSX.read). Devuelve los datos que
  // la Matriz necesita de UN comedor. Funciona con la plantilla normal y la FEGUA
  // porque localiza las columnas por su TEXTO de encabezado, no por letra fija.
  function extraerComedor(workbook, XLSX){
    const nombresHoja = workbook.SheetNames;
    let hojaNom = nombresHoja.find(n=>/CUADRO\s+INTEGRACI[OÓ]N/i.test(n));
    if(!hojaNom) hojaNom = nombresHoja.find(n=>/INTEGRACI[OÓ]N/i.test(n));
    if(!hojaNom) return { error:'No se encontró la hoja "CUADRO INTEGRACIÓN" en el Excel.' };
    const ws = workbook.Sheets[hojaNom];
    const rows = XLSX.utils.sheet_to_json(ws,{header:1, raw:true, defval:null});
    const txt  = XLSX.utils.sheet_to_json(ws,{header:1, raw:false, defval:null});
    const G=(r,c)=> (rows[r] && rows[r][c]!==undefined)? rows[r][c] : null;
    const T=(r,c)=> (txt[r]  && txt[r][c]!==undefined && txt[r][c]!==null)? String(txt[r][c]) : '';

    // 1a) localizar columnas por texto de encabezado
    let colNo=-1,colFecha=-1,colDes=-1,colAlmTot=-1,colAlmNorm=-1, hdrRow=-1;
    for(let r=0;r<Math.min(rows.length,40);r++){
      for(let c=0;c<(rows[r]?rows[r].length:0);c++){
        const s=norm(G(r,c));
        if(s==='NO.'&&colNo<0){ colNo=c; }
        if(s==='FECHA'&&colFecha<0){ colFecha=c; hdrRow=r; }
        if(/DESAYUNOS?\s+SERVIDOS/.test(s)&&colDes<0) colDes=c;
        if(/TOTAL\s+DE\s+ALMUERZOS\s+SERVIDOS/.test(s)&&colAlmTot<0) colAlmTot=c;
        if(/(NO\.?\s+DE\s+)?ALMUERZOS\s+SERVIDOS/.test(s)&&!/TOTAL/.test(s)&&colAlmNorm<0) colAlmNorm=c;
      }
    }
    // FEGUA trae "Total de Almuerzos Servidos" (normal + CAR); normal solo "Almuerzos Servidos".
    const esFegua = colAlmTot>=0;
    const colAlm = esFegua ? colAlmTot : colAlmNorm;
    if(colNo<0||colFecha<0||colDes<0||colAlm<0)
      return { error:'No se pudieron ubicar las columnas de días/desayunos/almuerzos en el CUADRO INTEGRACIÓN.' };

    // 1b) precios: celdas "PRECIO:" → valor de la derecha (1º desayuno, 2º almuerzo)
    const precios=[];
    for(let r=0;r<Math.min(rows.length,40);r++)
      for(let c=0;c<(rows[r]?rows[r].length:0);c++)
        if(/^PRECIO/i.test(String(G(r,c)||'').trim())) precios.push(toFloat(G(r,c+1)));
    const precioDes = precios[0]||0, precioAlm = precios[1]||0;

    // 1c) metadatos (título, tipo de contrato, lote, fecha/lugar, modalidad)
    let comedorRaw='', tipoContrato='', lote='', fechaLugar='';
    for(let r=0;r<Math.min(rows.length,14);r++)
      for(let c=0;c<(rows[r]?rows[r].length:0);c++){
        const raw=String(G(r,c)||'');
        const s=norm(raw);
        if(/CUADRO\s+DE\s+CONSOLIDACI[OÓ]N/.test(s)) comedorRaw=raw;
        if(s==='LICITACIÓN'||s==='LICITACION') tipoContrato='LICITACIÓN';
        if(s==='CONTRATO ABIERTO') tipoContrato='CONTRATO ABIERTO';
        const ml=raw.match(/LOTE\s*:?\s*(\d+)/i); if(ml&&!lote) lote=ml[1];
        const mr=raw.match(/REGI[OÓ]N\s*:?\s*(\d+)/i); if(mr&&!lote) lote=mr[1];
        if(/GUATEMALA\s+\d/i.test(raw)&&!fechaLugar) fechaLugar=raw.trim();
      }
    // nombre + modalidad a partir del título "...COMEDOR SOCIAL FIJO FEGUA, GUATEMALA"
    let comedor='Comedor', modalidad='Fijo';
    const mNom=comedorRaw.match(/COMEDOR\s+SOCIAL\s+(.*)$/i);
    if(mNom){
      let n=mNom[1].trim().replace(/[.\s]+$/,'');
      if(/^FIJO\b/i.test(n)){ modalidad='Fijo'; }
      else if(/^M[OÓ]VIL\b/i.test(n)){ modalidad='Móvil'; }
      comedor = tit(n);
    }

    // 1d) filas diarias
    const filas=[];
    for(let r=(hdrRow>=0?hdrRow+1:0); r<rows.length; r++){
      const noRaw=G(r,colNo);
      const noInt=parseInt(String(noRaw==null?'':noRaw).trim(),10);
      const fecha=aFecha(G(r,colFecha));
      if(!Number.isFinite(noInt)||!fecha){
        if(/TOTAL/.test(norm(G(r,colNo)||G(r,colFecha)||''))) break;
        continue;
      }
      filas.push({ fecha, des:toInt(G(r,colDes)), alm:toInt(G(r,colAlm)) });
    }
    if(!filas.length) return { error:'No se encontraron filas de días con datos en el CUADRO INTEGRACIÓN.' };

    const totDes=filas.reduce((a,f)=>a+f.des,0);
    const totAlm=filas.reduce((a,f)=>a+f.alm,0);
    return { esFegua, comedor, modalidad, precioDes, precioAlm, tipoContrato, lote,
             fechaLugar, filas, totDes, totAlm };
  }

  /* ─────────────── 2) Agrupación por ración idéntica ─────────────── */
  function agrupar(filas, incluirCeros){
    const g=[];
    for(const f of filas){
      if(!incluirCeros && f.des===0 && f.alm===0) continue;
      const u=g[g.length-1];
      if(u && u.des===f.des && u.alm===f.alm){ u.hasta=f.fecha; u.dias++; }
      else g.push({ desde:f.fecha, hasta:f.fecha, des:f.des, alm:f.alm, dias:1 });
    }
    return g;
  }

  /* ─────────────── 3) Modelo de la matriz ─────────────── */
  // comedores: [{ no, nombre, modalidad, precioDes, precioAlm, filas }]
  // meta: { depto, tipoContrato, lote, fechaLugar, integracion:[l1,l2?], incluirCeros }
  function construirModelo(comedores, meta){
    meta = meta || {};
    const inc = !!meta.incluirCeros;
    const out = comedores.map((c,i)=>{
      const grupos = agrupar(c.filas, inc);
      const filasM = grupos.map(g=>{
        const H=g.des*g.dias, I=g.alm*g.dias, J=H+I;
        const M=H*(c.precioDes||0), N=I*(c.precioAlm||0), O=M+N;
        return { desde:g.desde, hasta:g.hasta, eDes:g.des, fAlm:g.alm, dias:g.dias,
                 H, I, J, kDes:c.precioDes||0, lAlm:c.precioAlm||0, M, N, O };
      });
      const sub = filasM.reduce((a,f)=>({dias:a.dias+f.dias,H:a.H+f.H,I:a.I+f.I,J:a.J+f.J,M:a.M+f.M,N:a.N+f.N,O:a.O+f.O}),
                                 {dias:0,H:0,I:0,J:0,M:0,N:0,O:0});
      return { no:i+1, nombre:c.nombre||('Comedor '+(i+1)), modalidad:c.modalidad||'Fijo', filas:filasM, sub };
    });
    const total = out.reduce((a,c)=>({H:a.H+c.sub.H,I:a.I+c.sub.I,J:a.J+c.sub.J,M:a.M+c.sub.M,N:a.N+c.sub.N,O:a.O+c.sub.O}),
                             {H:0,I:0,J:0,M:0,N:0,O:0});
    return { comedores:out, total, meta };
  }

  // Texto de "Integración..." por defecto (editable por el usuario en la UI)
  function integracionDefault(comedores, meta){
    const fijos=comedores.filter(c=>/fijo/i.test(c.modalidad)).length;
    const movs =comedores.filter(c=>/m[oó]vil/i.test(c.modalidad)).length;
    const partes=[];
    if(fijos) partes.push(fijos+(fijos===1?' comedor social fijo':' comedores sociales fijos'));
    if(movs)  partes.push(movs +(movs===1 ?' comedor social móvil':' comedores sociales móviles'));
    const conteo = partes.join(' y ') || (comedores.length+' comedores sociales');
    // rango global de fechas
    let ini=null,fin=null;
    comedores.forEach(c=>c.filas.forEach(f=>{ if(!ini||f.fecha<ini)ini=f.fecha; if(!fin||f.fecha>fin)fin=f.fecha; }));
    const rango = (ini&&fin)? ('del '+ini.getDate()+' al '+fin.getDate()+' de '+MESES[fin.getMonth()]+' '+fin.getFullYear()) : 'del periodo';
    const depto = meta.depto? (' del departamento de '+tit(meta.depto)) : '';
    return 'Integración del servicio de raciones preparadas servidas '+rango+' de '+conteo+depto+'.';
  }

  /* ─────────────── 4) Generador XLSX (SheetJS / xlsx-js-style) ─────────────── */
  // Estilos (xlsx-js-style; si es SheetJS plano, los .s se ignoran sin romper nada)
  const THIN={style:'thin',color:{rgb:'000000'}};
  const BORD={top:THIN,bottom:THIN,left:THIN,right:THIN};
  const F_HDR={font:{name:'Arial',sz:11,bold:true},fill:{fgColor:{rgb:'B4C6E7'}},alignment:{horizontal:'center',vertical:'center',wrapText:true},border:BORD};
  const F_CEL={font:{name:'Arial',sz:11},alignment:{horizontal:'center',vertical:'center',wrapText:true},border:BORD};
  const F_CELr={font:{name:'Arial',sz:11},alignment:{horizontal:'right',vertical:'center'},border:BORD};
  const F_CELl={font:{name:'Arial',sz:11},alignment:{horizontal:'left',vertical:'center',wrapText:true},border:BORD};
  const F_BOLD={font:{name:'Arial',sz:11,bold:true},alignment:{horizontal:'center',vertical:'center',wrapText:true},border:BORD};
  const F_TIT={font:{name:'Arial',sz:11,bold:true},alignment:{horizontal:'center',vertical:'center',wrapText:true}};
  const F_SUB={font:{name:'Arial',sz:10},alignment:{horizontal:'center',vertical:'center',wrapText:true}};
  const QFMT='"Q"#,##0.00', NFMT='#,##0', DFMT='dd/mm/yyyy';

  function construirXLSX(modelo, XLSX){
    const ws={}; const merges=[];
    const set=(r,c,v,st,z,t)=>{ const ref=XLSX.utils.encode_cell({r:r-1,c:c-1});
      const cell={v:v,s:st}; if(z)cell.z=z; if(t)cell.t=t; else if(typeof v==='number')cell.t='n'; else if(v instanceof Date){cell.t='d';} else cell.t='s';
      ws[ref]=cell; };
    const mrg=(r1,c1,r2,c2)=>merges.push({s:{r:r1-1,c:c1-1},e:{r:r2-1,c:c2-1}});
    const m=modelo.meta||{};

    // Encabezado superior
    set(1,14,'Lote '+(m.lote||''),{font:{name:'Arial',sz:11,bold:true}});
    set(2,3,'CUADRO CONSOLIDADO DE CONSUMO POR '+(m.tipoContrato||'LICITACIÓN')+', DEPARTAMENTO DE '+String(m.depto||'').toUpperCase(),F_TIT); mrg(2,3,2,13);
    set(3,3,'Programa Comedor Social',F_SUB); mrg(3,3,3,13);
    set(3,14,m.fechaLugar||'',{font:{name:'Arial',sz:10}});
    const integ = m.integracion||[];
    set(5,2,integ[0]||'',F_SUB); mrg(5,2,5,14);
    let filaHdr=8;
    if(integ[1]){ set(6,3,integ[1],F_SUB); mrg(6,3,6,13); }

    // Encabezado de tabla (2 niveles)
    const H=(r,c,v)=>set(r,c,v,F_HDR);
    H(filaHdr,1,'No.');           mrg(filaHdr,1,filaHdr+1,1);
    H(filaHdr,2,'COMEDOR SOCIAL'); mrg(filaHdr,2,filaHdr+1,2);
    H(filaHdr,3,'PERIODO DEL SERVICIO'); mrg(filaHdr,3,filaHdr+1,4); H(filaHdr,4,'');
    H(filaHdr,5,'RACIONES DIARIO'); mrg(filaHdr,5,filaHdr,6);
    H(filaHdr+1,5,'DESAYUNO'); H(filaHdr+1,6,'ALMUERZO');
    H(filaHdr,7,'DÍAS DE SERVICIO'); mrg(filaHdr,7,filaHdr+1,7);
    H(filaHdr,8,'RACIONES DE ALIMENTOS'); mrg(filaHdr,8,filaHdr,10); H(filaHdr,9,''); H(filaHdr,10,'');
    H(filaHdr+1,8,'DESAYUNO'); H(filaHdr+1,9,'ALMUERZO'); H(filaHdr+1,10,'TOTAL');
    H(filaHdr,11,'COSTO POR RACIÓN'); mrg(filaHdr,11,filaHdr,12); H(filaHdr,12,'');
    H(filaHdr+1,11,'DESAYUNO'); H(filaHdr+1,12,'ALMUERZO');
    H(filaHdr,13,'SUBTOTALES'); mrg(filaHdr,13,filaHdr,14); H(filaHdr,14,'');
    H(filaHdr+1,13,'DESAYUNO'); H(filaHdr+1,14,'ALMUERZO');
    H(filaHdr,15,'TOTAL'); mrg(filaHdr,15,filaHdr+1,15);

    // Bloques por comedor
    let r=filaHdr+2;
    modelo.comedores.forEach(c=>{
      const r0=r;
      c.filas.forEach(f=>{
        set(r,3,f.desde,F_CELr,DFMT,'d'); set(r,4,f.hasta,F_CELr,DFMT,'d');
        set(r,5,f.eDes,F_CEL,NFMT); set(r,6,f.fAlm,F_CEL,NFMT); set(r,7,f.dias,F_CEL,NFMT);
        set(r,8,f.H,F_CEL,NFMT); set(r,9,f.I,F_CEL,NFMT); set(r,10,f.J,F_CEL,NFMT);
        set(r,11,f.kDes,F_CEL,QFMT); set(r,12,f.lAlm,F_CEL,QFMT);
        set(r,13,f.M,F_CEL,QFMT); set(r,14,f.N,F_CEL,QFMT); set(r,15,f.O,F_CEL,QFMT);
        r++;
      });
      // No. y Comedor combinados sobre el bloque
      set(r0,1,c.no,F_CEL); mrg(r0,1,r-1,1);
      set(r0,2,c.nombre,F_CELl); mrg(r0,2,r-1,2);
      // SUBTOTAL
      set(r,1,'SUBTOTAL',F_BOLD); mrg(r,1,r,6);
      for(let cc=2;cc<=6;cc++) if(!ws[XLSX.utils.encode_cell({r:r-1,c:cc-1})]) set(r,cc,'',F_BOLD);
      set(r,7,c.sub.dias,F_BOLD,NFMT);
      set(r,8,c.sub.H,F_BOLD,NFMT); set(r,9,c.sub.I,F_BOLD,NFMT); set(r,10,c.sub.J,F_BOLD,NFMT);
      set(r,11,'',F_BOLD); set(r,12,'',F_BOLD);
      set(r,13,c.sub.M,F_BOLD,QFMT); set(r,14,c.sub.N,F_BOLD,QFMT); set(r,15,c.sub.O,F_BOLD,QFMT);
      r++; r++; // blanco entre bloques
    });
    // TOTAL general
    set(r,1,'TOTAL',F_BOLD); mrg(r,1,r,7);
    for(let cc=2;cc<=7;cc++) if(!ws[XLSX.utils.encode_cell({r:r-1,c:cc-1})]) set(r,cc,'',F_BOLD);
    set(r,8,modelo.total.H,F_BOLD,NFMT); set(r,9,modelo.total.I,F_BOLD,NFMT); set(r,10,modelo.total.J,F_BOLD,NFMT);
    set(r,11,'',F_BOLD); set(r,12,'',F_BOLD);
    set(r,13,modelo.total.M,F_BOLD,QFMT); set(r,14,modelo.total.N,F_BOLD,QFMT); set(r,15,modelo.total.O,F_BOLD,QFMT);

    ws['!ref']=XLSX.utils.encode_range({s:{r:0,c:0},e:{r:r-1,c:14}});
    ws['!merges']=merges;
    ws['!cols']=[{wch:5},{wch:32},{wch:12},{wch:12},{wch:11},{wch:11},{wch:9},
                 {wch:12},{wch:12},{wch:12},{wch:12},{wch:12},{wch:14},{wch:14},{wch:15}];
    const wb=XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb, ws, 'CONSOLIDADO DE CONSUMO');
    return wb;
  }

  /* ─────────────── 5) Generador PDF (pdfmake) ─────────────── */
  function q(n){ return 'Q'+ (Number(n)||0).toLocaleString('es-GT',{minimumFractionDigits:2,maximumFractionDigits:2}); }
  function nfmt(n){ return (Number(n)||0).toLocaleString('es-GT'); }

  function construirPDFDoc(modelo, pdfMakeLib){
    if(pdfMakeLib){ pdfMakeLib.fonts=pdfMakeLib.fonts||{};
      const R={normal:'Roboto-Regular.ttf',bold:'Roboto-Medium.ttf',italics:'Roboto-Italic.ttf',bolditalics:'Roboto-MediumItalic.ttf'};
      if(!pdfMakeLib.fonts.Roboto) pdfMakeLib.fonts.Roboto=R;
    }
    const m=modelo.meta||{};
    const HDR='#B4C6E7', SUBF='#EAF0FA';
    // Anchos FIJOS (no '*') para que todos los cuadros queden perfectamente alineados
    // entre sí, aunque cada comedor sea una tabla independiente. Suman ≈ el ancho útil.
    const W=[20,205,46,46,38,38,38, 46,46,50, 48,48, 66,66, 76];
    const LAYOUT={ hLineWidth:()=>0.5, vLineWidth:()=>0.5, hLineColor:()=>'#555', vLineColor:()=>'#555',
                   paddingLeft:()=>2, paddingRight:()=>2, paddingTop:()=>1.5, paddingBottom:()=>1.5 };
    const cC=(t,al)=>({text:t,style:'td',alignment:al||'center'});
    const sb=(t,extra)=>Object.assign({text:t,style:'tb',alignment:'center'},extra||{});

    // Encabezado de tabla (2 niveles) — se repite al inicio de cada cuadro de comedor
    function encabezado(){
      const cH=(t,extra)=>Object.assign({text:t,style:'th',fillColor:HDR,alignment:'center'},extra||{});
      return [
        [ cH('No.',{rowSpan:2}), cH('COMEDOR SOCIAL',{rowSpan:2}),
          cH('PERIODO DEL SERVICIO',{colSpan:2}),{},
          cH('RACIONES DIARIO',{colSpan:2}),{},
          cH('DÍAS DE\nSERVICIO',{rowSpan:2}),
          cH('RACIONES DE ALIMENTOS',{colSpan:3}),{},{},
          cH('COSTO POR RACIÓN',{colSpan:2}),{},
          cH('SUBTOTALES',{colSpan:2}),{},
          cH('TOTAL',{rowSpan:2}) ],
        [ {},{}, cH('DESDE'),cH('HASTA'), cH('DES.'),cH('ALM.'), {},
          cH('DES.'),cH('ALM.'),cH('TOTAL'), cH('DES.'),cH('ALM.'), cH('DES.'),cH('ALM.'), {} ]
      ];
    }

    // Un CUADRO por comedor: encabezado + días + subtotal, marcado como "unbreakable"
    // para que pdfmake nunca lo parta; si no cabe, lo mueve entero a la página siguiente.
    function cuadroComedor(c){
      const body=encabezado();
      c.filas.forEach((f,idx)=>{
        const row=[];
        if(idx===0){ row.push(Object.assign(cC(String(c.no)),{rowSpan:c.filas.length}));
                     row.push(Object.assign({text:c.nombre,style:'td',alignment:'left'},{rowSpan:c.filas.length})); }
        else { row.push({}); row.push({}); }
        row.push(cC(fmtDMY(f.desde)),cC(fmtDMY(f.hasta)),cC(nfmt(f.eDes)),cC(nfmt(f.fAlm)),cC(nfmt(f.dias)),
                 cC(nfmt(f.H)),cC(nfmt(f.I)),cC(nfmt(f.J)),cC(q(f.kDes)),cC(q(f.lAlm)),cC(q(f.M)),cC(q(f.N)),cC(q(f.O)));
        body.push(row);
      });
      body.push([ sb('SUBTOTAL',{colSpan:6,fillColor:SUBF}),{},{},{},{},{},
        sb(nfmt(c.sub.dias),{fillColor:SUBF}), sb(nfmt(c.sub.H),{fillColor:SUBF}), sb(nfmt(c.sub.I),{fillColor:SUBF}),
        sb(nfmt(c.sub.J),{fillColor:SUBF}), {text:'',fillColor:SUBF},{text:'',fillColor:SUBF},
        sb(q(c.sub.M),{fillColor:SUBF}), sb(q(c.sub.N),{fillColor:SUBF}), sb(q(c.sub.O),{fillColor:SUBF}) ]);
      return { table:{ headerRows:2, widths:W, body:body }, layout:LAYOUT, unbreakable:true, margin:[0,0,0,10] };
    }

    // Bloque final: TOTAL general (también unbreakable)
    const totalTable={ table:{ widths:W, body:[[
        sb('TOTAL',{colSpan:7,fillColor:HDR}),{},{},{},{},{},{},
        sb(nfmt(modelo.total.H),{fillColor:HDR}), sb(nfmt(modelo.total.I),{fillColor:HDR}), sb(nfmt(modelo.total.J),{fillColor:HDR}),
        {text:'',fillColor:HDR},{text:'',fillColor:HDR},
        sb(q(modelo.total.M),{fillColor:HDR}), sb(q(modelo.total.N),{fillColor:HDR}), sb(q(modelo.total.O),{fillColor:HDR})
    ]] }, layout:LAYOUT, unbreakable:true };

    const integ=m.integracion||[];
    const content=[
      { columns:[ {text:'Programa Comedor Social',style:'sub'},
                  {text:'Lote '+(m.lote||''),style:'sub',alignment:'right'} ] },
      { text:'CUADRO CONSOLIDADO DE CONSUMO POR '+(m.tipoContrato||'LICITACIÓN')+', DEPARTAMENTO DE '+String(m.depto||'').toUpperCase(),
        style:'titulo', margin:[0,2,0,2] },
      { text:(m.fechaLugar||''), style:'sub', alignment:'right' },
      integ[0]? { text:integ[0], style:'integ', margin:[0,6,0,6] } : {},
      integ[1]? { text:integ[1], style:'integ', margin:[0,0,0,6] } : {}
    ];
    modelo.comedores.forEach(c=>content.push(cuadroComedor(c)));
    content.push(totalTable);

    return {
      pageSize:{width:936, height:612},          // Oficio horizontal
      pageOrientation:'landscape',
      pageMargins:[24,24,24,28],
      defaultStyle:{ font:'Roboto', fontSize:7 },
      styles:{
        titulo:{ fontSize:11, bold:true, alignment:'center' },
        sub:{ fontSize:9 },
        integ:{ fontSize:9, alignment:'center' },
        th:{ fontSize:6.5, bold:true },
        td:{ fontSize:7 },
        tb:{ fontSize:7, bold:true }
      },
      content: content
    };
  }

  return { extraerComedor, agrupar, construirModelo, integracionDefault,
           construirXLSX, construirPDFDoc, fmtDMY, fechaLarga, MESES, tit };
});

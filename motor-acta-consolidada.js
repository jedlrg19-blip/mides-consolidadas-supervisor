/* ═══════════════════════════════════════════════════════════════════════════
   MOTOR COMPARTIDO — Acta Consolidada MIDES · Subdirección de Comedores
   Replica fielmente generar_pdf_consolidada_simple() de app_web.py:
     - Página Oficio Guatemala: 8.5 x 13 in (612 x 936 pt)
     - Márgenes: izq 120pt, sup 90pt, der 45pt, inf 35pt
     - Fuente Helvetica 9.5pt / interlineado 1.5, tabla 7.5pt
     - Auto-ajuste (shrink) para caber en una sola hoja, igual que KeepInFrame
     - Marca de agua "BORRADOR" diagonal cuando no_acta === 'SN'
   Usado igual por la app de Supervisores (borrador) y la app de Carlos (final),
   así ambas producen visualmente el MISMO documento.
   ═══════════════════════════════════════════════════════════════════════════ */
(function (root, factory) {
  if (typeof module === 'object' && module.exports) module.exports = factory();
  else root.MotorConsolidada = factory();
})(typeof self !== 'undefined' ? self : this, function () {
  'use strict';

  /* ─────────────── Números a letras (equivalente a num2words 'es') ─────────────── */
  const U = ['cero','uno','dos','tres','cuatro','cinco','seis','siete','ocho','nueve','diez',
    'once','doce','trece','catorce','quince','dieciséis','diecisiete','dieciocho','diecinueve','veinte',
    'veintiuno','veintidós','veintitrés','veinticuatro','veinticinco','veintiséis','veintisiete','veintiocho','veintinueve'];
  const DEC = {3:'treinta',4:'cuarenta',5:'cincuenta',6:'sesenta',7:'setenta',8:'ochenta',9:'noventa'};
  const CEN = {2:'doscientos',3:'trescientos',4:'cuatrocientos',5:'quinientos',6:'seiscientos',7:'setecientos',8:'ochocientos',9:'novecientos'};
  function dos(n){ if(n<30) return U[n]; const t=Math.floor(n/10),u=n%10; return DEC[t]+(u?' y '+U[u]:''); }
  function tres(n){
    if(n===0) return ''; if(n===100) return 'cien'; if(n<100) return dos(n);
    const c=Math.floor(n/100),r=n%100; const base=(c===1)?'ciento':CEN[c];
    return base+(r?' '+dos(r):'');
  }
  function apocopar(w){
    if(w.endsWith('veintiuno')) return w.slice(0,-9)+'veintiún';
    if(w.endsWith('uno')) return w.slice(0,-3)+'un';
    return w;
  }
  function cardinal(n){
    n=Math.floor(Math.abs(Number(n)||0));
    if(n===0) return 'cero';
    const mill=Math.floor(n/1000000), resto=n%1000000;
    const miles=Math.floor(resto/1000), cien=resto%1000;
    const out=[];
    if(mill){ out.push(mill===1?'un millón':apocopar(cardinal(mill))+' millones'); }
    if(miles){ out.push(miles===1?'mil':apocopar(tres(miles))+' mil'); }
    if(cien) out.push(tres(cien));
    return out.join(' ');
  }
  function numLetras(n){ return cardinal(n); }
  function numLetrasFem(n){
    let t=cardinal(n).split('cientos').join('cientas').split('quinientos').join('quinientas');
    if(t.endsWith('uno')) t=t.slice(0,-3)+'una';
    return t;
  }
  function numLetrasMasc(n){
    let t=cardinal(n);
    if(t==='uno') return 'un';
    if(t.endsWith('y uno')) return t.slice(0,-3)+'un';
    return t;
  }

  /* ─────────────── Formato de texto ─────────────── */
  function capitalizar(t){
    if(!t) return '';
    const exc=['de','la','el','los','las','y','en','del','a'];
    const pal=String(t).toLowerCase().split(/\s+/).filter(Boolean);
    const res=pal.map((p,i)=>(i===0||!exc.includes(p))?(p.charAt(0).toUpperCase()+p.slice(1)):p);
    return res.join(' ').split('S.a.').join('S.A.');
  }
  // CUI legal: "2549 76123 2206" (o sin espacios) -> "dos mil... espacio ... espacio ... (2549 76123 2206)"
  function formatearCuiLegal(cui){
    let c=String(cui==null?'':cui).replace(/\D/g,'');
    c=c.padStart(13,'0').slice(0,13);
    const b1=c.slice(0,4),b2=c.slice(4,9),b3=c.slice(9);
    function bloque(b){
      const sin=b.replace(/^0+/,'');
      const nc=b.length-sin.length;
      const ceros='cero '.repeat(nc);
      const resto=sin||'0';
      return (ceros+numLetras(parseInt(resto,10))).trim();
    }
    return bloque(b1)+' espacio '+bloque(b2)+' espacio '+bloque(b3)+' ('+b1+' '+b2+' '+b3+')';
  }
  // Igual que arriba pero SOLO letras, sin el "(digitos)" final (para textos que ya
  // agregan los dígitos aparte, como el acta consolidada: "CUI {letras} ({digitos})").
  function cuiLetrasSolo(cui){
    const partes=String(cui==null?'':cui).trim().split(/\s+/).filter(p=>/^\d+$/.test(p));
    if(!partes.length) return '';
    return partes.map(p=>{
      const sin=p.replace(/^0+/,''); const nc=p.length-sin.length;
      const ceros='cero '.repeat(nc); const resto=sin||'0';
      return (ceros+numLetras(parseInt(resto,10))).trim();
    }).join(' espacio ');
  }
  function dos2(n){ return String(n).padStart(2,'0'); }
  function parseHora(s){ try{ const a=String(s).split(':'); return [parseInt(a[0],10)||0, parseInt(a[1],10)||0]; }catch(e){ return [8,0]; } }
  const MESES=['enero','febrero','marzo','abril','mayo','junio','julio','agosto','septiembre','octubre','noviembre','diciembre'];

  function fechaLegal(iso, anioLetrasCache){
    // iso: 'YYYY-MM-DD' -> "cuatro (4) de mayo del año dos mil veintiséis (2026)"
    try{
      const [y,m,d]=String(iso).split('-').map(x=>parseInt(x,10));
      const anioL = anioLetrasCache && anioLetrasCache[y] ? anioLetrasCache[y] : numLetras(y);
      return numLetras(d)+' ('+d+') de '+MESES[m-1]+' del año '+anioL+' ('+y+')';
    }catch(e){ return String(iso); }
  }

  /* ─────────────── Constantes institucionales fijas ─────────────── */
  const SUBDIRECTORA_NOMBRE = 'AZUCENA MAGDELY CORADO ORTEGA';
  const SUBDIRECTORA_CUI    = '2549 76123 2206';
  const SUBDIRECCION_DIR    = 'quinta avenida ocho guion setenta y ocho zona nueve ' +
    '(5 Av 8-78 zona 9) de la ciudad de Guatemala, Edificio Plaza Lauderdale, sexto nivel';

  /* ─────────────── Página / márgenes (idénticos a reportlab) ─────────────── */
  const PAGE_W = 612, PAGE_H = 936;                 // Oficio Guatemala: 8.5 x 13 in
  const LEFT_M = 120, RIGHT_M = 45, TOP_M = 90, BOT_M = 35;
  const FS_MAX = 9.5, FS_MIN = 6.0, FS_STEP = 0.25;  // igual criterio de auto-ajuste que el acta diaria

  /**
   * Arma el docDefinition de pdfmake para una Acta Consolidada.
   * @param {object} d - datos del acta (ver README del motor / abajo)
   * @param {number} fs - tamaño de fuente del cuerpo a usar en este intento
   * @param {number} lh - interlineado a usar en este intento
   */
  function construirDoc(d, fs, lh, sfx){
    sfx = sfx || '';   // sufijo para nombres de estilo (permite combinar varias actas en un solo PDF)
    const fst = Math.max(5.5, Math.round((fs*0.79)*100)/100);   // proporción 7.5/9.5 del original
    const ldt = Math.round(fst*1.5*100)/100;
    const ld  = Math.round(fs*lh*100)/100;

    const ep  = { fontSize:fs,  lineHeight:lh, alignment:'justify' };
    const et  = { fontSize:fst, alignment:'center' };
    const etb = { fontSize:fst, alignment:'center', bold:true };

    const comedor     = String(d.comedor||'').toUpperCase();
    const modalidad   = (String(d.modalidad||'FIJO').toUpperCase().trim()) || 'FIJO';
    const comedorFull = 'COMEDOR SOCIAL '+modalidad+' '+comedor;
    const supNom  = String(d.supervisor||'').toUpperCase();
    const supCui  = String(d.supervisor_cui||'').trim();
    const empresa = String(d.empresa||'Banquetes de Guatemala, S.A.');
    const incump  = String(d.incumplimiento||'NO').toUpperCase();

    const esPrev = (String(d.no_acta||'SN')==='SN');
    const fechaEmision = d.fecha_emision || null;  // 'YYYY-MM-DD'
    let anioN, mesN, diaN;
    if(fechaEmision){
      const [y,m,dd]=fechaEmision.split('-').map(x=>parseInt(x,10));
      anioN=y; mesN=MESES[m-1]; diaN=dd;
    } else {
      const now=new Date();
      anioN=now.getFullYear(); mesN=MESES[now.getMonth()]; diaN=now.getDate();
    }
    const diaL=numLetras(diaN), anioL=numLetras(anioN);
    const [hAc,mAc]=parseHora(d.hora_acta||'08:00');
    const hLetras=numLetrasFem(hAc), mLetras=numLetrasMasc(mAc);
    const lblM = mAc===1?'minuto':'minutos';

    const noLetras = esPrev ? 'BORRADOR - SIN NUMERO ASIGNADO' : numLetras(parseInt(d.no_acta,10)).toUpperCase();
    const numTxt   = esPrev ? 'SN' : String(d.no_acta);

    const subCuiLetras = cuiLetrasSolo(SUBDIRECTORA_CUI);
    const supCuiLetras = supCui ? cuiLetrasSolo(supCui) : '';

    const filas = Array.isArray(d.filas) ? d.filas : [];
    let totDes = d.tot_des, totAlm = d.tot_alm, totTotal = d.tot_total;
    if(totDes==null || totAlm==null){
      totDes = filas.reduce((a,f)=>a+(parseInt(f.des_tot,10)||0),0);
      totAlm = filas.reduce((a,f)=>a+(parseInt(f.alm_tot,10)||0),0);
    }
    if(totTotal==null) totTotal = (totDes||0)+(totAlm||0);
    const totDesL   = totDes   ? numLetrasFem(totDes)   : 'cero';
    const totAlmL   = totAlm   ? numLetrasFem(totAlm)   : 'cero';
    const totTotalL = totTotal ? numLetrasFem(totTotal) : 'cero';
    const corrTxt   = d.correlativo ? ('correlativo '+d.correlativo) : 'el correlativo correspondiente';
    const incumpTxt = incump.replace('Í','I').includes('SI') ? 'SI hubo incumplimiento' : 'NO hubo incumplimiento';

    let rangoTxt = 'del periodo correspondiente';
    if(filas.length){
      const [y1,m1,d1]=filas[0].fecha.split('-').map(Number);
      const inicioCorto = numLetras(d1)+' ('+d1+') de '+MESES[m1-1];
      rangoTxt = 'del '+inicioCorto+' al '+fechaLegal(filas[filas.length-1].fecha);
    }

    // ── Párrafo de comparecientes / intro (runs con negritas, como el acta diaria) ──
    const R=[]; const B=(t)=>R.push({text:t,bold:true}); const N=(t)=>R.push({text:t});
    B('ACTA NÚMERO '+noLetras+' GUION '+anioL.toUpperCase()+' ('+numTxt+'-'+anioN+').');
    N(' En el municipio y departamento de Guatemala, el día '+diaL+' ('+diaN+') de '+mesN+
      ' del año '+anioL+' ('+anioN+'), siendo las '+hLetras+' con '+mLetras+' '+lblM+' ('+
      dos2(hAc)+':'+dos2(mAc)+'), constituidos en las instalaciones de la Subdirección de '+
      'Comedores, ubicada en la '+SUBDIRECCION_DIR+', nos encontramos: A) ');
    B(SUBDIRECTORA_NOMBRE);
    N(', me identifico con el Documento Personal de Identificación (DPI), con Código Único de '+
      'Identificación (CUI) '+subCuiLetras+' ('+SUBDIRECTORA_CUI+') extendido por el Registro '+
      'Nacional de las Personas de la República de Guatemala, actúo en mi calidad de '+
      'Subdirectora de Comedores de la Dirección de Prevención Social del Viceministerio de '+
      'Prevención Social del Ministerio de Desarrollo Social, a cargo del Programa Social '+
      '\u201cComedor Social\u201d y; B) ');
    B(supNom);
    if(supCui){
      N(', me identifico con el Documento Personal de Identificación (DPI) con Código Único de '+
        'Identificación (CUI) '+supCuiLetras+' ('+supCui+') extendido por el Registro Nacional de '+
        'las Personas de la República de Guatemala,');
    }
    N(' actúo en mi calidad de Supervisor de Comedores, asignado al '+comedorFull+
      ' del Ministerio de Desarrollo Social. Estamos constituidos para suscribir la presente '+
      'Acta Consolidada, con el propósito de hacer constar lo siguiente: ');
    B('PRIMERO:');
    N(' El Manual Operativo del Programa Social \u201cComedor Social\u201d, establece el procedimiento '+
      'denominado CONSOLIDADO DE CONSUMO, que consiste en consolidar el total de raciones '+
      'servidas durante el periodo establecido, quedando conformado y confirmando la siguiente '+
      'integración:');

    // ── Tabla (12 columnas, spans, igual que reportlab) ──
    const rawCols=[0.045,0.13,0.10,0.09,0.075,0.075,0.065,0.09,0.075,0.075,0.065,0.10];
    const W = PAGE_W-LEFT_M-RIGHT_M;
    const sumCols=rawCols.reduce((a,b)=>a+b,0);
    const widths = rawCols.map(c=>Math.round((W*c/sumCols)*100)/100);

    const P=(t,bold)=>({text:String(t), style: (bold?'tabHdr':'tabCell')+sfx});
    const body=[];
    body.push([P('No.',1),P('Fecha',1),P('No. de\nActa',1),
      P('Desayunos',1),P('',1),P('',1),P('',1),
      P('Almuerzos',1),P('',1),P('',1),P('',1),
      P('Total\nRaciones',1)]);
    body.push([P('',1),P('',1),P('',1),
      P('No.',1),P('Inicio',1),P('Final',1),P('Total',1),
      P('No.',1),P('Inicio',1),P('Final',1),P('Total',1),P('',1)]);
    if(filas.length){
      filas.forEach(f=>{
        body.push([P(f.no),P(f.fecha_fmt||f.fecha),P(f.acta),
          P(f.des_envio),P(f.des_ini),P(f.des_fin),P(f.des_tot),
          P(f.alm_envio),P(f.alm_ini),P(f.alm_fin),P(f.alm_tot),P(f.total)]);
      });
    } else {
      body.push([{text:'Sin datos - adjunte el Excel (hoja: CONTROL DE SUPERVISOR)', style:'tabCell', colSpan:12, alignment:'center'},{},{},{},{},{},{},{},{},{},{},{}]);
    }
    body.push([P(''),P('Totales',1),P(''),P(''),P(''),P(''),P(String(totDes),1),
      P(''),P(''),P(''),P(String(totAlm),1),P(String(totTotal),1)]);

    const nRows=body.length;
    const tabla={
      table:{ headerRows:2, widths:widths, body:body },
      layout:{
        hLineWidth:()=>0.4, vLineWidth:()=>0.4, hLineColor:()=>'#000', vLineColor:()=>'#000',
        paddingLeft:()=>2, paddingRight:()=>2, paddingTop:()=>2, paddingBottom:()=>2,
        fillColor: function(rowIndex){
          if(rowIndex<2) return '#e6e6e6';               // encabezados
          if(rowIndex===nRows-1) return '#f2f2f2';        // totales
          return (rowIndex%2===0) ? '#f7f7f7' : null;      // filas alternadas
        }
      }
    };

    // ── SEGUNDO / TERCERO ──
    const p2=[
      {text:'SEGUNDO:  ',bold:true},
      {text:'La información que se integra, proviene de: a) la recepción de las raciones a través '+
        'de los envíos realizados al Comedor Social por parte de la entidad proveedora '},
      {text:empresa,bold:true},
      {text:'. b) la generación de los tickets electrónicos con los que hace constar el registro de '+
        'atención a usuarios en el comedor. c) la suscripción de las respectivas actas en la que '+
        'constan el cierre administrativo diario, suscritas por el Encargado del Comedor Social; y '+
        'd) el Cuadro de Consolidación de las raciones servidas, identificado con '+corrTxt+', '+
        'elaborado por el Encargado del Comedor y revisado por el infrascrito Supervisor, asignado '+
        'al '+comedorFull+'; registrando la cantidad de '+totDesL+' ('+totDes+') RACIONES DE '+
        'DESAYUNOS y '+totAlmL+' ('+totAlm+') RACIONES DE ALMUERZOS, PARA UN TOTAL DE '+totTotalL+
        ' ('+totTotal+') RACIONES SERVIDAS, durante el periodo que comprende '+rangoTxt+' a los '+
        'usuarios del Comedor Social descrito. Asimismo, se hace constar que, durante el periodo '+
        'descrito en el presente punto de acta, '+incumpTxt+' por parte de la empresa proveedora. '},
      {text:'TERCERO:',bold:true},
      {text:' Se finaliza la presente, en el mismo lugar y fecha veinte (20) minutos después de su '+
        'inicio, la que previa lectura hecha por los comparecientes y bien enterados de su '+
        'contenido, objeto, validez y efectos legales, la aceptan, ratifican y firman, la que '+
        'consta en una (1) hoja de papel bond oficio, impresa únicamente en su anverso, '+
        'debidamente autorizada por la Contraloría General de Cuentas.'}
    ];

    const content=[
      { text:R, ...ep },
      { text:'', margin:[0,5,0,0] },
      tabla,
      { text:p2, margin:[0,5,0,0], ...ep },
      { text:'', margin:[0,90,0,0] }   // espacio libre final para sellos, igual que reportlab
    ];

    const docDefinition={
      pageSize:{ width:PAGE_W, height:PAGE_H },
      pageMargins:[LEFT_M, TOP_M, RIGHT_M, BOT_M],
      defaultStyle:{ font:'ArialNarrow' },
      styles:{
        ['tabHdr'+sfx]:{ fontSize:fst, lineHeight:1.15, alignment:'center', bold:true },
        ['tabCell'+sfx]:{ fontSize:fst, lineHeight:1.15, alignment:'center' }
      },
      content: content
    };
    if(esPrev){
      docDefinition.background = function(currentPage, pageSize){
        return {
          svg: '<svg width="'+pageSize.width+'" height="'+pageSize.height+'" xmlns="http://www.w3.org/2000/svg">'+
               '<text x="'+(pageSize.width/2)+'" y="'+(pageSize.height/2)+'" '+
               'transform="rotate(-42 '+(pageSize.width/2)+' '+(pageSize.height/2)+')" '+
               'font-family="Helvetica, Arial, sans-serif" font-weight="bold" font-size="80" '+
               'fill="#d9d9d9" fill-opacity="0.5" text-anchor="middle">BORRADOR</text></svg>',
          width: pageSize.width, height: pageSize.height
        };
      };
    }
    return docDefinition;
  }

  /* Cuenta páginas renderizando sin comprimir (igual criterio que el acta diaria) */
  function contarPaginas(pdfMakeLib, doc){
    return new Promise(function(resolve){
      doc.compress=false;
      pdfMakeLib.createPdf(doc).getBuffer(function(buf){
        let pags=1;
        try{
          const s=new TextDecoder('latin1').decode(buf);
          const m=s.match(/\/Type\s*\/Page(?![sA-Za-z])/g);
          if(m && m.length) pags=m.length;
        }catch(e){}
        resolve(pags);
      });
    });
  }

  /**
   * Genera el docDefinition final ya ajustado para caber en una sola hoja Oficio.
   * @param {object} pdfMakeLib - la instancia global de pdfMake (window.pdfMake en el navegador)
   * @param {object} d - datos del acta consolidada
   * @returns {Promise<{doc:object, fs:number}>}
   */
  async function autoAjustar(pdfMakeLib, d, sfx){
    sfx = sfx || '';
    let lh=1.5;
    for(let fs=FS_MAX; fs>=FS_MIN-0.001; fs-=FS_STEP){
      const doc=construirDoc(d, fs, lh, sfx);
      const pags=await contarPaginas(pdfMakeLib, doc);
      if(pags<=1) return { doc: construirDoc(d, fs, lh, sfx), fs: Math.round(fs*100)/100 };
      lh=Math.max(1.05, lh-0.03);
    }
    return { doc: construirDoc(d, FS_MIN, 1.05, sfx), fs: FS_MIN };
  }

  /**
   * Combina varias actas consolidadas (borradores) en un solo docDefinition,
   * una acta por página. Cada acta se auto-ajusta a su propio tamaño de fuente
   * y sus estilos de tabla se aíslan con un sufijo para no colisionar.
   * @param {object} pdfMakeLib - instancia global de pdfMake
   * @param {Array<object>} items - lista de datos de actas
   * @returns {Promise<{doc:object}>}
   */
  async function combinar(pdfMakeLib, items){
    const lista = Array.isArray(items) ? items : [];
    const styles = {};
    let content = [];
    let base = null;
    for(let i=0;i<lista.length;i++){
      const { doc } = await autoAjustar(pdfMakeLib, lista[i], '_'+i);
      if(!base) base = doc;
      Object.assign(styles, doc.styles);
      const partes = doc.content.slice();
      if(content.length && partes.length){
        partes[0] = Object.assign({}, partes[0], { pageBreak:'before' });
      }
      content = content.concat(partes);
    }
    if(!base) base = construirDoc({}, FS_MAX, 1.5);
    const docDefinition = {
      pageSize: base.pageSize,
      pageMargins: base.pageMargins,
      defaultStyle: base.defaultStyle,
      styles: styles,
      content: content
    };
    if(base.background) docDefinition.background = base.background;
    return { doc: docDefinition };
  }

  return {
    PAGE_W, PAGE_H, LEFT_M, RIGHT_M, TOP_M, BOT_M,
    numLetras, numLetrasFem, numLetrasMasc, capitalizar, formatearCuiLegal,
    construirDoc, autoAjustar, combinar,
    SUBDIRECTORA_NOMBRE, SUBDIRECTORA_CUI, SUBDIRECCION_DIR
  };
});

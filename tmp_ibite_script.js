
/* =========================================================
   CONFIGURACIÓN — cambia esto si adaptas la plantilla a otro colegio
   ========================================================= */
const CONFIG = {
  EXPECTED_FILENAME: "IBITE Manizales (respuestas).xlsx",
  SHEET_NAME: "Puntuaciones",
  STORAGE_KEY: "ibite_dashboard_cache::Manizales_v1"
};
document.getElementById('expectedFilename').textContent = CONFIG.EXPECTED_FILENAME;

/* Column index map for the 'Puntuaciones' sheet (0-based, header row = row 0) */
const COL = {
  FILA:0, TS:1, RAMA:2, NOMBRE:3, TEL:4, EDAD:5, SEXO:6, INST:7, GRADO:8, COMUNA:9,
  CONVIV_HERIDO:10, SALIDA_VIVIENDA:11, DANOS_VIVIENDA:12, DIFICULTADES_SERV:13, CAMBIO_RESID:14,
  I1:15,I2:16,I3:17,I4:18,I5:19,I6:20,I7:21,I8:22,I9:23,I10:24,I11:25,I12:26,I13:27,I14:28,I15:29,I16:30,I17:31,
  POSTRAUM_SCORE:32, POSTRAUM_NIVEL:33,
  GENFUNC_SCORE:34, DETERIORO_FLAG:35,
  COMBINADO_SCORE:36, COMBINADO_NIVEL:37,
  ALERTA_SEGURIDAD:38,
  REINTEGRA_SCORE:39, REINTEGRA_NIVEL:40, DESESCOLAR_FLAG:41,
  RECURSOS_PERS_SCORE:42, RECURSOS_PERS_NIVEL:43,
  RECURSOS_FAM_SCORE:44, RECURSOS_FAM_NIVEL:45,
  RIESGO_MATERIAL:46, NECESIDADES_BASICAS:47,
  CUESTIONARIO_COMPLETO:48,
  PERFIL:49, INDICADOR_GENERAL:50, ACCION_RECOMENDADA:51, PLAZO:52, RUTA:53, ACCIONES_ESPECIFICAS:54
};

const INDICATOR_LABELS = ["Intrusión","Pesadillas/sueño","Evitación","Hiperalerta","Malestar general","Funcionamiento",
"Ideación (seguridad)","Autolesión (seguridad)","Miedo/rechazo escolar","Concentración","Ausentismo",
"Autorregulación","Disfrute","Búsqueda de ayuda","Contención familiar","Sobrecarga parental (inv.)","Red de apoyo"];

const URGENCY_META = {
 0:{label:'Prioridad máxima — actuar hoy mismo', short:'Hoy mismo', color:'#dc2626', bg:'#fee2e2'},
 1:{label:'Prioridad alta — 48–72 horas', short:'48–72 h', color:'#ea580c', bg:'#ffedd5'},
 2:{label:'Prioridad media — dentro de la semana', short:'Esta semana', color:'#d97706', bg:'#fef3c7'},
 3:{label:'Seguimiento habitual', short:'Habitual', color:'#16a34a', bg:'#dcfce7'},
 4:{label:'Sin clasificar — revisar manualmente', short:'Sin clasificar', color:'#6b7280', bg:'#f1f5f9'}
};

/* =========================================================
   ESTADO
   ========================================================= */
const state = {
  cases: [],
  meta: null,        // {savedAt, source: 'file'|'cache'|'demo'|'fetch'}
  filters: {institucion:'', grado:'', comuna:'', rama:'', desde:'', hasta:''},
  tab2: {urgency:null, search:'', onlySeguridad:false},
  charts: {}
};

/* =========================================================
   UTILIDADES
   ========================================================= */
function normalize(s){
  return (s===null||s===undefined?'':String(s)).normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase().trim();
}
function isBlank(v){ return v===null || v===undefined || String(v).trim()===''; }
function isFlagYes(v){ const s = normalize(v); return s.startsWith('si'); }
function toNum(v){ const n = typeof v === 'number' ? v : parseFloat(v); return isNaN(n) ? null : n; }
function toDate(v){
  if (v instanceof Date) return v;
  if (v===null || v===undefined || v==='') return null;
  const d = new Date(v);
  return isNaN(d.getTime()) ? null : d;
}
function fmtDate(d){
  if(!d) return '—';
  return d.toLocaleDateString('es-CO',{day:'2-digit',month:'short',year:'numeric'}) + ' ' + d.toLocaleTimeString('es-CO',{hour:'2-digit',minute:'2-digit'});
}
function groupCount(arr, keyFn){
  const map = new Map();
  arr.forEach(item=>{
    const k = keyFn(item);
    const key = (k===null||k===undefined||k==='') ? '(Sin dato)' : k;
    map.set(key, (map.get(key)||0)+1);
  });
  return map;
}
function gradeOrder(g){
  const s = normalize(g);
  if(!s) return 999;
  if(s.includes('transicion')||s.includes('preescolar')||s.includes('jardin')||s.includes('prejardin')) return -1;
  const m = s.match(/(\d+)/);
  if(m) return parseInt(m[1],10);
  return 998;
}
function csvEscape(v){
  v = v===null||v===undefined ? '' : String(v);
  if(/[",\n;]/.test(v)) v = '"' + v.replace(/"/g,'""') + '"';
  return v;
}
function dateStamp(){
  const d = new Date();
  return d.toISOString().slice(0,10);
}
function escapeHtml(value){
  return String(value ?? '').replace(/[&<>"']/g, ch => ({
    '&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'
  }[ch]));
}

/* =========================================================
   CLASIFICACIÓN DE URGENCIA (alineada con la hoja 'Acciones' del Excel)
   ========================================================= */
function classifyPlazo(plazo){
  const p = normalize(plazo);
  if(!p) return null;
  if(p.includes('mismo dia')) return 0;
  if(p.includes('48') || p.includes('72') || p.includes('hora')) return 1;
  if(p.includes('prioritario')) return 1;
  if(p.includes('semana')) return 2;
  if(p.includes('segun valoracion')) return 2;
  if(p.includes('segun senal')) return 2;
  if(p.includes('habitual')) return 3;
  return null;
}
function classifyIndicador(indicador){
  const s = normalize(indicador);
  if(!s) return null;
  if(s.includes('critico')) return 0;
  if(s.includes('alto')) return 1;
  if(s.includes('moderado')) return 2;
  if(s.includes('bajo')) return 3;
  return null;
}
function alertaClass(v){
  const s = normalize(v);
  if(!s) return 'ninguna';
  if(s.includes('no aclarada') || s.includes('aclarar')) return 'no_aclarada';
  if(s.startsWith('sin') || s.startsWith('no')) return 'ninguna';
  if(s.includes('alerta')) return 'activa';
  if(s.startsWith('si')) return 'activa';
  return 'ninguna';
}
function computeUrgency(c){
  let rank = classifyPlazo(c.plazo);
  if(rank===null) rank = classifyIndicador(c.indicadorGeneral);
  if(rank===null) rank = 4;
  const ac = alertaClass(c.alertaSeguridad);
  if(ac==='activa' || ac==='no_aclarada') rank = Math.min(rank,0);
  if(isFlagYes(c.riesgoMaterial)) rank = Math.min(rank,1);
  if(isFlagYes(c.desescolarFlag)) rank = Math.min(rank,1);
  if(isFlagYes(c.deterioroFlag)) rank = Math.min(rank,2);
  return rank;
}


/* =========================================================
   IMPORTADOR ROBUSTO — DETECCIÓN Y MAPEO DE COLUMNAS
   Todo permanece en este único HTML.
   ========================================================= */

const IBITE_FIELDS = {
  FILA: {label:'Fila / ID', required:false, aliases:['fila','id','numero','número','registro']},
  TS: {label:'Fecha / marca temporal', required:true, aliases:['marca temporal','timestamp','fecha','fecha de registro','fecha y hora','fecha/hora','hora']},
  RAMA: {label:'Rama / sede / programa', required:false, aliases:['rama','sede','programa']},
  NOMBRE: {label:'Nombre del estudiante', required:true, aliases:['nombre','nombre completo','nombre del estudiante','nombre estudiante','estudiante','nombres']},
  TEL: {label:'Teléfono', required:false, aliases:['telefono','teléfono','celular','telefono celular','teléfono celular','movil','móvil']},
  EDAD: {label:'Edad', required:false, aliases:['edad','edad del estudiante']},
  SEXO: {label:'Sexo / género', required:false, aliases:['sexo','genero','género']},
  INST: {label:'Institución educativa', required:true, aliases:['institucion','institución','institucion educativa','institución educativa','colegio','escuela','nombre institucion','nombre institución']},
  GRADO: {label:'Grado / curso', required:false, aliases:['grado','curso','nivel','grado escolar','curso escolar']},
  COMUNA: {label:'Comuna / localidad', required:false, aliases:['comuna','localidad','barrio','municipio']},
  CONVIV_HERIDO: {label:'Conviviente herido', required:false, aliases:['conviviente herido','alguien herido','familiar herido']},
  SALIDA_VIVIENDA: {label:'Salida de vivienda', required:false, aliases:['salida vivienda','salida de vivienda','tuvo que salir de vivienda']},
  DANOS_VIVIENDA: {label:'Daños en vivienda', required:false, aliases:['danos vivienda','daños vivienda','daños en vivienda']},
  DIFICULTADES_SERV: {label:'Dificultades de servicios', required:false, aliases:['dificultades servicios','dificultades de servicios','servicios']},
  CAMBIO_RESID: {label:'Cambio de residencia', required:false, aliases:['cambio residencia','cambio de residencia','cambio de vivienda']},
  POSTRAUM_SCORE: {label:'Puntaje postraumático', required:false, aliases:['postraum score','puntaje postraumatico','puntaje postraumático','postraumatico score']},
  POSTRAUM_NIVEL: {label:'Nivel postraumático', required:false, aliases:['postraum nivel','nivel postraumatico','nivel postraumático']},
  GENFUNC_SCORE: {label:'Puntaje de funcionamiento general', required:false, aliases:['genfunc score','funcionamiento general score','puntaje funcionamiento general']},
  DETERIORO_FLAG: {label:'Deterioro funcional', required:false, aliases:['deterioro flag','deterioro funcional','deterioro']},
  COMBINADO_SCORE: {label:'Puntaje combinado', required:false, aliases:['combinado score','puntaje combinado']},
  COMBINADO_NIVEL: {label:'Nivel combinado', required:false, aliases:['combinado nivel','nivel combinado']},
  ALERTA_SEGURIDAD: {label:'Alerta de seguridad', required:false, aliases:['alerta seguridad','alerta de seguridad','seguridad','alerta']},
  REINTEGRA_SCORE: {label:'Puntaje de reintegración', required:false, aliases:['reintegra score','puntaje reintegra','puntaje reintegracion','puntaje reintegración']},
  REINTEGRA_NIVEL: {label:'Nivel de reintegración', required:false, aliases:['reintegra nivel','nivel reintegra','nivel reintegracion','nivel reintegración']},
  DESESCOLAR_FLAG: {label:'Desescolarización', required:false, aliases:['desescolar flag','desescolarizacion','desescolarización','desescolarizado']},
  RECURSOS_PERS_SCORE: {label:'Puntaje recursos personales', required:false, aliases:['recursos pers score','recursos personales score','puntaje recursos personales']},
  RECURSOS_PERS_NIVEL: {label:'Nivel recursos personales', required:false, aliases:['recursos pers nivel','recursos personales nivel','nivel recursos personales']},
  RECURSOS_FAM_SCORE: {label:'Puntaje recursos familiares', required:false, aliases:['recursos fam score','recursos familiares score','puntaje recursos familiares']},
  RECURSOS_FAM_NIVEL: {label:'Nivel recursos familiares', required:false, aliases:['recursos fam nivel','recursos familiares nivel','nivel recursos familiares']},
  RIESGO_MATERIAL: {label:'Riesgo material', required:false, aliases:['riesgo material','riesgo material grave','riesgo material alto']},
  NECESIDADES_BASICAS: {label:'Necesidades básicas', required:false, aliases:['necesidades basicas','necesidades básicas']},
  CUESTIONARIO_COMPLETO: {label:'Cuestionario completo', required:false, aliases:['cuestionario completo','completo','estado cuestionario']},
  PERFIL: {label:'Perfil', required:false, aliases:['perfil']},
  INDICADOR_GENERAL: {label:'Indicador general', required:false, aliases:['indicador general','indicador de riesgo general','riesgo general','indicador']},
  ACCION_RECOMENDADA: {label:'Acción recomendada', required:false, aliases:['accion recomendada','acción recomendada','recomendacion','recomendación']},
  PLAZO: {label:'Plazo de acción', required:false, aliases:['plazo','plazo de accion','plazo de acción','tiempo de accion','tiempo de acción']},
  RUTA: {label:'Ruta / responsable', required:false, aliases:['ruta','ruta responsable','ruta/responsable','responsable']},
  ACCIONES_ESPECIFICAS: {label:'Acciones específicas', required:false, aliases:['acciones especificas','acciones específicas','acciones']},
  INDICATORS: {label:'Indicadores I1–I17', required:false, aliases:['i1','i2','i3','i4','i5','i6','i7','i8','i9','i10','i11','i12','i13','i14','i15','i16','i17']}
};

function headerKey(s){
  return normalize(s).replace(/[_\-\/\\.,:;()[\]{}]+/g,' ').replace(/\s+/g,' ').trim();
}
function headerCompact(s){
  return headerKey(s).replace(/\s+/g,'');
}
function similarityScore(a,b){
  a=headerCompact(a); b=headerCompact(b);
  if(!a || !b) return 0;
  if(a===b) return 1;
  if(a.includes(b) || b.includes(a)) return 0.88;
  const aa=new Set(a.split('')), bb=new Set(b.split(''));
  let inter=0; aa.forEach(x=>{if(bb.has(x))inter++});
  return inter/Math.max(aa.size,bb.size);
}
function expectedAliases(field){
  const f=IBITE_FIELDS[field];
  return f ? [f.label,...f.aliases] : [field];
}
function autoMapHeaders(headers){
  const map={};
  const used=new Set();
  const confidence={};
  Object.keys(IBITE_FIELDS).forEach(field=>{
    if(field==='INDICATORS') return;
    let best={idx:-1,score:0};
    headers.forEach((h,idx)=>{
      if(used.has(idx)) return;
      const candidates=expectedAliases(field);
      let score=0;
      candidates.forEach(a=>{
        const s=similarityScore(h,a);
        if(s>score) score=s;
      });
      // Calculated fields in this workbook often have highly distinctive names.
      if(score>best.score) best={idx,score};
    });
    if(best.idx>=0 && best.score>=0.86){
      map[field]=best.idx;
      confidence[field]=best.score;
      used.add(best.idx);
    }
  });

  // Indicators are a special group: detect I1..I17 independently.
  for(let k=1;k<=17;k++){
    const field='I'+k;
    let best={idx:-1,score:0};
    headers.forEach((h,idx)=>{
      if(used.has(idx)) return;
      const n=headerCompact(h);
      const score=(n===field || n===('indicador'+k) || n===('item'+k) || n===('i'+k)) ? 1 : 0;
      if(score>best.score) best={idx,score};
    });
    if(best.idx>=0){
      map[field]=best.idx; confidence[field]=best.score; used.add(best.idx);
    }
  }
  return {map,confidence};
}

let importContext=null;

function ensureMapperUI(){
  if(document.getElementById('columnMapperModal')) return;
  const style=document.createElement('style');
  style.textContent=`
    #columnMapperModal{position:fixed;inset:0;background:rgba(15,23,42,.72);z-index:9999;display:none;align-items:center;justify-content:center;padding:20px}
    #columnMapperModal.show{display:flex}
    .mapper-card{background:#fff;width:min(1120px,96vw);max-height:92vh;overflow:auto;border-radius:20px;box-shadow:0 25px 70px rgba(0,0,0,.25)}
    .mapper-head{padding:22px 26px;border-bottom:1px solid #e5e7eb;position:sticky;top:0;background:#fff;z-index:2}
    .mapper-head h2{margin:0 0 5px;font-size:22px;color:#0f172a}
    .mapper-head p{margin:0;color:#64748b;font-size:14px}
    .mapper-summary{display:flex;gap:10px;flex-wrap:wrap;padding:16px 26px;background:#f8fafc;border-bottom:1px solid #e5e7eb}
    .mapper-pill{padding:7px 11px;border-radius:999px;font-size:12px;font-weight:700}
    .pill-ok{background:#dcfce7;color:#166534}.pill-warn{background:#fef3c7;color:#92400e}.pill-bad{background:#fee2e2;color:#991b1b}
    .mapper-table{width:100%;border-collapse:collapse}
    .mapper-table th{background:#f8fafc;text-align:left;font-size:12px;text-transform:uppercase;letter-spacing:.04em;color:#64748b;padding:11px 18px;border-bottom:1px solid #e5e7eb;position:sticky;top:77px}
    .mapper-table td{padding:12px 18px;border-bottom:1px solid #eef2f7;vertical-align:middle}
    .mapper-select{width:100%;padding:9px 10px;border:1px solid #cbd5e1;border-radius:9px;background:#fff}
    .mapper-sample{font-size:12px;color:#64748b;max-width:330px}
    .mapper-status{font-size:12px;font-weight:700;white-space:nowrap}
    .status-auto{color:#15803d}.status-manual{color:#b45309}.status-missing{color:#dc2626}.status-optional{color:#64748b}
    .mapper-foot{padding:16px 26px;display:flex;justify-content:space-between;gap:12px;position:sticky;bottom:0;background:#fff;border-top:1px solid #e5e7eb}
    .mapper-btn{border:0;border-radius:10px;padding:11px 17px;font-weight:700;cursor:pointer}
    .mapper-btn.primary{background:#2563eb;color:#fff}.mapper-btn.secondary{background:#e2e8f0;color:#0f172a}
    .mapper-btn:disabled{opacity:.45;cursor:not-allowed}
    .mapper-help{padding:12px 26px;color:#475569;font-size:13px;background:#eff6ff}
  `;
  document.head.appendChild(style);
  document.body.insertAdjacentHTML('beforeend',`
    <div id="columnMapperModal">
      <div class="mapper-card">
        <div class="mapper-head">
          <h2>Revisar columnas del Excel</h2>
          <p>El sistema ha comparado los encabezados del archivo con los campos que necesita el tablero. Confirma o corrige los campos marcados.</p>
        </div>
        <div id="mapperSummary" class="mapper-summary"></div>
        <div class="mapper-help">Puedes dejar sin asignar los campos opcionales. Los campos obligatorios deben estar correctamente asignados antes de importar.</div>
        <table class="mapper-table">
          <thead><tr><th>Campo del tablero</th><th>Columna del Excel</th><th>Ejemplos encontrados</th><th>Estado</th></tr></thead>
          <tbody id="mapperBody"></tbody>
        </table>
        <div class="mapper-foot">
          <button id="mapperCancel" class="mapper-btn secondary">Cancelar importación</button>
          <button id="mapperApply" class="mapper-btn primary">Usar este mapeo e importar</button>
        </div>
      </div>
    </div>`);
  document.getElementById('mapperCancel').onclick=()=>document.getElementById('columnMapperModal').classList.remove('show');
  document.getElementById('mapperApply').onclick=applyColumnMapping;
}

function showColumnMapper(ctx){
  ensureMapperUI();
  importContext=ctx;
  const {headers, rows, map, confidence}=ctx;
  const body=document.getElementById('mapperBody');
  const used=new Set(Object.values(map));
  let html='';
  let missingRequired=0, auto=0, manual=0, optionalMissing=0;

  const fields=Object.keys(IBITE_FIELDS).filter(f=>f!=='INDICATORS');
  fields.push(...Array.from({length:17},(_,i)=>'I'+(i+1)));

  fields.forEach(field=>{
    const meta=IBITE_FIELDS[field] || {label:'Indicador '+field.replace('I',''),required:false};
    const idx=map[field];
    const isAuto=idx!==undefined && confidence[field]>=.86;
    if(isAuto) auto++;
    else if(idx!==undefined) manual++;
    else if(meta.required) missingRequired++;
    else optionalMissing++;

    const options=['<option value="">— No asignar —</option>'].concat(
      headers.map((h,i)=>{
        const selected=idx===i?' selected':'';
        const disabled=(used.has(i)&&idx!==i)?' disabled':'';
        return `<option value="${i}"${selected}${disabled}>${escapeHtml(h || ('Columna '+(i+1)))}</option>`;
      })
    ).join('');

    const sample=idx!==undefined ? rows.slice(1,6).map(r=>r[idx]).filter(v=>!isBlank(v)).slice(0,3).map(v=>escapeHtml(String(v))).join(' · ') : 'Sin datos de ejemplo';
    const status=idx===undefined
      ? (meta.required ? '<span class="mapper-status status-missing">❌ Obligatorio</span>':'<span class="mapper-status status-optional">Opcional</span>')
      : (isAuto ? '<span class="mapper-status status-auto">✓ Automático</span>':'<span class="mapper-status status-manual">⚠ Revisar</span>');

    html+=`<tr data-field="${field}">
      <td><strong>${escapeHtml(meta.label || field)}</strong>${meta.required?' <small style="color:#dc2626">*</small>':''}</td>
      <td><select class="mapper-select" data-map-field="${field}">${options}</select></td>
      <td class="mapper-sample">${sample}</td>
      <td>${status}</td>
    </tr>`;
  });

  body.innerHTML=html;
  document.getElementById('mapperSummary').innerHTML=
    `<span class="mapper-pill pill-ok">✓ ${auto} automáticas</span>
     <span class="mapper-pill pill-warn">⚠ ${manual} para revisar</span>
     <span class="mapper-pill ${missingRequired?'pill-bad':'pill-ok'}">${missingRequired?'❌ '+missingRequired+' obligatorias faltantes':'✓ Obligatorios completos'}</span>
     <span class="mapper-pill pill-warn">${optionalMissing} opcionales sin asignar</span>`;

  body.querySelectorAll('select').forEach(sel=>{
    sel.addEventListener('change',()=>{
      const f=sel.dataset.mapField;
      const v=sel.value;
      if(v==='') delete importContext.map[f];
      else importContext.map[f]=Number(v);
      // Re-rendering keeps status/sample accurate and preserves selections.
      const current={...importContext.map};
      const conf={};
      Object.keys(current).forEach(k=>conf[k]=0);
      importContext.confidence=conf;
      showColumnMapper(importContext);
    });
  });

  const requiredMissing=fields.some(f=>{
    const meta=IBITE_FIELDS[f] || {required:false};
    return meta.required && importContext.map[f]===undefined;
  });
  document.getElementById('mapperApply').disabled=requiredMissing;
  document.getElementById('columnMapperModal').classList.add('show');
}

function saveColumnMapping(map, headers){
  try{
    localStorage.setItem('ibite_column_mapping_v2',JSON.stringify({map,headers,savedAt:new Date().toISOString()}));
  }catch(e){}
}
function loadColumnMapping(){
  try{return JSON.parse(localStorage.getItem('ibite_column_mapping_v2')||'null');}catch(e){return null;}
}

function buildDynamicCOL(map){
  const C={};
  Object.keys(map).forEach(k=>{if(k.startsWith('I') && /^I\d+$/.test(k)) C[k]=map[k]; else C[k]=map[k];});
  return C;
}

function parseWorkbookDynamic(wb, suppliedMap=null){
  const wsname=wb.SheetNames.includes(CONFIG.SHEET_NAME)?CONFIG.SHEET_NAME:wb.SheetNames[0];
  const ws=wb.Sheets[wsname];
  if(!ws) throw new Error('No se encontró una hoja de cálculo válida.');
  const rows=XLSX.utils.sheet_to_json(ws,{header:1,raw:true,defval:null});
  if(!rows.length) throw new Error('La hoja seleccionada está vacía.');
  const headers=(rows[0]||[]).map((h,i)=>String(h??'').trim() || `Columna ${i+1}`);
  let mapping=suppliedMap;
  if(!mapping){
    const auto=autoMapHeaders(headers);
    mapping=auto.map;
    const cached=loadColumnMapping();
    if(cached && cached.map){
      const cachedAuto=autoMapHeaders(headers);
      Object.keys(cached.map).forEach(k=>{
        const idx=cached.map[k];
        if(Number.isInteger(idx) && idx<headers.length && !cachedAuto.map[k]) mapping[k]=idx;
      });
    }
    const required=Object.keys(IBITE_FIELDS).filter(k=>IBITE_FIELDS[k].required);
    const missing=required.filter(k=>mapping[k]===undefined);
    // Show mapping UI when required fields are missing OR automatic confidence is
    // incomplete for relevant fields.
    const needsReview=missing.length>0 || Object.keys(IBITE_FIELDS).some(k=>{
      if(k==='INDICATORS' || IBITE_FIELDS[k].required===false) return false;
      return mapping[k]===undefined;
    });
    if(needsReview){
      showColumnMapper({wb,ws,rows,headers,map:mapping,confidence:auto.confidence});
      return null;
    }
    saveColumnMapping(mapping,headers);
  }

  const get=k=>mapping[k]!==undefined ? rows.map(r=>r[mapping[k]]) : [];
  const val=(r,k)=>mapping[k]===undefined?null:r[mapping[k]];
  const cases=[];
  let totalAttempted=0,totalSinDatos=0;
  for(let i=1;i<rows.length;i++){
    const r=rows[i]; if(!r) continue;
    const fila=val(r,'FILA'), ts=toDate(val(r,'TS'));
    if(fila===null && !ts) continue;
    if(!ts) continue;
    totalAttempted++;
    const nombre=val(r,'NOMBRE');
    if(isBlank(nombre)){totalSinDatos++;continue;}
    const ind={}; for(let k=1;k<=17;k++) ind[k]=toNum(val(r,'I'+k));
    cases.push({
      fila,ts,rama:val(r,'RAMA'),nombre:String(nombre).trim(),telefono:val(r,'TEL'),
      edad:toNum(val(r,'EDAD')),sexo:val(r,'SEXO'),institucion:val(r,'INST'),grado:val(r,'GRADO'),comuna:val(r,'COMUNA'),
      convivienteHerido:val(r,'CONVIV_HERIDO'),salidaVivienda:val(r,'SALIDA_VIVIENDA'),danosVivienda:val(r,'DANOS_VIVIENDA'),
      dificultadesServ:val(r,'DIFICULTADES_SERV'),cambioResid:val(r,'CAMBIO_RESID'),ind,
      postraumScore:toNum(val(r,'POSTRAUM_SCORE')),postraumNivel:val(r,'POSTRAUM_NIVEL'),
      genFuncScore:toNum(val(r,'GENFUNC_SCORE')),deterioroFlag:val(r,'DETERIORO_FLAG'),
      combinadoScore:toNum(val(r,'COMBINADO_SCORE')),combinadoNivel:val(r,'COMBINADO_NIVEL'),
      alertaSeguridad:val(r,'ALERTA_SEGURIDAD'),
      reintegraScore:toNum(val(r,'REINTEGRA_SCORE')),reintegraNivel:val(r,'REINTEGRA_NIVEL'),desescolarFlag:val(r,'DESESCOLAR_FLAG'),
      recursosPersScore:toNum(val(r,'RECURSOS_PERS_SCORE')),recursosPersNivel:val(r,'RECURSOS_PERS_NIVEL'),
      recursosFamScore:toNum(val(r,'RECURSOS_FAM_SCORE')),recursosFamNivel:val(r,'RECURSOS_FAM_NIVEL'),
      riesgoMaterial:val(r,'RIESGO_MATERIAL'),necesidadesBasicas:val(r,'NECESIDADES_BASICAS'),
      cuestionarioCompleto:val(r,'CUESTIONARIO_COMPLETO'),perfil:val(r,'PERFIL'),indicadorGeneral:val(r,'INDICADOR_GENERAL'),
      accionRecomendada:val(r,'ACCION_RECOMENDADA'),plazo:val(r,'PLAZO'),ruta:val(r,'RUTA'),accionesEspecificas:val(r,'ACCIONES_ESPECIFICAS')
    });
  }
  return {cases,totalAttempted,totalSinDatos,headers,mapping,wsname};
}

function applyColumnMapping(){
  if(!importContext) return;
  const selects=document.querySelectorAll('#mapperBody select[data-map-field]');
  const mapping={};
  selects.forEach(s=>{if(s.value!=='')mapping[s.dataset.mapField]=Number(s.value);});
  const required=Object.keys(IBITE_FIELDS).filter(k=>IBITE_FIELDS[k].required);
  const missing=required.filter(k=>mapping[k]===undefined);
  if(missing.length){alert('Faltan campos obligatorios: '+missing.map(k=>IBITE_FIELDS[k].label).join(', '));return;}
  saveColumnMapping(mapping,importContext.headers);
  document.getElementById('columnMapperModal').classList.remove('show');
  try{
    const result=parseWorkbookDynamic(importContext.wb,mapping);
    if(!result) return;
    if(result.cases.length===0) alert('El archivo se leyó, pero no se encontraron registros válidos con la estructura seleccionada.');
    applyLoadedData(result.cases,'file',{totalAttempted:result.totalAttempted,totalSinDatos:result.totalSinDatos,mapping:result.mapping,headers:result.headers});
    saveToCache();
    setStatus('Excel importado · '+result.cases.length+' casos · columnas verificadas y mapeadas.');
  }catch(err){
    console.error(err);
    alert('No se pudo procesar el archivo después del mapeo.\\n\\n'+err.message);
  }
}

/* =========================================================
   PARSEO DEL EXCEL
   ========================================================= */
function parseWorkbook(wb){
  const wsname = wb.SheetNames.includes(CONFIG.SHEET_NAME) ? CONFIG.SHEET_NAME : wb.SheetNames[0];
  const ws = wb.Sheets[wsname];
  const rows = XLSX.utils.sheet_to_json(ws, {header:1, raw:true, defval:null});
  const cases = [];
  let totalAttempted = 0, totalSinDatos = 0;

  for(let i=1;i<rows.length;i++){
    const r = rows[i];
    if(!r) continue;
    const fila = r[COL.FILA];
    const ts = toDate(r[COL.TS]);
    if(fila===null && !ts) continue;   // fila de plantilla sin usar
    if(!ts) continue;                  // sin marca temporal = no es un envío real
    totalAttempted++;
    const nombre = r[COL.NOMBRE];
    if(isBlank(nombre)){
      totalSinDatos++;
      continue; // no aceptó participar / no continuó
    }
    const ind = {};
    for(let k=0;k<17;k++){ ind[k+1] = toNum(r[COL.I1+k]); }
    cases.push({
      fila, ts, rama: r[COL.RAMA], nombre: String(nombre).trim(), telefono: r[COL.TEL],
      edad: toNum(r[COL.EDAD]), sexo: r[COL.SEXO], institucion: r[COL.INST], grado: r[COL.GRADO], comuna: r[COL.COMUNA],
      convivienteHerido: r[COL.CONVIV_HERIDO], salidaVivienda: r[COL.SALIDA_VIVIENDA], danosVivienda: r[COL.DANOS_VIVIENDA],
      dificultadesServ: r[COL.DIFICULTADES_SERV], cambioResid: r[COL.CAMBIO_RESID],
      ind,
      postraumScore: toNum(r[COL.POSTRAUM_SCORE]), postraumNivel: r[COL.POSTRAUM_NIVEL],
      genFuncScore: toNum(r[COL.GENFUNC_SCORE]), deterioroFlag: r[COL.DETERIORO_FLAG],
      combinadoScore: toNum(r[COL.COMBINADO_SCORE]), combinadoNivel: r[COL.COMBINADO_NIVEL],
      alertaSeguridad: r[COL.ALERTA_SEGURIDAD],
      reintegraScore: toNum(r[COL.REINTEGRA_SCORE]), reintegraNivel: r[COL.REINTEGRA_NIVEL], desescolarFlag: r[COL.DESESCOLAR_FLAG],
      recursosPersScore: toNum(r[COL.RECURSOS_PERS_SCORE]), recursosPersNivel: r[COL.RECURSOS_PERS_NIVEL],
      recursosFamScore: toNum(r[COL.RECURSOS_FAM_SCORE]), recursosFamNivel: r[COL.RECURSOS_FAM_NIVEL],
      riesgoMaterial: r[COL.RIESGO_MATERIAL], necesidadesBasicas: r[COL.NECESIDADES_BASICAS],
      cuestionarioCompleto: r[COL.CUESTIONARIO_COMPLETO],
      perfil: r[COL.PERFIL], indicadorGeneral: r[COL.INDICADOR_GENERAL],
      accionRecomendada: r[COL.ACCION_RECOMENDADA], plazo: r[COL.PLAZO], ruta: r[COL.RUTA],
      accionesEspecificas: r[COL.ACCIONES_ESPECIFICAS]
    });
  }
  return {cases, totalAttempted, totalSinDatos};
}

/* =========================================================
   CARGA: fetch automático, cache, archivo manual, drag&drop, demo
   ========================================================= */
function setStatus(txt){ document.getElementById('statusLine').textContent = txt; }

function applyLoadedData(cases, sourceLabel, extra){
  state.cases = cases;
  state.meta = Object.assign({source: sourceLabel, savedAt: new Date()}, extra||{});
  document.getElementById('emptyState').classList.add('hidden');
  document.getElementById('tabsNav').classList.remove('hidden');
  document.getElementById('filtersBar').classList.remove('hidden');
  document.getElementById('demoBanner').classList.toggle('hidden', sourceLabel!=='demo');
  document.getElementById('cacheBanner').classList.toggle('hidden', sourceLabel!=='cache');
  if(sourceLabel==='cache' && extra && extra.savedAt){
    document.getElementById('cacheTimestamp').textContent = 'Guardado el ' + fmtDate(new Date(extra.savedAt)) + '.';
  }
  populateFilterOptions();
  renderAll();
  const iq=document.getElementById('importQualityBar');
  if(iq){
    const mapped=extra && extra.mapping ? Object.keys(extra.mapping).length : null;
    iq.style.display = sourceLabel==='file' ? 'block' : 'none';
    if(sourceLabel==='file'){
      iq.textContent = '✓ Excel verificado · ' + (mapped!==null ? mapped + ' campos mapeados' : 'columnas verificadas') +
        (extra && extra.totalAttempted!==undefined ? ' · ' + extra.totalAttempted + ' registros revisados' : '');
    }
  }
  const labelMap = {file:'Excel cargado', fetch:'Excel cargado automáticamente', cache:'datos guardados en este navegador', demo:'datos de EJEMPLO'};
  setStatus('Mostrando ' + (labelMap[sourceLabel]||sourceLabel) + ' · ' + state.cases.length + ' casos con cuestionario iniciado.');
}

function saveToCache(){
  try{
    const serial = state.cases.map(c=>Object.assign({}, c, {ts: c.ts ? c.ts.toISOString() : null}));
    localStorage.setItem(CONFIG.STORAGE_KEY, JSON.stringify({savedAt: new Date().toISOString(), cases: serial}));
  }catch(e){ console.warn('No se pudo guardar en caché local:', e); }
}
function loadFromCache(){
  try{
    const raw = localStorage.getItem(CONFIG.STORAGE_KEY);
    if(!raw) return null;
    const parsed = JSON.parse(raw);
    parsed.cases.forEach(c=>{ c.ts = c.ts ? new Date(c.ts) : null; });
    return parsed;
  }catch(e){ return null; }
}

async function tryAutoFetch(){
  if(typeof XLSX==='undefined'){ return false; }
  try{
    const res = await fetch(encodeURI('./' + CONFIG.EXPECTED_FILENAME), {cache:'no-store'});
    if(!res.ok) return false;
    const buf = await res.arrayBuffer();
    const wb = XLSX.read(buf, {type:'array', cellDates:true});
    const {cases} = parseWorkbook(wb);
    applyLoadedData(cases, 'fetch');
    saveToCache();
    return true;
  }catch(e){ return false; }
}

function handleFile(file){
  if(!file) return;
  const ext=(file.name.split('.').pop()||'').toLowerCase();
  if(!['xlsx','xls','xlsm'].includes(ext)){
    alert('Selecciona un archivo Excel (.xlsx, .xls o .xlsm).');
    return;
  }
  setStatus('Leyendo y verificando columnas de ' + file.name + '…');
  const reader=new FileReader();
  reader.onload=function(e){
    try{
      const wb=XLSX.read(e.target.result,{type:'array',cellDates:true});
      const wsname=wb.SheetNames.includes(CONFIG.SHEET_NAME)?CONFIG.SHEET_NAME:wb.SheetNames[0];
      const ws=wb.Sheets[wsname];
      if(!ws) throw new Error('No se encontró ninguna hoja utilizable.');
      const rows=XLSX.utils.sheet_to_json(ws,{header:1,raw:true,defval:null});
      const headers=(rows[0]||[]).map((h,i)=>String(h??'').trim()||`Columna ${i+1}`);
      if(headers.length===0) throw new Error('No se encontró una fila de encabezados.');

      const auto=autoMapHeaders(headers);
      const cached=loadColumnMapping();
      let mapping={...auto.map};

      // Reuse previous user mappings only when the cached column still exists
      // and the current automatic mapper has not already identified that field.
      if(cached && cached.map){
        Object.keys(cached.map).forEach(k=>{
          const idx=cached.map[k];
          if(mapping[k]===undefined && Number.isInteger(idx) && idx<headers.length) mapping[k]=idx;
        });
      }

      const required=Object.keys(IBITE_FIELDS).filter(k=>IBITE_FIELDS[k].required);
      const missingRequired=required.filter(k=>mapping[k]===undefined);

      // For a human-safe import, show the mapper for every user-uploaded file
      // when there are uncertain/missing fields. Fully matched files can load
      // immediately.
      const importantFields=['TS','NOMBRE','INST','GRADO','INDICADOR_GENERAL','PLAZO','ALERTA_SEGURIDAD'];
      const importantMissing=importantFields.some(k=>mapping[k]===undefined);
      const lowConfidence=importantFields.some(k=>mapping[k]!==undefined && (auto.confidence[k]||0)<.86);

      if(missingRequired.length || importantMissing || lowConfidence){
        showColumnMapper({wb,ws,rows,headers,map:mapping,confidence:auto.confidence});
        setStatus('Excel leído · revisa el mapeo de columnas antes de importar.');
        return;
      }

      saveColumnMapping(mapping,headers);
      const result=parseWorkbookDynamic(wb,mapping);
      if(!result) return;
      if(result.cases.length===0){
        alert('El archivo se leyó, pero no se encontraron registros válidos. Verifica la hoja y las columnas de fecha/nombre.');
      }
      applyLoadedData(result.cases,'file',{
        totalAttempted:result.totalAttempted,
        totalSinDatos:result.totalSinDatos,
        mapping:result.mapping,
        headers:result.headers
      });
      saveToCache();
    }catch(err){
      console.error(err);
      alert('No se pudo leer el archivo. Verifica que sea un Excel válido.\n\nDetalle técnico: '+err.message);
      setStatus('Error al leer el archivo.');
    }
  };
  reader.readAsArrayBuffer(file);
}

function generateDemoCases(){
  const escuelas = ['IE Ejemplo Bolívar','IE Ejemplo San José','IE Ejemplo La Aurora'];
  const comunas = ['Comuna 1 (ejemplo)','Comuna 3 (ejemplo)','Comuna 5 (ejemplo)','Corregimiento El Remanso (ejemplo)'];
  const grados = ['Transición','1.º','3.º','5.º','7.º','9.º','11.º'];
  const perfiles = ['1 - Bienestar preservado','2 - Recursos moderados','3 - Vulnerabilidad global','4 - Bajo bienestar individual/familiar','5 - Recursos mixtos'];
  const profiles = [
    {plazo:'MISMO DÍA', indicador:'Crítico - riesgo de seguridad', alerta:'Alerta - valoración el mismo día', material:'No', deter:'Sí', desesc:'Sí', accion:'VALORACIÓN DE SEGURIDAD EL MISMO DÍA', ruta:'Profesional capacitado / 123 opción 3; activar Código Dorado si se confirma', combinadoNivel:'Elevado', reintegraNivel:'Elevado'},
    {plazo:'48–72 HORAS', indicador:'Alto - desescolarización', alerta:'Sin alerta', material:'Sí', deter:'No', desesc:'Sí', accion:'Contacto con la institución educativa; plan de reingreso gradual', ruta:'Rectoría, docente orientador y Secretaría de Educación', combinadoNivel:'Intermedio', reintegraNivel:'Elevado'},
    {plazo:'DENTRO DE LA SEMANA', indicador:'Alto - malestar/deterioro', alerta:'Sin alerta', material:'No', deter:'Sí', desesc:'No', accion:'Valoración individual y acompañamiento psicosocial focalizado', ruta:'Docente orientador / Zona de Orientación Escolar', combinadoNivel:'Elevado', reintegraNivel:'Intermedio'},
    {plazo:'SEGUIMIENTO HABITUAL', indicador:'Moderado - recursos mixtos', alerta:'Sin alerta', material:'No', deter:'No', desesc:'No', accion:'Promoción y prevención universal o selectiva en el aula', ruta:'Entornos Educativos Saludables / docente orientador', combinadoNivel:'Intermedio', reintegraNivel:'Bajo'},
    {plazo:'SEGUIMIENTO HABITUAL', indicador:'Bajo - bienestar preservado', alerta:'Sin alerta', material:'No', deter:'No', desesc:'No', accion:'Promoción y prevención universal en el aula', ruta:'Entornos Educativos Saludables / docente orientador', combinadoNivel:'Bajo', reintegraNivel:'Bajo'}
  ];
  const cases = [];
  let idCounter = 1;
  const counts = [3,5,6,6,4]; // cuántos casos por perfil de urgencia (crítico->bajo)
  const startDate = new Date(); startDate.setDate(startDate.getDate()-14);
  profiles.forEach((p, idx)=>{
    for(let i=0;i<counts[idx];i++){
      const ts = new Date(startDate.getTime() + Math.random()*14*24*3600*1000);
      const ind = {};
      for(let k=1;k<=17;k++) ind[k] = Math.floor(Math.random()*4);
      cases.push({
        fila: idCounter, ts,
        rama: Math.random()>0.4 ? 'A' : 'B',
        nombre: 'Estudiante Ejemplo ' + idCounter,
        telefono: '300' + String(1000000+idCounter).slice(0,7),
        edad: Math.floor(Math.random()*13)+5,
        sexo: Math.random()>0.5 ? 'Femenino.' : 'Masculino.',
        institucion: escuelas[idCounter % escuelas.length],
        grado: grados[idCounter % grados.length],
        comuna: comunas[idCounter % comunas.length],
        convivienteHerido: 'No.', salidaVivienda: 'No.', danosVivienda: 'Sin daños visibles.',
        dificultadesServ: 'Ninguna.', cambioResid: 'No.',
        ind,
        postraumScore: Math.floor(Math.random()*12), postraumNivel: p.combinadoNivel,
        genFuncScore: Math.floor(Math.random()*6), deterioroFlag: p.deter,
        combinadoScore: Math.floor(Math.random()*18), combinadoNivel: p.combinadoNivel,
        alertaSeguridad: p.alerta,
        reintegraScore: Math.floor(Math.random()*9), reintegraNivel: p.reintegraNivel, desescolarFlag: p.desesc,
        recursosPersScore: Math.floor(Math.random()*9), recursosPersNivel: 'Intermedio',
        recursosFamScore: Math.floor(Math.random()*9), recursosFamNivel: 'Intermedio',
        riesgoMaterial: p.material, necesidadesBasicas: Math.random()>0.6?'Sí':'No',
        cuestionarioCompleto: 'Sí',
        perfil: perfiles[idx % perfiles.length], indicadorGeneral: p.indicador,
        accionRecomendada: p.accion, plazo: p.plazo, ruta: p.ruta,
        accionesEspecificas: 'Ejemplo de acciones específicas combinadas según las señales presentes en este caso ficticio.'
      });
      idCounter++;
    }
  });
  return cases;
}

/* =========================================================
   FILTROS
   ========================================================= */
function uniqueSorted(arr){ return Array.from(new Set(arr.filter(v=>!isBlank(v)))).sort((a,b)=>String(a).localeCompare(String(b),'es')); }

function populateFilterOptions(){
  const insts = uniqueSorted(state.cases.map(c=>c.institucion));
  const comunas = uniqueSorted(state.cases.map(c=>c.comuna));
  const grados = uniqueSorted(state.cases.map(c=>c.grado)).sort((a,b)=>gradeOrder(a)-gradeOrder(b));
  const ramas = uniqueSorted(state.cases.map(c=>c.rama));

  fillSelect('fInstitucion', insts, 'Todas');
  fillSelect('fComuna', comunas, 'Todas');
  fillSelect('fGrado', grados, 'Todos');
  fillSelect('fRama', ramas, 'Todos', v => v==='A' ? 'A · Cuidador (5–10 años)' : (v==='B' ? 'B · Adolescente (11–19 años)' : v));
}
function fillSelect(id, values, allLabel, labelFn){
  const sel = document.getElementById(id);
  const current = sel.value;
  sel.innerHTML = '<option value="">'+allLabel+'</option>' + values.map(v=>`<option value="${escapeHtml(v)}">${escapeHtml(labelFn?labelFn(v):v)}</option>`).join('');
  if(values.includes(current)) sel.value = current;
}

function getFilteredCases(){
  const f = state.filters;
  return state.cases.filter(c=>{
    if(f.institucion && c.institucion!==f.institucion) return false;
    if(f.grado && c.grado!==f.grado) return false;
    if(f.comuna && c.comuna!==f.comuna) return false;
    if(f.rama && c.rama!==f.rama) return false;
    if(f.desde && c.ts && c.ts < new Date(f.desde)) return false;
    if(f.hasta && c.ts && c.ts > new Date(f.hasta + 'T23:59:59')) return false;
    return true;
  });
}

/* =========================================================
   RENDER — TAB 1 (Resumen general)
   ========================================================= */
function kpiCard(cls, num, lbl){
  const colorMap = {
    info: { bg: '#ffffff', border: '#e2e8f0', num: '#0b5b8c' },
    ok: { bg: '#ffffff', border: '#dcfce7', num: '#16a34a' },
    crit: { bg: '#ffffff', border: '#fecaca', num: '#dc2626' },
    alto: { bg: '#ffffff', border: '#fed7aa', num: '#ea580c' }
  };
  const palette = colorMap[cls] || colorMap.info;
  return `
    <div class="kpi ${cls}" style="background:${palette.bg};border:1px solid ${palette.border};border-radius:12px;padding:14px 16px;box-shadow:0 1px 2px rgba(0,0,0,.03);">
      <div class="num" style="font-size:1.7rem;font-weight:800;line-height:1;color:${palette.num};">${num}</div>
      <div class="lbl" style="font-size:.75rem;color:#64748b;margin-top:6px;font-weight:600;">${lbl}</div>
    </div>
  `;
}

function renderKPIs(cases){
  const total = cases.length;
  const completos = cases.filter(c=>isFlagYes(c.cuestionarioCompleto)).length;
  const alertas = cases.filter(c=>alertaClass(c.alertaSeguridad)!=='ninguna').length;
  const materialGrave = cases.filter(c=>isFlagYes(c.riesgoMaterial)).length;
  const deterioro = cases.filter(c=>isFlagYes(c.deterioroFlag)).length;
  const desescolar = cases.filter(c=>isFlagYes(c.desescolarFlag)).length;
  const perfil3 = cases.filter(c=>normalize(c.perfil).startsWith('3')).length;
  const criticos = cases.filter(c=>computeUrgency(c)===0).length;

  const html = [
    kpiCard('info', total, 'Total de casos con cuestionario iniciado'),
    kpiCard('ok', completos, 'Cuestionarios completos'),
    kpiCard('crit', criticos, 'Casos críticos (acción hoy)'),
    kpiCard('crit', alertas, 'Alertas de seguridad (Í7/Í8)'),
    kpiCard('alto', materialGrave, 'Riesgo material grave (vivienda/desplazamiento)'),
    kpiCard('alto', deterioro, 'Bandera de deterioro funcional'),
    kpiCard('alto', desescolar, 'Bandera de desescolarización'),
    kpiCard('alto', perfil3, 'Perfil 3 — Vulnerabilidad global'),
  ].join('');
  document.getElementById('kpiGrid').innerHTML = html;
}

function destroyChart(id){ if(state.charts[id]){ state.charts[id].destroy(); delete state.charts[id]; } }
function makeChart(id, config){
  destroyChart(id);
  const ctx = document.getElementById(id).getContext('2d');
  state.charts[id] = new Chart(ctx, config);
}
const PALETTE = ['#0b5b8c','#0f9b8e','#d97706','#dc2626','#7c3aed','#0891b2','#65a30d','#db2777','#6b7280','#ea580c'];

function renderCharts(cases){
  // Urgencia
  const urgCounts = [0,0,0,0,0];
  cases.forEach(c=> urgCounts[computeUrgency(c)]++ );
  makeChart('chartUrgencia', {
    type:'bar',
    data:{ labels:[0,1,2,3,4].map(k=>URGENCY_META[k].label),
      datasets:[{ data:urgCounts, backgroundColor:[0,1,2,3,4].map(k=>URGENCY_META[k].color) }] },
    options:{ indexAxis:'y', plugins:{legend:{display:false}}, scales:{x:{beginAtZero:true, ticks:{precision:0}}} }
  });

  simpleCountChart('chartIndicador','doughnut', cases, c=>c.indicadorGeneral, {colorByRank: classifyIndicador});
  simpleCountChart('chartPerfil','bar', cases, c=>c.perfil, {});
  simpleCountChart('chartCombinado','bar', cases, c=>c.combinadoNivel, {});
  simpleCountChart('chartReintegra','bar', cases, c=>c.reintegraNivel, {});
  simpleCountChart('chartInstitucion','bar', cases, c=>c.institucion, {horizontal:true});
  simpleCountChart('chartGrado','bar', cases, c=>c.grado, {sortFn:(a,b)=>gradeOrder(a[0])-gradeOrder(b[0])});
  simpleCountChart('chartComuna','bar', cases, c=>c.comuna, {horizontal:true});

  // Tendencia por día
  const dayMap = groupCount(cases, c=> c.ts ? c.ts.toISOString().slice(0,10) : null);
  const days = Array.from(dayMap.keys()).filter(d=>d!=='(Sin dato)').sort();
  makeChart('chartTrend', {
    type:'line',
    data:{ labels: days, datasets:[{ label:'Respuestas', data: days.map(d=>dayMap.get(d)), borderColor:'#0b5b8c', backgroundColor:'rgba(11,91,140,.15)', fill:true, tension:.25 }] },
    options:{ plugins:{legend:{display:false}}, scales:{y:{beginAtZero:true, ticks:{precision:0}}} }
  });
}

function simpleCountChart(id, type, cases, keyFn, opts){
  opts = opts || {};
  const map = groupCount(cases, keyFn);
  let entries = Array.from(map.entries());
  if(opts.sortFn) entries.sort(opts.sortFn); else entries.sort((a,b)=>b[1]-a[1]);
  const labels = entries.map(e=>e[0]);
  const data = entries.map(e=>e[1]);
  let colors;
  if(opts.colorByRank){
    const rankColors = {0:'#dc2626',1:'#ea580c',2:'#d97706',3:'#16a34a'};
    colors = labels.map(l=>{ const r = opts.colorByRank(l); return r===null||r===undefined ? '#6b7280' : (rankColors[r]||'#6b7280'); });
  } else {
    colors = labels.map((_,i)=>PALETTE[i%PALETTE.length]);
  }
  makeChart(id, {
    type: type==='doughnut' ? 'doughnut' : 'bar',
    data:{ labels, datasets:[{ data, backgroundColor:colors }] },
    options: type==='doughnut'
      ? { plugins:{legend:{position:'right', labels:{boxWidth:12,font:{size:10}}}} }
      : { indexAxis: opts.horizontal ? 'y' : 'x', plugins:{legend:{display:false}}, scales:{ [opts.horizontal?'x':'y']:{beginAtZero:true, ticks:{precision:0}} } }
  });
}

/* =========================================================
   RENDER — TAB 2 (Casos prioritarios)
   ========================================================= */
function renderUrgencyChips(cases){
  const counts = [0,0,0,0,0];
  cases.forEach(c=>counts[computeUrgency(c)]++);
  const chips = [0,1,2,3,4].map(k=>{
    const m = URGENCY_META[k];
    const active = state.tab2.urgency===k ? 'active' : '';
    return `<button class="chip ${active}" data-urg="${k}" style="background:${m.bg};color:${m.color};border-radius:999px;padding:8px 14px;font-size:.8rem;font-weight:700;cursor:pointer;border:2px solid transparent;display:flex;gap:6px;align-items:center;${state.tab2.urgency===k ? 'outline:2px solid rgba(0,0,0,.15);' : ''}">${m.short}: ${counts[k]}</button>`;
  }).join('');
  const allActive = state.tab2.urgency===null ? 'active' : '';
  document.getElementById('urgencyChips').innerHTML =
    `<button class="chip ${allActive}" data-urg="all" style="background:#e2e8f0;color:#1e293b;border-radius:999px;padding:8px 14px;font-size:.8rem;font-weight:700;cursor:pointer;border:2px solid transparent;display:flex;gap:6px;align-items:center;${state.tab2.urgency===null ? 'outline:2px solid rgba(0,0,0,.15);' : ''}">Todos: ${cases.length}</button>` + chips;

  document.querySelectorAll('#urgencyChips .chip').forEach(btn=>{
    btn.addEventListener('click', ()=>{
      const v = btn.getAttribute('data-urg');
      state.tab2.urgency = v==='all' ? null : parseInt(v,10);
      renderTab2();
    });
  });
}

function badge(text, color, bg){
  return `<span class="badge" style="color:${color};background:${bg};border-radius:999px;padding:4px 10px;font-size:.72rem;font-weight:700;white-space:nowrap;">${escapeHtml(text)}</span>`;
}
function badgeOutline(text){
  return `<span class="badge outline" style="border-radius:999px;padding:4px 10px;font-size:.72rem;font-weight:700;white-space:nowrap;background:#fff;border:1px solid #e2e8f0;color:#64748b;">${escapeHtml(text)}</span>`;
}

function renderCaseList(cases){
  const container = document.getElementById('caseList');
  document.getElementById('casosCount').textContent = cases.length + ' caso(s) mostrados de ' + state.cases.length + ' en total.';
  if(cases.length===0){ container.innerHTML = '<p class="small-note">No hay casos que coincidan con los filtros seleccionados.</p>'; return; }

  container.innerHTML = cases.map(c=>{
    const urg = computeUrgency(c);
    const meta = URGENCY_META[urg];
    const ac = alertaClass(c.alertaSeguridad);
    const badges = [];
    if(ac==='activa') badges.push(badge('⚠ Alerta de seguridad', '#991b1b', '#fee2e2'));
    if(ac==='no_aclarada') badges.push(badge('❓ Seguridad no aclarada', '#9a3412', '#ffedd5'));
    if(isFlagYes(c.riesgoMaterial)) badges.push(badge('🏚 Riesgo material grave', '#9a3412', '#ffedd5'));
    if(isFlagYes(c.deterioroFlag)) badges.push(badge('📉 Deterioro funcional', '#92400e', '#fef3c7'));
    if(isFlagYes(c.desescolarFlag)) badges.push(badge('🏫 Desescolarización', '#92400e', '#fef3c7'));
    if(isFlagYes(c.necesidadesBasicas)) badges.push(badgeOutline('Necesidades básicas afectadas'));
    if(!isFlagYes(c.cuestionarioCompleto)) badges.push(badgeOutline('Cuestionario incompleto'));

    const indicatorRows = INDICATOR_LABELS.map((lbl,i)=>{
      const v = c.ind[i+1];
      return `<tr><td>Í${i+1} ${lbl}</td><td>${v===null||v===undefined?'—':v}</td></tr>`;
    }).join('');

    return `<details class="case-card" style="background:#fff;border-radius:12px;border:1px solid #e2e8f0;border-left:6px solid ${meta.color};overflow:hidden;">
      <summary style="list-style:none;cursor:pointer;padding:12px 14px;display:flex;justify-content:space-between;gap:12px;align-items:center;flex-wrap:wrap;">
        <div class="case-main" style="display:flex;flex-direction:column;gap:3px;">
          <span class="name" style="font-weight:700;font-size:1rem;">${escapeHtml(c.nombre)}</span>
          <span class="meta" style="font-size:.78rem;color:#64748b;">${c.edad?c.edad+' años · ':''}${escapeHtml(c.sexo||'')} · ${escapeHtml(c.institucion||'Sin institución')} · Grado ${escapeHtml(c.grado||'—')} · ${escapeHtml(c.comuna||'Sin comuna')} · Registrado: ${fmtDate(c.ts)}</span>
        </div>
        <div class="case-badges" style="display:flex;gap:6px;flex-wrap:wrap;align-items:center;">
          <span class="badge" style="background:${meta.bg};color:${meta.color};border-radius:999px;padding:4px 10px;font-size:.72rem;font-weight:700;white-space:nowrap;">${meta.label}</span>
          ${badges.join('')}
        </div>
      </summary>
      <div class="case-detail" style="padding:0 16px 16px;border-top:1px dashed #e2e8f0;">
        <div class="action-box ${urg===0?'crit':''}" style="margin-top:12px;background:${urg===0 ? '#fee2e2' : '#fff7ed'};border:1px solid ${urg===0 ? '#fecaca' : '#fed7aa'};border-radius:8px;padding:10px 12px;">
          <strong style="display:block;margin-bottom:2px;">Acción recomendada (${escapeHtml(c.plazo||'—')})</strong>
          ${escapeHtml(c.accionRecomendada||'—')}
          <div style="margin-top:6px;font-size:.82rem;color:#7c2d12"><strong>Ruta / responsable:</strong> ${escapeHtml(c.ruta||'—')}</div>
        </div>
        <div class="grid" style="display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:10px;margin-top:12px;">
          <div class="block" style="background:#eef1f4;border-radius:8px;padding:10px 12px;"><h4 style="margin:0 0 4px;font-size:.75rem;text-transform:uppercase;color:#64748b;letter-spacing:.03em;">Contacto</h4><p style="margin:0;font-size:.86rem;">📞 ${escapeHtml(c.telefono||'Sin teléfono')}${c.telefono?` — <a href="tel:${escapeHtml(String(c.telefono))}">llamar</a>`:''}</p></div>
          <div class="block" style="background:#eef1f4;border-radius:8px;padding:10px 12px;"><h4 style="margin:0 0 4px;font-size:.75rem;text-transform:uppercase;color:#64748b;letter-spacing:.03em;">Quién respondió</h4><p style="margin:0;font-size:.86rem;">${c.rama==='A'?'Cuidador (niño/a 5–10 años)':(c.rama==='B'?'Autoreporte (11–19 años)':escapeHtml(c.rama||'—'))}</p></div>
          <div class="block" style="background:#eef1f4;border-radius:8px;padding:10px 12px;"><h4 style="margin:0 0 4px;font-size:.75rem;text-transform:uppercase;color:#64748b;letter-spacing:.03em;">Perfil operativo</h4><p style="margin:0;font-size:.86rem;">${escapeHtml(c.perfil||'—')}</p></div>
          <div class="block" style="background:#eef1f4;border-radius:8px;padding:10px 12px;"><h4 style="margin:0 0 4px;font-size:.75rem;text-transform:uppercase;color:#64748b;letter-spacing:.03em;">Indicador de riesgo general</h4><p style="margin:0;font-size:.86rem;">${escapeHtml(c.indicadorGeneral||'—')}</p></div>
          <div class="block" style="background:#eef1f4;border-radius:8px;padding:10px 12px;"><h4 style="margin:0 0 4px;font-size:.75rem;text-transform:uppercase;color:#64748b;letter-spacing:.03em;">Malestar postraumático</h4><p style="margin:0;font-size:.86rem;">${c.postraumScore!==null?c.postraumScore+'/12 · ':''}${escapeHtml(c.postraumNivel||'—')}</p></div>
          <div class="block" style="background:#eef1f4;border-radius:8px;padding:10px 12px;"><h4 style="margin:0 0 4px;font-size:.75rem;text-transform:uppercase;color:#64748b;letter-spacing:.03em;">Malestar combinado</h4><p style="margin:0;font-size:.86rem;">${c.combinadoScore!==null?c.combinadoScore+'/18 · ':''}${escapeHtml(c.combinadoNivel||'—')}</p></div>
          <div class="block" style="background:#eef1f4;border-radius:8px;padding:10px 12px;"><h4 style="margin:0 0 4px;font-size:.75rem;text-transform:uppercase;color:#64748b;letter-spacing:.03em;">Reintegración escolar</h4><p style="margin:0;font-size:.86rem;">${c.reintegraScore!==null?c.reintegraScore+'/9 · ':''}${escapeHtml(c.reintegraNivel||'—')}</p></div>
          <div class="block" style="background:#eef1f4;border-radius:8px;padding:10px 12px;"><h4 style="margin:0 0 4px;font-size:.75rem;text-transform:uppercase;color:#64748b;letter-spacing:.03em;">Recursos personales</h4><p style="margin:0;font-size:.86rem;">${c.recursosPersScore!==null?c.recursosPersScore+'/9 · ':''}${escapeHtml(c.recursosPersNivel||'—')}</p></div>
          <div class="block" style="background:#eef1f4;border-radius:8px;padding:10px 12px;"><h4 style="margin:0 0 4px;font-size:.75rem;text-transform:uppercase;color:#64748b;letter-spacing:.03em;">Recursos familiares</h4><p style="margin:0;font-size:.86rem;">${c.recursosFamScore!==null?c.recursosFamScore+'/9 · ':''}${escapeHtml(c.recursosFamNivel||'—')}</p></div>
          <div class="block" style="background:#eef1f4;border-radius:8px;padding:10px 12px;"><h4 style="margin:0 0 4px;font-size:.75rem;text-transform:uppercase;color:#64748b;letter-spacing:.03em;">Contexto material</h4><p style="margin:0;font-size:.86rem;">Vivienda: ${escapeHtml(c.danosVivienda||'—')}<br>Salida temporal: ${escapeHtml(c.salidaVivienda||'—')}<br>Dificultades de servicios: ${escapeHtml(c.dificultadesServ||'—')}</p></div>
        </div>
        <details style="margin-top:10px;">
          <summary style="cursor:pointer;font-size:.82rem;color:#64748b;">Ver los 17 indicadores IBITE-17</summary>
          <table class="indicator-table" style="width:100%;border-collapse:collapse;margin-top:10px;font-size:.8rem;"><tbody>${indicatorRows}</tbody></table>
        </details>
        ${c.accionesEspecificas ? `<div class="action-box" style="margin-top:10px;background:#eff6ff;border:1px solid #bfdbfe;border-radius:8px;padding:10px 12px;"><strong style="display:block;margin-bottom:2px;">Acciones específicas por señal</strong>${escapeHtml(c.accionesEspecificas)}</div>` : ''}
      </div>
    </details>`;
  }).join('');
}

function renderTab2(){
  let cases = getFilteredCases();
  if(state.tab2.urgency!==null) cases = cases.filter(c=>computeUrgency(c)===state.tab2.urgency);
  if(state.tab2.onlySeguridad) cases = cases.filter(c=>alertaClass(c.alertaSeguridad)!=='ninguna');
  if(state.tab2.search){
    const s = normalize(state.tab2.search);
    cases = cases.filter(c=> normalize(c.nombre).includes(s) || normalize(c.institucion).includes(s) || normalize(c.grado).includes(s) || normalize(c.comuna).includes(s));
  }
  cases.sort((a,b)=> computeUrgency(a)-computeUrgency(b) || (b.ts - a.ts));
  renderUrgencyChips(getFilteredCases());
  renderCaseList(cases);
}

/* =========================================================
   RENDER ALL
   ========================================================= */
function renderAll(){
  const filtered = getFilteredCases();
  renderKPIs(filtered);
  renderCharts(filtered);
  renderTab2();
}

/* =========================================================
   EXPORTAR CSV
   ========================================================= */
function exportCSV(){
  let cases = getFilteredCases();
  if(state.tab2.urgency!==null) cases = cases.filter(c=>computeUrgency(c)===state.tab2.urgency);
  if(state.tab2.onlySeguridad) cases = cases.filter(c=>alertaClass(c.alertaSeguridad)!=='ninguna');
  cases.sort((a,b)=> computeUrgency(a)-computeUrgency(b));
  const headers = ['Nombre','Edad','Sexo','Institución','Grado','Comuna','Teléfono','Urgencia','Alerta de seguridad','Riesgo material grave','Deterioro funcional','Desescolarización','Indicador de riesgo general','Acción recomendada','Plazo de acción','Ruta/responsable','Fecha de registro'];
  const rows = cases.map(c=>[
    c.nombre, c.edad, c.sexo, c.institucion, c.grado, c.comuna, c.telefono,
    URGENCY_META[computeUrgency(c)].short, c.alertaSeguridad, c.riesgoMaterial, c.deterioroFlag, c.desescolarFlag,
    c.indicadorGeneral, c.accionRecomendada, c.plazo, c.ruta, c.ts ? c.ts.toISOString() : ''
  ]);
  const csv = [headers, ...rows].map(r=>r.map(csvEscape).join(';')).join('\r\n');
  const blob = new Blob(['\ufeff'+csv], {type:'text/csv;charset=utf-8;'});
  const url = URL.createObjectURL(blob);
  const a = document.createElement('a'); a.href = url; a.download = 'ibite_casos_prioritarios_' + dateStamp() + '.csv'; document.body.appendChild(a); a.click(); a.remove();
  URL.revokeObjectURL(url);
}

/* =========================================================
   EVENTOS UI
   ========================================================= */
document.querySelectorAll('nav.tabs .tab').forEach(btn=>{
  btn.addEventListener('click', ()=>{
    document.querySelectorAll('nav.tabs .tab').forEach(b=>b.classList.remove('active'));
    btn.classList.add('active');
    const tabId = btn.getAttribute('data-tab');
    document.querySelectorAll('.tabpanel').forEach(p=>p.classList.remove('active'));
    document.getElementById('tab-'+tabId).classList.add('active');
  });
});

document.getElementById('closeConfBanner').addEventListener('click', ()=>{
  document.getElementById('confBanner').classList.add('hidden');
});

['fInstitucion','fGrado','fComuna','fRama'].forEach(id=>{
  document.getElementById(id).addEventListener('change', e=>{
    const map = {fInstitucion:'institucion',fGrado:'grado',fComuna:'comuna',fRama:'rama'};
    state.filters[map[id]] = e.target.value;
    renderAll();
  });
});
document.getElementById('fDesde').addEventListener('change', e=>{ state.filters.desde = e.target.value; renderAll(); });
document.getElementById('fHasta').addEventListener('change', e=>{ state.filters.hasta = e.target.value; renderAll(); });
document.getElementById('btnClearFilters').addEventListener('click', ()=>{
  state.filters = {institucion:'', grado:'', comuna:'', rama:'', desde:'', hasta:''};
  ['fInstitucion','fGrado','fComuna','fRama'].forEach(id=>document.getElementById(id).value='');
  document.getElementById('fDesde').value=''; document.getElementById('fHasta').value='';
  renderAll();
});

document.getElementById('searchBox').addEventListener('input', e=>{ state.tab2.search = e.target.value; renderTab2(); });
document.getElementById('onlySeguridad').addEventListener('change', e=>{ state.tab2.onlySeguridad = e.target.checked; renderTab2(); });
document.getElementById('btnExportCSV').addEventListener('click', exportCSV);
document.getElementById('btnPrint').addEventListener('click', ()=>window.print());

const fileInput = document.getElementById('fileInput');
function openFilePicker(){ fileInput.click(); }
document.getElementById('btnLoad').addEventListener('click', openFilePicker);
document.getElementById('btnReloadFromCache').addEventListener('click', openFilePicker);
fileInput.addEventListener('change', e=>{ if(e.target.files[0]) handleFile(e.target.files[0]); });

function loadDemo(){
  applyLoadedData(generateDemoCases(), 'demo');
}
document.getElementById('btnDemo').addEventListener('click', loadDemo);
document.getElementById('btnDemo2').addEventListener('click', loadDemo);

// Drag & drop en toda la página
let dragCounter = 0;
window.addEventListener('dragenter', e=>{ e.preventDefault(); dragCounter++; document.getElementById('dragOverlay').style.display='flex'; });
window.addEventListener('dragleave', e=>{ e.preventDefault(); dragCounter--; if(dragCounter<=0){ document.getElementById('dragOverlay').style.display='none'; dragCounter=0; } });
window.addEventListener('dragover', e=> e.preventDefault());
window.addEventListener('drop', e=>{
  e.preventDefault();
  dragCounter = 0;
  document.getElementById('dragOverlay').style.display='none';
  if(e.dataTransfer.files && e.dataTransfer.files[0]) handleFile(e.dataTransfer.files[0]);
});

/* =========================================================
   INICIO
   ========================================================= */
async function init(){
  try{
    ensureMapperUI();

    // XLSX is loaded from the existing page dependency. If it isn't available,
    // don't leave the dashboard permanently stuck on "Inicializando".
    if(typeof XLSX === 'undefined'){
      setStatus('⚠️ No se pudo cargar el lector de Excel. Puedes revisar la conexión y volver a cargar la página.');
      const b=document.getElementById('importQualityBar');
      if(b){
        b.style.display='block';
        b.style.background='#fffbeb';
        b.style.borderColor='#fde68a';
        b.style.color='#92400e';
        b.textContent='⚠️ Lector de Excel no disponible. Recarga la página o verifica tu conexión a internet.';
      }
      // Demo data keeps the dashboard usable even if the CDN is unavailable.
      loadDemo();
      return;
    }

    setStatus('Buscando "' + CONFIG.EXPECTED_FILENAME + '" en la misma carpeta…');
    const okFetch = await tryAutoFetch();
    if(okFetch) return;

    const cached = loadFromCache();
    if(cached && cached.cases && cached.cases.length){
      applyLoadedData(cached.cases, 'cache', {savedAt: cached.savedAt});
      return;
    }

    setStatus('No se encontró el Excel automáticamente. Carga el archivo manualmente o mira los datos de ejemplo.');
    loadDemo();
  }catch(err){
    console.error('IBITE initialization error:',err);
    setStatus('⚠️ El tablero no pudo completar la inicialización. Puedes cargar un Excel manualmente.');
    try{ loadDemo(); }catch(e){ console.error(e); }
  }
}
init();

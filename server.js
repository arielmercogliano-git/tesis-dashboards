const express = require('express');
const cors = require('cors');
const { AsyncLocalStorage } = require('async_hooks');
const fetch = (...args) => import('node-fetch').then(({default: f}) => f(...args));

const auth = require('./auth');
const oauth = require('./oauth');

const app = express();
const PORT = process.env.PORT || 3737;

const TENANT_ID = process.env.BC_TENANT_ID;
const CLIENT_ID = process.env.BC_CLIENT_ID;
const CLIENT_SECRET = process.env.BC_CLIENT_SECRET;
const COMPANY = process.env.BC_COMPANY || 'My%20Company';

if (!TENANT_ID || !CLIENT_ID || !CLIENT_SECRET) { console.error('ERROR: Faltan variables de entorno'); process.exit(1); }

const ODATA = `https://api.businesscentral.dynamics.com/v2.0/${TENANT_ID}/Production/ODataV4/Company('${COMPANY}')`;
const V2 = `https://api.businesscentral.dynamics.com/v2.0/${TENANT_ID}/Production/api/v2.0`;

app.use(cors({ origin: '*', methods: ['GET','POST','DELETE'] }));
app.use(express.json());
app.use(express.urlencoded({ extended: true })); // NUEVO: /oauth/token puede llegar form-urlencoded
// NOTA: la vieja x-api-key fija se sacó de todas las rutas — quedaba expuesta en el JS de
// los dashboards y ya no aportaba nada real, porque la protección de verdad es el login por
// email+código (dashboards) y OAuth (conector MCP), ambos por persona.

// Advertencias por llamada: si alguna lectura a BC se corta por un tope, se registra acá y la
// respuesta de la herramienta las incluye en "advertencias" (antes se cortaba sin avisar).
const als = new AsyncLocalStorage();
function avisarTruncado(msg){ const s = als.getStore(); if (s) s.warns.push(msg); console.warn('[TRUNCADO]', msg); }

let cachedToken=null, tokenExp=0;
async function getToken() {
if (cachedToken && Date.now()<tokenExp-60000) return cachedToken;
const r=await fetch(`https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`,{method:'POST',headers:{'Content-Type':'application/x-www-form-urlencoded'},body:new URLSearchParams({grant_type:'client_credentials',client_id:CLIENT_ID,client_secret:CLIENT_SECRET,scope:'https://api.businesscentral.dynamics.com/.default'}).toString()});
const d=await r.json(); if(!r.ok) throw new Error(d.error_description||JSON.stringify(d));
cachedToken=d.access_token; tokenExp=Date.now()+d.expires_in*1000; return cachedToken;
}

async function bcGet(url) {
const t=await getToken();
const r=await fetch(url,{headers:{Authorization:`Bearer ${t}`,Accept:'application/json'}});
if(!r.ok){const e=await r.json().catch(()=>({}));throw new Error(e.error?.message||`HTTP ${r.status}`);}
return r.json();
}

// Sigue @odata.nextLink hasta traer TODO (BC pagina solo). max es solo una red de seguridad contra
// consultas descontroladas; si se llega a ese tope con páginas pendientes, se avisa en vez de callar.
async function bcGetAll(url,max=300000) {
const rows=[];let next=url;
while(next&&rows.length<max){const d=await bcGet(next);rows.push(...(d.value||[]));next=d['@odata.nextLink']||null;}
if(next) avisarTruncado(`Lectura cortada en ${rows.length} registros (tope de seguridad ${max}); hay más datos en Business Central.`);
return rows;
}

let _coId=null;
async function coId(){if(_coId)return _coId;const d=await bcGet(`${V2}/companies`);_coId=(d.value||[])[0]?.id;return _coId;}

let _empMap=null;
async function empMap(){
if(_empMap)return _empMap;
const id=await coId();
const dims=await bcGet(`${V2}/companies(${id})/dimensions`);
const rec=(dims.value||[]).find(d=>(d.code||'').toUpperCase()==='RECURSO');
if(!rec)return{};
const vals=await bcGet(`${V2}/companies(${id})/dimensions(${rec.id})/dimensionValues`);
_empMap={};(vals.value||[]).forEach(v=>{if(v.code)_empMap[v.code]=v.displayName||v.name||v.code;});
return _empMap;
}

async function v2Get(ep,params=''){const id=await coId();return bcGetAll(`${V2}/companies(${id})/${ep}${params}`);}

// Para consultas del usuario con "top" explícito (ej. facturas_venta top=20): BC devuelve como
// máximo N. Si devuelve exactamente N probablemente haya más: se avisa.
function avisarSiTopCompleto(nombre,rows,top){ if(top && rows.length>=top) avisarTruncado(`${nombre}: se devolvieron exactamente ${top} registros (el pedido pedía top=${top}); puede haber más. Pedí un top mayor para ver todo.`); }

async function queryMovs(year,entryType){
let f=year&&year!=='all'?`Posting_Date ge ${year}-01-01T00:00:00Z and Posting_Date le ${year}-12-31T23:59:59Z`:'';
const rows=await bcGetAll(`${ODATA}/MovsProyecto${f?'?$filter='+encodeURIComponent(f):''}`);
if(!entryType)return rows;
return rows.filter(r=>{const et=String(r.Entry_Type||'').toLowerCase();return entryType==='usage'?(et.includes('usage')||r.Entry_Type===1||r.Entry_Type==='1'):(et.includes('sale')||r.Entry_Type===0||r.Entry_Type==='0');});
}

// ── Costo directo / nómina por empleado — cruza G_LEntries con la dimensión RECURSO ──
// (mismo mecanismo ya validado en el dashboard estados_financieros.html)
const SOCIOS_RECURSO=['D 001','D 002']; // Ariel Mercogliano, Pablo Berman — excluidos (no son costo directo)
const CTA_MAQUILA_BASE='511007'; // Servicios de maquila — no es costo directo de personal
function accountBase(n){return (n||'').split('-')[0];}
function esCuentaCostoDirecto(accNo){
const base=accountBase(accNo);
// La maquila (511007) SÍ cuenta como costo directo acá: cuando está tageada a un empleado normal
// (no socio) representa el "Adicional Nómina" que se le paga por esa vía. Los socios (Ariel/Pablo)
// ya se excluyen aparte en costoPorEmpleadoGenerico, así que no hay riesgo de mezclar sus extracciones.
return base.startsWith('511')||base.startsWith('512')||base.startsWith('513')||base.startsWith('514')||base.startsWith('520');
}
const NOMINA_KEYWORDS=['SUELDO','SALARIO','HORAS EXTRA','GRATIFICACION','PREMIO','DESPENSA','AGUINALDO','PRIMA VACACIONAL','VACACIONES','CUOTAS IMSS','RETIRO','CESANTIA','INFONAVIT','IMPUESTO SOBRE NOMINA','IMPUESTO AL SALARIO','VALES DE','FOMENTO EDUCACION'];

async function dimSetsRecurso(){
const rows=await bcGetAll(`${ODATA}/DimensionSetEntries?$filter=${encodeURIComponent("Dimension_Code eq 'RECURSO'")}`);
const map={};rows.forEach(r=>{if(r.Dimension_Value_Code)map[r.Dimension_Set_ID]=r.Dimension_Value_Code;});
return map;
}
async function glEntriesAnio(year){
let f=year&&year!=='all'?`Posting_Date ge ${year}-01-01T00:00:00Z and Posting_Date le ${year}-12-31T23:59:59Z`:'';
return bcGetAll(`${ODATA}/G_LEntries${f?'?$filter='+encodeURIComponent(f):''}`);
}
async function cuentasNomina(){
const rows=await v2Get('accounts');
const set=new Set();
rows.forEach(a=>{const nombre=(a.displayName||'').toUpperCase();if(NOMINA_KEYWORDS.some(k=>nombre.includes(k)))set.add(accountBase(a.number));});
return set;
}
function mesKeyG(d){return (d||'').slice(0,7);}

async function costoPorEmpleadoGenerico(empleado_codigo,year,filtroFn){
const[gl,dimMap,m]=await Promise.all([glEntriesAnio(year),dimSetsRecurso(),empMap()]);
const porEmp={};
gl.forEach(e=>{
const accNo=e.G_L_Account_No;if(!accNo)return;
if(!filtroFn(accNo))return;
const empCod=dimMap[e.Dimension_Set_ID];
if(!empCod||SOCIOS_RECURSO.includes(empCod))return;
if(empleado_codigo&&empCod!==empleado_codigo)return;
const monto=(e.Debit_Amount||0)-(e.Credit_Amount||0);
const mk=mesKeyG(e.Posting_Date);if(!mk)return;
if(!porEmp[empCod])porEmp[empCod]={};
porEmp[empCod][mk]=(porEmp[empCod][mk]||0)+monto;
});
return Object.entries(porEmp).map(([codigo,meses])=>({
codigo,nombre:m[codigo]||codigo,
total:parseFloat(Object.values(meses).reduce((a,b)=>a+b,0).toFixed(2)),
por_mes:Object.keys(meses).sort().map(mk=>({mes:mk,monto:parseFloat(meses[mk].toFixed(2))}))
})).sort((a,b)=>b.total-a.total);
}

// Nómina "real" de cada empleado = nómina legal (cuentas de sueldo/prestaciones) + Adicional Nómina
// (parte de la maquila que Business Central tiene tageada a ese empleado vía RECURSO, no a un socio).
// Se devuelven ambas partes por separado además del total, para que quede claro de dónde sale cada peso.
async function nominaConAdicionalPorEmpleado(empleado_codigo,year){
const[gl,dimMap,m,cuentasNom]=await Promise.all([glEntriesAnio(year),dimSetsRecurso(),empMap(),cuentasNomina()]);
const porEmp={};
gl.forEach(e=>{
const accNo=e.G_L_Account_No;if(!accNo)return;
const empCod=dimMap[e.Dimension_Set_ID];
if(!empCod||SOCIOS_RECURSO.includes(empCod))return;
if(empleado_codigo&&empCod!==empleado_codigo)return;
const base=accountBase(accNo);
const esLegal=cuentasNom.has(base);
const esAdicionalMaquila=base===CTA_MAQUILA_BASE;
if(!esLegal&&!esAdicionalMaquila)return;
const monto=(e.Debit_Amount||0)-(e.Credit_Amount||0);
const mk=mesKeyG(e.Posting_Date);if(!mk)return;
if(!porEmp[empCod])porEmp[empCod]={};
if(!porEmp[empCod][mk])porEmp[empCod][mk]={legal:0,adicional_nomina:0};
if(esLegal)porEmp[empCod][mk].legal+=monto;else porEmp[empCod][mk].adicional_nomina+=monto;
});
return Object.entries(porEmp).map(([codigo,meses])=>{
const por_mes=Object.keys(meses).sort().map(mk=>{
const legal=parseFloat(meses[mk].legal.toFixed(2));
const adicional=parseFloat(meses[mk].adicional_nomina.toFixed(2));
return{mes:mk,nomina_legal:legal,adicional_nomina:adicional,total:parseFloat((legal+adicional).toFixed(2))};
});
const total_nomina_legal=parseFloat(por_mes.reduce((s,x)=>s+x.nomina_legal,0).toFixed(2));
const total_adicional_nomina=parseFloat(por_mes.reduce((s,x)=>s+x.adicional_nomina,0).toFixed(2));
return{codigo,nombre:m[codigo]||codigo,total:parseFloat((total_nomina_legal+total_adicional_nomina).toFixed(2)),total_nomina_legal,total_adicional_nomina,por_mes};
}).sort((a,b)=>b.total-a.total);
}

// Facturación (en pesos, no solo horas) por empleado, mes a mes. Mismo prorrateo que usan los
// dashboards: el monto de cada factura se reparte entre los empleados según cuántas horas de
// esa tarea trabajó cada uno (share = horas del empleado / horas totales de la tarea).
async function facturacionPorEmpleadoMensual(empleado_codigo,year){
const[all,m]=await Promise.all([queryMovs(year),empMap()]);
const isUsageRow=r=>{const et=String(r.Entry_Type||'').toLowerCase();return et.includes('usage')||r.Entry_Type===1||r.Entry_Type==='1';};
const usageRows=all.filter(isUsageRow);
const porEmp={};
usageRows.forEach(u=>{
const empCod=u.Global_Dimension_2_Code;if(!empCod)return;
if(empleado_codigo&&empCod!==empleado_codigo)return;
const qty=parseFloat(u.Quantity||0);
// Monto propio de ESTA línea (el valor/hora de este empleado en esta tarea), no un promedio
// prorrateado de la factura agregada — cada empleado puede tener una tarifa distinta.
const sinIVA=Math.abs(parseFloat(u.Total_Price||0));
const mk=(u.Posting_Date||'').slice(0,7);if(!mk)return;
if(!porEmp[empCod])porEmp[empCod]={};
if(!porEmp[empCod][mk])porEmp[empCod][mk]={horas:0,sinIVA:0};
porEmp[empCod][mk].horas+=qty;
porEmp[empCod][mk].sinIVA+=sinIVA;
});
return Object.entries(porEmp).map(([codigo,meses])=>{
const por_mes=Object.keys(meses).sort().map(mk=>{
const sinIVA=parseFloat(meses[mk].sinIVA.toFixed(2));
const horas=parseFloat(meses[mk].horas.toFixed(2));
return{mes:mk,horas,facturacion_sin_iva:sinIVA,facturacion_con_iva:parseFloat((sinIVA*1.16).toFixed(2)),valor_hora:horas>0?parseFloat((sinIVA/horas).toFixed(2)):null};
});
const totalHoras=parseFloat(por_mes.reduce((s,x)=>s+x.horas,0).toFixed(2));
const totalSinIVA=parseFloat(por_mes.reduce((s,x)=>s+x.facturacion_sin_iva,0).toFixed(2));
return{codigo,nombre:m[codigo]||codigo,total_horas:totalHoras,total_facturacion_sin_iva:totalSinIVA,total_facturacion_con_iva:parseFloat((totalSinIVA*1.16).toFixed(2)),valor_hora_promedio:totalHoras>0?parseFloat((totalSinIVA/totalHoras).toFixed(2)):null,por_mes};
}).sort((a,b)=>b.total_facturacion_sin_iva-a.total_facturacion_sin_iva);
}

// ── Período real de la tarea — vive en Job_Task_Lines.Description, no en el código T-N ──
// (mismo mecanismo ya validado en los dashboards). Distintos proyectos usan distintos formatos:
// "DEL DD/MM/YYYY AL DD/MM/YYYY", "MES AÑO", o a veces solo un texto sin período (ej. "TIPO 3"),
// en cuyo caso no hay período que resolver y se cae al código de tarea tal cual.
const MESES_ES={ENERO:1,FEBRERO:2,MARZO:3,ABRIL:4,MAYO:5,JUNIO:6,JULIO:7,AGOSTO:8,SEPTIEMBRE:9,OCTUBRE:10,NOVIEMBRE:11,DICIEMBRE:12};
const MESES_ABR={ENE:1,FEB:2,MAR:3,ABR:4,MAY:5,JUN:6,JUL:7,AGO:8,SEP:9,OCT:10,NOV:11,DIC:12};
function diaMayoriaMeses(d1,m1,y1,d2,m2,y2){
const inicio=new Date(y1,m1-1,d1), fin=new Date(y2,m2-1,d2);
if(isNaN(inicio)||isNaN(fin)||fin<inicio) return `${y1}-${String(m1).padStart(2,'0')}`;
const dias={}; let cur=new Date(inicio), guard=0;
while(cur<=fin&&guard<400){const k=`${cur.getFullYear()}-${String(cur.getMonth()+1).padStart(2,'0')}`;dias[k]=(dias[k]||0)+1;cur.setDate(cur.getDate()+1);guard++;}
let mejor=null,max=-1; Object.keys(dias).forEach(k=>{if(dias[k]>max){max=dias[k];mejor=k;}});
return mejor;
}
function parsearPeriodoTarea(desc){
if(!desc) return null;
const d=desc.toUpperCase();
let m=d.match(/DEL\s+(\d{1,2})\/(\d{1,2})\/(\d{4})\s+AL\s+(\d{1,2})\/(\d{1,2})\/(\d{4})/);
if(m){const mes=diaMayoriaMeses(+m[1],+m[2],+m[3],+m[4],+m[5],+m[6]);return{mes,label:`${m[1]}/${m[2]}/${m[3]} – ${m[4]}/${m[5]}/${m[6]}`};}
m=d.match(/\b([A-ZÑ]{3,12})-([A-ZÑ]{3,12})\s+(\d{4})\b/);
if(m){const mes1=MESES_ES[m[1]]||MESES_ABR[m[1]];if(mes1)return{mes:`${m[3]}-${String(mes1).padStart(2,'0')}`,label:`${m[1]}-${m[2]} ${m[3]}`};}
m=d.match(/\b([A-ZÑ]{3,12})\s+(\d{4})\b/);
if(m){const mes=MESES_ES[m[1]]||MESES_ABR[m[1]];if(mes)return{mes:`${m[2]}-${String(mes).padStart(2,'0')}`,label:`${m[1]} ${m[2]}`};}
return null;
}
let taskPeriodMap=null;
async function cargarJobTaskLines(){
if(taskPeriodMap) return taskPeriodMap;
const rows=await bcGetAll(`${ODATA}/Job_Task_Lines`);
taskPeriodMap={};
rows.forEach(t=>{const p=parsearPeriodoTarea(t.Description);if(p)taskPeriodMap[`${t.Job_No}||${t.Job_Task_No}`]=p;});
return taskPeriodMap;
}
// Etiqueta de período legible para un registro de horas: prioridad Job_Task_Lines (más confiable),
// después la propia Description del registro, y si nada matchea, el código de tarea tal cual
// (proyectos de alcance fijo, como los que usan "TIPO 3", no tienen período mensual — no hay
// forma de inventarlo, así que se muestra lo que hay).
function labelPeriodo(jobNo,taskNo,descRegistro,mapa){
const fromTask=mapa[`${jobNo}||${taskNo}`];
if(fromTask) return fromTask.label;
const fromDesc=parsearPeriodoTarea(descRegistro);
if(fromDesc) return fromDesc.label;
return descRegistro || taskNo || '—';
}

const TOOLS={
horas_por_empleado:async({year='2026',top=10})=>{const[rows,m]=await Promise.all([queryMovs(year,'usage'),empMap()]);const b={};rows.forEach(r=>{const c=r.Global_Dimension_2_Code||'?';if(!b[c])b[c]={codigo:c,nombre:m[c]||c,horas:0,registros:0};b[c].horas+=parseFloat(r.Quantity||0);b[c].registros++;});return{año:year,total_registros:rows.length,ranking:Object.values(b).sort((a,x)=>x.horas-a.horas).slice(0,top)};},
horas_por_proyecto:async({year='2026',top=10})=>{const rows=await queryMovs(year,'usage');const b={};rows.forEach(r=>{const j=r.Job_No||'?';if(!b[j])b[j]={proyecto:j,horas:0};b[j].horas+=parseFloat(r.Quantity||0);});return{año:year,ranking:Object.values(b).sort((a,x)=>x.horas-a.horas).slice(0,top)};},
detalle_empleado:async({empleado_codigo,year='2026'})=>{if(!empleado_codigo)throw new Error('Se requiere empleado_codigo');const[all,m,mapaPeriodos]=await Promise.all([queryMovs(year),empMap(),cargarJobTaskLines()]);const u=all.filter(r=>{const et=String(r.Entry_Type||'').toLowerCase();return(et.includes('usage')||r.Entry_Type===1)&&r.Global_Dimension_2_Code===empleado_codigo;});const sm={};all.filter(r=>{const et=String(r.Entry_Type||'').toLowerCase();return et.includes('sale')||r.Entry_Type===0;}).forEach(s=>{sm[`${s.Job_No}||${s.Job_Task_No}`]=s;});return{empleado:{codigo:empleado_codigo,nombre:m[empleado_codigo]||empleado_codigo},año:year,total_horas:parseFloat(u.reduce((s,r)=>s+parseFloat(r.Quantity||0),0).toFixed(2)),proyectos:[...new Set(u.map(r=>r.Job_No))],facturas:[...new Set(u.map(r=>sm[`${r.Job_No}||${r.Job_Task_No}`]?.Document_No).filter(Boolean))],registros:u.map(x=>({fecha:x.Posting_Date,proyecto:x.Job_No,tarea:x.Job_Task_No,periodo:labelPeriodo(x.Job_No,x.Job_Task_No,x.Description,mapaPeriodos),descripcion:x.Description,horas:parseFloat(x.Quantity||0),factura:sm[`${x.Job_No}||${x.Job_Task_No}`]?.Document_No||'—'}))};},
facturas_por_empleado:async({year='2026',empleado_codigo})=>{const[all,m]=await Promise.all([queryMovs(year),empMap()]);const u=all.filter(r=>{const et=String(r.Entry_Type||'').toLowerCase();return(et.includes('usage')||r.Entry_Type===1)&&(!empleado_codigo||r.Global_Dimension_2_Code===empleado_codigo);});const sm={};all.filter(r=>{const et=String(r.Entry_Type||'').toLowerCase();return et.includes('sale')||r.Entry_Type===0;}).forEach(s=>{sm[`${s.Job_No}||${s.Job_Task_No}`]=s;});const b={};u.forEach(x=>{const c=x.Global_Dimension_2_Code||'?';const s=sm[`${x.Job_No}||${x.Job_Task_No}`];if(!b[c])b[c]={codigo:c,nombre:m[c]||c,horas:0,facturas:new Set()};b[c].horas+=parseFloat(x.Quantity||0);if(s?.Document_No)b[c].facturas.add(s.Document_No);});return Object.values(b).map(e=>({codigo:e.codigo,nombre:e.nombre,horas:parseFloat(e.horas.toFixed(2)),facturas:[...e.facturas],num_facturas:e.facturas.size})).sort((a,x)=>x.horas-a.horas);},
facturacion_por_empleado:async({empleado_codigo,year='2026'})=>{const empleados=await facturacionPorEmpleadoMensual(empleado_codigo,year);return{año:year,nota:'Facturación en MXN = Total_Price propio de cada línea de horas del empleado en MovsProyecto (no un promedio prorrateado — cada empleado puede tener su propio valor/hora). Incluye valor_hora calculado. Si no se pasa empleado_codigo, devuelve todos.',empleados};},
resumen_proyectos:async({year='2026'})=>{const[all,m]=await Promise.all([queryMovs(year),empMap()]);const u=all.filter(r=>{const et=String(r.Entry_Type||'').toLowerCase();return et.includes('usage')||r.Entry_Type===1;});const s=all.filter(r=>{const et=String(r.Entry_Type||'').toLowerCase();return et.includes('sale')||r.Entry_Type===0;});return{año:year,total_horas:parseFloat(u.reduce((a,r)=>a+parseFloat(r.Quantity||0),0).toFixed(2)),total_facturado_mxn:parseFloat(s.reduce((a,r)=>a+Math.abs(parseFloat(r.Total_Price||0)),0).toFixed(2)),empleados_activos:new Set(u.map(r=>r.Global_Dimension_2_Code).filter(Boolean)).size,proyectos_activos:new Set(u.map(r=>r.Job_No).filter(Boolean)).size,facturas_generadas:new Set(s.map(r=>r.Document_No).filter(Boolean)).size};},
buscar_empleado:async({nombre})=>{const m=await empMap();const q=(nombre||'').toUpperCase();return{resultados:Object.entries(m).filter(([c,n])=>n.toUpperCase().includes(q)||c.toUpperCase().includes(q)).map(([codigo,nombre])=>({codigo,nombre}))};},
facturas_venta:async({top=20,filter='',year})=>{let f=filter;if(year&&!f)f=`invoiceDate ge ${year}-01-01 and invoiceDate le ${year}-12-31`;const rows=await v2Get('salesInvoices',`?$top=${top}${f?'&$filter='+encodeURIComponent(f):''}&$orderby=invoiceDate desc`);avisarSiTopCompleto('facturas_venta',rows,top);return{total:rows.length,facturas:rows.map(r=>({numero:r.number,cliente:r.customerName,fecha:r.invoiceDate,vencimiento:r.dueDate,total:r.totalAmountIncludingTax,moneda:r.currencyCode,estado:r.status}))};},
resumen_ventas:async({year='2026'})=>{const rows=await v2Get('salesInvoices',`?$filter=${encodeURIComponent(`invoiceDate ge ${year}-01-01 and invoiceDate le ${year}-12-31`)}`);const total=rows.reduce((s,r)=>s+(r.totalAmountIncludingTax||0),0);const pc={};rows.forEach(r=>{const c=r.customerName||'?';if(!pc[c])pc[c]={cliente:c,total:0,facturas:0};pc[c].total+=r.totalAmountIncludingTax||0;pc[c].facturas++;});return{año:year,total_facturas:rows.length,total_mxn:parseFloat(total.toFixed(2)),pagadas:rows.filter(r=>r.status==='Paid').length,abiertas:rows.filter(r=>r.status==='Open').length,top_clientes:Object.values(pc).sort((a,b)=>b.total-a.total).slice(0,10)};},
notas_credito_venta:async({top=20,year})=>{let f=year?`postingDate ge ${year}-01-01 and postingDate le ${year}-12-31`:'';const rows=await v2Get('salesCreditMemos',`?$top=${top}${f?'&$filter='+encodeURIComponent(f):''}&$orderby=postingDate desc`);avisarSiTopCompleto('notas_credito_venta',rows,top);return{total:rows.length,notas:rows.map(r=>({numero:r.number,cliente:r.customerName,fecha:r.postingDate,total:r.totalAmountIncludingTax,estado:r.status}))};},
clientes:async({top=50,filter=''})=>{const rows=await v2Get('customers',`?$top=${top}${filter?'&$filter='+encodeURIComponent(filter):''}`);avisarSiTopCompleto('clientes',rows,top);return{total:rows.length,clientes:rows.map(r=>({id:r.number,nombre:r.displayName,email:r.email,telefono:r.phoneNumber,ciudad:r.city,pais:r.country,saldo:r.balance,limite_credito:r.creditLimit,terminos_pago:r.paymentTermsId}))};},
buscar_cliente:async({nombre})=>{const rows=await v2Get('customers',`?$filter=${encodeURIComponent(`contains(displayName,'${nombre}')`)}&$top=10`);return{resultados:rows.map(r=>({id:r.number,nombre:r.displayName,email:r.email,saldo:r.balance}))};},
facturas_compra:async({top=20,year})=>{let f=year?`invoiceDate ge ${year}-01-01 and invoiceDate le ${year}-12-31`:'';const rows=await v2Get('purchaseInvoices',`?$top=${top}${f?'&$filter='+encodeURIComponent(f):''}&$orderby=invoiceDate desc`);avisarSiTopCompleto('facturas_compra',rows,top);return{total:rows.length,facturas:rows.map(r=>({numero:r.number,proveedor:r.vendorName,fecha:r.invoiceDate,total:r.totalAmountIncludingTax,estado:r.status}))};},
resumen_compras:async({year='2026'})=>{const rows=await v2Get('purchaseInvoices',`?$filter=${encodeURIComponent(`invoiceDate ge ${year}-01-01 and invoiceDate le ${year}-12-31`)}`);const total=rows.reduce((s,r)=>s+(r.totalAmountIncludingTax||0),0);const pp={};rows.forEach(r=>{const p=r.vendorName||'?';if(!pp[p])pp[p]={proveedor:p,total:0,facturas:0};pp[p].total+=r.totalAmountIncludingTax||0;pp[p].facturas++;});return{año:year,total_facturas:rows.length,total_mxn:parseFloat(total.toFixed(2)),top_proveedores:Object.values(pp).sort((a,b)=>b.total-a.total).slice(0,10)};},
proveedores:async({top=50,filter=''})=>{const rows=await v2Get('vendors',`?$top=${top}${filter?'&$filter='+encodeURIComponent(filter):''}`);avisarSiTopCompleto('proveedores',rows,top);return{total:rows.length,proveedores:rows.map(r=>({id:r.number,nombre:r.displayName,email:r.email,saldo:r.balance}))};},
asientos_contables:async({top=50,year,cuenta})=>{let f='';if(year)f=`postingDate ge ${year}-01-01 and postingDate le ${year}-12-31`;if(cuenta)f=f?`${f} and accountNumber eq '${cuenta}'`:`accountNumber eq '${cuenta}'`;const rows=await v2Get('generalLedgerEntries',`?$top=${top}${f?'&$filter='+encodeURIComponent(f):''}&$orderby=postingDate desc`);avisarSiTopCompleto('asientos_contables',rows,top);return{total:rows.length,asientos:rows.map(r=>({fecha:r.postingDate,cuenta:r.accountNumber,nombre_cuenta:r.accountName,descripcion:r.description,debe:r.debitAmount,haber:r.creditAmount,importe:r.amount,documento:r.documentNumber}))};},
plan_cuentas:async({filter=''})=>{const rows=await v2Get('accounts',filter?`?$filter=${encodeURIComponent(filter)}`:'');return{total:rows.length,cuentas:rows.map(r=>({numero:r.number,nombre:r.displayName,tipo:r.accountType,bloqueada:r.blocked}))};},
resumen_financiero:async({year='2026'})=>{const[v,c,a]=await Promise.all([v2Get('salesInvoices',`?$filter=${encodeURIComponent(`invoiceDate ge ${year}-01-01 and invoiceDate le ${year}-12-31`)}`),v2Get('purchaseInvoices',`?$filter=${encodeURIComponent(`invoiceDate ge ${year}-01-01 and invoiceDate le ${year}-12-31`)}`),v2Get('generalLedgerEntries',`?$filter=${encodeURIComponent(`postingDate ge ${year}-01-01 and postingDate le ${year}-12-31`)}`)]);return{año:year,ventas:{total_facturas:v.length,total_mxn:parseFloat(v.reduce((s,r)=>s+(r.totalAmountIncludingTax||0),0).toFixed(2)),pagadas:v.filter(r=>r.status==='Paid').length,abiertas:v.filter(r=>r.status==='Open').length},compras:{total_facturas:c.length,total_mxn:parseFloat(c.reduce((s,r)=>s+(r.totalAmountIncludingTax||0),0).toFixed(2))},contabilidad:{total_asientos:a.length,total_debitos:parseFloat(a.reduce((s,r)=>s+(r.debitAmount||0),0).toFixed(2)),total_creditos:parseFloat(a.reduce((s,r)=>s+(r.creditAmount||0),0).toFixed(2))}};},
empleados:async({top=100})=>{const rows=await v2Get('employees',`?$top=${top}`);avisarSiTopCompleto('empleados',rows,top);const m=await empMap();return{total:rows.length>0?rows.length:Object.keys(m).length,empleados:rows.length>0?rows.map(r=>({id:r.number||r.id,nombre:r.displayName||`${r.givenName||''} ${r.surname||''}`.trim(),email:r.email,puesto:r.jobTitle,departamento:r.department,estado:r.status})):Object.entries(m).map(([codigo,nombre])=>({id:codigo,nombre}))};},
items:async({top=50,filter=''})=>{const rows=await v2Get('items',`?$top=${top}${filter?'&$filter='+encodeURIComponent(filter):''}`);avisarSiTopCompleto('items',rows,top);return{total:rows.length,items:rows.map(r=>({id:r.number,nombre:r.displayName,tipo:r.type,inventario:r.inventory,precio_venta:r.unitPrice,costo:r.unitCost}))};},
cuentas_bancarias:async()=>{const rows=await v2Get('bankAccounts');return{total:rows.length,cuentas:rows.map(r=>({id:r.number,nombre:r.displayName,moneda:r.currencyId}))};},
dimensiones:async()=>{const rows=await v2Get('dimensions');return{dimensiones:rows.map(r=>({id:r.id,codigo:r.code,nombre:r.displayName}))};},
valores_dimension:async({dimension_codigo})=>{if(!dimension_codigo)throw new Error('Se requiere dimension_codigo');const dims=await v2Get('dimensions');const dim=dims.find(d=>(d.code||'').toUpperCase()===dimension_codigo.toUpperCase());if(!dim)throw new Error(`Dimensión ${dimension_codigo} no encontrada`);const id=await coId();const vals=await bcGetAll(`${V2}/companies(${id})/dimensions(${dim.id})/dimensionValues`);return{dimension:{codigo:dim.code,nombre:dim.displayName},valores:vals.map(v=>({codigo:v.code,nombre:v.displayName||v.name}))};},
contactos:async({top=50,filter=''})=>{const rows=await v2Get('contacts',`?$top=${top}${filter?'&$filter='+encodeURIComponent(filter):''}`);avisarSiTopCompleto('contactos',rows,top);return{total:rows.length,contactos:rows.map(r=>({id:r.number,nombre:r.displayName,tipo:r.type,email:r.email,empresa:r.companyName,ciudad:r.city}))};},
monedas:async()=>{const rows=await v2Get('currencies');return{monedas:rows.map(r=>({codigo:r.code,nombre:r.displayName,simbolo:r.symbol}))};},
terminos_pago:async()=>{const rows=await v2Get('paymentTerms');return{terminos:rows.map(r=>({codigo:r.code,nombre:r.displayName,dias:r.dueDateCalculation}))};},
costo_directo_por_empleado:async({empleado_codigo,year='2026'})=>{const empleados=await costoPorEmpleadoGenerico(empleado_codigo,year,esCuentaCostoDirecto);return{año:year,nota:'Costo directo = débitos en cuentas 511/512/513/514/520, dimensionado por RECURSO. Incluye la maquila (511007) cuando está tageada a un empleado (no a un socio) — ese caso corresponde al Adicional Nómina que se paga por esa vía. Positivo = costo.',empleados};},
nomina_por_empleado:async({empleado_codigo,year='2026'})=>{const empleados=await nominaConAdicionalPorEmpleado(empleado_codigo,year);return{año:year,nota:'Nómina de cada empleado = nomina_legal (cuentas de sueldo/prestaciones) + adicional_nomina (parte de la maquila 511007 tageada a ese empleado, no a un socio). Se devuelven ambas partes por separado y el total. Positivo = costo.',empleados};},
};

const MCP_TOOLS=[
{name:'horas_por_empleado',description:'Ranking de empleados por horas facturadas en proyectos.',inputSchema:{type:'object',properties:{year:{type:'string',default:'2026'},top:{type:'number',default:10}}}},
{name:'horas_por_proyecto',description:'Ranking de proyectos por horas registradas.',inputSchema:{type:'object',properties:{year:{type:'string',default:'2026'},top:{type:'number',default:10}}}},
{name:'detalle_empleado',description:'Horas, proyectos y facturas detalladas de un empleado específico, incluyendo el período real de cada tarea (resuelto vía Job_Task_Lines cuando existe). Requiere código de empleado.',inputSchema:{type:'object',required:['empleado_codigo'],properties:{empleado_codigo:{type:'string'},year:{type:'string',default:'2026'}}}},
{name:'facturas_por_empleado',description:'Qué facturas de venta generaron las horas de cada empleado.',inputSchema:{type:'object',properties:{year:{type:'string'},empleado_codigo:{type:'string'}}}},
{name:'facturacion_por_empleado',description:'Facturación en MXN por empleado, mes a mes, prorrateada por horas trabajadas en cada tarea. Si no se pasa empleado_codigo, devuelve todos.',inputSchema:{type:'object',properties:{empleado_codigo:{type:'string'},year:{type:'string',default:'2026'}}}},
{name:'resumen_proyectos',description:'Resumen ejecutivo de proyectos: horas totales, facturación, empleados y proyectos activos.',inputSchema:{type:'object',properties:{year:{type:'string',default:'2026'}}}},
{name:'buscar_empleado',description:'Busca un empleado por nombre para obtener su código. Útil antes de llamar detalle_empleado.',inputSchema:{type:'object',required:['nombre'],properties:{nombre:{type:'string'}}}},
{name:'facturas_venta',description:'Lista facturas de venta. Filtrar por año.',inputSchema:{type:'object',properties:{top:{type:'number',default:20},year:{type:'string'},filter:{type:'string'}}}},
{name:'resumen_ventas',description:'Resumen anual de ventas: total MXN, pagadas, abiertas y top 10 clientes.',inputSchema:{type:'object',properties:{year:{type:'string',default:'2026'}}}},
{name:'notas_credito_venta',description:'Notas de crédito de venta emitidas.',inputSchema:{type:'object',properties:{top:{type:'number',default:20},year:{type:'string'}}}},
{name:'clientes',description:'Lista de clientes con saldo, límite de crédito y datos de contacto.',inputSchema:{type:'object',properties:{top:{type:'number',default:50},filter:{type:'string'}}}},
{name:'buscar_cliente',description:'Busca un cliente por nombre.',inputSchema:{type:'object',required:['nombre'],properties:{nombre:{type:'string'}}}},
{name:'facturas_compra',description:'Lista facturas de compra por año.',inputSchema:{type:'object',properties:{top:{type:'number',default:20},year:{type:'string'}}}},
{name:'resumen_compras',description:'Resumen anual de compras y top proveedores.',inputSchema:{type:'object',properties:{year:{type:'string',default:'2026'}}}},
{name:'proveedores',description:'Lista de proveedores con saldos.',inputSchema:{type:'object',properties:{top:{type:'number',default:50},filter:{type:'string'}}}},
{name:'asientos_contables',description:'Entradas del libro mayor general. Filtrar por año y/o número de cuenta.',inputSchema:{type:'object',properties:{top:{type:'number',default:50},year:{type:'string'},cuenta:{type:'string'}}}},
{name:'plan_cuentas',description:'Plan de cuentas contables completo.',inputSchema:{type:'object',properties:{filter:{type:'string'}}}},
{name:'resumen_financiero',description:'Resumen financiero completo del año: ventas, compras y contabilidad.',inputSchema:{type:'object',properties:{year:{type:'string',default:'2026'}}}},
{name:'empleados',description:'Lista de empleados con datos básicos (nombre, puesto, departamento). Sin datos salariales.',inputSchema:{type:'object',properties:{top:{type:'number',default:100}}}},
{name:'items',description:'Catálogo de artículos y servicios con precios e inventario.',inputSchema:{type:'object',properties:{top:{type:'number',default:50},filter:{type:'string'}}}},
{name:'cuentas_bancarias',description:'Cuentas bancarias de la empresa.',inputSchema:{type:'object',properties:{}}},
{name:'dimensiones',description:'Dimensiones contables disponibles (RECURSO, AREA, CONTRATO, etc.).',inputSchema:{type:'object',properties:{}}},
{name:'valores_dimension',description:'Todos los valores de una dimensión específica.',inputSchema:{type:'object',required:['dimension_codigo'],properties:{dimension_codigo:{type:'string'}}}},
{name:'contactos',description:'Contactos de la empresa.',inputSchema:{type:'object',properties:{top:{type:'number',default:50},filter:{type:'string'}}}},
{name:'monedas',description:'Monedas configuradas en BC.',inputSchema:{type:'object',properties:{}}},
{name:'terminos_pago',description:'Términos de pago disponibles.',inputSchema:{type:'object',properties:{}}},
{name:'costo_directo_por_empleado',description:'Costo directo de servicio por empleado, mes a mes (cuentas 511-520 excluyendo maquila, dimensionado por RECURSO). Si no se pasa empleado_codigo, devuelve todos.',inputSchema:{type:'object',properties:{empleado_codigo:{type:'string'},year:{type:'string',default:'2026'}}}},
{name:'nomina_por_empleado',description:'Costo de nómina por empleado, mes a mes, separando nómina legal (sueldo/prestaciones) de Adicional Nómina (pagos vía cuenta de maquila tageados a ese empleado). Si no se pasa empleado_codigo, devuelve todos.',inputSchema:{type:'object',properties:{empleado_codigo:{type:'string'},year:{type:'string',default:'2026'}}}},
];

app.get('/ping',(req,res)=>res.json({status:'ok',message:`BC MCP v2.0 ✓ | ${MCP_TOOLS.length} herramientas`}));

// Dashboards: requireApiKey (como ya tenías) + requireSession (login por email+código, NUEVO)
app.post('/token',auth.requireSession,async(req,res)=>{try{res.json({access_token:await getToken(),expires_in:3600});}catch(e){res.status(500).json({error:e.message});}});
app.get('/bc',auth.requireSession,async(req,res)=>{const{token,url}=req.query;if(!token||!url)return res.status(400).json({error:'Faltan parámetros'});try{const r=await fetch(decodeURIComponent(url),{headers:{Authorization:`Bearer ${token}`,Accept:'application/json'}});const d=await r.json();if(!r.ok)return res.status(r.status).json(d);res.json(d);}catch(e){res.status(500).json({error:e.message});}});

// Login por email+código para los dashboards (NUEVO)
app.post('/auth/request-code',auth.handleRequestCode);
app.post('/auth/verify-code',auth.handleVerifyCode);

// OAuth para el conector MCP: cada persona se loguea con su propio email (NUEVO)
app.get('/authorize',oauth.handleAuthorizePage);
app.post('/oauth/request-code',oauth.handleOAuthRequestCode);
app.post('/oauth/verify-code',oauth.handleOAuthVerifyCode);
app.post('/oauth/token',oauth.handleOAuthToken);
app.get('/.well-known/oauth-authorization-server',oauth.handleAuthServerMetadata);
app.get('/.well-known/oauth-protected-resource',oauth.handleProtectedResourceMetadata);

// Antes SIN protección alguna — ahora exige el Bearer token que entrega el flujo OAuth de arriba (NUEVO)
app.get('/mcp/sse',oauth.requireBearer,(req,res)=>{res.setHeader('Content-Type','text/event-stream');res.setHeader('Cache-Control','no-cache');res.setHeader('Connection','keep-alive');res.setHeader('X-Accel-Buffering','no');const id=Date.now().toString();res.write(`event: endpoint\ndata: /mcp/messages?sessionId=${id}\n\n`);const ka=setInterval(()=>res.write(': ping\n\n'),15000);req.on('close',()=>clearInterval(ka));app.locals.sessions=app.locals.sessions||{};app.locals.sessions[id]=res;});

app.post('/mcp/messages',oauth.requireBearer,async(req,res)=>{const{sessionId}=req.query;const msg=req.body;const sse=(app.locals.sessions||{})[sessionId];const send=(d)=>{if(sse){sse.write(`event: message\ndata: ${JSON.stringify(d)}\n\n`);res.status(200).json({});}else{res.json(d);}};try{if(msg.method==='initialize'){send({jsonrpc:'2.0',id:msg.id,result:{protocolVersion:'2024-11-05',serverInfo:{name:'bc-tesis',version:'2.0.0'},capabilities:{tools:{}}}});}else if(msg.method==='tools/list'){send({jsonrpc:'2.0',id:msg.id,result:{tools:MCP_TOOLS}});}else if(msg.method==='tools/call'){const{name,arguments:args}=msg.params;if(!TOOLS[name])throw new Error(`Tool desconocida: ${name}`);const store={warns:[]};const result=await als.run(store,()=>TOOLS[name](args||{}));if(store.warns.length&&result&&typeof result==='object'&&!Array.isArray(result))result.advertencias=[...new Set(store.warns)];send({jsonrpc:'2.0',id:msg.id,result:{content:[{type:'text',text:JSON.stringify(result,null,2)}]}});}else if(msg.method==='notifications/initialized'){res.status(200).json({});}else{send({jsonrpc:'2.0',id:msg.id,result:{}});}}catch(e){send({jsonrpc:'2.0',id:msg?.id,error:{code:-32000,message:e.message}});}});

app.listen(PORT,()=>console.log(`BC MCP v2.0 en puerto ${PORT} | ${MCP_TOOLS.length} herramientas`));

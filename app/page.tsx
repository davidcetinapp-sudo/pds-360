'use client';
export const dynamic = 'force-dynamic';
// Powerchina PDS 360 v2.2
import { useCallback, useEffect, useMemo, useRef, useState, type ReactNode } from 'react';
import { supabase, type Profile, type UserRole } from '@/lib/supabase';
import { ResponsiveContainer, RadialBarChart, RadialBar, PolarAngleAxis, LineChart, Line, XAxis, YAxis, CartesianGrid, Tooltip, BarChart, Bar, PieChart, Pie, Cell, Legend } from 'recharts';
import * as XLSX from 'xlsx';

// ── TIPOS ─────────────────────────────────────────────────────────
type AppView = 'home'|'planear'|'reporte'|'aprobacion'|'solicitudes'|'dashboard'|'informes'|'catalogos'|'maquinaria'|'config_act'|'usuarios';

interface EspAct  { id:string; especialidad_es:string; especialidad_en:string; actividad_es:string; actividad_en:string; activo?:boolean; }
interface Area     { id:string; area_es:string; area_en:string; activo?:boolean; }
interface Lider    { id:string; nombre:string; documento:string; cargo_es:string; activo?:boolean; }
interface Personal { id:string; nombre:string; documento:string; cargo_es:string; tipo?:string; empresa?:string; activo?:boolean; }
interface Maq      { id:string; tipo:string; item_id:string; nombre:string; estado:string; horas_acum_operativas?:number; horas_acum_standby?:number; }
interface ConfigAct{ id:string; actividad_id:string; especialidad_id:string; tipo:string; unidad_es:string; unidad_en?:string; meta_total?:number; tiene_meta?:boolean; es_medible?:boolean; rendimiento_esperado?:number; rendimiento_por?:string; acumulado_previo?:number; tiene_items_unicos?:boolean; }
interface ItemDatabase { id:string; config_actividad_id:string; nombre:string; columnas:{nombre:string;tipo:string}[]; bloqueo_tipo:'permanente'|'temporal'; activo:boolean; }
interface ItemDB { id:string; database_id:string; datos:Record<string,string>; bloqueado:boolean; bloqueado_fecha:string|null; bloqueado_en_reporte:string|null; }
interface Catalogs { especialidades_actividades:EspAct[]; areas:Area[]; lideres:Lider[]; personal:Personal[]; }
interface Notif    { id:string; titulo:string; mensaje:string; leida:boolean; created_at:string; }
interface SuspItem { uid:string; tipo_susp:string; otro_desc:string; hora_inicio:string; hora_fin:string; descripcion:string; }
interface SuspActividad { uid:string; actividad_id:string; tipo:string; otro_desc:string; parcial:boolean; hora_inicio:string; hora_fin:string; observacion:string; }
interface AsistItem{ personal_id:string; documento_personal:string; nombre:string; cargo_es:string; asistio:boolean; motivo_ausencia:string; ausencia_parcial:boolean; hora_ausencia_ini:string; hora_ausencia_fin:string; jornada_parcial:boolean; hora_jornada_ini:string; hora_jornada_fin:string; es_adicional?:boolean; }
interface AreaRep  { uid:string; area_id:string; cantidad:string; }
interface ActRep   { uid:string; actividad_id:string; areas:AreaRep[]; descripcion_cualitativa:string; observacion_es:string; items_seleccionados:string[]; }
interface ActAdicCat { id:string; nombre:string; veces_usada:number; }
interface ActAdicItem { uid:string; nombre:string; descripcion:string; catalogoId:string|null; }
interface PersSel  { personal_id:string; documento_personal:string; }
interface ActForm  {
  uid:string; especialidad_id:string; actividad_id:string;
  area_id:string; areas_adicionales:string[];
  lider_id:string; maquinaria_ids:string[];
  rendimiento_esperado:string; observacion_es:string; observacion_en:string;
  personal:PersSel[];
}

// ── UTILS ─────────────────────────────────────────────────────────
function today():string {
  const d = new Date();
  return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
}
function gid():string { return Math.random().toString(36).slice(2)+Date.now().toString(36); }
function sem(p:number):string { return p>=70?'🟢':p>=50?'🟡':'🔴'; }

function uniqueEsp(rows:EspAct[]):EspAct[] {
  const s = new Set<string>();
  return rows.filter(r=>{ const k=(r.especialidad_es||'').toLowerCase(); if(s.has(k))return false; s.add(k); return true; });
}
function actsForEsp(rows:EspAct[], espId:string):EspAct[] {
  const e = rows.find(r=>r.id===espId);
  if(!e) return [];
  const t = (e.especialidad_es||'').toLowerCase();
  return rows.filter(r=>(r.especialidad_es||'').toLowerCase()===t);
}
function resumenPorActividad(avances:Record<string,unknown>[], configActs:ConfigAct[], catalogs:Catalogs|null){
  const porAct:Record<string,{nombre:string;unidad:string;meta:number|null;totalRango:number;acumuladoHistorico:number;pct:number|null}> = {};
  avances.forEach(av=>{
    const id=av.actividad_id as string;
    const aR=catalogs?.especialidades_actividades.find(e=>e.id===id);
    const cfg=configActs.find(c=>c.actividad_id===id);
    if(!porAct[id]) porAct[id]={nombre:aR?.actividad_es||id,unidad:cfg?.unidad_es||'',meta:cfg?.meta_total||null,totalRango:0,acumuladoHistorico:0,pct:null};
    porAct[id].totalRango+=parseFloat(String(av.cantidad||0));
    porAct[id].acumuladoHistorico=Math.max(porAct[id].acumuladoHistorico,parseFloat(String(av.acumulado_total_real||0)));
  });
  Object.values(porAct).forEach(a=>{if(a.meta)a.pct=Math.min(100,Math.round(a.acumuladoHistorico/a.meta*100));});
  return porAct;
}
function horasLost(ss:SuspItem[]):number {
  return ss.reduce((a,s)=>{
    if(!s.hora_inicio||!s.hora_fin) return a;
    const[ih,im]=s.hora_inicio.split(':').map(Number);
    const[fh,fm]=s.hora_fin.split(':').map(Number);
    return a+Math.max(0,(fh+fm/60)-(ih+im/60));
  },0);
}
function calcHoras(ini:string,fin:string):number{
  const[h1,m1]=ini.split(':').map(Number);
  const[h2,m2]=fin.split(':').map(Number);
  return Math.max(0,(h2*60+m2-h1*60-m1)/60);
}
function emptyAct():ActForm {
  return { uid:gid(), especialidad_id:'', actividad_id:'', area_id:'', areas_adicionales:[], lider_id:'', maquinaria_ids:[], rendimiento_esperado:'', observacion_es:'', observacion_en:'', personal:[] };
}

// ── APP ROOT ──────────────────────────────────────────────────────
export default function App() {
  const[profile,setProfile]=useState<Profile|null>(null);
  const[view,setView]=useState<AppView>('home');
  const[catalogs,setCatalogs]=useState<Catalogs|null>(null);
  const[maquinaria,setMaq]=useState<Maq[]>([]);
  const[configActs,setCA]=useState<ConfigAct[]>([]);
  const[notifs,setNotifs]=useState<Notif[]>([]);
  const[loading,setLoading]=useState(true);
  const[toast,setToast]=useState<{k:'ok'|'err'|'info';m:string}|null>(null);

  useEffect(()=>{
    (async()=>{
      try {
        const{data:{session}}=await supabase.auth.getSession();
        if(session?.user){
          const{data:p}=await supabase.from('profiles').select('*').eq('id',session.user.id).single();
          if(p) setProfile(p as Profile);
        }
      } catch(e){ console.error('Auth:',e); }
      finally{ setLoading(false); }
    })();
    const{data:{subscription}}=supabase.auth.onAuthStateChange((_,session)=>{
      if(!session){ setProfile(null); setView('home'); }
    });
    return()=>subscription.unsubscribe();
  },[]);

  useEffect(()=>{
    if(toast){ const t=setTimeout(()=>setToast(null),4500); return()=>clearTimeout(t); }
  },[toast]);

  const showToast=useCallback((k:'ok'|'err'|'info',m:string)=>setToast({k,m}),[]);

  const loadCatalogs=useCallback(async()=>{
    try{
      const[ea,ar,li,pe,ma,ca]=await Promise.all([
        supabase.from('especialidades_actividades').select('*').order('especialidad_es'),
        supabase.from('areas').select('*').order('area_es'),
        supabase.from('lideres').select('*').order('nombre'),
        supabase.from('personal').select('*').order('nombre'),
        supabase.from('maquinaria').select('*').order('item_id'),
        supabase.from('config_actividades').select('*'),
      ]);
      setCatalogs({ especialidades_actividades:(ea.data||[]) as EspAct[], areas:(ar.data||[]) as Area[], lideres:(li.data||[]) as Lider[], personal:(pe.data||[]) as Personal[] });
      setMaq((ma.data||[]) as Maq[]);
      setCA((ca.data||[]) as ConfigAct[]);
    } catch(e){ console.error('Catálogos:',e); }
  },[]);

  const loadNotifs=useCallback(async(uid:string)=>{
    try{
      const{data}=await supabase.from('notificaciones').select('*').eq('usuario_id',uid).eq('leida',false).order('created_at',{ascending:false}).limit(20);
      setNotifs((data||[]) as Notif[]);
    } catch{ setNotifs([]); }
  },[]);

  useEffect(()=>{ if(profile){ loadCatalogs(); loadNotifs(profile.id); } },[profile,loadCatalogs,loadNotifs]);

  async function marcarLeidas(){
    if(!profile) return;
    try{ await supabase.from('notificaciones').update({leida:true}).eq('usuario_id',profile.id).eq('leida',false); } catch{}
    setNotifs([]);
  }

  async function handleLogout(){
    await supabase.auth.signOut();
    setProfile(null); setCatalogs(null); setView('home');
  }

  if(loading) return(
    <div className="min-h-screen flex flex-col items-center justify-center gap-3 text-slate-500">
      <div className="w-8 h-8 border-4 border-[#003b7a] border-t-transparent rounded-full animate-spin"/>
      <span className="text-sm">Iniciando PDS 360…</span>
    </div>
  );
  if(!profile) return <LoginScreen onLogin={setProfile} showToast={showToast} toast={toast}/>;

  const u=profile;
  const canEdit=['admin','lider','tecnico'].includes(u.rol);
  const canPlanear=['admin','lider','tecnico','gerencia','cliente'].includes(u.rol);
  const canReporte=['admin','lider','tecnico','gerencia'].includes(u.rol);
  const canApprove=['admin','lider'].includes(u.rol);
  const canViewReports=['admin','lider','tecnico','gerencia','cliente'].includes(u.rol);
  const canViewDashboard=['admin','lider','tecnico','gerencia','cliente','visualizador'].includes(u.rol);

  function cambiarVista(v:AppView){
    if(u.rol==='visualizador'&&v!=='dashboard'&&v!=='home'){
      showToast('info','Tu rol solo permite ver el Dashboard');return;
    }
    setView(v);
  }

  return(
    <div className="min-h-screen flex flex-col">
      <Header user={u} onLogout={handleLogout} setView={cambiarVista} currentView={view} notifs={notifs} onReadNotifs={marcarLeidas}/>
      <main className="flex-1 p-3 sm:p-5 max-w-7xl mx-auto w-full">
        {view==='home'       &&<HomeScreen user={u} setView={cambiarVista} notifs={notifs}/>}
        {view==='planear'    &&canPlanear&&<PlaneacionModule user={u} catalogs={catalogs} maquinaria={maquinaria} showToast={showToast} readOnly={u.rol==='cliente'}/>}
        {view==='reporte'    &&canReporte&&<ReporteModule user={u} catalogs={catalogs} maquinaria={maquinaria} configActs={configActs} showToast={showToast}/>}
        {view==='aprobacion' &&canApprove&&<AprobacionModule user={u} catalogs={catalogs} configActs={configActs} showToast={showToast} onRefreshNotifs={()=>loadNotifs(u.id)}/>}
        {view==='solicitudes'&&canApprove&&<SolicitudesModule user={u} catalogs={catalogs} showToast={showToast}/>}
        {view==='dashboard'  &&canViewDashboard&&<DashboardModule catalogs={catalogs} configActs={configActs} showToast={showToast}/>}
        {view==='informes'   &&canViewReports&&<InformesModule user={u} catalogs={catalogs} configActs={configActs} maquinaria={maquinaria} showToast={showToast}/>}
        {view==='catalogos'  &&u.rol==='admin'&&<CatalogosModule catalogs={catalogs} onRefresh={loadCatalogs} showToast={showToast}/>}
        {view==='maquinaria' &&u.rol==='admin'&&<MaquinariaModule maquinaria={maquinaria} onRefresh={loadCatalogs} showToast={showToast}/>}
        {view==='config_act' &&u.rol==='admin'&&<ConfigActModule user={u} configActs={configActs} catalogs={catalogs} onRefresh={loadCatalogs} showToast={showToast}/>}
        {view==='usuarios'   &&u.rol==='admin'&&<UsuariosModule showToast={showToast}/>}
      </main>
      <Toast toast={toast}/>
      <footer className="bg-slate-100 border-t text-center text-xs text-slate-500 py-2 no-print">Powerchina · PDS 360 · {new Date().getFullYear()}</footer>
    </div>
  );
}

// ── HEADER ────────────────────────────────────────────────────────
function Header({user,onLogout,setView,currentView,notifs,onReadNotifs}:{
  user:Profile; onLogout:()=>void; setView:(v:AppView)=>void;
  currentView:AppView; notifs:Notif[]; onReadNotifs:()=>void;
}){
  const[showN,setShowN]=useState(false);
  const tabs=[
    {key:'home'      as AppView,label:'Inicio',     labelEn:'Home',       show:true},
    {key:'planear'   as AppView,label:'Planear',    labelEn:'Planning',   show:['admin','lider','tecnico','gerencia','cliente'].includes(user.rol)},
    {key:'reporte'   as AppView,label:'Reporte',    labelEn:'Report',     show:['admin','lider','tecnico','gerencia'].includes(user.rol)},
    {key:'aprobacion'as AppView,label:'Aprobar',    labelEn:'Approve',    show:['admin','lider'].includes(user.rol)},
    {key:'solicitudes'as AppView,label:'Solicitudes',labelEn:'Requests',  show:['admin','lider'].includes(user.rol)},
    {key:'dashboard' as AppView,label:'Dashboard',  labelEn:'Dashboard',  show:['admin','lider','tecnico','gerencia','cliente','visualizador'].includes(user.rol)},
    {key:'informes'  as AppView,label:'Informes',   labelEn:'Reports',    show:['admin','lider','tecnico','gerencia','cliente'].includes(user.rol)},
    {key:'catalogos' as AppView,label:'Catálogos',  labelEn:'Catalogs',   show:user.rol==='admin'},
    {key:'maquinaria'as AppView,label:'Maquinaria', labelEn:'Machinery',  show:user.rol==='admin'},
    {key:'config_act'as AppView,label:'Config. Act.',labelEn:'Act. Config',show:user.rol==='admin'},
    {key:'usuarios'  as AppView,label:'Usuarios',   labelEn:'Users',      show:user.rol==='admin'},
  ];
  return(
    <header className="bg-[#003b7a] text-white shadow no-print">
      <div className="max-w-7xl mx-auto px-4 py-3 flex items-center justify-between gap-2 flex-wrap">
        <div className="flex items-center gap-3">
          <img src="/icons/icon-192.png" alt="PDS" className="h-8 w-auto bg-white rounded p-0.5"/>
          <div><div className="font-bold text-base leading-tight">Powerchina · PDS 360</div><div className="text-xs text-blue-200">Plan · Develop · Succeed</div></div>
        </div>
        <div className="flex items-center gap-3">
          <div className="relative">
            <button onClick={()=>{setShowN(!showN);if(!showN&&notifs.length)onReadNotifs();}} className="relative p-1.5 rounded-full hover:bg-blue-700 transition-colors">
              <span className="text-lg">🔔</span>
              {notifs.length>0&&<span className="absolute -top-0.5 -right-0.5 bg-rose-500 text-white text-xs w-4 h-4 rounded-full flex items-center justify-center font-bold">{notifs.length>9?'9+':notifs.length}</span>}
            </button>
            {showN&&(
              <div className="absolute right-0 top-10 w-80 bg-white rounded-lg shadow-xl border border-slate-200 z-50 max-h-80 overflow-y-auto">
                <div className="p-3 border-b font-semibold text-slate-700 text-sm">Notificaciones ({notifs.length})</div>
                {!notifs.length?<div className="p-4 text-sm text-slate-500 text-center">Sin notificaciones</div>:
                  notifs.map(n=>(
                    <div key={n.id} className="p-3 border-b hover:bg-slate-50">
                      <div className="font-medium text-sm text-slate-800">{n.titulo}</div>
                      <div className="text-xs text-slate-500 mt-0.5">{n.mensaje}</div>
                      <div className="text-xs text-slate-400 mt-1">{new Date(n.created_at).toLocaleString('es-CO')}</div>
                    </div>
                  ))
                }
              </div>
            )}
          </div>
          <div className="hidden sm:block text-right"><div className="font-medium text-sm">{user.nombre}</div><div className="text-xs text-blue-200 uppercase">{user.rol}</div></div>
          <button className="btn-secondary text-xs" onClick={onLogout}>
            <span className="block">Salir</span>
            <span className="block text-blue-300 opacity-70" style={{fontSize:'8px'}}>Log out</span>
          </button>
        </div>
      </div>
      <nav className="max-w-7xl mx-auto px-2 overflow-x-auto">
        <div className="flex gap-0.5">{tabs.filter(t=>t.show).map(t=>(
          <button key={t.key} onClick={()=>setView(t.key)}
            className={`px-3 py-1.5 text-xs font-medium border-b-2 whitespace-nowrap transition-colors flex flex-col items-center ${currentView===t.key?'border-white text-white':'border-transparent text-blue-200 hover:text-white'}`}>
            <span>{t.label}</span>
            <span className="text-blue-300 opacity-70" style={{fontSize:'9px'}}>{t.labelEn}</span>
          </button>
        ))}</div>
      </nav>
    </header>
  );
}

function Toast({toast}:{toast:{k:string;m:string}|null}){
  if(!toast) return null;
  const c=toast.k==='ok'?'bg-emerald-600':toast.k==='err'?'bg-rose-600':'bg-slate-700';
  return <div className={`fixed bottom-4 left-4 right-4 sm:left-auto sm:right-4 sm:max-w-sm ${c} text-white px-4 py-3 rounded-lg shadow-lg z-50 text-sm no-print`}>{toast.m}</div>;
}

// ── LOGIN ─────────────────────────────────────────────────────────
function LoginScreen({onLogin,showToast,toast}:{
  onLogin:(p:Profile)=>void;
  showToast:(k:'ok'|'err'|'info',m:string)=>void;
  toast:{k:string;m:string}|null;
}){
  const[correo,setCorreo]=useState('');
  const[clave,setClave]=useState('');
  const[loading,setLoading]=useState(false);
  async function submit(e:React.FormEvent){
    e.preventDefault();
    if(!correo||!clave){showToast('err','Ingresa correo y clave');return;}
    setLoading(true);
    try{
      const{data,error}=await supabase.auth.signInWithPassword({email:correo,password:clave});
      if(error||!data.user) throw new Error(error?.message||'Credenciales inválidas');
      const{data:p}=await supabase.from('profiles').select('*').eq('id',data.user.id).single();
      if(!p) throw new Error('Perfil no encontrado. Contacta al administrador.');
      if(!(p as Profile).activo) throw new Error('Usuario inactivo. Contacta al administrador.');
      onLogin(p as Profile);
    } catch(e:unknown){ showToast('err',(e as Error)?.message||'Error de red'); }
    finally{ setLoading(false); }
  }
  return(
    <div className="min-h-screen flex items-center justify-center p-4 bg-gradient-to-br from-[#003b7a] to-[#002752]">
      <div className="card p-6 sm:p-8 w-full max-w-md">
        <div className="flex flex-col items-center mb-6">
          <img src="/icons/icon-192.png" alt="PDS" className="h-20 w-auto mb-3"/>
          <h1 className="text-2xl font-bold text-[#003b7a]">Powerchina · PDS 360</h1>
          <p className="text-sm text-slate-500">Plan · Develop · Succeed</p>
        </div>
        <form onSubmit={submit} className="space-y-4">
          <div><label className="label">Correo</label><input className="input" type="email" autoComplete="username" value={correo} onChange={e=>setCorreo(e.target.value)}/></div>
          <div><label className="label">Clave</label><input className="input" type="password" autoComplete="current-password" value={clave} onChange={e=>setClave(e.target.value)}/></div>
          <button className="btn-primary w-full py-3" disabled={loading}>{loading?'Ingresando…':'Ingresar'}</button>
        </form>
      </div>
      <Toast toast={toast}/>
    </div>
  );
}

// ── HOME ──────────────────────────────────────────────────────────
function HomeScreen({user,setView,notifs}:{user:Profile;setView:(v:AppView)=>void;notifs:Notif[]}){
  const canEdit=['admin','lider','tecnico'].includes(user.rol);
  const canApprove=['admin','lider'].includes(user.rol);
  return(
    <div className="space-y-6">
      <div className="text-center py-4">
        <h2 className="text-2xl font-bold text-[#003b7a]">Bienvenido, {user.nombre.split(' ')[0]}</h2>
        <p className="text-slate-500 text-sm mt-1">{today()} · <span className="font-semibold uppercase">{user.rol}</span></p>
      </div>
      {notifs.length>0&&(
        <div className="card p-3 border-l-4 border-amber-400 bg-amber-50">
          <div className="font-semibold text-amber-800 text-sm mb-1">🔔 {notifs.length} notificaciones</div>
          {notifs.slice(0,3).map(n=><div key={n.id} className="text-xs text-amber-700">• {n.titulo}: {n.mensaje}</div>)}
        </div>
      )}
      {canEdit&&(
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 max-w-2xl mx-auto">
          <button onClick={()=>setView('planear')} className="card p-6 text-left hover:shadow-md transition-shadow border-2 border-transparent hover:border-[#003b7a]">
            <div className="text-3xl mb-3">📋</div><div className="font-bold text-[#003b7a] text-lg">Planear actividades</div>
            <p className="text-sm text-slate-500 mt-1">Programa actividades con personal y maquinaria.</p>
          </button>
          <button onClick={()=>setView('reporte')} className="card p-6 text-left hover:shadow-md transition-shadow border-2 border-transparent hover:border-emerald-600">
            <div className="text-3xl mb-3">📊</div><div className="font-bold text-emerald-700 text-lg">Reporte de avance</div>
            <p className="text-sm text-slate-500 mt-1">Registra el avance real del día.</p>
          </button>
        </div>
      )}
      <div className="grid grid-cols-2 sm:grid-cols-4 gap-3 max-w-2xl mx-auto">
        {canApprove&&<QBtn icon="✅" label="Aprobar" onClick={()=>setView('aprobacion')}/>}
        <QBtn icon="📈" label="Dashboard" onClick={()=>setView('dashboard')}/>
        <QBtn icon="📁" label="Informes" onClick={()=>setView('informes')}/>
        <QBtn icon="📝" label="Solicitudes" onClick={()=>setView('solicitudes')}/>
        {user.rol==='admin'&&<QBtn icon="⚙️" label="Catálogos" onClick={()=>setView('catalogos')}/>}
        {user.rol==='admin'&&<QBtn icon="🔧" label="Maquinaria" onClick={()=>setView('maquinaria')}/>}
        {user.rol==='admin'&&<QBtn icon="📐" label="Config. Act." onClick={()=>setView('config_act')}/>}
        {user.rol==='admin'&&<QBtn icon="👤" label="Usuarios" onClick={()=>setView('usuarios')}/>}
      </div>
    </div>
  );
}
function QBtn({icon,label,onClick}:{icon:string;label:string;onClick:()=>void}){
  return <button onClick={onClick} className="card p-3 text-center hover:shadow-md transition-shadow"><div className="text-2xl">{icon}</div><div className="text-xs font-medium text-slate-700 mt-1">{label}</div></button>;
}

// ── SELECTOR MÚLTIPLE DESPLEGABLE (checklist) ──────────────────────
function MultiSelectDropdown<T extends {id:string}>({label,options,selected,onChange,renderRow,placeholder='Todas'}:{
  label:string; options:T[]; selected:string[]; onChange:(v:string[])=>void;
  renderRow:(opt:T)=>ReactNode; placeholder?:string;
}){
  const[open,setOpen]=useState(false);
  const ref=useRef<HTMLDivElement>(null);
  useEffect(()=>{
    function onDoc(e:MouseEvent){ if(ref.current&&!ref.current.contains(e.target as Node)) setOpen(false); }
    document.addEventListener('mousedown',onDoc);
    return ()=>document.removeEventListener('mousedown',onDoc);
  },[]);
  function toggleOne(id:string){ onChange(selected.includes(id)?selected.filter(x=>x!==id):[...selected,id]); }
  return(
    <div className="relative" ref={ref}>
      <label className="label">{label}</label>
      <button type="button" onClick={()=>setOpen(o=>!o)} className="select w-full text-left flex items-center justify-between gap-2">
        <span className="truncate">{selected.length?`${selected.length} seleccionada(s)`:placeholder}</span>
        <span className="text-slate-400 flex-shrink-0">{open?'▲':'▼'}</span>
      </button>
      {open&&(
        <div className="absolute z-20 mt-1 w-full min-w-[260px] bg-white border border-slate-300 rounded-lg shadow-lg max-h-64 overflow-y-auto">
          <div className="flex justify-between px-3 py-2 border-b border-slate-100 sticky top-0 bg-white">
            <button type="button" className="text-xs text-blue-600 hover:underline" onClick={()=>onChange(options.map(o=>o.id))}>Todas</button>
            <button type="button" className="text-xs text-slate-500 hover:underline" onClick={()=>onChange([])}>Ninguna</button>
          </div>
          {options.length===0&&<div className="px-3 py-2 text-xs text-slate-400">Sin opciones</div>}
          {options.map(opt=>(
            <label key={opt.id} className="flex items-center gap-2 px-3 py-2 hover:bg-slate-50 cursor-pointer text-sm border-b border-slate-50 last:border-0">
              <input type="checkbox" checked={selected.includes(opt.id)} onChange={()=>toggleOne(opt.id)}/>
              <span className="flex-1 min-w-0">{renderRow(opt)}</span>
            </label>
          ))}
        </div>
      )}
    </div>
  );
}

// ── PLANEACIÓN ────────────────────────────────────────────────────
function PlaneacionModule({user,catalogs,maquinaria,showToast,readOnly=false}:{
  user:Profile; catalogs:Catalogs|null; maquinaria:Maq[];
  showToast:(k:'ok'|'err'|'info',m:string)=>void;
  readOnly?:boolean;
}){
  const[fecha,setFecha]=useState(today());
  const[actividades,setActividades]=useState<ActForm[]>([emptyAct()]);
  const[blocked,setBlocked]=useState<{documento_personal:string;usuario_nombre:string}[]>([]);
  const[estado,setEstado]=useState<'nuevo'|'borrador'|'enviado'>('nuevo');
  const[saving,setSaving]=useState(false);
  const[progId,setProgId]=useState<string|null>(null);

  // Export por rango
  const[showRange,setShowRange]=useState(false);
  const[rangeIni,setRangeIni]=useState(today());
  const[rangeFin,setRangeFin]=useState(today());
  const[rangeEsps,setRangeEsps]=useState<string[]>([]);
  const[rangeLoading,setRangeLoading]=useState(false);
  const espListRange=useMemo(()=>catalogs?uniqueEsp(catalogs.especialidades_actividades):[], [catalogs]);

  async function fetchRangeData(){
    setRangeLoading(true);
    try{
      let qP=supabase.from('programaciones').select('*').gte('fecha',rangeIni).lte('fecha',rangeFin);
      if(rangeEsps.length) qP=qP.in('especialidad_id',rangeEsps);
      const{data:progs}=await qP;
      if(!progs?.length){showToast('info','Sin planeaciones en ese rango');return null;}
      const progIds=(progs as Record<string,unknown>[]).map(p=>p.id as string);
      const{data:acts}=await supabase.from('actividades_programadas').select('*').in('programacion_id',progIds);
      if(!acts?.length){showToast('info','Sin actividades en ese rango');return null;}
      const actIds=(acts as Record<string,unknown>[]).map(a=>a.id as string);
      const{data:paRaw}=await supabase.from('personal_asignado').select('*').in('actividad_programada_id',actIds);
      const persIds=[...new Set((paRaw||[]).map((p:Record<string,unknown>)=>p.personal_id as string))];
      let persMap:Record<string,{nombre:string;cargo_es:string}>={};
      if(persIds.length){
        const{data:pData}=await supabase.from('personal').select('id,nombre,cargo_es').in('id',persIds);
        (pData||[]).forEach((p:Record<string,unknown>)=>{persMap[p.id as string]={nombre:p.nombre as string,cargo_es:p.cargo_es as string};});
      }
      const pa=(paRaw||[]).map((p:Record<string,unknown>)=>({...p,personal:persMap[p.personal_id as string]||{nombre:p.documento_personal as string,cargo_es:''}}));
      return{progs:progs as Record<string,unknown>[],acts:acts as Record<string,unknown>[],pa:pa as Record<string,unknown>[]};
    } catch(e:unknown){showToast('err',(e as Error)?.message||'Error cargando datos');return null;}
    finally{setRangeLoading(false);}
  }

  async function exportRangeExcel(){
    const d=await fetchRangeData(); if(!d||!catalogs) return;
    const rows:Record<string,unknown>[]=[];
    d.acts.forEach(a=>{
      const prog=d.progs.find(p=>p.id===a.programacion_id as string);
      const espRow=catalogs.especialidades_actividades.find(e=>e.id===a.especialidad_id as string);
      const actRow=catalogs.especialidades_actividades.find(e=>e.id===a.actividad_id as string);
      const areaRow=catalogs.areas.find(ar=>ar.id===a.area_id as string);
      const liderRow=catalogs.lideres.find(l=>l.id===a.lider_id as string);
      const areasAd=((a.areas_adicionales as string[])||[]).map(id=>catalogs.areas.find(ar=>ar.id===id)?.area_es||id).join(', ');
      const maqStr=((a.maquinaria_ids as string[])||[]).map(id=>{const mq=maquinaria.find(x=>x.id===id);return mq?(mq.nombre||`${mq.item_id} (${mq.tipo})`):id;}).join(', ');
      const persAct=d.pa.filter(p=>p.actividad_programada_id===a.id as string);
      const persNombres=persAct.map(p=>{const ps=p.personal as Record<string,unknown>;return`${ps?.nombre||p.documento_personal} (${ps?.cargo_es||''})`;}).join(', ');
      rows.push({
        Fecha:prog?.fecha||'',Especialidad:espRow?.especialidad_es||'',Actividad:actRow?.actividad_es||'',
        'Área Principal':areaRow?.area_es||'','Áreas Adicionales':areasAd,
        Líder:liderRow?.nombre||'',Personal:persNombres,Maquinaria:maqStr,
        'Rendimiento Esperado':a.rendimiento_esperado||'',Observaciones:a.observacion_es||'',
      });
    });
    const wb=XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(rows),'Planeacion');
    XLSX.writeFile(wb,`Planeacion_${rangeIni}_${rangeFin}.xlsx`);
    showToast('ok','Excel descargado');
  }

  async function exportRangePDF(){
    const d=await fetchRangeData(); if(!d||!catalogs) return;
    const css=`
      *{box-sizing:border-box;margin:0;padding:0}
      body{font-family:Arial,sans-serif;font-size:11px;color:#1e293b;padding:20px}
      .hdr{border-bottom:3px solid #003b7a;padding-bottom:12px;margin-bottom:16px}
      .hdr h1{font-size:18px;color:#003b7a;font-weight:900}
      .hdr-meta{font-size:10px;color:#64748b;margin-top:4px}
      .hdr-meta strong{color:#1e293b}
      .fecha-bloque{margin-bottom:20px}
      .fecha-titulo{background:#f1f5f9;border-left:4px solid #003b7a;padding:6px 12px;font-weight:800;font-size:12px;color:#003b7a;margin-bottom:10px}
      .act-card{border:1px solid #e2e8f0;border-radius:8px;margin-bottom:12px;overflow:hidden;page-break-inside:avoid}
      .act-header{background:#003b7a;color:white;padding:7px 14px;font-size:11px;font-weight:700;display:flex;align-items:center;gap:8px}
      .act-header .act-num{font-size:10px;opacity:0.8;font-weight:400}
      .act-body{padding:12px 14px}
      .meta-grid{display:grid;grid-template-columns:1fr 1fr 1fr;gap:10px;margin-bottom:12px}
      .meta-item label{font-size:9px;color:#94a3b8;text-transform:uppercase;letter-spacing:0.5px;display:block;margin-bottom:2px}
      .meta-item span{font-size:11px;font-weight:700;color:#1e293b}
      .meta-item.wide{grid-column:span 2}
      .maq-row{background:#fef9c3;border:1px solid #fde68a;border-radius:6px;padding:6px 10px;font-size:10px;margin-bottom:10px}
      .maq-row label{color:#92400e;font-weight:700;margin-right:6px}
      .personal-titulo{font-size:9px;color:#64748b;text-transform:uppercase;letter-spacing:0.5px;font-weight:700;margin-bottom:4px}
      .personal-table{width:100%;border-collapse:collapse;font-size:10px}
      .personal-table th{background:#f8fafc;text-align:left;padding:4px 8px;font-size:9px;color:#94a3b8;text-transform:uppercase;letter-spacing:0.3px;border-bottom:1px solid #e2e8f0}
      .personal-table td{padding:4px 8px;border-bottom:1px solid #f8fafc}
      .personal-table tr:last-child td{border-bottom:none}
      .num-cell{color:#94a3b8;font-size:9px}
      .ftr{text-align:center;font-size:9px;color:#94a3b8;margin-top:24px;border-top:1px solid #e2e8f0;padding-top:8px}
    `;
    const byDate:Record<string,Record<string,unknown>[]>={};
    d.acts.forEach(a=>{
      const prog=d.progs.find(p=>p.id===a.programacion_id as string);
      const f2=prog?.fecha as string||'';
      if(!byDate[f2]) byDate[f2]=[];
      byDate[f2].push(a);
    });
    const bloquesHTML=Object.entries(byDate).sort(([a],[b])=>a.localeCompare(b)).map(([fch,acts])=>{
      const tarjetas=acts.map((act,idx)=>{
        const actRow=catalogs.especialidades_actividades.find(e=>e.id===act.actividad_id as string);
        const areaRow=catalogs.areas.find(ar=>ar.id===act.area_id as string);
        const liderRow=catalogs.lideres.find(l=>l.id===act.lider_id as string);
        const areasAdic=((act.areas_adicionales as string[])||[]).map(id=>catalogs.areas.find(ar=>ar.id===id)?.area_es||id);
        const areasAdicStr=areasAdic.length?` + ${areasAdic.join(', ')}`: '';
        const maqList=((act.maquinaria_ids as string[])||[]).map(id=>{const mq=maquinaria.find(x=>x.id===id);return mq?(mq.nombre||`${mq.item_id} (${mq.tipo})`):id;});
        const maqHTML=maqList.length?`<div class="maq-row"><label>🚜 Maquinaria:</label>${maqList.join(' · ')}</div>`:'';
        const persAct=d.pa.filter(p=>p.actividad_programada_id===act.id as string);
        const personalFilas=persAct.map((p,i)=>{
          const ps=p.personal as Record<string,unknown>;
          return `<tr><td class="num-cell">${i+1}</td><td><strong>${ps?.nombre as string||'—'}</strong></td><td>${p.documento_personal as string||'—'}</td><td>${ps?.cargo_es as string||'—'}</td></tr>`;
        }).join('');
        return `
          <div class="act-card">
            <div class="act-header">
              <span>${actRow?.actividad_es||'Actividad'}</span>
              <span class="act-num">Actividad ${idx+1}</span>
            </div>
            <div class="act-body">
              <div class="meta-grid">
                <div class="meta-item"><label>Especialidad</label><span>${actRow?.especialidad_es||'—'}</span></div>
                <div class="meta-item"><label>Área principal</label><span>${areaRow?.area_es||'—'}${areasAdicStr}</span></div>
                <div class="meta-item"><label>Líder</label><span>${liderRow?.nombre||'—'}</span></div>
                ${act.rendimiento_esperado?`<div class="meta-item"><label>Rendimiento esperado</label><span>${act.rendimiento_esperado as string}</span></div>`:''}
                ${act.observacion_es?`<div class="meta-item wide"><label>Observación</label><span style="font-weight:400">${act.observacion_es as string}</span></div>`:''}
              </div>
              ${maqHTML}
              ${persAct.length?`
              <div class="personal-titulo">👥 Personal asignado</div>
              <table class="personal-table">
                <thead><tr><th>N</th><th>Nombre</th><th>Documento</th><th>Cargo</th></tr></thead>
                <tbody>${personalFilas}</tbody>
              </table>`:'<div style="font-size:10px;color:#94a3b8;font-style:italic">Sin personal asignado</div>'}
            </div>
          </div>`;
      }).join('');
      return `<div class="fecha-bloque"><div class="fecha-titulo">📅 ${fch}</div>${tarjetas}</div>`;
    }).join('');
    const totalPersonal=new Set(d.pa.map(p=>p.documento_personal as string)).size;
    const html=`<!DOCTYPE html><html lang="es"><head><meta charset="UTF-8"><title>Planeación PDS 360</title><style>${css}</style></head><body>
      <div class="hdr">
        <h1>Powerchina · PDS 360 — Planeación</h1>
        <div class="hdr-meta">Período: <strong>${rangeIni}</strong> al <strong>${rangeFin}</strong> &nbsp;·&nbsp; Actividades: <strong>${d.acts.length}</strong> &nbsp;·&nbsp; Personal único: <strong>${totalPersonal}</strong></div>
      </div>
      ${bloquesHTML}
      <div class="ftr">Powerchina PDS 360 · Generado: ${new Date().toLocaleString('es-CO')}</div>
    </body></html>`;
    const win=window.open('','_blank');
    if(win){win.document.write(html);win.document.close();setTimeout(()=>win.print(),600);}
  }

  const loadFecha=useCallback(async()=>{
    if(!fecha) return;
    const{data:bl}=await supabase.from('personal_asignado')
      .select('documento_personal,usuario_id')
      .eq('fecha',fecha).neq('usuario_id',user.id);
    const docsBlocked=[...new Set((bl||[]).map((b:Record<string,unknown>)=>b.documento_personal as string))];
    let blockedNames:Record<string,string>={};
    if(docsBlocked.length){
      const{data:persB}=await supabase.from('personal').select('documento,nombre').in('documento',docsBlocked);
      (persB||[]).forEach((p:Record<string,unknown>)=>{blockedNames[p.documento as string]=p.nombre as string;});
    }
    setBlocked((bl||[]).map((b:Record<string,unknown>)=>({
      documento_personal:b.documento_personal as string,
      usuario_nombre:blockedNames[b.documento_personal as string]||b.documento_personal as string,
    })));
    const{data:prog}=await supabase.from('programaciones').select('*').eq('fecha',fecha).eq('usuario_id',user.id).maybeSingle();
    if(prog){
      setEstado(prog.estado as 'nuevo'|'borrador'|'enviado');
      setProgId(prog.id as string);
      const{data:acts}=await supabase.from('actividades_programadas').select('*').eq('programacion_id',prog.id);
      if(acts?.length){
        const{data:pa}=await supabase.from('personal_asignado').select('*').eq('programacion_id',prog.id);
        setActividades((acts||[]).map((a:Record<string,unknown>)=>({
          uid:gid(), especialidad_id:(a.especialidad_id as string)||'', actividad_id:(a.actividad_id as string)||'',
          area_id:(a.area_id as string)||'', areas_adicionales:(a.areas_adicionales as string[])||[],
          lider_id:(a.lider_id as string)||'', maquinaria_ids:(a.maquinaria_ids as string[])||[],
          rendimiento_esperado:(a.rendimiento_esperado as string)||'', observacion_es:(a.observacion_es as string)||'',
          observacion_en:(a.observacion_en as string)||'',
          personal:(pa||[])
            .filter((p:Record<string,unknown>)=>p.actividad_programada_id===a.id)
            .map((p:Record<string,unknown>)=>({personal_id:p.personal_id as string,documento_personal:p.documento_personal as string})),
        })));
        return;
      }
    }
    setEstado('nuevo'); setProgId(null); setActividades([emptyAct()]);
  },[fecha,user.id]);

  useEffect(()=>{loadFecha();},[loadFecha]);

  const blockedMap=useMemo(()=>{
    const m:Record<string,string>={};
    blocked.forEach(b=>{m[b.documento_personal]=b.usuario_nombre;});
    return m;
  },[blocked]);

  const isRO=estado==='enviado'||readOnly;

  async function save(est:'borrador'|'enviado'){
    for(let i=0;i<actividades.length;i++){
      const a=actividades[i];
      if(!a.especialidad_id||!a.actividad_id||!a.area_id||!a.lider_id){ showToast('err',`Actividad ${i+1}: completa todos los campos`); return; }
      if(!a.personal.length){ showToast('err',`Actividad ${i+1}: agrega al menos una persona`); return; }
    }
    if(est==='enviado'&&!window.confirm('¿Enviar planeación?')) return;
    setSaving(true);
    try{
      const{data:prog,error:pe}=await supabase.from('programaciones').upsert(
        {id:progId||undefined,fecha,usuario_id:user.id,usuario_nombre:user.nombre,estado:est,updated_at:new Date().toISOString()},
        {onConflict:'fecha,usuario_id'}
      ).select().single();
      if(pe||!prog) throw new Error(pe?.message||'Error');
      await supabase.from('actividades_programadas').delete().eq('programacion_id',prog.id);
      for(const act of actividades){
        const{data:aR,error:ae}=await supabase.from('actividades_programadas').insert({
          programacion_id:prog.id, fecha, usuario_id:user.id,
          especialidad_id:act.especialidad_id, actividad_id:act.actividad_id,
          area_id:act.area_id, areas_adicionales:act.areas_adicionales,
          lider_id:act.lider_id, maquinaria_ids:act.maquinaria_ids,
          rendimiento_esperado:act.rendimiento_esperado, observacion_es:act.observacion_es, observacion_en:act.observacion_en,
        }).select().single();
        if(ae||!aR) throw new Error(ae?.message||'Error en actividad');
        if(act.personal.length){
          const{error:paE}=await supabase.from('personal_asignado').insert(
            act.personal.map(p=>({
              programacion_id:prog.id, actividad_programada_id:aR.id,
              fecha, usuario_id:user.id, personal_id:p.personal_id, documento_personal:p.documento_personal,
            }))
          );
          if(paE?.code==='23505'){
            // Identificar quién está duplicado
            const docs=act.personal.map(p=>p.documento_personal);
            const{data:dup}=await supabase.from('personal_asignado')
              .select('documento_personal,personal!inner(nombre)').eq('fecha',fecha).in('documento_personal',docs);
            const nombres=(dup||[]).map((d:Record<string,unknown>)=>{
              const p=d.personal as Record<string,unknown>;
              return p?.nombre as string||d.documento_personal as string;
            }).join(', ');
            throw new Error(`Personal duplicado: ${nombres} ya está asignado hoy por otro técnico.`);
          }
          if(paE) throw new Error(paE.message);
        }
      }
      setProgId(prog.id as string); setEstado(est);
      showToast('ok',est==='enviado'?'Planeación enviada ✓':'Borrador guardado ✓');
      await loadFecha();
    } catch(e:unknown){ showToast('err',(e as Error)?.message||'Error'); }
    finally{ setSaving(false); }
  }

  function exportarExcel(){
    if(!catalogs) return;
    const rows:Record<string,unknown>[]=[];
    actividades.forEach((a,idx)=>{
      const espRow=catalogs.especialidades_actividades.find(e=>e.id===a.especialidad_id);
      const actRow=catalogs.especialidades_actividades.find(e=>e.id===a.actividad_id);
      const areaRow=catalogs.areas.find(ar=>ar.id===a.area_id);
      const liderRow=catalogs.lideres.find(l=>l.id===a.lider_id);
      const areasAd=(a.areas_adicionales||[]).map(id=>catalogs.areas.find(ar=>ar.id===id)?.area_es||id).join(", ");
      const maqNombres=(a.maquinaria_ids||[]).map(id=>{const mq=maquinaria.find(x=>x.id===id);return mq?(mq.nombre||`${mq.item_id} (${mq.tipo})`):id;}).join(", ");
      if(!a.personal.length){
        rows.push({"N°Act":idx+1,Fecha:fecha,Especialidad:espRow?.especialidad_es||"",Actividad:actRow?.actividad_es||"","Area principal":areaRow?.area_es||"","Areas adicionales":areasAd,Lider:liderRow?.nombre||"",Maquinaria:maqNombres,Nombre:"",Documento:"",Cargo:""});
      } else {
        a.personal.forEach(ps=>{
          const pers=catalogs.personal.find(p=>p.id===ps.personal_id);
          rows.push({"N°Act":idx+1,Fecha:fecha,Especialidad:espRow?.especialidad_es||"",Actividad:actRow?.actividad_es||"","Area principal":areaRow?.area_es||"","Areas adicionales":areasAd,Lider:liderRow?.nombre||"",Maquinaria:maqNombres,Nombre:pers?.nombre||ps.documento_personal,Documento:pers?.documento||ps.documento_personal,Cargo:pers?.cargo_es||""});
        });
      }
    });
    const wb=XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(rows),"Planeacion");
    XLSX.writeFile(wb,"Planeacion_"+fecha+".xlsx");
  }

  function imprimirPlan(){
    if(!catalogs) return;
    const lineas=actividades.map((a,idx)=>{
      const espRow=catalogs.especialidades_actividades.find(e=>e.id===a.especialidad_id);
      const actRow=catalogs.especialidades_actividades.find(e=>e.id===a.actividad_id);
      const areaRow=catalogs.areas.find(ar=>ar.id===a.area_id);
      const liderRow=catalogs.lideres.find(l=>l.id===a.lider_id);
      const areasAd=(a.areas_adicionales||[]).map(id=>catalogs.areas.find(ar=>ar.id===id)?.area_es||id).join(", ");
      const maqNombres=(a.maquinaria_ids||[]).map(id=>{const mq=maquinaria.find(x=>x.id===id);return mq?(mq.nombre||`${mq.item_id} (${mq.tipo})`):id;}).join(", ");
      const persRows=a.personal.map((ps,pi)=>{
        const p=catalogs.personal.find(x=>x.id===ps.personal_id);
        return "<tr><td>"+(pi+1)+"</td><td>"+(p?.nombre||ps.documento_personal)+"</td><td>"+(p?.documento||ps.documento_personal)+"</td><td>"+(p?.cargo_es||"")+"</td></tr>";
      }).join("");
      return "<div class=act><div class=act-h>Actividad "+(idx+1)+" — "+(actRow?.actividad_es||a.actividad_id)+"</div><div class=act-b><div class=grid><div class=field><label>Especialidad</label><span>"+(espRow?.especialidad_es||"—")+"</span></div><div class=field><label>Area</label><span>"+(areaRow?.area_es||"—")+"</span></div><div class=field><label>Lider</label><span>"+(liderRow?.nombre||"—")+" · "+(liderRow?.cargo_es||"")+"</span></div>"+(areasAd?"<div class=field><label>Areas adicionales</label><span>"+areasAd+"</span></div>":"")+(maqNombres?"<div class=field><label>Maquinaria</label><span>"+maqNombres+"</span></div>":"")+"</div>"+(a.personal.length?"<table class=pt><thead><tr><th>N</th><th>Nombre</th><th>Documento</th><th>Cargo</th></tr></thead><tbody>"+persRows+"</tbody></table>":"<p style=color:#94a3b8>Sin personal asignado</p>")+"</div></div>";
    }).join("");
    const css="body{font-family:Arial,sans-serif;font-size:11px;color:#111;margin:20px}.hdr{display:flex;align-items:center;gap:12px;border-bottom:2px solid #003b7a;padding-bottom:10px;margin-bottom:16px}.hdr h1{font-size:16px;color:#003b7a;margin:0}.hdr p{margin:2px 0;font-size:10px;color:#555}.act{border:1px solid #cbd5e1;border-radius:6px;margin-bottom:12px;page-break-inside:avoid;overflow:hidden}.act-h{background:#003b7a;color:white;padding:6px 10px;font-weight:bold;font-size:11px}.act-b{padding:8px 10px}.grid{display:grid;grid-template-columns:1fr 1fr 1fr;gap:6px;margin-bottom:8px}.field label{font-size:9px;color:#666;display:block;margin-bottom:1px}.field span{font-weight:bold}.pt{width:100%;border-collapse:collapse;font-size:10px;margin-top:6px}.pt th{background:#e2e8f0;text-align:left;padding:3px 6px;font-size:9px}.pt td{padding:3px 6px;border-bottom:1px solid #f1f5f9}.ftr{text-align:center;font-size:9px;color:#94a3b8;margin-top:20px;border-top:1px solid #e2e8f0;padding-top:8px}";
    const html="<!DOCTYPE html><html lang=es><head><meta charset=UTF-8><title>Planeacion "+fecha+"</title><style>"+css+"</style></head><body><div class=hdr><div><h1>Powerchina PDS 360 Planeacion</h1><p>Fecha: <strong>"+fecha+"</strong> Estado: <strong>"+estado.toUpperCase()+"</strong> Actividades: <strong>"+actividades.length+"</strong> Personal total: <strong>"+actividades.reduce((s,a)=>s+a.personal.length,0)+"</strong></p></div></div>"+lineas+"<div class=ftr>Powerchina PDS 360 Generado: "+new Date().toLocaleString("es-CO")+"</div></body></html>";
    const win=window.open("","_blank");
    if(win){win.document.write(html);win.document.close();setTimeout(()=>win.print(),500);}
  }

  return(
    <div className="space-y-4">
      {readOnly&&(
        <div className="bg-blue-50 border border-blue-200 rounded-xl px-4 py-3 flex items-center gap-3">
          <span className="text-blue-500 text-lg">👁️</span>
          <div>
            <div className="font-semibold text-sm text-blue-800">Vista de solo lectura / Read-only view</div>
            <div className="text-xs text-blue-600">Puedes consultar la planeación pero no puedes editarla · You can view the planning but not edit it</div>
          </div>
        </div>
      )}
      <div className="card p-4 flex flex-col sm:flex-row gap-3 items-end justify-between flex-wrap">
        <div className="flex-1 min-w-[160px]"><label className="label">Fecha</label><input type="date" className="input" value={fecha} onChange={e=>setFecha(e.target.value)}/></div>
        <div className="flex items-center gap-3 flex-wrap">
          {estado==='nuevo'&&<span className="badge bg-slate-200 text-slate-700">Nuevo</span>}
          {estado==='borrador'&&<span className="badge-borrador">Borrador</span>}
          {estado==='enviado'&&<span className="badge-enviado">Enviado</span>}
          <button className="btn-secondary text-xs" onClick={exportarExcel}>📥 Excel día</button>
          <button className="btn-secondary text-xs" onClick={imprimirPlan}>🖨️ PDF día</button>
          <button className="btn-secondary text-xs" onClick={()=>setShowRange(v=>!v)}>📅 Exportar por rango</button>
        </div>
      </div>

      {showRange&&(
        <div className="card p-4 space-y-3 border-2 border-blue-200 bg-blue-50">
          <h4 className="font-bold text-[#003b7a] text-sm">📅 Exportar planeación por rango</h4>
          <div className="grid grid-cols-2 sm:grid-cols-4 gap-3 items-end">
            <div><label className="label">Desde</label><input type="date" className="input" value={rangeIni} onChange={e=>setRangeIni(e.target.value)}/></div>
            <div><label className="label">Hasta</label><input type="date" className="input" value={rangeFin} onChange={e=>setRangeFin(e.target.value)}/></div>
          </div>
          <div>
            <label className="label">Especialidades (vacío = todas)</label>
            <div className="flex flex-wrap gap-2">
              {espListRange.map(e=>(
                <button key={e.id} onClick={()=>setRangeEsps(prev=>prev.includes(e.id)?prev.filter(x=>x!==e.id):[...prev,e.id])}
                  className={`text-xs px-2 py-1 rounded border transition-colors ${rangeEsps.includes(e.id)?'bg-[#003b7a] text-white border-[#003b7a]':'border-slate-300 text-slate-600 hover:border-[#003b7a]'}`}>
                  {e.especialidad_es}
                </button>
              ))}
              {rangeEsps.length>0&&<button className="text-xs text-slate-500 underline" onClick={()=>setRangeEsps([])}>Limpiar</button>}
            </div>
          </div>
          <div className="flex gap-2 flex-wrap">
            <button className="btn-primary text-xs" disabled={rangeLoading} onClick={exportRangeExcel}>{rangeLoading?'Cargando…':'📥 Excel por rango'}</button>
            <button className="btn-secondary text-xs" disabled={rangeLoading} onClick={exportRangePDF}>{rangeLoading?'Cargando…':'🖨️ PDF por rango'}</button>
          </div>
        </div>
      )}

      {actividades.map((a,idx)=>(
        <ActCard key={a.uid} index={idx} act={a} catalogs={catalogs} maquinaria={maquinaria}
          blockedMap={blockedMap} readOnly={isRO}
          otherUsedPersonalIds={actividades.filter(x=>x.uid!==a.uid).flatMap(x=>x.personal.map(p=>p.personal_id))}
          onChange={p=>setActividades(arr=>arr.map(x=>x.uid===a.uid?{...x,...p}:x))}
          onRemove={()=>setActividades(arr=>arr.length<=1?arr:arr.filter(x=>x.uid!==a.uid))}/>
      ))}

      <div className="flex flex-wrap gap-2">
        {!isRO&&<button className="btn-secondary" onClick={()=>setActividades(a=>[...a,emptyAct()])}>+ Actividad</button>}
        {!isRO&&<><button className="btn-primary" disabled={saving} onClick={()=>save('borrador')}>{saving?'Guardando…':'Guardar borrador'}</button><button className="btn-success" disabled={saving} onClick={()=>save('enviado')}>{saving?'Enviando…':'Enviar planeación'}</button></>}
        {isRO&&<button className="btn-secondary" disabled={saving} onClick={async()=>{if(!window.confirm('¿Reabrir como borrador?'))return;setSaving(true);await supabase.from('programaciones').update({estado:'borrador'}).eq('id',progId!);setEstado('borrador');setSaving(false);}}>Reabrir</button>}
      </div>
    </div>
  );
}

// ── ActCard CON PANEL RESUMEN ─────────────────────────────────────
function ActCard({index,act,catalogs,maquinaria,blockedMap,readOnly,otherUsedPersonalIds,onChange,onRemove}:{
  index:number; act:ActForm; catalogs:Catalogs|null; maquinaria:Maq[];
  blockedMap:Record<string,string>; readOnly:boolean; otherUsedPersonalIds:string[];
  onChange:(p:Partial<ActForm>)=>void; onRemove:()=>void;
}){
  const[search,setSearch]=useState('');
  const[liderSearch,setLiderSearch]=useState('');

  const esps=useMemo(()=>catalogs?uniqueEsp(catalogs.especialidades_actividades.filter(e=>e.activo!==false)):[],[catalogs]);
  const acts=useMemo(()=>catalogs&&act.especialidad_id?actsForEsp(catalogs.especialidades_actividades.filter(e=>e.activo!==false),act.especialidad_id):[],[catalogs,act.especialidad_id]);
  const persActivos=useMemo(()=>(catalogs?.personal||[]).filter(p=>p.activo!==false),[catalogs]);
  const persFilt=useMemo(()=>{
    const q=search.trim().toLowerCase();
    return q?persActivos.filter(p=>p.nombre.toLowerCase().includes(q)||p.documento.toLowerCase().includes(q)||p.cargo_es.toLowerCase().includes(q)):persActivos;
  },[persActivos,search]);

  const lideresActivos=useMemo(()=>(catalogs?.lideres||[]).filter(l=>l.activo!==false),[catalogs]);
  const lideresFilt=useMemo(()=>{
    const q=liderSearch.trim().toLowerCase();
    return q?lideresActivos.filter(l=>l.nombre.toLowerCase().includes(q)||l.documento.toLowerCase().includes(q)):lideresActivos;
  },[lideresActivos,liderSearch]);

  const liderSel=useMemo(()=>lideresActivos.find(l=>l.id===act.lider_id),[lideresActivos,act.lider_id]);
  const areaSel=useMemo(()=>(catalogs?.areas||[]).find(a=>a.id===act.area_id),[catalogs,act.area_id]);
  const actNombre=useMemo(()=>acts.find(a=>a.id===act.actividad_id)?.actividad_es||null,[acts,act.actividad_id]);

  // Panel resumen — personal seleccionado
  const persSelDetalle=useMemo(()=>
    act.personal.map(ps=>{
      const p=persActivos.find(x=>x.id===ps.personal_id);
      return p?{nombre:p.nombre,cargo:p.cargo_es}:{nombre:ps.documento_personal,cargo:''};
    }),
  [act.personal,persActivos]);

  const maqSel=useMemo(()=>
    act.maquinaria_ids.map(id=>maquinaria.find(m=>m.id===id)).filter(Boolean) as Maq[],
  [act.maquinaria_ids,maquinaria]);

  return(
    <div className="card p-0 overflow-hidden">
      <div className="flex items-center justify-between px-4 py-3 bg-[#003b7a] text-white">
        <span className="font-semibold">Actividad {index+1}</span>
        {!readOnly&&<button className="text-red-300 hover:text-white text-xs" onClick={onRemove}>✕ Eliminar</button>}
      </div>

      <div className="flex flex-col lg:flex-row gap-0">
        {/* FORMULARIO */}
        <div className="flex-1 p-4 space-y-3">
          <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
            <div><label className="label">Especialidad</label>
              <select className="select" disabled={readOnly} value={act.especialidad_id} onChange={e=>onChange({especialidad_id:e.target.value,actividad_id:''})}>
                <option value="">— Seleccionar —</option>
                {esps.map(e=><option key={e.id} value={e.id}>{e.especialidad_es}</option>)}
              </select>
            </div>
            <div><label className="label">Actividad</label>
              <select className="select" disabled={readOnly||!act.especialidad_id} value={act.actividad_id} onChange={e=>onChange({actividad_id:e.target.value})}>
                <option value="">— Seleccionar —</option>
                {acts.map(a=><option key={a.id} value={a.id}>{a.actividad_es}</option>)}
              </select>
            </div>
            <div><label className="label">Área principal</label>
              <select className="select" disabled={readOnly} value={act.area_id} onChange={e=>onChange({area_id:e.target.value})}>
                <option value="">— Seleccionar —</option>
                {(catalogs?.areas||[]).map(a=><option key={a.id} value={a.id}>{a.area_es}</option>)}
              </select>
            </div>
          </div>

          <div><label className="label">Áreas adicionales</label>
            <div className="flex flex-wrap gap-2">
              {(catalogs?.areas||[]).filter(a=>a.id!==act.area_id).map(a=>{
                const sel=act.areas_adicionales.includes(a.id);
                return <button key={a.id} disabled={readOnly}
                  onClick={()=>onChange({areas_adicionales:sel?act.areas_adicionales.filter(x=>x!==a.id):[...act.areas_adicionales,a.id]})}
                  className={`text-xs px-2 py-1 rounded border transition-colors ${sel?'bg-[#003b7a] text-white border-[#003b7a]':'border-slate-300 text-slate-600 hover:border-[#003b7a]'}`}>
                  {a.area_es}
                </button>;
              })}
            </div>
          </div>

          {/* BÚSQUEDA DE LÍDER */}
          <div><label className="label">Líder de esta actividad</label>
            {act.lider_id&&<div className="mb-1 flex items-center gap-2">
              <span className="badge bg-[#003b7a] text-white">{liderSel?.nombre||act.lider_id}</span>
              {!readOnly&&<button className="text-xs text-rose-500 underline" onClick={()=>onChange({lider_id:''})}>Cambiar</button>}
            </div>}
            {!act.lider_id&&<>
              <input className="input mb-1" placeholder="🔎 Buscar líder…" value={liderSearch} onChange={e=>setLiderSearch(e.target.value)}/>
              <div className="border border-slate-200 rounded-md max-h-32 overflow-y-auto">
                {lideresFilt.map(l=>(
                  <button key={l.id} onClick={()=>{onChange({lider_id:l.id});setLiderSearch('');}}
                    className="w-full text-left px-3 py-2 border-b border-slate-100 hover:bg-blue-50 text-sm">
                    <span className="font-medium">{l.nombre}</span>
                    <span className="text-xs text-slate-500 ml-2">{l.cargo_es}</span>
                  </button>
                ))}
              </div>
            </>}
          </div>

          <div><label className="label">Maquinaria</label>
            <div className="flex flex-wrap gap-2">
              {maquinaria.filter(m=>m.estado==='activo').map(m=>{
                const sel=act.maquinaria_ids.includes(m.id);
                return <button key={m.id} disabled={readOnly}
                  onClick={()=>onChange({maquinaria_ids:sel?act.maquinaria_ids.filter(x=>x!==m.id):[...act.maquinaria_ids,m.id]})}
                  className={`text-xs px-2 py-1 rounded border transition-colors ${sel?'bg-orange-500 text-white border-orange-500':'border-slate-300 text-slate-600 hover:border-orange-400'}`}>
                  {m.nombre||`${m.item_id} (${m.tipo})`}
                </button>;
              })}
            </div>
          </div>

          <div>
            <div className="flex items-center justify-between mb-2">
              <label className="label !mb-0">Personal ({act.personal.length} sel.)</label>
              {!readOnly&&act.personal.length>0&&<button className="text-xs text-rose-500 underline" onClick={()=>onChange({personal:[]})}>Limpiar</button>}
            </div>
            <input className="input mb-2" placeholder="🔎 Buscar nombre, documento o cargo…" value={search} onChange={e=>setSearch(e.target.value)}/>
            <div className="border border-slate-200 rounded-md max-h-52 overflow-y-auto">
              {persFilt.map(p=>{
                const sel=!!act.personal.find(x=>x.documento_personal===p.documento);
                const bl=blockedMap[p.documento];
                const inOtherAct=!bl&&otherUsedPersonalIds.includes(p.id);
                return(
                  <div key={p.id} className="relative group">
                    <label className={`flex items-center gap-2 px-3 py-2 border-b border-slate-100 cursor-pointer text-sm ${bl?'bg-rose-50 cursor-not-allowed':inOtherAct&&!sel?'bg-orange-50':sel?'bg-blue-50':'hover:bg-slate-50'}`}>
                      <input type="checkbox" checked={sel} disabled={readOnly||!!bl}
                        onChange={()=>onChange({personal:sel?act.personal.filter(x=>x.documento_personal!==p.documento):[...act.personal,{personal_id:p.id,documento_personal:p.documento}]})}/>
                      <div className="flex-1 min-w-0">
                        <div className={`font-medium truncate ${bl?'text-slate-400':''}`}>{p.nombre}</div>
                        <div className="text-xs text-slate-500">{p.documento} · {p.cargo_es}</div>
                      </div>
                      {bl&&<span className="text-xs text-rose-500 font-medium flex-shrink-0">🔒 {bl}</span>}
                      {inOtherAct&&!sel&&<span className="text-xs text-orange-500 font-medium flex-shrink-0">⚡ Otra act.</span>}
                      {sel&&!bl&&<span className="badge bg-blue-100 text-blue-800 flex-shrink-0">✓</span>}
                    </label>
                    {bl&&<div className="absolute left-0 bottom-full mb-1 hidden group-hover:block z-20 bg-slate-800 text-white text-xs rounded p-2 shadow-lg whitespace-nowrap">Asignado por: <strong>{bl}</strong></div>}
                    {inOtherAct&&!bl&&<div className="absolute left-0 bottom-full mb-1 hidden group-hover:block z-20 bg-orange-700 text-white text-xs rounded p-2 shadow-lg whitespace-nowrap">Ya planificado en otra actividad de este día</div>}
                  </div>
                );
              })}
            </div>
          </div>
        </div>

        {/* PANEL RESUMEN — derecho */}
        <div className="lg:w-64 border-t lg:border-t-0 lg:border-l border-slate-200 bg-slate-50 p-4 flex-shrink-0 lg:sticky lg:top-4 lg:self-start">
          <div className="text-xs font-bold text-[#003b7a] uppercase mb-3">📋 Resumen</div>

          <div className="space-y-2 text-xs">
            {actNombre&&<div><span className="text-slate-400">Actividad:</span> <span className="font-medium text-[#003b7a]">{actNombre}</span></div>}
            <div><span className="text-slate-400">Área:</span> <span className="font-medium">{areaSel?.area_es||'—'}</span></div>
            <div><span className="text-slate-400">Líder:</span> <span className="font-medium">{liderSel?.nombre||'—'}</span>{liderSel&&<div className="text-slate-500">{liderSel.cargo_es}</div>}</div>

            {maqSel.length>0&&(
              <div><span className="text-slate-400">Maquinaria:</span>
                <div className="flex flex-wrap gap-1 mt-1">{maqSel.map(m=><span key={m.id} className="badge bg-orange-100 text-orange-700">{m.nombre||`${m.item_id} (${m.tipo})`}</span>)}</div>
              </div>
            )}

            <div className="border-t border-slate-200 pt-2">
              <span className="text-slate-400">Personal ({persSelDetalle.length}):</span>
              {!persSelDetalle.length&&<div className="text-slate-400 italic mt-1">Ninguno seleccionado</div>}
              <div className="mt-1 max-h-48 overflow-y-auto space-y-1">
                {persSelDetalle.map((p,i)=>(
                  <div key={i} className="bg-white rounded p-1.5 border border-slate-200">
                    <div className="font-medium text-slate-800 truncate">{p.nombre}</div>
                    <div className="text-slate-500 truncate">{p.cargo}</div>
                  </div>
                ))}
              </div>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}

// ── REPORTE DIARIO ────────────────────────────────────────────────
function ReporteModule({user,catalogs,maquinaria,configActs,showToast}:{
  user:Profile; catalogs:Catalogs|null; maquinaria:Maq[];
  configActs:ConfigAct[]; showToast:(k:'ok'|'err'|'info',m:string)=>void;
}){
  const[step,setStep]=useState(1);
  const[fecha,setFecha]=useState(today());
  const[espId,setEspId]=useState(user.especialidad_id||'');
  const[jornadaHrs,setJornadaHrs]=useState(9);
  const[clima,setClima]=useState('despejado');
  const[susps,setSusps]=useState<SuspItem[]>([]);
  const[charla,setCharla]=useState(true);
  const[charlaTema,setCharlaTema]=useState('');
  const[asistencia,setAsistencia]=useState<AsistItem[]>([]);
  const[maqDia,setMaqDia]=useState<{uid:string;maquinaria_id:string;nombre:string;uso:'si'|'no'|null;planeada:boolean;es_adicional:boolean;parcial:boolean;hora_inicio:string;hora_fin:string;tiene_novedad:boolean;novedad:string}[]>([]);
  const[suspPorActividad,setSuspPorActividad]=useState<SuspActividad[]>([]);
  const[actReps,setActReps]=useState<ActRep[]>([]);
  const[incidente,setIncidente]=useState({tipo:'sin_novedad',descripcion:'',medidas:''});
  const[notaBit,setNotaBit]=useState('');
  const[saving,setSaving]=useState(false);
  const[solicitudOk,setSolicitudOk]=useState(false);
  const[itemDbs,setItemDbs]=useState<Record<string,{db:ItemDatabase;items:ItemDB[]}>>({}); // keyed by config_actividad_id
  const[itemDbSearch,setItemDbSearch]=useState<Record<string,string>>({});
  const[actAdicionales,setActAdicionales]=useState<ActAdicItem[]>([]);
  const[catAdicionales,setCatAdicionales]=useState<ActAdicCat[]>([]);
  const[mostrarFormAdicional,setMostrarFormAdicional]=useState(false);
  const[nuevaAdicionalNombre,setNuevaAdicionalNombre]=useState('');
  const[nuevaAdicionalDesc,setNuevaAdicionalDesc]=useState('');

  const horasClima=useMemo(()=>horasLost(susps),[susps]);
  const horasReal=Math.max(0,jornadaHrs-horasClima);
  const esDiaAnt=fecha<today();
  const STEPS=['Encabezado','Condiciones','Charla','Asistencia','Maquinaria','Avance','Seguridad','Enviar'];

  useEffect(()=>{
    if(!fecha||!espId) return;
    if(esDiaAnt){
      supabase.from('solicitudes_reporte_pasado').select('*')
        .eq('tecnico_id',user.id).eq('fecha_reporte',fecha).eq('estado','aprobado').maybeSingle()
        .then(({data})=>setSolicitudOk(!!data));
    } else { setSolicitudOk(false); }
    (async()=>{
      const espRow=catalogs?.especialidades_actividades.find(e=>e.id===espId);
      const espName=(espRow?.especialidad_es||'').trim().toLowerCase();
      const espIds=espName&&catalogs?catalogs.especialidades_actividades.filter(e=>(e.especialidad_es||'').trim().toLowerCase()===espName).map(e=>e.id):[espId];
      const{data:acts}=await supabase.from('actividades_programadas').select('id').eq('fecha',fecha).in('especialidad_id',espIds.length?espIds:[espId]);
      if(!acts?.length){setAsistencia([]);return;}
      const actIds=(acts as Record<string,unknown>[]).map(a=>a.id as string);
      const{data:paData}=await supabase.from('personal_asignado').select('personal_id,documento_personal').in('actividad_programada_id',actIds);
      const personalIds=[...new Set((paData||[]).map(p=>(p as Record<string,unknown>).personal_id as string))];
      let personalMap:Record<string,{nombre:string;cargo_es:string}>={};
      if(personalIds.length){
        const{data:pers}=await supabase.from('personal').select('id,nombre,cargo_es').in('id',personalIds);
        (pers||[]).forEach((p:Record<string,unknown>)=>{personalMap[p.id as string]={nombre:p.nombre as string,cargo_es:p.cargo_es as string};});
      }
      const seen=new Set<string>();
      setAsistencia((paData||[]).filter((r:Record<string,unknown>)=>{
        const doc=r.documento_personal as string;
        if(seen.has(doc)) return false; seen.add(doc); return true;
      }).map((r:Record<string,unknown>)=>{
        const info=personalMap[r.personal_id as string]||{nombre:r.documento_personal as string,cargo_es:''};
        return { personal_id:r.personal_id as string, documento_personal:r.documento_personal as string, nombre:info.nombre, cargo_es:info.cargo_es, asistio:true, motivo_ausencia:'', ausencia_parcial:false, hora_ausencia_ini:'', hora_ausencia_fin:'', jornada_parcial:false, hora_jornada_ini:'', hora_jornada_fin:'', es_adicional:false };
      }));
    })();
  },[fecha,espId,user.id,esDiaAnt,catalogs]);

  useEffect(()=>{
    if(!fecha||!espId) return;
    async function loadMaqPlaneada(){
      const espRow=catalogs?.especialidades_actividades.find(e=>e.id===espId);
      const espName=(espRow?.especialidad_es||'').trim().toLowerCase();
      const espIds=espName&&catalogs?catalogs.especialidades_actividades.filter(e=>(e.especialidad_es||'').trim().toLowerCase()===espName).map(e=>e.id):[espId];
      const{data:plans}=await supabase.from('actividades_programadas').select('maquinaria_ids').eq('fecha',fecha).in('especialidad_id',espIds.length?espIds:[espId]);
      const allIds=[...new Set((plans||[]).flatMap(p=>(p.maquinaria_ids||[]) as string[]))];
      if(allIds.length){
        const{data:maqRows}=await supabase.from('maquinaria').select('*').in('id',allIds);
        setMaqDia((maqRows||[]).map(m=>({uid:gid(),maquinaria_id:m.id as string,nombre:(m.nombre as string)||`${m.item_id as string} (${m.tipo as string})`,uso:null,planeada:true,es_adicional:false,parcial:false,hora_inicio:'',hora_fin:'',tiene_novedad:false,novedad:''})));
      } else {
        setMaqDia([]);
      }
    }
    loadMaqPlaneada();
  },[fecha,espId,catalogs]);

  useEffect(()=>{
    if(!catalogs||!espId) return;
    const espRow=catalogs.especialidades_actividades.find(e=>e.id===espId);
    if(!espRow) return;
    const t=(espRow.especialidad_es||'').trim().toLowerCase();
    const acts=catalogs.especialidades_actividades.filter(e=>(e.especialidad_es||'').trim().toLowerCase()===t&&e.activo!==false);
    setActReps(acts.map(a=>({uid:gid(),actividad_id:a.id,areas:[{uid:gid(),area_id:'',cantidad:''}],descripcion_cualitativa:'',observacion_es:'',items_seleccionados:[]})));
    setSuspPorActividad([]);
    setActAdicionales([]);
    setMaqDia([]);
  },[catalogs,espId]);

  useEffect(()=>{
    if(step!==6) return;
    supabase.from('actividades_adicionales_catalogo').select('*').eq('activo',true).order('veces_usada',{ascending:false}).limit(50)
      .then(({data})=>{ if(data) setCatAdicionales(data as ActAdicCat[]); });
    const cfgsConItems=configActs.filter(c=>c.tiene_items_unicos);
    if(!cfgsConItems.length) return;
    cfgsConItems.forEach(async(cfg)=>{
      if(itemDbs[cfg.id]) return; // ya cargado
      try{
        const{data:dbs}=await supabase.from('item_databases').select('*').eq('config_actividad_id',cfg.id).eq('activo',true);
        if(!dbs?.length) return;
        const db=dbs[0] as ItemDatabase;
        const{data:items}=await supabase.from('items_database').select('*').eq('database_id',db.id);
        setItemDbs(prev=>({...prev,[cfg.id]:{db,items:(items||[]) as ItemDB[]}}));
      } catch{}
    });
  // eslint-disable-next-line react-hooks/exhaustive-deps
  },[step]);

  async function submit(){
    if(!fecha||!espId){showToast('err','Falta fecha o especialidad');return;}
    if(esDiaAnt&&!solicitudOk&&user.rol!=='admin'){showToast('err','Necesitas aprobación para reportar día anterior. Ve a Solicitudes.');return;}
    setSaving(true);
    try{
      const{data:rep,error:re}=await supabase.from('reportes_avance').insert({
        fecha,usuario_id:user.id,usuario_nombre:user.nombre,especialidad_id:espId,
        jornada_horas:jornadaHrs,clima,charla_preturno:charla,charla_tema:charlaTema,estado:'borrador',
      }).select().single();
      if(re||!rep) throw new Error(re?.message||'Error');
      const rid=rep.id as string;

      await Promise.all([
        susps.length?supabase.from('suspensiones_clima').insert(susps.map(s=>({
          reporte_id:rid,fecha,usuario_id:user.id,
          hora_inicio:s.hora_inicio||null,hora_fin:s.hora_fin||null,
          horas_perdidas:horasLost([s]),
          descripcion:`[${s.tipo_susp}${s.tipo_susp==='otro'?': '+s.otro_desc:''}] ${s.descripcion}`,
        }))):null,
        asistencia.length?supabase.from('asistencia_real').insert(asistencia.map(a=>({
          reporte_id:rid,fecha,usuario_id:user.id,
          personal_id:a.personal_id,documento_personal:a.documento_personal,
          asistio:a.asistio,motivo_ausencia:a.motivo_ausencia||null,
          horas_trabajadas:a.asistio?(a.jornada_parcial&&a.hora_jornada_ini&&a.hora_jornada_fin?calcHoras(a.hora_jornada_ini,a.hora_jornada_fin):horasReal):0,
          ausencia_parcial:a.ausencia_parcial||false,
          hora_ausencia_ini:a.ausencia_parcial&&a.hora_ausencia_ini?a.hora_ausencia_ini:null,
          hora_ausencia_fin:a.ausencia_parcial&&a.hora_ausencia_fin?a.hora_ausencia_fin:null,
          jornada_parcial:a.jornada_parcial||false,
          hora_jornada_ini:a.jornada_parcial&&a.hora_jornada_ini?a.hora_jornada_ini:null,
          hora_jornada_fin:a.jornada_parcial&&a.hora_jornada_fin?a.hora_jornada_fin:null,
          es_adicional:a.es_adicional||false,
        }))):null,
        maqDia.filter(m=>m.uso==='si').length?supabase.from('novedades_maquinaria').insert(
          maqDia.filter(m=>m.uso==='si').map(m=>({
            reporte_id:rid,fecha,usuario_id:user.id,maquinaria_id:m.maquinaria_id,
            descripcion:m.tiene_novedad?m.novedad:'',
            hora_inicio:m.parcial&&m.hora_inicio?m.hora_inicio:null,
            hora_fin:m.parcial&&m.hora_fin?m.hora_fin:null,
            horas_standby:0,es_adicional:m.es_adicional||false,
          }))
        ):null,
        incidente.tipo!=='sin_novedad'?supabase.from('incidentes_seg').insert({
          reporte_id:rid,fecha,usuario_id:user.id,
          tipo:incidente.tipo,descripcion:incidente.descripcion,medidas_tomadas:incidente.medidas,
        }):null,
        notaBit?supabase.from('bitacora_decisiones').insert({fecha,usuario_id:user.id,descripcion:notaBit,especialidad_id:espId}):null,
      ]);

      for(const ar of actReps){
        const cfg=configActs.find(c=>c.actividad_id===ar.actividad_id);
        if(!cfg||cfg.tipo==='D'){
          if(ar.descripcion_cualitativa.trim()){
            const{error:cualErr}=await supabase.from('avance_diario').insert({
              reporte_id:rid,fecha,usuario_id:user.id,actividad_id:ar.actividad_id,
              especialidad_id:espId,cantidad:0,unidad:'cualitativo',
              acumulado_anterior:0,acumulado_total:0,observacion_es:ar.descripcion_cualitativa,
            });
            if(cualErr){
              showToast('err','Error guardando cualitativa: '+cualErr.message);
              console.error('Error cualitativa:',cualErr);
            }
          }
          continue;
        }
        if(cfg?.es_medible===false) continue;
        let avanceId:string|null=null;
        for(const area of ar.areas.filter(a=>a.area_id&&parseFloat(a.cantidad)>0)){
          const{data:prev}=await supabase.from('avance_diario').select('cantidad')
            .eq('actividad_id',ar.actividad_id).lte('fecha',fecha);
          const acumPrev=(prev||[]).reduce((s:number,r:Record<string,unknown>)=>s+parseFloat(String(r.cantidad||0)),0);
          const cantidad=parseFloat(area.cantidad);
          const{data:avRow,error:avErr}=await supabase.from('avance_diario').insert({
            reporte_id:rid,fecha,usuario_id:user.id,actividad_id:ar.actividad_id,
            especialidad_id:espId,area_id:area.area_id,cantidad,
            unidad:cfg?.unidad_es||'',acumulado_anterior:acumPrev,acumulado_total:acumPrev+cantidad,
            observacion_es:ar.observacion_es,
          }).select().single();
          if(avErr){ showToast('err','Error guardando avance: '+avErr.message); console.error('avance_diario insert error:',avErr); }
          if(avRow&&!avanceId) avanceId=(avRow as Record<string,unknown>).id as string;
        }
        // Guardar ítems seleccionados si la actividad tiene base de datos
        if(cfg?.tiene_items_unicos&&ar.items_seleccionados.length>0&&avanceId){
          const dbEntry=itemDbs[cfg.id];
          if(dbEntry){
            const bloqueoTipo=dbEntry.db.bloqueo_tipo;
            await supabase.from('items_usados').insert(ar.items_seleccionados.map(itemId=>({
              reporte_id:rid,avance_id:avanceId,item_id:itemId,fecha,
            })));
            await supabase.from('items_database').upsert(ar.items_seleccionados.map(itemId=>({
              id:itemId,bloqueado:true,
              bloqueado_fecha:bloqueoTipo==='temporal'?fecha:null,
              bloqueado_en_reporte:rid,
            })));
          }
        }
      }
      if(actAdicionales.length>0){
        const{error:adErr}=await supabase.from('actividades_adicionales_reporte').insert(
          actAdicionales.map(ad=>({reporte_id:rid,catalogo_id:ad.catalogoId||null,nombre:ad.nombre,descripcion_ejecutado:ad.descripcion,fecha}))
        );
        if(adErr){ showToast('err','Error guardando adicionales: '+adErr.message); console.error('adicionales insert error:',adErr); }
      }
      // Suspensiones generales (sin actividad)
      // ya se guardan arriba en susps
      // Suspensiones por actividad — guardar en suspensiones_clima con actividad_id
      if(suspPorActividad.length>0){
        try{
          await supabase.from('suspensiones_clima').insert(
            suspPorActividad.map(s=>({
              reporte_id:rid,fecha,usuario_id:user.id,
              actividad_id:s.actividad_id,
              es_general:false,
              tipo_susp:s.tipo,
              otro_desc:s.otro_desc||null,
              hora_inicio:s.parcial&&s.hora_inicio?s.hora_inicio:null,
              hora_fin:s.parcial&&s.hora_fin?s.hora_fin:null,
              horas_perdidas:horasLost([s] as unknown as SuspItem[]),
              descripcion:`[${s.tipo}${s.tipo==='otro'?': '+s.otro_desc:''}] ${s.observacion||''}`,
            }))
          );
        } catch{}
      }
      try{
        const{data:la}=await supabase.from('profiles').select('id').in('rol',['admin','lider']);
        if(la?.length) await supabase.from('notificaciones').insert((la as {id:string}[]).map(x=>({
          usuario_id:x.id,tipo:'reporte_enviado',titulo:'Nuevo reporte por aprobar',
          mensaje:`${user.nombre} envió su reporte del ${fecha}`,data:{reporte_id:rid},
        })));
      } catch{}
      showToast('ok','Reporte enviado ✓');setStep(1);
    } catch(e:unknown){ showToast('err',(e as Error)?.message||'Error'); }
    finally{ setSaving(false); }
  }

  const asistio=asistencia.filter(a=>a.asistio).length;
  const efic=asistencia.length>0?Math.round((asistio/asistencia.length)*100):100;

  const TIPOS_SUSP=[
    {v:'clima',l:'🌧️ Clima / lluvia'},
    {v:'protesta',l:'🚫 Orden público / protesta'},
    {v:'logistica',l:'🚛 Logística / transporte'},
    {v:'falla_equipo',l:'🔧 Falla de equipo'},
    {v:'decision',l:'📋 Decisión de dirección'},
    {v:'otro',l:'📝 Otro (especificar)'},
  ];

  if(esDiaAnt&&!solicitudOk&&step===1&&user.rol!=='admin'){
    return(
      <div className="card p-6 text-center space-y-4">
        <div className="text-4xl">🔒</div>
        <h3 className="font-bold text-[#003b7a] text-lg">Reporte de día anterior</h3>
        <p className="text-sm text-slate-600">Para reportar un día anterior necesitas autorización del líder o admin.</p>
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 text-left max-w-md mx-auto">
          <div><label className="label">Fecha anterior</label><input type="date" className="input" value={fecha} onChange={e=>setFecha(e.target.value)} max={new Date(Date.now()-86400000).toISOString().split('T')[0]}/></div>
          <div><label className="label">Especialidad</label><select className="select" value={espId} onChange={e=>setEspId(e.target.value)}><option value="">— Seleccionar —</option>{catalogs&&uniqueEsp(catalogs.especialidades_actividades).map(e=><option key={e.id} value={e.id}>{e.especialidad_es}</option>)}</select></div>
        </div>
        <p className="text-xs text-slate-400">Ve a <strong>Solicitudes</strong> para pedir autorización.</p>
      </div>
    );
  }

  return(
    <div className="space-y-4">
      <div className="card p-3">
        <div className="flex items-center gap-1 overflow-x-auto">
          {STEPS.map((s,i)=>{
            const n=i+1;
            return(
              <div key={n} className="flex items-center gap-1 flex-shrink-0">
                <button onClick={()=>setStep(n)} className={n<step?'step-done':n===step?'step-active':'step-pending'}>{n<step?'✓':n}</button>
                <span className={`text-xs whitespace-nowrap hidden sm:inline ${n===step?'font-semibold text-[#003b7a]':'text-slate-500'}`}>{s}</span>
                {i<STEPS.length-1&&<span className="text-slate-300 mx-1">›</span>}
              </div>
            );
          })}
        </div>
      </div>

      {step===1&&(
        <div className="card p-4 space-y-3">
          <h3 className="font-bold text-[#003b7a]">Paso 1 — Encabezado</h3>
          <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
            <div><label className="label">Fecha</label><input type="date" className="input" value={fecha} onChange={e=>setFecha(e.target.value)}/></div>
            <div><label className="label">Especialidad</label><select className="select" value={espId} onChange={e=>setEspId(e.target.value)}><option value="">— Seleccionar —</option>{catalogs&&uniqueEsp(catalogs.especialidades_actividades.filter(e=>e.activo!==false)).map(e=><option key={e.id} value={e.id}>{e.especialidad_es}</option>)}</select></div>
            <div><label className="label">Horas de jornada</label><input type="number" className="input" min={1} max={24} value={jornadaHrs} onChange={e=>setJornadaHrs(parseFloat(e.target.value)||9)}/></div>
          </div>
          <button className="btn-primary" onClick={()=>{if(!fecha||!espId){showToast('err','Completa fecha y especialidad');return;}setStep(2);}}>Siguiente →</button>
        </div>
      )}

      {step===2&&(
        <div className="card p-4 space-y-3">
          <h3 className="font-bold text-[#003b7a]">Paso 2 — Condiciones del día</h3>
          <div className="flex gap-2 flex-wrap">
            {['despejado','nublado','lluvia','tormenta','suspendido'].map(c=>(
              <button key={c} onClick={()=>setClima(c)} className={`text-sm px-3 py-1.5 rounded-full border transition-colors ${clima===c?'bg-[#003b7a] text-white border-[#003b7a]':'border-slate-300 text-slate-600'}`}>
                {c==='despejado'?'☀️ Despejado':c==='nublado'?'☁️ Nublado':c==='lluvia'?'🌧️ Lluvia':c==='tormenta'?'⛈️ Tormenta':'🚫 Suspendido'}
              </button>
            ))}
          </div>

          <div className="space-y-3">
            {susps.map((s,i)=>(
              <div key={s.uid} className="border border-slate-200 rounded-lg p-3 bg-slate-50 space-y-2">
                <div className="flex items-center justify-between">
                  <span className="text-sm font-medium text-slate-700">Suspensión {i+1}</span>
                  <button className="text-rose-500 text-xs" onClick={()=>setSusps(a=>a.filter(x=>x.uid!==s.uid))}>✕ Eliminar</button>
                </div>
                <div className="grid grid-cols-2 sm:grid-cols-3 gap-2">
                  <div><label className="label">Tipo</label>
                    <select className="select" value={s.tipo_susp} onChange={e=>setSusps(a=>a.map((x,j)=>j===i?{...x,tipo_susp:e.target.value}:x))}>
                      {TIPOS_SUSP.map(t=><option key={t.v} value={t.v}>{t.l}</option>)}
                    </select>
                  </div>
                  <div><label className="label">Hora inicio</label><input type="time" className="input" value={s.hora_inicio} onChange={e=>setSusps(a=>a.map((x,j)=>j===i?{...x,hora_inicio:e.target.value}:x))}/></div>
                  <div><label className="label">Hora fin</label><input type="time" className="input" value={s.hora_fin} onChange={e=>setSusps(a=>a.map((x,j)=>j===i?{...x,hora_fin:e.target.value}:x))}/></div>
                </div>
                {s.tipo_susp==='otro'&&<div><label className="label">¿Cuál otro motivo?</label><input className="input" value={s.otro_desc} onChange={e=>setSusps(a=>a.map((x,j)=>j===i?{...x,otro_desc:e.target.value}:x))} placeholder="Describe el motivo…"/></div>}
                <div><label className="label">Descripción adicional (opcional)</label><input className="input" value={s.descripcion} onChange={e=>setSusps(a=>a.map((x,j)=>j===i?{...x,descripcion:e.target.value}:x))} placeholder="Detalles…"/></div>
              </div>
            ))}
          </div>
          <button className="btn-secondary text-xs" onClick={()=>setSusps(a=>[...a,{uid:gid(),tipo_susp:'clima',otro_desc:'',hora_inicio:'',hora_fin:'',descripcion:''}])}>+ Agregar suspensión</button>
          {horasClima>0&&<p className="text-sm text-amber-600">⏱ {horasClima.toFixed(1)}h perdidas → {horasReal.toFixed(1)}h operativas</p>}
          <div className="flex gap-2"><button className="btn-secondary" onClick={()=>setStep(1)}>← Anterior</button><button className="btn-primary" onClick={()=>setStep(3)}>Siguiente →</button></div>
        </div>
      )}

      {step===3&&(
        <div className="card p-4 space-y-3">
          <h3 className="font-bold text-[#003b7a]">Paso 3 — Charla preturno</h3>
          <label className="flex items-center gap-3 cursor-pointer"><input type="checkbox" checked={charla} onChange={e=>setCharla(e.target.checked)} className="w-5 h-5"/><span className="text-sm font-medium">Se realizó la charla preturno</span></label>
          {charla&&<div><label className="label">Tema</label><input className="input" value={charlaTema} onChange={e=>setCharlaTema(e.target.value)} placeholder="Tema de la charla…"/></div>}
          <div className="flex gap-2"><button className="btn-secondary" onClick={()=>setStep(2)}>← Anterior</button><button className="btn-primary" onClick={()=>setStep(4)}>Siguiente →</button></div>
        </div>
      )}

      {step===4&&(
        <div className="card p-4 space-y-3">
          <h3 className="font-bold text-[#003b7a]">Paso 4 — Asistencia real</h3>
          <div className="flex gap-4 text-sm flex-wrap">
            <span>Planeado: <strong>{asistencia.length}</strong></span>
            <span className="text-emerald-600">Asistió: <strong>{asistio}</strong></span>
            <span>Cumplimiento: <strong>{efic}%</strong> {sem(efic)}</span>
          </div>
          <div className="space-y-2">
            {asistencia.map((a,i)=>(
              <div key={a.personal_id} className={`p-3 border rounded-lg space-y-2 ${a.es_adicional?'border-amber-200 bg-amber-50':'border-slate-200'}`}>
                <div className="flex items-center gap-3 flex-wrap">
                  <input type="checkbox" checked={a.asistio} onChange={e=>setAsistencia(arr=>arr.map((x,j)=>j===i?{...x,asistio:e.target.checked,motivo_ausencia:'',ausencia_parcial:false}:x))}/>
                  <div className="flex-1"><div className={`text-sm font-medium ${a.asistio?'':'text-slate-400 line-through'}`}>{a.nombre}{a.es_adicional&&<span className="ml-1 text-xs text-amber-600">★ Adicional</span>}</div><div className="text-xs text-slate-400">{a.cargo_es}</div></div>
                  {a.asistio&&<span className="text-xs text-emerald-600">{horasReal.toFixed(1)}h</span>}
                  {a.es_adicional&&<button className="text-rose-400 text-xs" onClick={()=>setAsistencia(arr=>arr.filter((_,j)=>j!==i))}>✕</button>}
                </div>
                {!a.asistio&&(
                  <div className="flex gap-2 flex-wrap ml-6">
                    <select className="select text-xs w-auto" value={a.motivo_ausencia} onChange={e=>setAsistencia(arr=>arr.map((x,j)=>j===i?{...x,motivo_ausencia:e.target.value}:x))}>
                      <option value="">— Motivo —</option>
                      <option value="injustificada">Injustificada</option>
                      <option value="incapacidad">Incapacidad</option>
                      <option value="permiso">Permiso</option>
                    </select>
                    <label className="flex items-center gap-1 text-xs cursor-pointer">
                      <input type="checkbox" checked={a.ausencia_parcial} onChange={e=>setAsistencia(arr=>arr.map((x,j)=>j===i?{...x,ausencia_parcial:e.target.checked}:x))}/>
                      Ausencia parcial
                    </label>
                    {a.ausencia_parcial&&<>
                      <input type="time" className="input text-xs w-24" value={a.hora_ausencia_ini} onChange={e=>setAsistencia(arr=>arr.map((x,j)=>j===i?{...x,hora_ausencia_ini:e.target.value}:x))} placeholder="Desde"/>
                      <input type="time" className="input text-xs w-24" value={a.hora_ausencia_fin} onChange={e=>setAsistencia(arr=>arr.map((x,j)=>j===i?{...x,hora_ausencia_fin:e.target.value}:x))} placeholder="Hasta"/>
                    </>}
                  </div>
                )}
                {a.asistio&&(
                  <div className="flex gap-2 flex-wrap ml-6">
                    <label className="flex items-center gap-1 text-xs cursor-pointer">
                      <input type="checkbox" checked={a.jornada_parcial} onChange={e=>setAsistencia(arr=>arr.map((x,j)=>j===i?{...x,jornada_parcial:e.target.checked}:x))}/>
                      Jornada parcial
                    </label>
                    {a.jornada_parcial&&<>
                      <input type="time" className="input text-xs w-24" value={a.hora_jornada_ini} onChange={e=>setAsistencia(arr=>arr.map((x,j)=>j===i?{...x,hora_jornada_ini:e.target.value}:x))} placeholder="Entrada"/>
                      <input type="time" className="input text-xs w-24" value={a.hora_jornada_fin} onChange={e=>setAsistencia(arr=>arr.map((x,j)=>j===i?{...x,hora_jornada_fin:e.target.value}:x))} placeholder="Salida"/>
                    </>}
                  </div>
                )}
              </div>
            ))}
          </div>
          {!asistencia.length&&<p className="text-sm text-slate-500">Sin personal planeado para este día.</p>}
          <div>
            <label className="label text-xs">+ Agregar personal adicional</label>
            <select className="select text-sm" value="" onChange={e=>{
              if(!e.target.value) return;
              const p=catalogs?.personal.find(x=>x.id===e.target.value);
              if(!p||asistencia.some(a=>a.personal_id===p.id)) return;
              setAsistencia(arr=>[...arr,{personal_id:p.id,documento_personal:p.documento,nombre:p.nombre,cargo_es:p.cargo_es,asistio:true,motivo_ausencia:'',ausencia_parcial:false,hora_ausencia_ini:'',hora_ausencia_fin:'',jornada_parcial:false,hora_jornada_ini:'',hora_jornada_fin:'',es_adicional:true}]);
            }}>
              <option value="">— Seleccionar personal —</option>
              {(catalogs?.personal||[]).filter(p=>p.activo!==false&&!asistencia.some(a=>a.personal_id===p.id)).map(p=><option key={p.id} value={p.id}>{p.nombre} — {p.cargo_es}</option>)}
            </select>
          </div>
          <div className="flex gap-2"><button className="btn-secondary" onClick={()=>setStep(3)}>← Anterior</button><button className="btn-primary" onClick={()=>setStep(5)}>Siguiente →</button></div>
        </div>
      )}

      {step===5&&(
        <div className="card p-4 space-y-3">
          <h3 className="font-bold text-[#003b7a]">🚜 Paso 5 — Maquinaria del día</h3>
          <p className="text-xs text-blue-700 bg-blue-50 border border-blue-200 rounded p-2">Maquinaria asignada en la planeación de hoy. Indica cuál se usó realmente.</p>
          {!maqDia.filter(m=>m.planeada).length&&<p className="text-sm text-amber-600 bg-amber-50 border border-amber-200 rounded p-2">⚠️ No había maquinaria planeada para hoy. Puedes agregar maquinaria adicional abajo.</p>}
          {maqDia.map((m,i)=>(
            <div key={m.uid} className={`border rounded-xl p-4 ${m.es_adicional?'border-amber-200 bg-amber-50':'border-slate-200'}`}>
              <div className="flex items-center justify-between mb-3">
                <div className="font-semibold text-[#003b7a]">{m.nombre}{m.es_adicional&&<span className="ml-2 text-xs text-amber-600 font-normal">★ Adicional</span>}</div>
                {m.es_adicional&&<button className="text-rose-400 text-sm" onClick={()=>setMaqDia(a=>a.filter((_,j)=>j!==i))}>✕ Quitar</button>}
              </div>
              <div className="flex gap-4 mb-2 items-center flex-wrap">
                <span className="text-sm text-slate-600">¿Se usó?</span>
                <label className="flex items-center gap-1 cursor-pointer text-sm">
                  <input type="radio" name={`uso-${m.uid}`} checked={m.uso==='si'} onChange={()=>setMaqDia(a=>a.map((x,j)=>j===i?{...x,uso:'si' as const}:x))}/> Sí
                </label>
                <label className="flex items-center gap-1 cursor-pointer text-sm">
                  <input type="radio" name={`uso-${m.uid}`} checked={m.uso==='no'} onChange={()=>setMaqDia(a=>a.map((x,j)=>j===i?{...x,uso:'no' as const,tiene_novedad:false,parcial:false}:x))}/> No
                </label>
              </div>
              {m.uso==='si'&&(
                <div className="space-y-2 mt-2">
                  <div className="flex gap-4 items-center flex-wrap">
                    <span className="text-sm text-slate-600">Jornada:</span>
                    <label className="flex items-center gap-1 cursor-pointer text-sm">
                      <input type="radio" name={`jorn-${m.uid}`} checked={!m.parcial} onChange={()=>setMaqDia(a=>a.map((x,j)=>j===i?{...x,parcial:false,hora_inicio:'',hora_fin:''}:x))}/> Completa
                    </label>
                    <label className="flex items-center gap-1 cursor-pointer text-sm">
                      <input type="radio" name={`jorn-${m.uid}`} checked={m.parcial} onChange={()=>setMaqDia(a=>a.map((x,j)=>j===i?{...x,parcial:true}:x))}/> Parcial
                    </label>
                    {m.parcial&&<>
                      <input type="time" className="input text-xs w-24" value={m.hora_inicio} onChange={e=>setMaqDia(a=>a.map((x,j)=>j===i?{...x,hora_inicio:e.target.value}:x))} placeholder="Inicio"/>
                      <input type="time" className="input text-xs w-24" value={m.hora_fin} onChange={e=>setMaqDia(a=>a.map((x,j)=>j===i?{...x,hora_fin:e.target.value}:x))} placeholder="Fin"/>
                    </>}
                  </div>
                  <label className="flex items-center gap-2 text-sm">
                    <input type="checkbox" checked={m.tiene_novedad} onChange={e=>setMaqDia(a=>a.map((x,j)=>j===i?{...x,tiene_novedad:e.target.checked}:x))}/>
                    ¿Tiene novedad o standby?
                  </label>
                  {m.tiene_novedad&&(
                    <textarea className="input w-full text-sm" rows={2} placeholder="Describe la novedad de esta máquina…"
                      value={m.novedad} onChange={e=>setMaqDia(a=>a.map((x,j)=>j===i?{...x,novedad:e.target.value}:x))}/>
                  )}
                </div>
              )}
            </div>
          ))}
          <div>
            <label className="label">+ Agregar maquinaria no planeada</label>
            <select className="select" value="" onChange={e=>{
              if(!e.target.value) return;
              const maqRow=maquinaria.find(m=>m.id===e.target.value);
              if(!maqRow||maqDia.some(m=>m.maquinaria_id===maqRow.id)) return;
              setMaqDia(a=>[...a,{uid:gid(),maquinaria_id:maqRow.id,nombre:maqRow.nombre||`${maqRow.item_id} (${maqRow.tipo})`,uso:'si' as const,planeada:false,es_adicional:true,parcial:false,hora_inicio:'',hora_fin:'',tiene_novedad:false,novedad:''}]);
            }}>
              <option value="">— Seleccionar equipo —</option>
              {maquinaria.filter(m=>m.estado==='activo'&&!maqDia.some(x=>x.maquinaria_id===m.id)).map(m=><option key={m.id} value={m.id}>{m.nombre||`${m.item_id} (${m.tipo})`}</option>)}
            </select>
          </div>
          <div className="flex gap-2"><button className="btn-secondary" onClick={()=>setStep(4)}>← Anterior</button><button className="btn-primary" onClick={()=>setStep(6)}>Siguiente →</button></div>
        </div>
      )}

      {step===6&&(
        <div className="card p-4 space-y-4">
          <h3 className="font-bold text-[#003b7a]">Paso 6 — Avance de actividades</h3>
          <p className="text-xs text-blue-700 bg-blue-50 border border-blue-200 rounded p-2">💡 Puedes reportar múltiples áreas por actividad. El personal no se bloquea entre técnicos en el reporte.</p>
          {actReps.map((ar,ai)=>{
            const actRow=catalogs?.especialidades_actividades.find(e=>e.id===ar.actividad_id);
            const cfg=configActs.find(c=>c.actividad_id===ar.actividad_id);

            // Tipo D — Cualitativa
            if(cfg?.tipo==='D') return(
              <div key={ar.uid} className="border border-purple-200 rounded-lg p-3 bg-purple-50">
                <div className="font-medium text-sm text-purple-800 mb-2">{actRow?.actividad_es} <span className="text-xs font-normal">— Actividad cualitativa</span></div>
                <label className="label">Descripción de lo ejecutado hoy</label>
                <textarea className="textarea" rows={3} value={ar.descripcion_cualitativa}
                  onChange={e=>setActReps(arr=>arr.map((x,j)=>j===ai?{...x,descripcion_cualitativa:e.target.value}:x))}
                  placeholder="Describe qué se hizo hoy en esta actividad…"/>
              </div>
            );

            if(cfg?.es_medible===false) return(
              <div key={ar.uid} className="border border-slate-200 rounded-lg p-3 bg-slate-50">
                <div className="text-sm text-slate-500">{actRow?.actividad_es} — <em>Actividad no medible</em></div>
              </div>
            );

            if(!cfg) return(
              <div key={ar.uid} className="border border-amber-200 rounded-lg p-3 bg-amber-50">
                <div className="text-sm text-amber-700 mb-2">⚠️ {actRow?.actividad_es} — Aún no configurada por el admin en <strong>Config. Act.</strong> Mientras tanto, describe aquí lo ejecutado para que no se pierda:</div>
                <label className="label">Descripción de lo ejecutado hoy</label>
                <textarea className="textarea" rows={3} value={ar.descripcion_cualitativa}
                  onChange={e=>setActReps(arr=>arr.map((x,j)=>j===ai?{...x,descripcion_cualitativa:e.target.value}:x))}
                  placeholder="Describe qué se hizo hoy en esta actividad…"/>
              </div>
            );

            return(
              <div key={ar.uid} className="border border-slate-200 rounded-lg p-3 space-y-3">
                <div className="font-medium text-sm text-[#003b7a]">{actRow?.actividad_es}
                  <span className="ml-2 text-xs text-slate-400">Tipo {cfg.tipo} · {cfg.unidad_es}{cfg.tiene_meta&&cfg.meta_total?` · Meta: ${cfg.meta_total}`:''}</span>
                </div>
                {ar.areas.map((area,areai)=>(
                  <div key={area.uid} className="flex gap-2 items-end flex-wrap bg-slate-50 p-2 rounded border border-slate-100">
                    <div className="flex-1 min-w-[120px]">
                      <label className="label">Área trabajada</label>
                      <select className="select" value={area.area_id}
                        onChange={e=>setActReps(arr=>arr.map((x,j)=>j===ai?{...x,areas:x.areas.map((a2,k)=>k===areai?{...a2,area_id:e.target.value}:a2)}:x))}>
                        <option value="">— Área —</option>
                        {(catalogs?.areas||[]).map(a=><option key={a.id} value={a.id}>{a.area_es}</option>)}
                      </select>
                    </div>
                    <div className="w-32">
                      <label className="label">Cantidad ({cfg.unidad_es})</label>
                      <input type="number" className="input" min={0} value={area.cantidad}
                        onChange={e=>setActReps(arr=>arr.map((x,j)=>j===ai?{...x,areas:x.areas.map((a2,k)=>k===areai?{...a2,cantidad:e.target.value}:a2)}:x))}/>
                    </div>
                    {ar.areas.length>1&&(
                      <button className="btn-ghost text-rose-500 text-xs pb-2"
                        onClick={()=>setActReps(arr=>arr.map((x,j)=>j===ai?{...x,areas:x.areas.filter((_,k)=>k!==areai)}:x))}>
                        ✕ Quitar
                      </button>
                    )}
                  </div>
                ))}
                <button className="btn-secondary text-xs"
                  onClick={()=>setActReps(arr=>arr.map((x,j)=>j===ai?{...x,areas:[...x.areas,{uid:gid(),area_id:'',cantidad:''}]}:x))}>
                  + Agregar otra área
                </button>
                <div><label className="label">Observación</label>
                  <textarea className="textarea" rows={1} value={ar.observacion_es}
                    onChange={e=>setActReps(arr=>arr.map((x,j)=>j===ai?{...x,observacion_es:e.target.value}:x))}/>
                </div>
                {/* Suspensiones de esta actividad */}
                <div className="border-t border-slate-100 pt-3">
                  <div className="flex items-center justify-between mb-2">
                    <h5 className="text-xs font-semibold text-slate-600">⚠️ Suspensiones de esta actividad</h5>
                    <button className="text-xs text-blue-600 hover:underline"
                      onClick={()=>setSuspPorActividad(p=>[...p,{uid:gid(),actividad_id:ar.actividad_id,tipo:'clima',otro_desc:'',parcial:false,hora_inicio:'',hora_fin:'',observacion:''}])}>
                      + Agregar suspensión
                    </button>
                  </div>
                  {suspPorActividad.filter(s=>s.actividad_id===ar.actividad_id).map((s)=>{
                    const si=suspPorActividad.findIndex(x=>x.uid===s.uid);
                    return(
                      <div key={s.uid} className="bg-amber-50 border border-amber-200 rounded-lg p-3 mb-2 space-y-2">
                        <div className="flex gap-2 flex-wrap">
                          <select className="select text-xs flex-1" value={s.tipo}
                            onChange={e=>setSuspPorActividad(p=>p.map((x,j)=>j===si?{...x,tipo:e.target.value}:x))}>
                            <option value="clima">🌧 Clima / lluvia</option>
                            <option value="logistica">📦 Logística</option>
                            <option value="protesta">🚧 Orden público</option>
                            <option value="mecanico">🔧 Falla mecánica</option>
                            <option value="otro">📝 Otro</option>
                          </select>
                          {s.tipo==='otro'&&<input className="input text-xs flex-1" placeholder="Especificar…" value={s.otro_desc} onChange={e=>setSuspPorActividad(p=>p.map((x,j)=>j===si?{...x,otro_desc:e.target.value}:x))}/>}
                        </div>
                        <div className="flex gap-3 items-center flex-wrap">
                          <label className="flex items-center gap-1 text-xs cursor-pointer"><input type="radio" checked={!s.parcial} onChange={()=>setSuspPorActividad(p=>p.map((x,j)=>j===si?{...x,parcial:false}:x))}/> Todo el día</label>
                          <label className="flex items-center gap-1 text-xs cursor-pointer"><input type="radio" checked={s.parcial} onChange={()=>setSuspPorActividad(p=>p.map((x,j)=>j===si?{...x,parcial:true}:x))}/> Parcial</label>
                          {s.parcial&&<>
                            <input type="time" className="input text-xs w-24" value={s.hora_inicio} onChange={e=>setSuspPorActividad(p=>p.map((x,j)=>j===si?{...x,hora_inicio:e.target.value}:x))}/>
                            <input type="time" className="input text-xs w-24" value={s.hora_fin} onChange={e=>setSuspPorActividad(p=>p.map((x,j)=>j===si?{...x,hora_fin:e.target.value}:x))}/>
                          </>}
                        </div>
                        <textarea className="input w-full text-xs" rows={1} placeholder="Observación…" value={s.observacion} onChange={e=>setSuspPorActividad(p=>p.map((x,j)=>j===si?{...x,observacion:e.target.value}:x))}/>
                        <button className="text-xs text-rose-500 hover:underline" onClick={()=>setSuspPorActividad(p=>p.filter(x=>x.uid!==s.uid))}>Quitar suspensión</button>
                      </div>
                    );
                  })}
                </div>
                {cfg?.tiene_items_unicos&&(()=>{
                  const dbEntry=itemDbs[cfg.id];
                  if(!dbEntry) return <p className="text-xs text-amber-600">Cargando base de datos de ítems…</p>;
                  const srch=(itemDbSearch[cfg.id]||'').toLowerCase();
                  const totalCantidad=ar.areas.reduce((s,a)=>s+(parseFloat(a.cantidad)||0),0);
                  const disponibles=dbEntry.items.filter(it=>{
                    if(it.bloqueado&&it.bloqueado_fecha!==fecha) return false;
                    if(it.bloqueado&&!it.bloqueado_fecha) return false;
                    return true;
                  });
                  const filtrados=disponibles.filter(it=>srch?Object.values(it.datos).some(v=>String(v).toLowerCase().includes(srch)):true);
                  const bloqueados=dbEntry.items.filter(it=>it.bloqueado&&(!it.bloqueado_fecha||it.bloqueado_fecha!==fecha));
                  const cols=dbEntry.db.columnas.slice(0,3);
                  return(
                    <div className="border border-indigo-200 rounded-lg p-3 bg-indigo-50 space-y-2">
                      <div className="flex items-center justify-between flex-wrap gap-2">
                        <span className="text-xs font-semibold text-indigo-800">🗂️ {dbEntry.db.nombre}</span>
                        <span className={`text-xs px-2 py-0.5 rounded-full font-medium ${ar.items_seleccionados.length===totalCantidad&&totalCantidad>0?'bg-emerald-100 text-emerald-700':'bg-amber-100 text-amber-700'}`}>
                          Seleccionados: {ar.items_seleccionados.length} / {totalCantidad}
                        </span>
                      </div>
                      <input className="input text-xs" placeholder="🔍 Buscar ítem…" value={itemDbSearch[cfg.id]||''} onChange={e=>setItemDbSearch(p=>({...p,[cfg.id]:e.target.value}))}/>
                      <div className="max-h-48 overflow-y-auto space-y-1">
                        {filtrados.map(it=>{
                          const sel=ar.items_seleccionados.includes(it.id);
                          return(
                            <label key={it.id} className={`flex items-center gap-2 p-2 rounded border cursor-pointer text-xs transition-colors ${sel?'bg-indigo-100 border-indigo-400':'bg-white border-slate-200 hover:border-indigo-300'}`}>
                              <input type="checkbox" checked={sel} onChange={()=>{
                                setActReps(arr=>arr.map((x,j)=>j===ai?{...x,items_seleccionados:sel?x.items_seleccionados.filter(id=>id!==it.id):[...x.items_seleccionados,it.id]}:x));
                              }}/>
                              <span className="flex-1">{cols.map(c=>it.datos[c.nombre]).filter(Boolean).join(' — ')}</span>
                              <span className="text-emerald-600">disponible</span>
                            </label>
                          );
                        })}
                        {bloqueados.slice(0,3).map(it=>(
                          <div key={it.id} className="flex items-center gap-2 p-2 rounded border border-slate-200 bg-slate-50 text-xs text-slate-400">
                            <span className="w-4">🔒</span>
                            <span className="flex-1">{cols.map(c=>it.datos[c.nombre]).filter(Boolean).join(' — ')}</span>
                            <span>ya ejecutado</span>
                          </div>
                        ))}
                        {!filtrados.length&&!bloqueados.length&&<p className="text-xs text-slate-400 text-center py-2">Sin ítems disponibles.</p>}
                      </div>
                    </div>
                  );
                })()}
              </div>
            );
          })}
          {!actReps.length&&<p className="text-sm text-slate-500">Selecciona una especialidad en el Paso 1.</p>}

          {/* Actividades adicionales no planeadas */}
          <div className="mt-4 border-t border-slate-200 pt-4">
            <div className="flex items-center justify-between mb-3">
              <h4 className="font-bold text-[#003b7a] text-sm">➕ Actividades adicionales (no planeadas)</h4>
              <button className="btn-secondary text-sm" onClick={()=>setMostrarFormAdicional(true)}>+ Agregar actividad adicional</button>
            </div>
            {mostrarFormAdicional&&(
              <div className="border border-blue-200 bg-blue-50 rounded-xl p-4 mb-3 space-y-3">
                <div>
                  <label className="label">Actividad realizada</label>
                  <input list="catalogo-adicionales" className="input w-full" placeholder="Escribe o selecciona una actividad…"
                    value={nuevaAdicionalNombre} onChange={e=>setNuevaAdicionalNombre(e.target.value)}/>
                  <datalist id="catalogo-adicionales">
                    {catAdicionales.map(c=><option key={c.id} value={c.nombre}/>)}
                  </datalist>
                  {catAdicionales.some(c=>c.nombre.toLowerCase().includes(nuevaAdicionalNombre.toLowerCase())&&nuevaAdicionalNombre.length>2)&&(
                    <div className="text-xs text-blue-600 mt-1">💡 Selecciona del listado para mantener consistencia de nombres</div>
                  )}
                </div>
                <div>
                  <label className="label">Descripción de lo ejecutado</label>
                  <textarea className="textarea w-full" rows={2} placeholder="Describe qué se hizo exactamente…"
                    value={nuevaAdicionalDesc} onChange={e=>setNuevaAdicionalDesc(e.target.value)}/>
                </div>
                <div className="flex gap-2">
                  <button className="btn-primary text-sm" disabled={!nuevaAdicionalNombre.trim()||!nuevaAdicionalDesc.trim()} onClick={async()=>{
                    const nombre=nuevaAdicionalNombre.trim();
                    const existente=catAdicionales.find(c=>c.nombre.toLowerCase()===nombre.toLowerCase());
                    let catalogoId:string|null=existente?.id||null;
                    if(existente){
                      await supabase.from('actividades_adicionales_catalogo').update({veces_usada:existente.veces_usada+1}).eq('id',existente.id);
                    } else {
                      const{data:nueva}=await supabase.from('actividades_adicionales_catalogo').insert({nombre,veces_usada:1,activo:true}).select().single();
                      if(nueva){ catalogoId=(nueva as Record<string,unknown>).id as string; setCatAdicionales(p=>[...p,nueva as ActAdicCat]); }
                    }
                    setActAdicionales(p=>[...p,{uid:gid(),nombre,descripcion:nuevaAdicionalDesc.trim(),catalogoId}]);
                    setNuevaAdicionalNombre('');setNuevaAdicionalDesc('');setMostrarFormAdicional(false);
                  }}>Guardar actividad adicional</button>
                  <button className="btn-secondary text-sm" onClick={()=>{setMostrarFormAdicional(false);setNuevaAdicionalNombre('');setNuevaAdicionalDesc('');}}>Cancelar</button>
                </div>
              </div>
            )}
            {actAdicionales.map((ad,ai)=>(
              <div key={ad.uid} className="flex items-start justify-between bg-white border border-slate-200 rounded-lg p-3 mb-2">
                <div>
                  <div className="font-medium text-sm text-[#003b7a]">{ad.nombre}</div>
                  <div className="text-xs text-slate-500 mt-0.5">{ad.descripcion}</div>
                </div>
                <button className="text-rose-500 text-xs hover:underline ml-3 shrink-0" onClick={()=>setActAdicionales(a=>a.filter((_,j)=>j!==ai))}>Quitar</button>
              </div>
            ))}
            {!actAdicionales.length&&!mostrarFormAdicional&&<div className="text-xs text-slate-400 italic">Sin actividades adicionales este día.</div>}
          </div>

          <div className="flex gap-2"><button className="btn-secondary" onClick={()=>setStep(5)}>← Anterior</button><button className="btn-primary" onClick={()=>setStep(7)}>Siguiente →</button></div>
        </div>
      )}

      {step===7&&(
        <div className="card p-4 space-y-3">
          <h3 className="font-bold text-[#003b7a]">Paso 7 — Seguridad</h3>
          <div className="flex gap-2 flex-wrap">
            {['sin_novedad','casi_accidente','incidente','accidente'].map(t=>(
              <button key={t} onClick={()=>setIncidente(i=>({...i,tipo:t}))}
                className={`text-xs px-3 py-1.5 rounded-full border transition-colors ${incidente.tipo===t?t==='sin_novedad'?'bg-emerald-500 text-white border-emerald-500':t==='accidente'?'bg-rose-600 text-white border-rose-600':'bg-amber-500 text-white border-amber-500':'border-slate-300 text-slate-600'}`}>
                {t==='sin_novedad'?'✅ Sin novedad':t==='casi_accidente'?'⚠️ Casi accidente':t==='incidente'?'🔶 Incidente':'🚨 Accidente'}
              </button>
            ))}
          </div>
          {incidente.tipo!=='sin_novedad'&&(
            <>
              <div><label className="label">Descripción</label><textarea className="textarea" rows={2} value={incidente.descripcion} onChange={e=>setIncidente(i=>({...i,descripcion:e.target.value}))}/></div>
              <div><label className="label">Medidas tomadas</label><textarea className="textarea" rows={2} value={incidente.medidas} onChange={e=>setIncidente(i=>({...i,medidas:e.target.value}))}/></div>
            </>
          )}
          <div><label className="label">Nota de bitácora (opcional)</label><textarea className="textarea" rows={2} value={notaBit} onChange={e=>setNotaBit(e.target.value)} placeholder="Decisiones importantes del día…"/></div>
          <div className="flex gap-2"><button className="btn-secondary" onClick={()=>setStep(6)}>← Anterior</button><button className="btn-primary" onClick={()=>setStep(8)}>Siguiente →</button></div>
        </div>
      )}

      {step===8&&(
        <div className="card p-4 space-y-4">
          <h3 className="font-bold text-[#003b7a]">Paso 8 — Resumen y envío</h3>
          <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">
            <div className="card p-3 text-center"><div className="text-2xl font-bold text-[#003b7a]">{asistio}</div><div className="text-xs text-slate-500">Personal activo</div></div>
            <div className="card p-3 text-center"><div className="text-2xl font-bold text-emerald-600">{(asistio*horasReal).toFixed(0)}h</div><div className="text-xs text-slate-500">Horas-hombre</div></div>
            <div className="card p-3 text-center"><div className="text-2xl font-bold text-amber-500">{horasClima.toFixed(1)}h</div><div className="text-xs text-slate-500">Perdidas</div></div>
            <div className="card p-3 text-center"><div className="text-2xl font-bold text-[#003b7a]">{actReps.filter(a=>a.areas.some(ar=>parseFloat(ar.cantidad)>0)||a.descripcion_cualitativa.trim().length>0).length}</div><div className="text-xs text-slate-500">Actividades</div></div>
          </div>
          <div className="bg-slate-50 rounded-lg p-3 text-sm space-y-1">
            <div><strong>Fecha:</strong> {fecha}</div>
            <div><strong>Clima:</strong> {clima}</div>
            <div><strong>Asistencia:</strong> {asistio}/{asistencia.length} ({efic}%) {sem(efic)}</div>
            <div><strong>Suspensiones:</strong> {susps.length} ({horasClima.toFixed(1)}h perdidas)</div>
            <div><strong>Incidente:</strong> {incidente.tipo}</div>
          </div>
          <div className="flex gap-2">
            <button className="btn-secondary" onClick={()=>setStep(7)}>← Anterior</button>
            <button className="btn-success" disabled={saving} onClick={submit}>{saving?'Enviando…':'✅ Enviar reporte'}</button>
          </div>
        </div>
      )}
    </div>
  );
}

// ── APROBACIÓN ────────────────────────────────────────────────────
function AprobacionModule({user,catalogs,configActs,showToast,onRefreshNotifs}:{
  user:Profile; catalogs:Catalogs|null; configActs:ConfigAct[];
  showToast:(k:'ok'|'err'|'info',m:string)=>void; onRefreshNotifs:()=>void;
}){
  const[fecha,setFecha]=useState(today());
  const[espId,setEspId]=useState('');
  const[reportes,setReportes]=useState<Record<string,unknown>[]>([]);
  const[avances,setAvances]=useState<Record<string,unknown>[]>([]);
  const[cualAprob,setCualAprob]=useState<Record<string,unknown>[]>([]);
  const[adicAprob,setAdicAprob]=useState<Record<string,unknown>[]>([]);
  const[aprobadas,setAprobadas]=useState<Set<string>>(new Set());
  const[rechazadas,setRechazadas]=useState<Set<string>>(new Set());
  const[motivos,setMotivos]=useState<Record<string,string>>({});
  const[saving,setSaving]=useState(false);
  const[loading,setLoading]=useState(false);
  const[pendientes,setPendientes]=useState<Record<string,unknown>[]>([]);
  const[loadingPend,setLoadingPend]=useState(false);
  const espList=useMemo(()=>catalogs?uniqueEsp(catalogs.especialidades_actividades):[],[catalogs]);

  async function cargar(){
    if(!fecha||!espId){showToast('err','Selecciona fecha y especialidad');return;}
    setLoading(true);
    try{
      const{data:reps}=await supabase.from('reportes_avance').select('*').eq('fecha',fecha).eq('especialidad_id',espId).eq('estado','borrador');
      const repIds=(reps||[]).map((r:Record<string,unknown>)=>r.id as string);
      const avs=repIds.length?(await supabase.from('avance_diario').select('*').in('reporte_id',repIds)).data||[]:[];
      const adics=repIds.length?(await supabase.from('actividades_adicionales_reporte').select('*').in('reporte_id',repIds)).data||[]:[];
      const avD=avs as Record<string,unknown>[];
      setReportes(reps||[]);
      setAvances(avD.filter(a=>a.unidad!=='cualitativo'));
      setCualAprob(avD.filter(a=>a.unidad==='cualitativo'));
      setAdicAprob(adics as Record<string,unknown>[]);
      setAprobadas(new Set()); setRechazadas(new Set()); setMotivos({});
    } catch(e:unknown){ showToast('err',(e as Error)?.message||'Error'); }
    finally{ setLoading(false); }
  }

  async function cargarPendientes(){
    setLoadingPend(true);
    try{
      const{data:pends}=await supabase.from('reportes_avance').select('id,fecha,especialidad_id,estado,usuario_nombre,created_at,especialidades_actividades(especialidad_es)').in('estado',['borrador','enviado']).order('fecha',{ascending:false});
      setPendientes((pends||[]) as Record<string,unknown>[]);
    } catch{ setPendientes([]); }
    finally{ setLoadingPend(false); }
  }
  useEffect(()=>{cargarPendientes();},[]);

  async function aprobarMultiples(ids:string[]){
    if(!window.confirm(`¿Aprobar ${ids.length} reportes enviados?`)) return;
    await supabase.from('reportes_avance').update({estado:'aprobado',aprobado_por:user.id,aprobado_en:new Date().toISOString()}).in('id',ids);
    await supabase.from('aprobacion_informes').insert(ids.map(id=>({reporte_id:id,aprobado_por:user.id,estado:'aprobado',version:1})));
    showToast('ok',`${ids.length} reportes aprobados`);
    cargarPendientes();
  }

  const actividadesConReps=useMemo(()=>{
    const acts=new Map<string,{actividad_id:string;cuadrillas:{reporte_id:string;usuario_nombre:string;areas:{area_id:string;cantidad:number;unidad:string}[];total:number}[]}>();
    avances.forEach((av)=>{
      const actId=av.actividad_id as string;
      if(!acts.has(actId)) acts.set(actId,{actividad_id:actId,cuadrillas:[]});
      const entry=acts.get(actId)!;
      const rep=reportes.find(r=>r.id===av.reporte_id);
      const repId=av.reporte_id as string;
      let cuad=entry.cuadrillas.find(c=>c.reporte_id===repId);
      if(!cuad){cuad={reporte_id:repId,usuario_nombre:(rep?.usuario_nombre as string)||'Técnico',areas:[],total:0};entry.cuadrillas.push(cuad);}
      cuad.areas.push({area_id:av.area_id as string,cantidad:parseFloat(String(av.cantidad||0)),unidad:av.unidad as string});
      cuad.total+=parseFloat(String(av.cantidad||0));
    });
    reportes.forEach(r=>{
      if(!avances.some(av=>av.reporte_id===r.id)){
        if(!acts.has('sin_avance')) acts.set('sin_avance',{actividad_id:'',cuadrillas:[]});
        acts.get('sin_avance')!.cuadrillas.push({reporte_id:r.id as string,usuario_nombre:r.usuario_nombre as string,areas:[],total:0});
      }
    });
    return Array.from(acts.values());
  },[avances,reportes]);

  const consolidado=useMemo(()=>{
    const c:Record<string,number>={};
    actividadesConReps.forEach(act=>{
      act.cuadrillas.filter(cu=>aprobadas.has(cu.reporte_id)).forEach(cu=>{c[act.actividad_id]=(c[act.actividad_id]||0)+cu.total;});
    });
    return c;
  },[actividadesConReps,aprobadas]);

  function toggleAp(id:string){setAprobadas(p=>{const n=new Set(p);if(n.has(id))n.delete(id);else{n.add(id);setRechazadas(r=>{const nr=new Set(r);nr.delete(id);return nr;});}return n;});}
  function toggleRe(id:string){setRechazadas(p=>{const n=new Set(p);if(n.has(id))n.delete(id);else{n.add(id);setAprobadas(a=>{const na=new Set(a);na.delete(id);return na;});}return n;});}

  async function enviar(){
    const toAp=Array.from(aprobadas);const toRe=Array.from(rechazadas);
    if(!toAp.length&&!toRe.length){showToast('err','Sin decisiones');return;}
    for(const id of toRe){if(!motivos[id]){showToast('err','Indica el motivo de rechazo para todos los rechazados');return;}}
    if(!window.confirm(`¿Confirmar? ${toAp.length} aprobados · ${toRe.length} rechazados`)) return;
    setSaving(true);
    try{
      if(toAp.length){
        await supabase.from('reportes_avance').update({estado:'aprobado',aprobado_por:user.id,aprobado_en:new Date().toISOString()}).in('id',toAp);
        await supabase.from('aprobacion_informes').insert(toAp.map(id=>({reporte_id:id,aprobado_por:user.id,estado:'aprobado',version:1})));
        const rA=reportes.filter(r=>toAp.includes(r.id as string));
        if(rA.length) try{await supabase.from('notificaciones').insert(rA.map(r=>({usuario_id:r.usuario_id,tipo:'aprobado',titulo:'Reporte aprobado ✅',mensaje:`Tu reporte del ${fecha} fue aprobado por ${user.nombre}`,data:{reporte_id:r.id}})));}catch{}
      }
      if(toRe.length){
        await supabase.from('aprobacion_informes').insert(toRe.map(id=>({reporte_id:id,aprobado_por:user.id,estado:'rechazado',version:1,comentarios:motivos[id]||''})));
        const rR=reportes.filter(r=>toRe.includes(r.id as string));
        if(rR.length) try{await supabase.from('notificaciones').insert(rR.map(r=>({usuario_id:r.usuario_id,tipo:'rechazado',titulo:'Reporte rechazado ❌',mensaje:`Tu reporte del ${fecha} fue rechazado. Motivo: ${motivos[r.id as string]||'Ver detalles'}`,data:{reporte_id:r.id}})));}catch{}
      }
      showToast('ok',`✅ ${toAp.length} aprobados · ❌ ${toRe.length} rechazados`);
      onRefreshNotifs(); await cargar();
    } catch(e:unknown){ showToast('err',(e as Error)?.message||'Error'); }
    finally{ setSaving(false); }
  }

  const areaMap=useMemo(()=>{const m:Record<string,string>={};(catalogs?.areas||[]).forEach(a=>{m[a.id]=a.area_es;});return m;},[catalogs]);

  return(
    <div className="space-y-4">
      {/* REPORTES PENDIENTES */}
      <div className="card p-4">
        <div className="flex items-center justify-between mb-3">
          <h3 className="font-bold text-[#003b7a]">🕐 Reportes pendientes de aprobación{pendientes.length>0&&<span className="ml-2 bg-red-500 text-white text-xs font-bold px-2 py-0.5 rounded-full">{pendientes.length}</span>}</h3>
          <button className="btn-secondary text-xs" onClick={cargarPendientes}>🔄 Actualizar</button>
        </div>
        {loadingPend&&<div className="text-sm text-slate-500">Cargando...</div>}
        {!loadingPend&&pendientes.length===0&&<div className="text-sm text-slate-500 italic text-center py-4">✅ No hay reportes pendientes de aprobación</div>}
        {!loadingPend&&pendientes.length>0&&(
          <div className="space-y-2">
            {pendientes.map(rep=>{
              const esp=(rep.especialidades_actividades as Record<string,string>)?.especialidad_es||'—';
              const estado=rep.estado as string;
              return(
                <div key={rep.id as string} className="flex items-center justify-between border border-slate-200 rounded-xl p-3 hover:border-[#003b7a] hover:bg-blue-50 transition-all cursor-pointer" onClick={()=>{setFecha(rep.fecha as string);setEspId(rep.especialidad_id as string);}}>
                  <div>
                    <div className="font-bold text-sm text-[#003b7a]">{esp}</div>
                    <div className="text-xs text-slate-500 mt-0.5">📅 {rep.fecha as string} · 👤 {rep.usuario_nombre as string}</div>
                  </div>
                  <div className="flex items-center gap-2">
                    <span className={`text-xs font-bold px-2 py-1 rounded-full ${estado==='enviado'?'bg-blue-100 text-blue-700':'bg-amber-100 text-amber-700'}`}>{estado==='enviado'?'📤 Enviado':'📝 Borrador'}</span>
                    <span className="text-xs text-blue-600 font-medium">Ver →</span>
                  </div>
                </div>
              );
            })}
          </div>
        )}
        {pendientes.filter(p=>(p as Record<string,unknown>).estado==='enviado').length>1&&(
          <div className="mt-3 pt-3 border-t border-slate-200">
            <button className="btn-primary text-sm w-full" onClick={()=>{const ids=pendientes.filter(p=>(p as Record<string,unknown>).estado==='enviado').map(p=>(p as Record<string,unknown>).id as string);aprobarMultiples(ids);}}>
              ✅ Aprobar todos los enviados ({pendientes.filter(p=>(p as Record<string,unknown>).estado==='enviado').length})
            </button>
          </div>
        )}
      </div>

      <div className="card p-4 flex gap-3 items-end flex-wrap">
        <div className="flex-1 min-w-[140px]"><label className="label">Fecha</label><input type="date" className="input" value={fecha} onChange={e=>setFecha(e.target.value)}/></div>
        <div className="flex-1 min-w-[180px]"><label className="label">Especialidad</label><select className="select" value={espId} onChange={e=>setEspId(e.target.value)}><option value="">— Seleccionar —</option>{espList.map(e=><option key={e.id} value={e.id}>{e.especialidad_es}</option>)}</select></div>
        <button className="btn-primary" onClick={cargar} disabled={loading}>{loading?'Cargando…':'Cargar reportes'}</button>
      </div>
      {reportes.length>0&&(
        <>
          <div className="card p-4 border-2 border-[#003b7a]">
            <h3 className="font-bold text-[#003b7a] mb-3">📊 Consolidado en tiempo real</h3>
            <div className="flex items-center gap-4 mb-4">
              <span className="text-sm font-medium"><strong>{aprobadas.size}</strong> de <strong>{reportes.length}</strong> aprobados</span>
              <div className="flex-1 progress-bar"><div className="progress-fill-blue transition-all duration-300" style={{width:`${reportes.length>0?Math.round(aprobadas.size/reportes.length*100):0}%`}}/></div>
              <span className="text-sm font-bold text-[#003b7a]">{reportes.length>0?Math.round(aprobadas.size/reportes.length*100):0}%</span>
            </div>
            <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-3">
              {actividadesConReps.filter(a=>a.actividad_id).map(act=>{
                const actRow=catalogs?.especialidades_actividades.find(e=>e.id===act.actividad_id);
                const cfg=configActs.find(c=>c.actividad_id===act.actividad_id);
                const total=consolidado[act.actividad_id]||0;
                const pct=cfg?.meta_total&&cfg.tiene_meta?Math.min(100,Math.round(total/cfg.meta_total*100)):null;
                return(
                  <div key={act.actividad_id} className="bg-[#e6eef7] rounded-lg p-3">
                    <div className="text-sm font-semibold text-[#003b7a] truncate">{actRow?.actividad_es||act.actividad_id}</div>
                    <div className="text-2xl font-bold text-[#003b7a] my-1">{total.toFixed(0)} <span className="text-sm font-normal text-slate-500">{cfg?.unidad_es||'und'}</span></div>
                    {pct!==null&&(<><div className="progress-bar"><div className="progress-fill-blue" style={{width:`${pct}%`}}/></div><div className="text-xs text-slate-500 mt-1">{pct}% de {cfg?.meta_total} {cfg?.unidad_es}</div></>)}
                    <div className="text-xs text-emerald-600 mt-1">{act.cuadrillas.filter(c=>aprobadas.has(c.reporte_id)).length}/{act.cuadrillas.length} cuadrillas</div>
                  </div>
                );
              })}
            </div>
          </div>
          {actividadesConReps.map(act=>{
            const actRow=catalogs?.especialidades_actividades.find(e=>e.id===act.actividad_id);
            const cfg=configActs.find(x=>x.actividad_id===act.actividad_id);
            return(
              <div key={act.actividad_id||'sin'} className="card p-4">
                <h4 className="font-bold text-[#003b7a] mb-3">{actRow?.actividad_es||'Reportes sin avance'}</h4>
                <div className="space-y-3">
                  {act.cuadrillas.map(c=>{
                    const isAp=aprobadas.has(c.reporte_id);const isRe=rechazadas.has(c.reporte_id);
                    return(
                      <div key={c.reporte_id} className={`border rounded-lg p-3 transition-all ${isAp?'border-emerald-400 bg-emerald-50':isRe?'border-rose-400 bg-rose-50':'border-slate-200'}`}>
                        <div className="flex items-start justify-between gap-2 flex-wrap mb-2">
                          <div>
                            <div className="font-medium text-sm">{c.usuario_nombre}</div>
                            {c.areas.length>0&&<div className="text-xs text-slate-500 mt-0.5">{c.areas.map(a=>`${areaMap[a.area_id]||a.area_id}: ${a.cantidad} ${a.unidad}`).join(' · ')}</div>}
                            {c.total>0&&<div className="text-base font-bold text-[#003b7a] mt-1">Total: {c.total} {cfg?.unidad_es||'und'}</div>}
                          </div>
                          <div className="flex gap-4">
                            <label className="flex items-center gap-1.5 cursor-pointer select-none"><input type="checkbox" checked={isAp} onChange={()=>toggleAp(c.reporte_id)} className="w-4 h-4 accent-emerald-600"/><span className="text-xs font-semibold text-emerald-700">✓ Aprobar</span></label>
                            <label className="flex items-center gap-1.5 cursor-pointer select-none"><input type="checkbox" checked={isRe} onChange={()=>toggleRe(c.reporte_id)} className="w-4 h-4 accent-rose-600"/><span className="text-xs font-semibold text-rose-700">✗ Rechazar</span></label>
                          </div>
                        </div>
                        {cualAprob.filter(a=>a.reporte_id===c.reporte_id).length>0&&(
                          <div className="mt-2 space-y-1">
                            <div className="text-xs font-semibold text-purple-700 mb-1">📋 Actividades cualitativas</div>
                            {cualAprob.filter(a=>a.reporte_id===c.reporte_id).map((a,i)=>{
                              const aRq=catalogs?.especialidades_actividades.find(e=>e.id===(a.actividad_id as string));
                              return(<div key={i} className="bg-purple-50 border border-purple-200 rounded p-2">
                                <div className="font-semibold text-xs text-purple-800">{aRq?.actividad_es||'Actividad'}</div>
                                <div className="text-xs text-slate-600 mt-0.5">{a.observacion_es as string}</div>
                              </div>);
                            })}
                          </div>
                        )}
                        {adicAprob.filter(a=>a.reporte_id===c.reporte_id).length>0&&(
                          <div className="mt-2 space-y-1">
                            <div className="text-xs font-semibold text-indigo-700 mb-1">➕ Actividades adicionales</div>
                            {adicAprob.filter(a=>a.reporte_id===c.reporte_id).map((a,i)=>(
                              <div key={i} className="bg-indigo-50 border border-indigo-200 rounded p-2">
                                <div className="font-semibold text-xs text-indigo-800">{a.nombre as string}</div>
                                <div className="text-xs text-slate-600 mt-0.5">{a.descripcion_ejecutado as string}</div>
                              </div>
                            ))}
                          </div>
                        )}
                        {isRe&&<input className="input text-xs mt-2" placeholder="Motivo del rechazo (obligatorio)…" value={motivos[c.reporte_id]||''} onChange={e=>setMotivos(m=>({...m,[c.reporte_id]:e.target.value}))}/>}
                      </div>
                    );
                  })}
                </div>
              </div>
            );
          })}
          <div className="card p-4 flex gap-3 flex-wrap items-center justify-between">
            <div className="text-sm text-slate-600"><strong>{aprobadas.size}</strong> aprobar · <strong>{rechazadas.size}</strong> rechazar</div>
            <button className="btn-success" disabled={saving||(!aprobadas.size&&!rechazadas.size)} onClick={enviar}>{saving?'Procesando…':'📤 Enviar decisiones'}</button>
          </div>
        </>
      )}
      {!loading&&!reportes.length&&espId&&<div className="card p-6 text-center text-slate-500">Sin reportes pendientes.</div>}
      {!espId&&<div className="card p-6 text-center text-slate-500">Selecciona una especialidad y fecha.</div>}
    </div>
  );
}

// ── SOLICITUDES ───────────────────────────────────────────────────
function SolicitudesModule({user,catalogs,showToast}:{user:Profile;catalogs:Catalogs|null;showToast:(k:'ok'|'err'|'info',m:string)=>void}){
  const[solicitudes,setSolicitudes]=useState<Record<string,unknown>[]>([]);
  const[form,setForm]=useState({fecha_reporte:'',especialidad_id:'',motivo:''});
  const[saving,setSaving]=useState(false);
  const canApprove=user.rol==='admin'||user.rol==='lider';
  const espList=useMemo(()=>catalogs?uniqueEsp(catalogs.especialidades_actividades):[],[catalogs]);

  async function load(){
    try{
      let q=supabase.from('solicitudes_reporte_pasado').select('*').order('created_at',{ascending:false});
      if(!canApprove) q=q.eq('tecnico_id',user.id);
      const{data}=await q;setSolicitudes((data||[]) as Record<string,unknown>[]);
    } catch{ setSolicitudes([]); }
  }
  useEffect(()=>{load();},[]);

  async function crear(){
    if(!form.fecha_reporte||!form.especialidad_id||!form.motivo){showToast('err','Completa todos los campos');return;}
    if(form.fecha_reporte>=today()){showToast('err','Solo para días anteriores');return;}
    setSaving(true);
    try{
      const{error}=await supabase.from('solicitudes_reporte_pasado').insert({tecnico_id:user.id,tecnico_nombre:user.nombre,fecha_reporte:form.fecha_reporte,especialidad_id:form.especialidad_id,motivo:form.motivo});
      if(error) throw error;
      try{const{data:admins}=await supabase.from('profiles').select('id').in('rol',['admin','lider']);if(admins?.length)await supabase.from('notificaciones').insert((admins as {id:string}[]).map(a=>({usuario_id:a.id,tipo:'solicitud',titulo:'Solicitud de reporte pasado',mensaje:`${user.nombre} solicita reportar el ${form.fecha_reporte}`,data:{}})));}catch{}
      showToast('ok','Solicitud enviada.');setForm({fecha_reporte:'',especialidad_id:'',motivo:''});await load();
    } catch(e:unknown){ showToast('err',(e as Error)?.message||'Error'); }
    finally{ setSaving(false); }
  }

  async function decidir(id:string,estado:'aprobado'|'rechazado'){
    const mot=estado==='rechazado'?window.prompt('Motivo del rechazo:'):null;
    if(estado==='rechazado'&&!mot) return;
    await supabase.from('solicitudes_reporte_pasado').update({estado,aprobado_por:user.id,aprobado_en:new Date().toISOString(),comentario:mot||''}).eq('id',id);
    const sol=solicitudes.find(s=>s.id===id);
    if(sol) try{await supabase.from('notificaciones').insert({usuario_id:sol.tecnico_id,tipo:`solicitud_${estado}`,titulo:`Solicitud ${estado==='aprobado'?'aprobada ✅':'rechazada ❌'}`,mensaje:`Tu solicitud para reportar el ${sol.fecha_reporte as string} fue ${estado}`,data:{}});}catch{}
    showToast('ok',`Solicitud ${estado}`);await load();
  }

  return(
    <div className="space-y-4">
      {(user.rol==='tecnico'||user.rol==='lider')&&(
        <div className="card p-4 space-y-3">
          <h3 className="font-bold text-[#003b7a]">📝 Solicitar permiso — reportar día anterior</h3>
          <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
            <div><label className="label">Fecha anterior</label><input type="date" className="input" value={form.fecha_reporte} max={new Date(Date.now()-86400000).toISOString().split('T')[0]} onChange={e=>setForm({...form,fecha_reporte:e.target.value})}/></div>
            <div><label className="label">Especialidad</label><select className="select" value={form.especialidad_id} onChange={e=>setForm({...form,especialidad_id:e.target.value})}><option value="">— Seleccionar —</option>{espList.map(e=><option key={e.id} value={e.id}>{e.especialidad_es}</option>)}</select></div>
            <div><label className="label">Motivo</label><input className="input" value={form.motivo} onChange={e=>setForm({...form,motivo:e.target.value})} placeholder="¿Por qué no reportaste ese día?"/></div>
          </div>
          <button className="btn-primary text-sm" disabled={saving} onClick={crear}>{saving?'Enviando…':'Enviar solicitud'}</button>
        </div>
      )}
      <div className="card p-4">
        <h3 className="font-bold text-[#003b7a] mb-3">{canApprove?'Solicitudes recibidas':'Mis solicitudes'}</h3>
        {!solicitudes.length&&<p className="text-sm text-slate-500">Sin solicitudes.</p>}
        <div className="space-y-3">{solicitudes.map(s=>(
          <div key={s.id as string} className={`border rounded-lg p-3 ${s.estado==='aprobado'?'border-emerald-400 bg-emerald-50':s.estado==='rechazado'?'border-rose-400 bg-rose-50':'border-slate-200'}`}>
            <div className="flex items-start justify-between gap-2 flex-wrap">
              <div>
                <div className="font-medium text-sm">{s.tecnico_nombre as string} — <strong>{s.fecha_reporte as string}</strong></div>
                <div className="text-xs text-slate-500 mt-0.5">Motivo: {s.motivo as string}</div>
                {(s.comentario as string)&&<div className="text-xs text-slate-500">Resp: {s.comentario as string}</div>}
              </div>
              <div className="flex items-center gap-2 flex-wrap">
                <span className={`badge ${s.estado==='aprobado'?'bg-emerald-100 text-emerald-800':s.estado==='rechazado'?'bg-rose-100 text-rose-800':'badge-borrador'}`}>{s.estado as string}</span>
                {canApprove&&s.estado==='pendiente'&&<><button className="btn-success text-xs" onClick={()=>decidir(s.id as string,'aprobado')}>✓ Aprobar</button><button className="btn-danger text-xs" onClick={()=>decidir(s.id as string,'rechazado')}>✗ Rechazar</button></>}
              </div>
            </div>
          </div>
        ))}</div>
      </div>
    </div>
  );
}

// ── DASHBOARD MEJORADO ────────────────────────────────────────────
function DashboardModule({catalogs,configActs,showToast}:{catalogs:Catalogs|null;configActs:ConfigAct[];showToast:(k:'ok'|'err'|'info',m:string)=>void}){
  const[fechaIni,setFechaIni]=useState(new Date(Date.now()-7*86400000).toISOString().split('T')[0]);
  const[fechaFin,setFechaFin]=useState(today());
  const[espIds,setEspIds]=useState<string[]>([]);
  const[actIds,setActIds]=useState<string[]>([]);
  const[incluirPersonal,setIncluirPersonal]=useState(false);
  const[incluirMaquinaria,setIncluirMaquinaria]=useState(false);
  const[data,setData]=useState<Record<string,unknown>|null>(null);
  const[loading,setLoading]=useState(false);
  const espList=useMemo(()=>catalogs?uniqueEsp(catalogs.especialidades_actividades):[],[catalogs]);
  const actList=useMemo(()=>{
    if(!catalogs) return [];
    if(!espIds.length) return catalogs.especialidades_actividades.filter(a=>a.activo!==false);
    const nombres=new Set(espIds.map(id=>espList.find(e=>e.id===id)?.especialidad_es?.toLowerCase()).filter(Boolean));
    return catalogs.especialidades_actividades.filter(a=>a.activo!==false&&nombres.has((a.especialidad_es||'').toLowerCase()));
  },[catalogs,espIds,espList]);

  async function load(){
    setLoading(true);
    try{
      let qR=supabase.from('reportes_avance').select('*').gte('fecha',fechaIni).lte('fecha',fechaFin);
      if(espIds.length) qR=qR.in('especialidad_id',espIds);
      const[reps,incid,maqD]=await Promise.all([
        qR,
        supabase.from('incidentes_seg').select('*').gte('fecha',fechaIni).lte('fecha',fechaFin).neq('tipo','sin_novedad'),
        incluirMaquinaria?supabase.from('maquinaria').select('*'):Promise.resolve({data:[]}),
      ]);
      const repIds=(reps.data||[]).map((r:Record<string,unknown>)=>r.id as string);
      // Cargar avances, asistencia, cualitativas y adicionales filtrados por reportes
      const[avances,asist,adics,susps]=await Promise.all([
        repIds.length?supabase.from('avance_diario').select('*').in('reporte_id',repIds):Promise.resolve({data:[]}),
        repIds.length&&incluirPersonal?supabase.from('asistencia_real').select('*').in('reporte_id',repIds):Promise.resolve({data:[]}),
        repIds.length?supabase.from('actividades_adicionales_reporte').select('*').in('reporte_id',repIds):Promise.resolve({data:[]}),
        repIds.length?supabase.from('suspensiones_clima').select('*').in('reporte_id',repIds):Promise.resolve({data:[]}),
      ]);
      const aD=(asist.data||[]) as Record<string,unknown>[];
      let avD=(avances.data||[]) as Record<string,unknown>[];
      if(actIds.length) avD=avD.filter(a=>actIds.includes(a.actividad_id as string));
      const adicD=(adics.data||[]) as Record<string,unknown>[];
      const suspD=(susps.data||[]) as Record<string,unknown>[];
      const horasH=aD.filter(a=>a.asistio).reduce((s,a)=>s+parseFloat(String(a.horas_trabajadas||0)),0);
      const horasP=suspD.reduce((s,a)=>s+parseFloat(String(a.horas_perdidas||0)),0);
      const pl=aD.length,re=aD.filter(a=>a.asistio).length;

      // Avance por actividad (solo cuantitativas)
      const avPorAct:Record<string,number>={};
      avD.filter(av=>av.unidad!=='cualitativo').forEach(av=>{
        const id=av.actividad_id as string;
        avPorAct[id]=(avPorAct[id]||0)+parseFloat(String(av.cantidad||0));
      });
      // Cualitativas
      const cualD=avD.filter(av=>av.unidad==='cualitativo') as Record<string,unknown>[];

      // Avance por día (para la línea de tendencia)
      const avPorDia:Record<string,number>={};
      avD.filter(av=>av.unidad!=='cualitativo').forEach(av=>{
        const f=av.fecha as string;
        avPorDia[f]=(avPorDia[f]||0)+parseFloat(String(av.cantidad||0));
      });

      // Incidentes por tipo (para el donut)
      const incPorTipo:Record<string,number>={};
      ((incid.data||[]) as Record<string,unknown>[]).forEach(i=>{
        const t=(i.tipo as string)||'otro';
        incPorTipo[t]=(incPorTipo[t]||0)+1;
      });

      // Horas maquinaria por novedades
      const{data:novedades}=incluirMaquinaria?await supabase.from('novedades_maquinaria').select('*').gte('fecha',fechaIni).lte('fecha',fechaFin):{data:[]};
      const horasSB=(novedades||[]).reduce((s:number,n:Record<string,unknown>)=>s+parseFloat(String(n.horas_standby||0)),0);

      setData({reportes:reps.data||[],horas_hombre:Math.round(horasH),horas_perdidas:Math.round(horasP),eficiencia_personal:pl>0?Math.round(re/pl*100):100,incidentes:incid.data||[],maquinaria:maqD.data||[],avance_por_actividad:avPorAct,avance_por_dia:avPorDia,incidentes_por_tipo:incPorTipo,cualitativas:cualD,adicionales:adicD,asistencia:aD,suspensiones:suspD,total_personal_dias:aD.length,horas_standby_total:horasSB.toFixed(1)});
    } catch(e:unknown){ showToast('err',(e as Error)?.message||'Error'); }
    finally{ setLoading(false); }
  }

  useEffect(()=>{load();},[]);

  const actividadesConConfig=useMemo(()=>{
    if(!catalogs||!data) return [];
    const avPorAct=data.avance_por_actividad as Record<string,number>;
    return configActs
      .filter(c=>!espIds.length||espIds.some(id=>{const e=espList.find(x=>x.id===id);return e&&(catalogs.especialidades_actividades.find(x=>x.id===c.actividad_id)?.especialidad_es||'').toLowerCase()===e.especialidad_es.toLowerCase();}))
      .filter(c=>!actIds.length||actIds.includes(c.actividad_id))
      .map(c=>{
        const actRow=catalogs.especialidades_actividades.find(e=>e.id===c.actividad_id);
        const avanceHoy=avPorAct[c.actividad_id]||0;
        const acumPrevio=c.acumulado_previo||0;
        const total=avanceHoy+acumPrevio;
        const pct=c.meta_total&&c.tiene_meta?Math.min(100,Math.round(total/c.meta_total*100)):null;
        return{...c,actividad_nombre:actRow?.actividad_es||c.actividad_id,avance_periodo:avanceHoy,total_acumulado:total,pct};
      });
  },[catalogs,configActs,data,espIds,actIds,espList]);

  const avanceGeneral=useMemo(()=>{
    const conMeta=actividadesConConfig.filter(a=>a.meta_total&&a.tiene_meta);
    const sumTotal=conMeta.reduce((s,a)=>s+(a.total_acumulado||0),0);
    const sumMeta=conMeta.reduce((s,a)=>s+(a.meta_total||0),0);
    return sumMeta>0?Math.min(100,Math.round(sumTotal/sumMeta*100)):null;
  },[actividadesConConfig]);

  const tendenciaDiaria=useMemo(()=>{
    if(!data) return [];
    const porDia=(data.avance_por_dia||{}) as Record<string,number>;
    return Object.entries(porDia).sort(([a],[b])=>a.localeCompare(b)).map(([fecha,cantidad])=>({fecha:fecha.slice(5),cantidad}));
  },[data]);

  const incidentesDonut=useMemo(()=>{
    if(!data) return [];
    const porTipo=(data.incidentes_por_tipo||{}) as Record<string,number>;
    const nombres:Record<string,string>={protesta:'Orden público',logistica:'Logística',falla_equipo:'Falla de equipo',decision:'Decisión dirección',otro:'Otro',accidente:'Accidente',incidente:'Incidente'};
    return Object.entries(porTipo).map(([tipo,cantidad])=>({tipo:nombres[tipo]||tipo,cantidad}));
  },[data]);

  const asistenciaPorEsp=useMemo(()=>{
    if(!data) return [];
    const asistD=(data.asistencia||[]) as Record<string,unknown>[];
    const repsD=(data.reportes||[]) as Record<string,unknown>[];
    const repEspMap:Record<string,string>={};
    repsD.forEach(r=>{repEspMap[r.id as string]=r.especialidad_id as string;});
    const porEsp:Record<string,{asistio:number;ausente:number}>={};
    asistD.forEach(a=>{
      const espId=repEspMap[a.reporte_id as string]||'';
      const nom=catalogs?.especialidades_actividades.find(e=>e.id===espId)?.especialidad_es||'Sin especialidad';
      if(!porEsp[nom]) porEsp[nom]={asistio:0,ausente:0};
      if(a.asistio) porEsp[nom].asistio++; else porEsp[nom].ausente++;
    });
    delete porEsp['Sin especialidad'];
    return Object.entries(porEsp).map(([especialidad,v])=>({especialidad,...v}));
  },[data,catalogs]);

  return(
    <div className="space-y-4">
      <div className="card p-4 space-y-3 no-print">
        <div className="flex gap-3 items-end flex-wrap">
          <div className="flex-1 min-w-[130px]"><label className="label">Desde</label><input type="date" className="input" value={fechaIni} onChange={e=>setFechaIni(e.target.value)}/></div>
          <div className="flex-1 min-w-[130px]"><label className="label">Hasta</label><input type="date" className="input" value={fechaFin} onChange={e=>setFechaFin(e.target.value)}/></div>
          <button className="btn-primary" onClick={load} disabled={loading}>{loading?'Cargando…':'Actualizar'}</button>
        </div>
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
          <MultiSelectDropdown label="Especialidades" options={espList} selected={espIds} onChange={v=>{setEspIds(v);setActIds([]);}} renderRow={e=><span>{e.especialidad_es}</span>} placeholder="Todas las especialidades"/>
          <MultiSelectDropdown label="Actividades" options={actList} selected={actIds} onChange={setActIds} renderRow={a=><span className="grid grid-cols-2 gap-2"><span className="text-xs text-slate-400 truncate">{a.especialidad_es}</span><span className="font-medium truncate">{a.actividad_es}</span></span>} placeholder="Todas las actividades"/>
        </div>
        <div className="flex items-center gap-4 flex-wrap">
          <label className="flex items-center gap-2 text-sm cursor-pointer"><input type="checkbox" checked={incluirPersonal} onChange={e=>setIncluirPersonal(e.target.checked)}/> Incluir personal/asistencia</label>
          <label className="flex items-center gap-2 text-sm cursor-pointer"><input type="checkbox" checked={incluirMaquinaria} onChange={e=>setIncluirMaquinaria(e.target.checked)}/> Incluir maquinaria</label>
        </div>
      </div>

      {data&&(
        <>
          <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">
            <MC label="Eficiencia personal" value={incluirPersonal?`${data.eficiencia_personal as number}%`:'—'} sub="asistencia" color={!incluirPersonal?'text-slate-400':(data.eficiencia_personal as number)>=90?'text-emerald-600':(data.eficiencia_personal as number)>=70?'text-amber-500':'text-rose-600'}/>
            <MC label="Horas-hombre" value={incluirPersonal?`${data.horas_hombre as number}h`:'—'} sub="productivas"/>
            <MC label="Horas stand-by" value={incluirMaquinaria?`${data.horas_standby_total as string}h`:'—'} sub="maquinaria"/>
            <MC label="Incidentes" value={(data.incidentes as unknown[]).length} sub="seguridad" color={(data.incidentes as unknown[]).length>0?'text-rose-600':'text-emerald-600'}/>
          </div>

          {/* GAUGES Y DONUT */}
          <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
            {incluirPersonal&&<GaugeCard label="Eficiencia personal" value={data.eficiencia_personal as number} sub="asistencia del período" color={(data.eficiencia_personal as number)>=90?'#10b981':(data.eficiencia_personal as number)>=70?'#f59e0b':'#ef4444'}/>}
            {avanceGeneral!==null&&<GaugeCard label="Avance general del proyecto" value={avanceGeneral} sub="acumulado vs. meta" color="#003b7a"/>}
            {incidentesDonut.length>0&&<IncidentDonut data={incidentesDonut}/>}
          </div>

          {/* LÍNEA DE TENDENCIA DIARIA */}
          {tendenciaDiaria.length>=2&&<TrendLineCard title="📈 Avance diario del período" data={tendenciaDiaria} dataKey="cantidad" color="#003b7a"/>}

          {/* AVANCE POR ACTIVIDAD — GRÁFICAS */}
          {actividadesConConfig.length>0&&(
            <div className="card p-4">
              <h3 className="font-bold text-[#003b7a] mb-4">📊 Avance por actividad</h3>
              <div className="space-y-5">
                {actividadesConConfig.map((c,i)=>{
                  const pct=c.pct??null;
                  const color=pct===null?'bg-slate-400':pct>=90?'bg-emerald-500':pct>=50?'bg-blue-500':'bg-amber-500';
                  const textColor=pct===null?'text-slate-500':pct>=90?'text-emerald-700':pct>=50?'text-blue-700':'text-amber-700';
                  const bgLight=pct===null?'bg-slate-50':pct>=90?'bg-emerald-50':pct>=50?'bg-blue-50':'bg-amber-50';
                  const emoji=pct===null?'📋':pct>=90?'✅':pct>=50?'🔵':'⚠️';
                  return(
                    <div key={i} className={`rounded-xl border border-slate-200 overflow-hidden`}>
                      {/* cabecera */}
                      <div className={`${bgLight} px-4 py-3 flex items-center justify-between flex-wrap gap-2`}>
                        <div>
                          <div className="font-bold text-sm text-[#003b7a]">{emoji} {c.actividad_nombre}</div>
                          <div className="text-xs text-slate-500 mt-0.5">{c.unidad_es||'—'}</div>
                        </div>
                        <div className="text-right">
                          <div className={`text-2xl font-black ${textColor}`}>{c.total_acumulado}</div>
                          {c.meta_total&&c.tiene_meta&&<div className="text-xs text-slate-500">de {c.meta_total} {c.unidad_es}</div>}
                        </div>
                      </div>
                      {/* barra de progreso */}
                      {c.meta_total&&c.tiene_meta&&pct!==null&&(
                        <div className="px-4 pb-3 pt-2 bg-white">
                          <div className="flex justify-between text-xs text-slate-500 mb-1.5">
                            <span>Acumulado previo: <strong>{c.acumulado_previo||0}</strong></span>
                            <span>Este período: <strong>+{c.avance_periodo}</strong></span>
                            <span className={`font-bold text-sm ${textColor}`}>{pct}%</span>
                          </div>
                          <div className="w-full bg-slate-200 rounded-full h-5 overflow-hidden relative">
                            <div className={`h-5 rounded-full transition-all duration-700 ${color}`} style={{width:`${Math.max(2,pct)}%`}}/>
                            <span className="absolute inset-0 flex items-center justify-center text-xs font-bold text-white drop-shadow">{pct}% completado</span>
                          </div>
                          <div className="flex justify-between text-xs text-slate-400 mt-1">
                            <span>0</span>
                            <span>{c.meta_total} {c.unidad_es}</span>
                          </div>
                        </div>
                      )}
                      {c.tipo==='D'&&<div className="px-4 pb-3 pt-1 bg-white text-xs text-purple-600 italic">Actividad cualitativa — sin meta numérica</div>}
                      {(!c.meta_total||!c.tiene_meta)&&c.tipo!=='D'&&<div className="px-4 pb-3 pt-1 bg-white text-xs text-slate-500 italic">Acumulativo sin meta · Total ejecutado: {c.total_acumulado} {c.unidad_es}</div>}
                    </div>
                  );
                })}
              </div>
            </div>
          )}

          {/* ACTIVIDADES CUALITATIVAS EN DASHBOARD */}
          {data&&((data.cualitativas||[]) as unknown[]).length>0&&(
            <div className="card p-4">
              <h3 className="font-bold text-[#003b7a] mb-3">📋 Actividades cualitativas reportadas</h3>
              <div className="space-y-2">
                {((data.cualitativas||[]) as Record<string,unknown>[]).map((a,i)=>{
                  const aR=catalogs?.especialidades_actividades.find(e=>e.id===(a.actividad_id as string));
                  return(
                    <div key={i} className="bg-purple-50 border border-purple-200 rounded-lg p-3">
                      <div className="font-semibold text-sm text-purple-800">{aR?.actividad_es||'Actividad'}</div>
                      <div className="text-xs text-slate-500 mt-0.5">{a.fecha as string}</div>
                      <div className="text-xs text-slate-600 mt-1 italic">"{a.observacion_es as string||'Sin descripción'}"</div>
                    </div>
                  );
                })}
              </div>
            </div>
          )}

          {/* ACTIVIDADES ADICIONALES EN DASHBOARD */}
          {data&&((data.adicionales||[]) as unknown[]).length>0&&(
            <div className="card p-4">
              <h3 className="font-bold text-[#003b7a] mb-3">➕ Actividades adicionales</h3>
              <div className="space-y-2">
                {((data.adicionales||[]) as Record<string,unknown>[]).map((a,i)=>(
                  <div key={i} className="bg-indigo-50 border border-indigo-200 rounded-lg p-3">
                    <div className="font-semibold text-sm text-indigo-800">{a.nombre as string}</div>
                    <div className="text-xs text-slate-500 mt-0.5">{a.fecha as string}</div>
                    <div className="text-xs text-slate-600 mt-1 italic">"{a.descripcion_ejecutado as string||''}"</div>
                  </div>
                ))}
              </div>
            </div>
          )}

          {/* BARRA APILADA ASISTENCIA */}
          {incluirPersonal&&asistenciaPorEsp.length>0&&<StackedAttendanceChart data={asistenciaPorEsp}/>}

          {/* PERSONAL POR ESPECIALIDAD EN DASHBOARD */}
          {data&&((data.asistencia||[]) as unknown[]).length>0&&(()=>{
            const asistD=(data.asistencia as Record<string,unknown>[]);
            const repsD=(data.reportes as Record<string,unknown>[]);
            const repEspMap:Record<string,string>={};
            repsD.forEach(r=>{repEspMap[r.id as string]=r.especialidad_id as string;});
            const porEsp:Record<string,Record<string,{nombre:string;cargo:string;totalHoras:number;asistio:boolean;motivo:string}>>={};
            asistD.forEach(a=>{
              const espId=repEspMap[a.reporte_id as string]||'';
              const espNom=catalogs?.especialidades_actividades.find(e=>e.id===espId)?.especialidad_es||'Sin especialidad';
              if(!porEsp[espNom]) porEsp[espNom]={};
              const doc=a.documento_personal as string;
              const persInfo=catalogs?.personal.find(p=>p.documento===doc);
              if(!porEsp[espNom][doc]) porEsp[espNom][doc]={nombre:persInfo?.nombre||doc,cargo:persInfo?.cargo_es||'—',totalHoras:0,asistio:false,motivo:''};
              porEsp[espNom][doc].totalHoras+=parseFloat(String(a.horas_trabajadas||0));
              if(a.asistio) porEsp[espNom][doc].asistio=true;
              if(a.motivo_ausencia) porEsp[espNom][doc].motivo=a.motivo_ausencia as string;
            });
            delete porEsp['Sin especialidad'];
            if(!Object.keys(porEsp).length) return null;
            return(
              <div className="card p-4">
                <h3 className="font-bold text-[#003b7a] mb-3">👥 Personal por especialidad</h3>
                <div className="space-y-4">
                  {Object.entries(porEsp).map(([espNom,docs])=>{
                    const lista=Object.values(docs);
                    return(
                      <div key={espNom}>
                        <div className="text-xs font-bold text-[#003b7a] uppercase tracking-wide mb-2 bg-blue-50 px-3 py-1.5 rounded-lg">🌿 {espNom}</div>
                        <table className="w-full text-xs border-collapse">
                          <thead><tr className="bg-slate-50"><th className="text-left px-3 py-2 font-medium text-slate-500">Nombre</th><th className="text-left px-3 py-2 font-medium text-slate-500">Cargo</th><th className="text-center px-3 py-2 font-medium text-slate-500">Estado</th><th className="text-center px-3 py-2 font-medium text-slate-500">Horas</th><th className="text-left px-3 py-2 font-medium text-slate-500">Motivo</th></tr></thead>
                          <tbody>{lista.map((p,i)=>(
                            <tr key={i} className="border-t border-slate-100">
                              <td className="px-3 py-1.5 font-medium">{p.nombre}</td>
                              <td className="px-3 py-1.5 text-slate-500">{p.cargo}</td>
                              <td className="px-3 py-1.5 text-center">{p.asistio?<span className="text-emerald-600 font-bold">✅</span>:<span className="text-rose-600 font-bold">❌</span>}</td>
                              <td className="px-3 py-1.5 text-center"><span className="bg-blue-50 text-blue-700 font-bold px-2 py-0.5 rounded">{p.totalHoras}h</span></td>
                              <td className="px-3 py-1.5 text-slate-400 italic">{p.motivo||'—'}</td>
                            </tr>
                          ))}</tbody>
                        </table>
                      </div>
                    );
                  })}
                </div>
              </div>
            );
          })()}

          {/* MAQUINARIA */}
          {(data.maquinaria as unknown[]).length>0&&(
            <div className="card p-4">
              <h3 className="font-bold text-[#003b7a] mb-3">🔧 Maquinaria acumulada</h3>
              <div className="overflow-x-auto">
                <table className="table w-full">
                  <thead><tr><th>Equipo</th><th>Tipo</th><th>Estado</th><th>Horas op.</th><th>Stand-by</th><th>Eficiencia</th></tr></thead>
                  <tbody>
                    {(data.maquinaria as Record<string,unknown>[]).map(m=>{
                      const op=m.horas_acum_operativas as number||0;
                      const sb=m.horas_acum_standby as number||0;
                      const t=op+sb;
                      const ef=t>0?Math.round(op/t*100):100;
                      return(
                        <tr key={m.id as string}>
                          <td className="font-medium">{m.item_id as string}</td>
                          <td>{m.tipo as string}</td>
                          <td><span className={`badge ${m.estado==='activo'?'bg-emerald-100 text-emerald-700':m.estado==='mantenimiento'?'bg-amber-100 text-amber-700':'bg-slate-100 text-slate-600'}`}>{m.estado as string}</span></td>
                          <td className="text-emerald-600">{op.toFixed(1)}h</td>
                          <td className="text-amber-500">{sb.toFixed(1)}h</td>
                          <td>
                            <div className="flex items-center gap-2">
                              <div className="flex-1 bg-slate-200 rounded-full h-2"><div className={`h-2 rounded-full ${ef>=80?'bg-emerald-500':ef>=50?'bg-amber-500':'bg-rose-500'}`} style={{width:`${ef}%`}}/></div>
                              <span className="text-xs font-medium">{ef}%</span>
                            </div>
                          </td>
                        </tr>
                      );
                    })}
                  </tbody>
                </table>
              </div>
            </div>
          )}
        </>
      )}
    </div>
  );
}
function MC({label,value,sub,color='text-[#003b7a]'}:{label:string;value:unknown;sub:string;color?:string}){
  return <div className="card p-4 text-center"><div className={`text-2xl sm:text-3xl font-bold ${color}`}>{String(value)}</div><div className="text-xs font-semibold text-slate-700 mt-1">{label}</div><div className="text-xs text-slate-400">{sub}</div></div>;
}

// ── GAUGE (semicírculo) ─────────────────────────────────────────────
function GaugeCard({label,value,sub,color='#003b7a'}:{label:string;value:number|null;sub:string;color?:string}){
  const v=value??0;
  const data=[{name:label,value:v,fill:color}];
  return(
    <div className="card p-4">
      <div className="text-xs font-semibold text-slate-700 uppercase tracking-wide text-center mb-1">{label}</div>
      <div className="relative h-32">
        <ResponsiveContainer width="100%" height="100%">
          <RadialBarChart innerRadius="70%" outerRadius="100%" barSize={14} data={data} startAngle={180} endAngle={0} cy="85%">
            <PolarAngleAxis type="number" domain={[0,100]} angleAxisId={0} tick={false}/>
            <RadialBar background={{fill:'#f1f5f9'}} dataKey="value" cornerRadius={8} angleAxisId={0}/>
          </RadialBarChart>
        </ResponsiveContainer>
        <div className="absolute inset-0 flex items-end justify-center pb-2">
          <div className="text-center">
            <div className="text-3xl font-black" style={{color}}>{value===null?'—':`${value}%`}</div>
          </div>
        </div>
      </div>
      <div className="text-xs text-slate-400 text-center">{sub}</div>
    </div>
  );
}

// ── LÍNEA DE TENDENCIA ────────────────────────────────────────────
function TrendLineCard({title,data,dataKey,color='#003b7a'}:{title:string;data:Record<string,unknown>[];dataKey:string;color?:string}){
  return(
    <div className="card p-4">
      <h3 className="font-bold text-[#003b7a] mb-3">{title}</h3>
      <ResponsiveContainer width="100%" height={220}>
        <LineChart data={data} margin={{top:5,right:10,left:-20,bottom:0}}>
          <CartesianGrid strokeDasharray="3 3" stroke="#e2e8f0"/>
          <XAxis dataKey="fecha" tick={{fontSize:11,fill:'#64748b'}}/>
          <YAxis tick={{fontSize:11,fill:'#64748b'}}/>
          <Tooltip contentStyle={{fontSize:12,borderRadius:8,border:'1px solid #e2e8f0'}}/>
          <Line type="monotone" dataKey={dataKey} stroke={color} strokeWidth={2.5} dot={{r:3,fill:color}}/>
        </LineChart>
      </ResponsiveContainer>
    </div>
  );
}

// ── BARRA APILADA ASISTENCIA ──────────────────────────────────────
function StackedAttendanceChart({data}:{data:{especialidad:string;asistio:number;ausente:number}[]}){
  return(
    <div className="card p-4">
      <h3 className="font-bold text-[#003b7a] mb-3">👥 Asistencia por especialidad</h3>
      <ResponsiveContainer width="100%" height={Math.max(180,data.length*46)}>
        <BarChart data={data} layout="vertical" margin={{top:5,right:20,left:10,bottom:0}}>
          <CartesianGrid strokeDasharray="3 3" stroke="#e2e8f0" horizontal={false}/>
          <XAxis type="number" tick={{fontSize:11,fill:'#64748b'}} allowDecimals={false}/>
          <YAxis type="category" dataKey="especialidad" tick={{fontSize:11,fill:'#334155'}} width={110}/>
          <Tooltip contentStyle={{fontSize:12,borderRadius:8,border:'1px solid #e2e8f0'}}/>
          <Legend wrapperStyle={{fontSize:12}}/>
          <Bar dataKey="asistio" name="Asistió" stackId="a" fill="#10b981" radius={[0,4,4,0]}/>
          <Bar dataKey="ausente" name="Ausente" stackId="a" fill="#f43f5e"/>
        </BarChart>
      </ResponsiveContainer>
    </div>
  );
}

// ── DONUT INCIDENTES ──────────────────────────────────────────────
const DONUT_COLORS=['#003b7a','#f59e0b','#ef4444','#10b981','#8b5cf6','#64748b'];
function IncidentDonut({data}:{data:{tipo:string;cantidad:number}[]}){
  return(
    <div className="card p-4">
      <h3 className="font-bold text-[#003b7a] mb-3">⚠️ Incidentes por tipo</h3>
      <ResponsiveContainer width="100%" height={220}>
        <PieChart>
          <Pie data={data} dataKey="cantidad" nameKey="tipo" innerRadius={50} outerRadius={80} paddingAngle={2}>
            {data.map((_,i)=><Cell key={i} fill={DONUT_COLORS[i%DONUT_COLORS.length]}/>)}
          </Pie>
          <Tooltip contentStyle={{fontSize:12,borderRadius:8,border:'1px solid #e2e8f0'}}/>
          <Legend wrapperStyle={{fontSize:11}}/>
        </PieChart>
      </ResponsiveContainer>
    </div>
  );
}

// ── INFORMES ──────────────────────────────────────────────────────
function InformesModule({user,catalogs,configActs,maquinaria,showToast}:{
  user:Profile; catalogs:Catalogs|null; configActs:ConfigAct[]; maquinaria:Maq[];
  showToast:(k:'ok'|'err'|'info',m:string)=>void;
}){
  const[fechaIni,setFechaIni]=useState(new Date(Date.now()-7*86400000).toISOString().split('T')[0]);
  const[fechaFin,setFechaFin]=useState(today());
  const[espIds,setEspIds]=useState<string[]>([]);
  const[areaIds,setAreaIds]=useState<string[]>([]);
  const[actIds,setActIds]=useState<string[]>([]);
  const[incluirPersonal,setIncluirPersonal]=useState(false);
  const[incluirMaquinaria,setIncluirMaquinaria]=useState(false);
  const[modo,setModo]=useState<'resumen'|'detallado'>('resumen');
  const[soloAp,setSoloAp]=useState(false);
  const[data,setData]=useState<Record<string,unknown>|null>(null);
  const[loading,setLoading]=useState(false);
  const espList=useMemo(()=>catalogs?uniqueEsp(catalogs.especialidades_actividades):[],[catalogs]);
  function nombreMaquina(id:string){
    const m=maquinaria.find(x=>x.id===id);
    return m?`${m.item_id||m.nombre||''} (${m.tipo||''})`:'Máquina no encontrada';
  }
  const actList=useMemo(()=>{
    if(!catalogs) return [];
    if(!espIds.length) return catalogs.especialidades_actividades.filter(a=>a.activo!==false);
    const nombres=new Set(espIds.map(id=>espList.find(e=>e.id===id)?.especialidad_es?.toLowerCase()).filter(Boolean));
    return catalogs.especialidades_actividades.filter(a=>a.activo!==false&&nombres.has((a.especialidad_es||'').toLowerCase()));
  },[catalogs,espIds,espList]);

  useEffect(()=>{ if(catalogs) fetchData(); },[catalogs]);

  async function fetchData(){
    setLoading(true);
    try{
      let qR=supabase.from('reportes_avance').select('*').gte('fecha',fechaIni).lte('fecha',fechaFin);
      if(espIds.length) qR=qR.in('especialidad_id',espIds);
      if(soloAp) qR=qR.eq('estado','aprobado');
      if(user.rol==='cliente') qR=qR.eq('estado','aprobado');
      const{data:reps}=await qR;
      const repIds=(reps||[]).map((r:Record<string,unknown>)=>r.id as string);
      const[av,as2,sc,maqNov,adics]=await Promise.all([
        repIds.length?supabase.from('avance_diario').select('*').in('reporte_id',repIds):Promise.resolve({data:[]}),
        repIds.length&&incluirPersonal?supabase.from('asistencia_real').select('*').in('reporte_id',repIds):Promise.resolve({data:[]}),
        repIds.length?supabase.from('suspensiones_clima').select('*').in('reporte_id',repIds):Promise.resolve({data:[]}),
        repIds.length&&incluirMaquinaria?supabase.from('novedades_maquinaria').select('*').in('reporte_id',repIds):Promise.resolve({data:[]}),
        repIds.length?supabase.from('actividades_adicionales_reporte').select('*').in('reporte_id',repIds):Promise.resolve({data:[]}),
      ]);
      let avD=(av.data||[]) as Record<string,unknown>[];
      if(areaIds.length) avD=avD.filter(a=>areaIds.includes(a.area_id as string));
      if(actIds.length) avD=avD.filter(a=>actIds.includes(a.actividad_id as string));
      const aD=(as2.data||[]) as Record<string,unknown>[];
      const horasH=aD.filter(a=>a.asistio).reduce((s,a)=>s+parseFloat(String(a.horas_trabajadas||0)),0);
      const horasC=((sc.data||[]) as Record<string,unknown>[]).reduce((s,a)=>s+parseFloat(String(a.horas_perdidas||0)),0);

      // Calcular acumulado real incluyendo acumulado previo de config
      const avConMeta:Record<string,unknown>[]=avD.map(av=>{
        const cfg=configActs.find(c=>c.actividad_id===(av.actividad_id as string));
        const acumPrevio=cfg?.acumulado_previo||0;
        return{...av,acumulado_total_real:(av.acumulado_total as number)+acumPrevio};
      });
      const avancesNorm=avConMeta.filter(a=>a.unidad!=='cualitativo');
      const cualitativas=avConMeta.filter(a=>a.unidad==='cualitativo');
      const adicionales=(adics.data||[]) as Record<string,unknown>[];
      const suspensiones=(sc.data||[]) as Record<string,unknown>[];
      const maquinaria=(maqNov.data||[]) as Record<string,unknown>[];

      setData({reportes:reps||[],avances:avancesNorm,cualitativas,adicionales,asistencia:aD,suspensiones,maquinaria,totales:{horas_hombre:Math.round(horasH),horas_perdidas_clima:Math.round(horasC),dias:repIds.length}});
    } catch(e:unknown){ showToast('err',(e as Error)?.message||'Error'); }
    finally{ setLoading(false); }
  }

  function imprimirPlan(){
    try{
    if(!data){showToast('err','No hay datos. Consulta primero.');return;}
    if(!(data.avances as unknown[]).length&&!((data.cualitativas||[]) as unknown[]).length&&!((data.adicionales||[]) as unknown[]).length){showToast('info','No hay actividades en el período seleccionado');return;}
    const avances=data.avances as Record<string,unknown>[];
    const cual=((data.cualitativas||[]) as Record<string,unknown>[]);
    const adic=((data.adicionales||[]) as Record<string,unknown>[]);
    const susps=((data.suspensiones||[]) as Record<string,unknown>[]);
    const maqNovs=((data.maquinaria||[]) as Record<string,unknown>[]);
    const reps=data.reportes as Record<string,unknown>[];
    const resumen=data.totales as Record<string,unknown>;
    const repEspMap:Record<string,string>={};
    reps.forEach(r=>{repEspMap[r.id as string]=r.especialidad_id as string;});
    const getEspNom=(eId:string)=>catalogs?.especialidades_actividades.find(e=>e.id===eId)?.especialidad_es||eId;
    const isCliente=user.rol==='cliente';

    // Agrupar todo por especialidad
    const porEsp:Record<string,{
      avances:Record<string,unknown>[];cual:Record<string,unknown>[];adic:Record<string,unknown>[];
      personal:Record<string,unknown>[];susps:Record<string,unknown>[];maq:Record<string,unknown>[];
    }>={};
    const addE=(espId:string,tipo:'avances'|'cual'|'adic'|'personal'|'susps'|'maq',item:Record<string,unknown>)=>{
      if(!espId) return;
      const n=getEspNom(espId);
      if(!porEsp[n]) porEsp[n]={avances:[],cual:[],adic:[],personal:[],susps:[],maq:[]};
      porEsp[n][tipo].push(item);
    };
    avances.forEach(av=>addE(av.especialidad_id as string,'avances',av));
    cual.forEach(av=>addE(av.especialidad_id as string,'cual',av));
    adic.forEach(ad=>addE(repEspMap[ad.reporte_id as string]||'','adic',ad));
    ((data.asistencia||[]) as Record<string,unknown>[]).forEach(a=>addE(repEspMap[a.reporte_id as string]||'','personal',a));
    susps.forEach(s=>addE(repEspMap[s.reporte_id as string]||'','susps',s));
    maqNovs.forEach(m=>addE(repEspMap[m.reporte_id as string]||'','maq',m));
    delete porEsp[''];

    const css=`body{font-family:Arial,sans-serif;font-size:11px;color:#1e293b;margin:20px}
    h1{font-size:18px;color:#003b7a;margin:0;font-weight:900}
    .sub{font-size:10px;color:#555;margin:2px 0}
    .kpis{display:grid;grid-template-columns:repeat(3,1fr);gap:10px;margin:14px 0}
    .kpi{border:1px solid #e2e8f0;border-radius:8px;padding:8px;text-align:center}
    .kpi-val{font-size:20px;font-weight:900;color:#003b7a}
    .kpi-lbl{font-size:9px;color:#64748b}
    .esp-block{margin-bottom:24px;page-break-inside:avoid}
    .esp-tit{background:#003b7a;color:white;padding:7px 12px;border-radius:6px;font-weight:800;font-size:12px;margin-bottom:10px}
    .fecha-block{border-left:3px solid #003b7a;padding-left:10px;margin-bottom:10px}
    .fecha-tit{font-weight:700;color:#003b7a;font-size:10px;margin-bottom:5px}
    .act-row{padding:4px 0;border-bottom:1px solid #f1f5f9;font-size:10px;display:flex;justify-content:space-between}
    .act-nom{font-weight:600}.act-area{font-size:9px;color:#94a3b8}
    .act-val{text-align:right;font-weight:700;color:#003b7a}
    .act-obs{font-size:9px;color:#64748b;font-style:italic;margin-top:2px}
    .cual-row{padding:4px 0;border-bottom:1px solid #f1f5f9;font-size:10px}
    .cual-badge{background:#ede9fe;color:#7c3aed;font-size:8px;padding:1px 5px;border-radius:10px;margin-right:4px;font-weight:700}
    .adic-badge{background:#e0e7ff;color:#4338ca;font-size:8px;padding:1px 5px;border-radius:10px;margin-right:4px;font-weight:700}
    .obs{font-size:9px;color:#64748b;font-style:italic;padding-left:8px;margin-top:2px}
    .sec-tit{font-size:9px;color:#475569;font-weight:700;text-transform:uppercase;letter-spacing:0.4px;margin:10px 0 4px;border-top:1px solid #e2e8f0;padding-top:8px}
    .tbl{width:100%;border-collapse:collapse;font-size:10px}
    .tbl th{background:#f8fafc;text-align:left;padding:3px 8px;font-size:9px;color:#94a3b8;text-transform:uppercase;border-bottom:1px solid #e2e8f0}
    .tbl td{padding:3px 8px;border-bottom:1px solid #f8fafc}
    .tbl tr:last-child td{border-bottom:none}
    .susp-row{background:#fef9c3;border:1px solid #fde68a;border-radius:4px;padding:4px 8px;margin-bottom:4px;font-size:9px}
    .maq-row{background:#f1f5f9;border:1px solid #e2e8f0;border-radius:4px;padding:4px 8px;margin-bottom:4px;font-size:9px}
    .ftr{text-align:center;font-size:9px;color:#94a3b8;margin-top:20px;border-top:1px solid #e2e8f0;padding-top:8px}`;

    const gruposHTML=Object.entries(porEsp).map(([espNom,items])=>{
      // Resumen por actividad — total del rango seleccionado (siempre incluido)
      const porActEsp=resumenPorActividad(items.avances,configActs,catalogs);
      const resumenHTML=Object.keys(porActEsp).length?`<div class="sec-tit">📊 Resumen del período</div><table class="tbl"><thead><tr><th>Actividad</th><th>Total rango</th><th>Acumulado histórico</th><th>Meta</th><th>%</th></tr></thead><tbody>${Object.values(porActEsp).map(a=>`<tr><td><strong>${a.nombre}</strong></td><td>${a.totalRango} ${a.unidad}</td><td>${a.acumuladoHistorico} ${a.unidad}</td><td>${a.meta||'—'}</td><td>${a.pct!==null?a.pct+'%':'—'}</td></tr>`).join('')}</tbody></table>`:'';

      // Avances por fecha — solo en modo Detallado
      const porFecha:Record<string,Record<string,unknown>[]>={};
      items.avances.forEach(av=>{const f=av.fecha as string;if(!porFecha[f])porFecha[f]=[];porFecha[f].push(av);});
      const fechasHTML=modo!=='detallado'?'':Object.entries(porFecha).sort(([fa],[fb])=>fa.localeCompare(fb)).map(([fecha,avs])=>{
        const avHTML=avs.map(av=>{
          const aR=catalogs?.especialidades_actividades.find(e=>e.id===(av.actividad_id as string));
          const arR=catalogs?.areas.find(a=>a.id===(av.area_id as string));
          const cfg=configActs.find(c=>c.actividad_id===(av.actividad_id as string));
          const pct=cfg?.meta_total&&cfg.tiene_meta?Math.min(100,Math.round((av.acumulado_total_real as number)/cfg.meta_total*100)):null;
          const esAjuste=String(av.observacion_es||'').startsWith('[AJUSTE]');
          const obsTexto=esAjuste?String(av.observacion_es||'').replace('[AJUSTE] ',''):(av.observacion_es as string||'');
          const ajusteBadge=esAjuste?'<span class="adic-badge" style="background:#fef3c7;color:#92400e">⚖️ AJUSTE</span> ':'';
          const obsHTML=obsTexto&&!isCliente?`<div class="act-obs">${esAjuste?'':'💬 '}"${obsTexto}"</div>`:'';
          const areaHTML=esAjuste?'':`<div class="act-area">📍 ${arR?.area_es||'—'}</div>`;
          return `<div class="act-row"><div><div class="act-nom">${ajusteBadge}${aR?.actividad_es||''}</div>${areaHTML}${obsHTML}</div><div><div class="act-val">${(av.cantidad as number)>0?'+':''}${av.cantidad as number} ${av.unidad as string}</div><div style="font-size:9px;color:#64748b">Acum: ${av.acumulado_total_real as number}${pct!==null?` · ${pct}%`:''}</div></div></div>`;
        }).join('');
        return `<div class="fecha-block"><div class="fecha-tit">📅 ${fecha}</div>${avHTML}</div>`;
      }).join('');

      // Cualitativas
      const cualHTML=items.cual.length?`<div class="sec-tit">📋 Actividades cualitativas</div>${items.cual.map(a=>{
        const aR=catalogs?.especialidades_actividades.find(e=>e.id===(a.actividad_id as string));
        return `<div class="cual-row"><div><span class="cual-badge">CUALITATIVA</span><strong>${aR?.actividad_es||''}</strong> <span style="font-size:9px;color:#94a3b8">${a.fecha as string}</span></div><div class="obs">"${(a.observacion_es as string)||''}"</div></div>`;
      }).join('')}`:'';

      // Adicionales
      const adicHTML=items.adic.length?`<div class="sec-tit">➕ Actividades adicionales</div>${items.adic.map(a=>`<div class="cual-row"><div><span class="adic-badge">ADICIONAL</span><strong>${a.nombre as string}</strong> <span style="font-size:9px;color:#94a3b8">${a.fecha as string}</span></div><div class="obs">"${(a.descripcion_ejecutado as string)||''}"</div></div>`).join('')}`:'';

      // Personal
      const porDoc:Record<string,{nombre:string;cargo:string;horas:number;asistio:boolean;motivo:string;novedad:string}>={};
      items.personal.forEach(a=>{
        const doc=a.documento_personal as string;
        const pi=catalogs?.personal.find(p=>p.documento===doc);
        if(!porDoc[doc]) porDoc[doc]={nombre:pi?.nombre||doc,cargo:pi?.cargo_es||'—',horas:0,asistio:false,motivo:'',novedad:''};
        porDoc[doc].horas+=parseFloat(String(a.horas_trabajadas||0));
        if(a.asistio) porDoc[doc].asistio=true;
        if(a.motivo_ausencia) porDoc[doc].motivo=a.motivo_ausencia as string;
        if(a.novedad) porDoc[doc].novedad=a.novedad as string;
      });
      const listaP=Object.values(porDoc);
      const personalHTML=listaP.length?`<div class="sec-tit">👥 Personal — ${espNom}</div><table class="tbl"><thead><tr><th>N</th><th>Nombre</th><th>Cargo</th><th>Estado</th><th>Horas</th><th>Motivo / Novedad</th></tr></thead><tbody>${listaP.map((p,i)=>`<tr><td>${i+1}</td><td><strong>${p.nombre}</strong></td><td>${p.cargo}</td><td>${p.asistio?'<span style="color:#16a34a;font-weight:700">✅ Asistió</span>':'<span style="color:#dc2626;font-weight:700">❌ Ausente</span>'}</td><td><span style="background:#eff6ff;color:#1d4ed8;font-weight:700;padding:1px 5px;border-radius:4px">${p.horas}h</span></td><td style="color:#64748b;font-style:italic">${p.motivo||p.novedad||'—'}</td></tr>`).join('')}</tbody></table>`:'';

      // Suspensiones
      const suspsHTML=items.susps.length?`<div class="sec-tit">⚠️ Suspensiones</div>${items.susps.map(s=>`<div class="susp-row"><strong>${s.es_general?'General':'Por actividad'}</strong> · ${s.descripcion as string}${s.hora_inicio?` · ${s.hora_inicio as string}–${s.hora_fin as string}`:''}</div>`).join('')}`:'';

      // Maquinaria
      const maqHTML=items.maq.length?`<div class="sec-tit">🚜 Maquinaria</div>${items.maq.map(m=>{
        const mn=nombreMaquina(m.maquinaria_id as string);
        return `<div class="maq-row"><strong>${mn}</strong>${m.hora_inicio?` · ${m.hora_inicio as string}–${m.hora_fin as string}`:''}${m.descripcion?` · ${m.descripcion as string}`:''}${m.novedad?` · Novedad: ${m.novedad as string}`:''}</div>`;
      }).join('')}`:'';

      return `<div class="esp-block"><div class="esp-tit">🌿 ${espNom}</div>${resumenHTML}${fechasHTML}${cualHTML}${adicHTML}${isCliente?'':personalHTML}${isCliente?'':suspsHTML}${isCliente?'':maqHTML}</div>`;
    }).join('');

    const html=`<!DOCTYPE html><html lang="es"><head><meta charset="UTF-8"><title>Informe PDS 360</title><style>${css}</style></head><body>
    <h1>Powerchina · PDS 360 — Informe de Avance</h1>
    <p class="sub">Período: <strong>${fechaIni}</strong> al <strong>${fechaFin}</strong>${soloAp?' · Solo aprobados':''}</p>
    <div class="kpis">
      <div class="kpi"><div class="kpi-val">${(resumen.dias as number)||0}</div><div class="kpi-lbl">Reportes</div></div>
      <div class="kpi"><div class="kpi-val">${(resumen.horas_hombre as number)||0}h</div><div class="kpi-lbl">Horas-hombre</div></div>
      <div class="kpi"><div class="kpi-val">${(resumen.horas_perdidas_clima as number)||0}h</div><div class="kpi-lbl">Horas perdidas</div></div>
    </div>
    ${gruposHTML}
    <div class="ftr">Powerchina PDS 360 · Generado: ${new Date().toLocaleString('es-CO')}</div>
    </body></html>`;
    const win=window.open('','_blank');
    if(win){win.document.write(html);win.document.close();setTimeout(()=>win.print(),600);}
    }catch(e:unknown){console.error('Error PDF:',e);showToast('err','Error generando PDF: '+((e as Error)?.message||'Error desconocido'));}
  }

  function exportExcel(){
    if(!data) return;
    const isC=user.rol==='cliente';
    const avances=data.avances as Record<string,unknown>[];
    const reportes=data.reportes as Record<string,unknown>[];
    const asistencia=data.asistencia as Record<string,unknown>[];
    const rows=avances.map(av=>{
      const aR=catalogs?.especialidades_actividades.find(e=>e.id===(av.actividad_id as string));
      const arR=catalogs?.areas.find(a=>a.id===(av.area_id as string));
      const rep=reportes.find(r=>r.id===av.reporte_id);
      const cfg=configActs.find(c=>c.actividad_id===(av.actividad_id as string));
      const pct=cfg?.meta_total&&cfg.tiene_meta?Math.round((av.acumulado_total_real as number)/cfg.meta_total*100):null;
      const o:Record<string,unknown>={
        Fecha:av.fecha, Especialidad:aR?.especialidad_es||'', Actividad:aR?.actividad_es||'',
        Área:arR?.area_es||'', Cantidad_Hoy:av.cantidad, Unidad:av.unidad,
        Acumulado:av.acumulado_total_real,
        Meta:cfg?.meta_total||'—', Avance_Pct:pct!==null?`${pct}%`:'—',
      };
      if(!isC) o.Técnico=rep?.usuario_nombre||'';
      return o;
    });
    const cual=((data.cualitativas||[]) as Record<string,unknown>[]);
    const adic=((data.adicionales||[]) as Record<string,unknown>[]);
    const cualRows=cual.map(a=>{
      const aR=catalogs?.especialidades_actividades.find(e=>e.id===(a.actividad_id as string));
      return{Fecha:a.fecha,Especialidad:aR?.especialidad_es||'',Actividad:aR?.actividad_es||'',Tipo:'Cualitativa','Descripción / Observación':a.observacion_es||''};
    });
    const adicRows=adic.map(a=>({'Fecha':a.fecha,'Actividad adicional':a.nombre,Tipo:'Adicional (no planeada)','Descripción ejecutado':a.descripcion_ejecutado||''}));
    const porAct=resumenPorActividad(avances,configActs,catalogs);
    const resumenRows=Object.values(porAct).map(a=>({
      Actividad:a.nombre,Unidad:a.unidad,Total_Rango:a.totalRango,
      Acumulado_Historico:a.acumuladoHistorico,Meta:a.meta||'—',
      Avance_Pct:a.pct!==null?`${a.pct}%`:'—',
    }));
    const wb=XLSX.utils.book_new();
    if(resumenRows.length) XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(resumenRows),'Resumen');
    if(modo==='detallado') XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(rows),'Avance');
    if(cualRows.length) XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(cualRows),'Cualitativas');
    if(adicRows.length) XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(adicRows),'Adicionales');
    // Personal con nombres
    const aRows=asistencia.map(a=>{
      const pi=catalogs?.personal.find(p=>p.documento===(a.documento_personal as string));
      return{
        Fecha:a.fecha,
        Nombre:pi?.nombre||a.documento_personal,
        Documento:a.documento_personal,
        Cargo:pi?.cargo_es||'—',
        Asistió:(a.asistio as boolean)?'Sí':'No',
        Horas_trabajadas:a.horas_trabajadas,
        Motivo:a.motivo_ausencia||'—',
        Novedad:(a.novedad as string)||'—',
        Ausencia_parcial:(a.ausencia_parcial as boolean)?'Sí':'No',
        Hora_ausencia_ini:(a.hora_ausencia_ini as string)||'—',
        Hora_ausencia_fin:(a.hora_ausencia_fin as string)||'—',
      };
    });
    if(aRows.length) XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(aRows),'Personal');
    // Suspensiones
    const suspRows=((data.suspensiones||[]) as Record<string,unknown>[]).map(s=>({
      Fecha:s.fecha,Tipo:s.tipo_susp||'clima',General:(s.es_general as boolean)?'Sí':'No',
      Descripcion:s.descripcion,Hora_inicio:s.hora_inicio||'—',Hora_fin:s.hora_fin||'—',
      Horas_perdidas:s.horas_perdidas||0,
    }));
    if(suspRows.length) XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(suspRows),'Suspensiones');
    // Maquinaria
    const maqRows=((data.maquinaria||[]) as Record<string,unknown>[]).map(m=>({
      Fecha:m.fecha,Maquinaria:nombreMaquina(m.maquinaria_id as string),
      Hora_inicio:m.hora_inicio||'—',Hora_fin:m.hora_fin||'—',
      Descripcion:m.descripcion||'—',Novedad:(m.novedad as string)||'—',
    }));
    if(maqRows.length) XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(maqRows),'Maquinaria');
    if(!wb.SheetNames.length) XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(rows),'Avance');
    const espsSel=espIds.length>0?'_seleccion':'_todas';
    XLSX.writeFile(wb,`PDS360_${fechaIni}_${fechaFin}${espsSel}.xlsx`);
    showToast('ok','Excel descargado');
  }

  return(
    <div className="space-y-4">
      <div className="card p-4 space-y-3 no-print">
        <div className="grid grid-cols-1 sm:grid-cols-3 gap-3 items-end">
          <div><label className="label">Desde</label><input type="date" className="input" value={fechaIni} onChange={e=>setFechaIni(e.target.value)}/></div>
          <div><label className="label">Hasta</label><input type="date" className="input" value={fechaFin} onChange={e=>setFechaFin(e.target.value)}/></div>
          <button className="btn-primary" onClick={fetchData} disabled={loading}>{loading?'Cargando…':'Consultar'}</button>
        </div>
        <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
          <MultiSelectDropdown label="Especialidades" options={espList} selected={espIds} onChange={setEspIds} renderRow={e=><span>{e.especialidad_es}</span>} placeholder="Todas las especialidades"/>
          <MultiSelectDropdown label="Áreas" options={catalogs?.areas||[]} selected={areaIds} onChange={setAreaIds} renderRow={a=><span>{a.area_es}</span>} placeholder="Todas las áreas"/>
          <MultiSelectDropdown label="Actividades" options={actList} selected={actIds} onChange={setActIds} renderRow={a=><span className="grid grid-cols-2 gap-2"><span className="text-xs text-slate-400 truncate">{a.especialidad_es}</span><span className="font-medium truncate">{a.actividad_es}</span></span>} placeholder="Todas las actividades"/>
        </div>
        <div className="flex items-center gap-4 flex-wrap">
          <label className="flex items-center gap-2 text-sm cursor-pointer"><input type="checkbox" checked={incluirPersonal} onChange={e=>setIncluirPersonal(e.target.checked)}/> Incluir personal/asistencia</label>
          <label className="flex items-center gap-2 text-sm cursor-pointer"><input type="checkbox" checked={incluirMaquinaria} onChange={e=>setIncluirMaquinaria(e.target.checked)}/> Incluir maquinaria</label>
        </div>
        <div>
          <label className="label">Nivel de detalle</label>
          <div className="flex gap-2">
            <button type="button" onClick={()=>setModo('resumen')} className={`text-xs px-3 py-1.5 rounded-lg border transition-colors ${modo==='resumen'?'bg-[#003b7a] text-white border-[#003b7a]':'border-slate-300 text-slate-600'}`}>📋 Resumen (solo totales del rango)</button>
            <button type="button" onClick={()=>setModo('detallado')} className={`text-xs px-3 py-1.5 rounded-lg border transition-colors ${modo==='detallado'?'bg-[#003b7a] text-white border-[#003b7a]':'border-slate-300 text-slate-600'}`}>📅 Detallado (día a día)</button>
          </div>
        </div>
        <div className="flex items-center gap-3 flex-wrap">
          <label className="flex items-center gap-2 text-sm cursor-pointer"><input type="checkbox" checked={soloAp} onChange={e=>setSoloAp(e.target.checked)}/> Solo aprobados</label>
          <button className="btn-secondary text-xs" onClick={exportExcel} disabled={!data}>📥 Excel</button>
          <button className="btn-secondary text-xs" onClick={imprimirPlan}>🖨️ PDF / Imprimir</button>
        </div>
      </div>

      <div className="print-area space-y-4">
        {data&&(
          <>
            {/* Encabezado de impresión */}
            <div className="hidden print:flex items-center gap-3 mb-4 pb-4 border-b-2 border-[#003b7a]">
              <img src="/icons/icon-192.png" alt="PC" className="h-12 w-auto"/>
              <div><div className="font-bold text-xl text-[#003b7a]">Powerchina · PDS 360</div><div className="text-sm text-slate-500">Informe de avance · {fechaIni} a {fechaFin}</div></div>
            </div>

            <div className="card p-4">
              <h3 className="font-bold text-[#003b7a] mb-3">Resumen del período</h3>
              <div className="grid grid-cols-3 gap-3 text-center">
                <div><div className="text-xl font-bold text-[#003b7a]">{(data.totales as Record<string,unknown>).dias as number}</div><div className="text-xs text-slate-500">Reportes</div></div>
                <div><div className="text-xl font-bold text-emerald-600">{(data.totales as Record<string,unknown>).horas_hombre as number}h</div><div className="text-xs text-slate-500">Horas-hombre</div></div>
                <div><div className="text-xl font-bold text-amber-500">{(data.totales as Record<string,unknown>).horas_perdidas_clima as number}h</div><div className="text-xs text-slate-500">Perdidas</div></div>
              </div>
            </div>

            {/* AVANCE POR ACTIVIDAD — resumen del rango, siempre visible */}
            {(data.avances as unknown[]).length>0&&(()=>{
              const avances=data.avances as Record<string,unknown>[];
              const porAct=resumenPorActividad(avances,configActs,catalogs);
              return(
                <div className="card p-4">
                  <h3 className="font-bold text-[#003b7a] mb-3">📊 Avance por actividad — {fechaIni} a {fechaFin}</h3>
                  <div className="space-y-4">
                    {Object.values(porAct).map((a,i)=>{
                      const color=a.pct===null?'bg-slate-400':a.pct>=90?'bg-emerald-500':a.pct>=50?'bg-blue-500':'bg-amber-500';
                      const tc=a.pct===null?'text-slate-600':a.pct>=90?'text-emerald-700':a.pct>=50?'text-blue-700':'text-amber-700';
                      return(
                        <div key={i} className="border border-slate-200 rounded-xl overflow-hidden">
                          <div className="bg-slate-50 px-4 py-3 flex items-center justify-between flex-wrap gap-2">
                            <div>
                              <div className="font-bold text-sm text-[#003b7a]">{a.nombre}</div>
                              <div className="text-xs text-slate-500">{a.unidad}</div>
                            </div>
                            <div className="text-right flex gap-6">
                              <div className="text-center"><div className={`text-xl font-black ${tc}`}>{a.totalRango}</div><div className="text-xs text-slate-400">Total del rango</div></div>
                              <div className="text-center"><div className="text-xl font-black text-slate-500">{a.acumuladoHistorico}</div><div className="text-xs text-slate-400">Acumulado histórico</div></div>
                              {a.meta&&<><div className="text-center"><div className="text-xl font-black text-slate-500">{a.meta}</div><div className="text-xs text-slate-400">Meta</div></div>
                              <div className="text-center"><div className={`text-xl font-black ${tc}`}>{a.pct}%</div><div className="text-xs text-slate-400">Avance</div></div></>}
                            </div>
                          </div>
                          {a.meta&&a.pct!==null&&(
                            <div className="px-4 pb-3 pt-2 bg-white">
                              <div className="w-full bg-slate-200 rounded-full h-4 overflow-hidden">
                                <div className={`h-4 rounded-full ${color}`} style={{width:`${Math.max(2,a.pct)}%`}}/>
                              </div>
                              <div className="flex justify-between text-xs text-slate-400 mt-1"><span>0</span><span>{a.meta} {a.unidad}</span></div>
                            </div>
                          )}
                        </div>
                      );
                    })}
                  </div>
                </div>
              );
            })()}

            {/* Vista agrupada por especialidad > fecha — el día a día respeta el modo, personal/maquinaria no */}
            {(()=>{
              const repsG=data.reportes as Record<string,unknown>[];
              const avG=data.avances as Record<string,unknown>[];
              const cualG=((data.cualitativas||[]) as Record<string,unknown>[]);
              const adicG=((data.adicionales||[]) as Record<string,unknown>[]);
              const asistG=((data.asistencia||[]) as Record<string,unknown>[]);
              const suspG=((data.suspensiones||[]) as Record<string,unknown>[]);
              const maqG=((data.maquinaria||[]) as Record<string,unknown>[]);
              const repEspMap:Record<string,string>={};
              repsG.forEach(r=>{repEspMap[r.id as string]=r.especialidad_id as string;});

              // Agrupar por especialidad
              const porEsp:Record<string,{
                avances:Record<string,unknown>[];
                cual:Record<string,unknown>[];
                adic:Record<string,unknown>[];
                personal:Record<string,unknown>[];
                susps:Record<string,unknown>[];
                maq:Record<string,unknown>[];
              }>={};

              const getEspNom=(espId:string)=>catalogs?.especialidades_actividades.find(e=>e.id===espId)?.especialidad_es||espId;

              const addToEsp=(espId:string,tipo:'avances'|'cual'|'adic'|'personal'|'susps'|'maq',item:Record<string,unknown>)=>{
                if(!espId) return;
                const nom=getEspNom(espId);
                if(!porEsp[nom]) porEsp[nom]={avances:[],cual:[],adic:[],personal:[],susps:[],maq:[]};
                porEsp[nom][tipo].push(item);
              };

              avG.forEach(av=>addToEsp(av.especialidad_id as string,'avances',av));
              cualG.forEach(av=>addToEsp(av.especialidad_id as string,'cual',av));
              adicG.forEach(ad=>addToEsp(repEspMap[ad.reporte_id as string]||'','adic',ad));
              asistG.forEach(a=>addToEsp(repEspMap[a.reporte_id as string]||'','personal',a));
              suspG.forEach(s=>addToEsp(repEspMap[s.reporte_id as string]||'','susps',s));
              maqG.forEach(m=>addToEsp(repEspMap[m.reporte_id as string]||'','maq',m));

              delete porEsp[''];
              if(!Object.keys(porEsp).length) return null;

              return(
                <div className="space-y-4">
                  {Object.entries(porEsp).map(([espNom,items])=>(
                    <div key={espNom} className="card p-4">
                      {/* Cabecera especialidad */}
                      <div className="bg-[#003b7a] text-white px-4 py-2 rounded-lg mb-4 font-bold text-sm">🌿 {espNom}</div>

                      {/* AVANCES CUANTITATIVOS agrupados por fecha — solo en modo Detallado */}
                      {modo==='detallado'&&(()=>{
                        const porFecha:Record<string,Record<string,unknown>[]>={};
                        items.avances.forEach(av=>{
                          const f=av.fecha as string;
                          if(!porFecha[f]) porFecha[f]=[];
                          porFecha[f].push(av);
                        });
                        return Object.entries(porFecha).sort(([a],[b])=>a.localeCompare(b)).map(([fecha2,avs])=>{
                          const repDia=(data.reportes as Record<string,unknown>[]).find(r=>r.fecha===fecha2&&(catalogs?.especialidades_actividades.find(e=>e.id===(r.especialidad_id as string))?.especialidad_es||'')=== espNom);
                          return(
                          <div key={fecha2} className="mb-3 border-l-4 border-[#003b7a] pl-4">
                            <div className="font-bold text-xs text-[#003b7a] mb-2">📅 {fecha2}</div>
                            {repDia&&<div className="bg-slate-50 border border-slate-200 rounded-lg p-3 mb-3 text-xs"><div className="grid grid-cols-2 sm:grid-cols-4 gap-2"><div><span className="text-slate-400">Jornada</span><div className="font-medium">{repDia.jornada_horas as number}h</div></div><div><span className="text-slate-400">Clima</span><div className="font-medium">{(repDia.clima as string)||'—'}</div></div>{(repDia.charla_preturno as boolean)&&<div className="col-span-2"><span className="text-slate-400">Charla pre-turno</span><div className="font-medium">✅ {(repDia.charla_tema as string)||'Charla realizada'}</div></div>}</div></div>}
                            {avs.map((av,i)=>{
                              const aR=catalogs?.especialidades_actividades.find(e=>e.id===(av.actividad_id as string));
                              const arR=catalogs?.areas.find(a=>a.id===(av.area_id as string));
                              const cfg=configActs.find(c=>c.actividad_id===(av.actividad_id as string));
                              const pct=cfg?.meta_total&&cfg.tiene_meta?Math.min(100,Math.round((av.acumulado_total_real as number)/cfg.meta_total*100)):null;
                              return(
                                <div key={i} className="py-2 border-b border-slate-100">
                                  <div className="flex items-start justify-between flex-wrap gap-2">
                                    <div>
                                      <div className="font-medium text-sm text-[#003b7a]">{aR?.actividad_es||''}</div>
                                      <div className="text-xs text-slate-400">📍 {arR?.area_es||'—'}</div>
                                    </div>
                                    <div className="text-right">
                                      <div className="font-bold text-sm text-[#003b7a]">{av.cantidad as number} {av.unidad as string}</div>
                                      <div className="text-xs text-slate-400">Acum: {av.acumulado_total_real as number}{pct!==null&&<span className={`ml-1 font-bold ${pct>=90?'text-emerald-600':pct>=50?'text-blue-600':'text-amber-600'}`}>({pct}%)</span>}</div>
                                    </div>
                                  </div>
                                  {(av.observacion_es as string)&&user.rol!=='cliente'&&<div className="text-xs text-slate-500 mt-1 italic pl-2">💬 "{av.observacion_es as string}"</div>}
                                </div>
                              );
                            })}
                          </div>
                        );});
                      })()}

                      {/* CUALITATIVAS */}
                      {items.cual.length>0&&(
                        <div className="mt-3 mb-3">
                          <div className="text-xs font-bold text-purple-700 uppercase tracking-wide mb-2">📋 Actividades cualitativas</div>
                          {items.cual.map((a,i)=>{
                            const aR=catalogs?.especialidades_actividades.find(e=>e.id===(a.actividad_id as string));
                            return(
                              <div key={i} className="bg-purple-50 border border-purple-200 rounded-lg p-3 mb-2">
                                <div className="font-semibold text-sm text-purple-800">{aR?.actividad_es||'Actividad'}</div>
                                <div className="text-xs text-slate-500 mt-0.5">{a.fecha as string}</div>
                                <div className="text-xs text-slate-600 mt-1 italic">"{a.observacion_es as string||'Sin descripción'}"</div>
                              </div>
                            );
                          })}
                        </div>
                      )}

                      {/* ADICIONALES */}
                      {items.adic.length>0&&(
                        <div className="mt-3 mb-3">
                          <div className="text-xs font-bold text-indigo-700 uppercase tracking-wide mb-2">➕ Actividades adicionales</div>
                          {items.adic.map((a,i)=>(
                            <div key={i} className="bg-indigo-50 border border-indigo-200 rounded-lg p-3 mb-2">
                              <div className="font-semibold text-sm text-indigo-800">{a.nombre as string}</div>
                              <div className="text-xs text-slate-500 mt-0.5">{a.fecha as string}</div>
                              <div className="text-xs text-slate-600 mt-1 italic">"{a.descripcion_ejecutado as string||''}"</div>
                            </div>
                          ))}
                        </div>
                      )}

                      {/* PERSONAL al final de la especialidad — oculto para cliente */}
                      {items.personal.length>0&&user.rol!=='cliente'&&(()=>{
                        const porDoc:Record<string,{nombre:string;cargo:string;totalHoras:number;asistio:boolean;motivo:string;novedad:string}>={};
                        items.personal.forEach(a=>{
                          const doc=a.documento_personal as string;
                          const persInfo=catalogs?.personal.find(p=>p.documento===doc);
                          if(!porDoc[doc]) porDoc[doc]={
                            nombre:persInfo?.nombre||doc,
                            cargo:persInfo?.cargo_es||'—',
                            totalHoras:0,asistio:false,motivo:'',novedad:''
                          };
                          porDoc[doc].totalHoras+=parseFloat(String(a.horas_trabajadas||0));
                          if(a.asistio) porDoc[doc].asistio=true;
                          if(a.motivo_ausencia) porDoc[doc].motivo=a.motivo_ausencia as string;
                          if(a.novedad) porDoc[doc].novedad=a.novedad as string;
                        });
                        const lista=Object.values(porDoc);
                        return(
                          <div className="mt-4 border-t border-slate-200 pt-3">
                            <div className="text-xs font-bold text-slate-600 uppercase tracking-wide mb-2">👥 Personal — {espNom}</div>
                            <table className="w-full text-xs border-collapse">
                              <thead><tr className="bg-slate-50">
                                <th className="text-left px-3 py-2 font-medium text-slate-500">Nombre</th>
                                <th className="text-left px-3 py-2 font-medium text-slate-500">Cargo</th>
                                <th className="text-center px-3 py-2 font-medium text-slate-500">Estado</th>
                                <th className="text-center px-3 py-2 font-medium text-slate-500">Horas</th>
                                <th className="text-left px-3 py-2 font-medium text-slate-500">Motivo / Novedad</th>
                              </tr></thead>
                              <tbody>{lista.map((p,i)=>(
                                <tr key={i} className="border-t border-slate-100">
                                  <td className="px-3 py-1.5 font-medium">{p.nombre}</td>
                                  <td className="px-3 py-1.5 text-slate-500">{p.cargo}</td>
                                  <td className="px-3 py-1.5 text-center">{p.asistio?<span className="text-emerald-600 font-bold">✅</span>:<span className="text-rose-600 font-bold">❌</span>}</td>
                                  <td className="px-3 py-1.5 text-center"><span className="bg-blue-50 text-blue-700 font-bold px-2 py-0.5 rounded">{p.totalHoras}h</span></td>
                                  <td className="px-3 py-1.5 text-slate-400 italic text-xs">{p.motivo||p.novedad||'—'}</td>
                                </tr>
                              ))}</tbody>
                            </table>
                          </div>
                        );
                      })()}

                      {/* SUSPENSIONES Y MAQUINARIA al final de la especialidad — oculto para cliente */}
                      {(items.susps.length>0||items.maq.length>0)&&user.rol!=='cliente'&&(
                        <div className="mt-4 border-t border-slate-200 pt-3">
                          <div className="text-xs font-bold text-slate-600 uppercase tracking-wide mb-2">⚠️ Suspensiones y maquinaria — {espNom}</div>

                          {items.susps.length>0&&(
                            <div className="mb-3">
                              <div className="text-xs font-semibold text-amber-700 mb-1">Suspensiones</div>
                              {items.susps.map((s,i)=>(
                                <div key={i} className="bg-amber-50 border border-amber-200 rounded-lg p-2 mb-1 text-xs">
                                  <span className="font-medium">{s.es_general?'🌐 General':'📌 Por actividad'}</span>
                                  {(s.hora_inicio as string)&&<span className="ml-2 text-slate-500">{s.hora_inicio as string} — {s.hora_fin as string}</span>}
                                  <div className="text-slate-600 mt-0.5">{s.descripcion as string}</div>
                                </div>
                              ))}
                            </div>
                          )}

                          {items.maq.length>0&&(
                            <div>
                              <div className="text-xs font-semibold text-slate-600 mb-1">🚜 Maquinaria</div>
                              {items.maq.map((m,i)=>{
                                const maqNom=nombreMaquina(m.maquinaria_id as string);
                                return(
                                  <div key={i} className="bg-slate-50 border border-slate-200 rounded-lg p-2 mb-1 text-xs">
                                    <span className="font-medium">{maqNom}</span>
                                    {(m.hora_inicio as string)&&<span className="ml-2 text-slate-500">{m.hora_inicio as string} — {m.hora_fin as string}</span>}
                                    {(m.descripcion as string)&&<div className="text-slate-600 mt-0.5">{m.descripcion as string}</div>}
                                    {(m.novedad as string)&&<div className="text-slate-600 mt-0.5">Novedad: {m.novedad as string}</div>}
                                  </div>
                                );
                              })}
                            </div>
                          )}
                        </div>
                      )}
                    </div>
                  ))}
                </div>
              );
            })()}

            {!(data.avances as unknown[]).length&&!((data.cualitativas||[]) as unknown[]).length&&!((data.adicionales||[]) as unknown[]).length&&<div className="card p-6 text-center text-slate-500">Sin datos de avance para este período. Consulta primero.</div>}
          </>
        )}
        {!data&&<div className="card p-6 text-center text-slate-500">Selecciona el período y haz clic en Consultar.</div>}
      </div>
    </div>
  );
}

// ── CATÁLOGOS — con botones siempre visibles ───────────────────────
function ItemDatabasesSection(){
  const[databases,setDatabases]=useState<Record<string,unknown>[]>([]);
  const[dbSeleccionada,setDbSeleccionada]=useState<string|null>(null);
  const[items,setItems]=useState<Record<string,unknown>[]>([]);
  const[filtro,setFiltro]=useState<'todos'|'disponibles'|'bloqueados'>('todos');

  useEffect(()=>{
    supabase.from('item_databases')
      .select('*, config_actividades(actividad_id, especialidades_actividades(actividad_es))')
      .eq('activo',true)
      .then(({data})=>setDatabases(data||[]));
  },[]);

  async function cargarItems(dbId:string){
    setDbSeleccionada(dbId);
    const{data}=await supabase.from('items_database').select('*').eq('database_id',dbId).order('created_at');
    setItems(data||[]);
  }

  async function desbloquearItem(itemId:string){
    await supabase.from('items_database').update({bloqueado:false,bloqueado_fecha:null,bloqueado_en_reporte:null}).eq('id',itemId);
    setItems(prev=>prev.map(i=>i.id===itemId?{...i,bloqueado:false}:i));
  }

  const itemsFiltrados=items.filter(i=>{
    if(filtro==='disponibles') return !i.bloqueado;
    if(filtro==='bloqueados') return i.bloqueado;
    return true;
  });

  return(
    <div className="card p-4 mt-2">
      <h3 className="font-bold text-[#003b7a] mb-3">📦 Bases de datos de ítems</h3>
      {databases.length===0&&(
        <div className="text-sm text-slate-500 italic">No hay bases de datos configuradas. Ve a Config. Act. para agregar una.</div>
      )}
      {databases.length>0&&(
        <div className="mb-4">
          <select className="select w-full" value={dbSeleccionada||''} onChange={e=>cargarItems(e.target.value)}>
            <option value="">Selecciona una base de datos…</option>
            {databases.map(db=>(
              <option key={db.id as string} value={db.id as string}>
                {db.nombre as string} — {((db.config_actividades as Record<string,unknown>)?.especialidades_actividades as Record<string,string>)?.actividad_es||''}
              </option>
            ))}
          </select>
        </div>
      )}
      {dbSeleccionada&&(
        <>
          <div className="flex gap-2 mb-3">
            {(['todos','disponibles','bloqueados'] as const).map(f=>(
              <button key={f} className={`text-xs px-3 py-1.5 rounded-lg font-medium ${filtro===f?'bg-[#003b7a] text-white':'bg-slate-100 text-slate-600'}`} onClick={()=>setFiltro(f)}>
                {f==='todos'?'Todos':f==='disponibles'?'✅ Disponibles':'🔒 Bloqueados'}
              </button>
            ))}
            <span className="text-xs text-slate-500 ml-auto self-center">{itemsFiltrados.length} ítems</span>
          </div>
          <div className="overflow-x-auto">
            <table className="w-full text-xs">
              <thead>
                <tr className="bg-slate-50">
                  {items[0]&&Object.keys(items[0].datos as Record<string,string>).map(col=>(
                    <th key={col} className="text-left px-3 py-2 font-medium text-slate-600">{col}</th>
                  ))}
                  <th className="px-3 py-2 text-left font-medium text-slate-600">Estado</th>
                  <th className="px-3 py-2"></th>
                </tr>
              </thead>
              <tbody>
                {itemsFiltrados.map(item=>{
                  const datos=item.datos as Record<string,string>;
                  const bloqueado=item.bloqueado as boolean;
                  return(
                    <tr key={item.id as string} className="border-t border-slate-100">
                      {Object.values(datos).map((v,i)=>(
                        <td key={i} className="px-3 py-2 text-slate-700">{v}</td>
                      ))}
                      <td className="px-3 py-2">
                        {bloqueado?<span className="text-rose-600 font-medium">🔒 Bloqueado</span>:<span className="text-emerald-600 font-medium">✅ Disponible</span>}
                      </td>
                      <td className="px-3 py-2">
                        {bloqueado&&(
                          <button className="text-xs text-blue-600 hover:underline" onClick={()=>desbloquearItem(item.id as string)}>Desbloquear</button>
                        )}
                      </td>
                    </tr>
                  );
                })}
              </tbody>
            </table>
          </div>
        </>
      )}
    </div>
  );
}

function CatalogosModule({catalogs,onRefresh,showToast}:{catalogs:Catalogs|null;onRefresh:()=>void;showToast:(k:'ok'|'err'|'info',m:string)=>void}){
  return(
    <div className="space-y-4">
      <CatMgr title="Especialidades y Actividades" table="especialidades_actividades" nameField="actividad_es"
        fields={[{n:'especialidad_es',l:'Especialidad (ES)'},{n:'especialidad_en',l:'Especialidad (EN)'},{n:'actividad_es',l:'Actividad (ES)'},{n:'actividad_en',l:'Actividad (EN)'}]}
        rows={(catalogs?.especialidades_actividades||[]) as unknown as Record<string,unknown>[]} onChanged={onRefresh} showToast={showToast}/>
      <CatMgr title="Áreas" table="areas" nameField="area_es"
        fields={[{n:'area_es',l:'Área (ES)'},{n:'area_en',l:'Área (EN)'}]}
        rows={(catalogs?.areas||[]) as unknown as Record<string,unknown>[]} onChanged={onRefresh} showToast={showToast}/>
      <CatMgr title="Líderes" table="lideres" nameField="nombre"
        fields={[{n:'nombre',l:'Nombre'},{n:'documento',l:'Documento'},{n:'cargo_es',l:'Cargo (ES)'},{n:'cargo_en',l:'Cargo (EN)'}]}
        rows={(catalogs?.lideres||[]) as unknown as Record<string,unknown>[]} onChanged={onRefresh} showToast={showToast}/>
      <CatMgr title="Personal" table="personal" nameField="nombre"
        fields={[{n:'nombre',l:'Nombre'},{n:'documento',l:'Documento'},{n:'cargo_es',l:'Cargo (ES)'},{n:'cargo_en',l:'Cargo (EN)'},{n:'tipo',l:'Tipo'},{n:'empresa',l:'Empresa'}]}
        rows={(catalogs?.personal||[]) as unknown as Record<string,unknown>[]} onChanged={onRefresh} showToast={showToast}/>
      <ItemDatabasesSection />
    </div>
  );
}

function CatMgr({title,table,nameField,fields,rows,onChanged,showToast}:{
  title:string; table:string; nameField:string; fields:{n:string;l:string}[];
  rows:Record<string,unknown>[]; onChanged:()=>void;
  showToast:(k:'ok'|'err'|'info',m:string)=>void;
}){
  const[form,setForm]=useState<Record<string,string>>({});
  const[search,setSearch]=useState('');
  const[filtro,setFiltro]=useState<'activos'|'inactivos'|'todos'>('activos');
  const[busy,setBusy]=useState(false);

  const filtered=useMemo(()=>{
    let list=rows;
    if(filtro==='activos') list=list.filter(r=>r.activo!==false);
    else if(filtro==='inactivos') list=list.filter(r=>r.activo===false);
    const q=search.trim().toLowerCase();
    return q?list.filter(r=>fields.some(f=>String(r[f.n]||'').toLowerCase().includes(q))):list;
  },[rows,search,fields,filtro]);

  async function addOne(){
    if(!form[fields[0].n]){showToast('err',`Falta ${fields[0].l}`);return;}
    setBusy(true);
    try{
      const{error}=await supabase.from(table).insert({...form,activo:true});
      if(error) throw error;
      showToast('ok','Agregado'); setForm({}); onChanged();
    } catch(e:unknown){ showToast('err',(e as Error)?.message||'Error'); }
    finally{ setBusy(false); }
  }

  async function toggleActivo(row:Record<string,unknown>){
    const nuevo=row.activo===false;
    if(!window.confirm(`¿${nuevo?'Activar':'Desactivar'} "${row[nameField]}"?\n${!nuevo?'Ya no estará disponible para nuevas asignaciones. Los registros históricos se mantienen.':''}`)) return;
    const{error}=await supabase.from(table).update({activo:nuevo}).eq('id',row.id);
    if(error){showToast('err',error.message);return;}
    showToast('ok',nuevo?'Activado ✓':'Desactivado ✓'); onChanged();
  }

  async function eliminar(row:Record<string,unknown>){
    if(!window.confirm(`⚠️ ¿ELIMINAR PERMANENTEMENTE "${row[nameField]}"?\n\nSe recomienda DESACTIVAR para mantener el historial.`)) return;
    setBusy(true);
    try{
      const{error}=await supabase.from(table).delete().eq('id',row.id);
      if(error) throw error;
      showToast('ok','Eliminado'); onChanged();
    } catch(e:unknown){ showToast('err',`No se puede eliminar: ${(e as Error)?.message}. Intenta desactivarlo.`); }
    finally{ setBusy(false); }
  }

  async function loadExcel(file:File){
    setBusy(true);
    try{
      const buf=await file.arrayBuffer();
      const wb=XLSX.read(buf);
      const xlData=XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{defval:''}) as Record<string,unknown>[];
      const norm=xlData.map(row=>{
        const o:Record<string,unknown>={activo:true};
        for(const k in row){ o[String(k).trim().toLowerCase().replace(/\s+/g,'_')]=row[k]; }
        return o;
      });
      if(!norm.length){showToast('err','Excel vacío');return;}
      if(!window.confirm(`¿Reemplazar "${title}" con ${norm.length} filas del Excel?`)) return;
      await supabase.from(table).delete().neq('id','00000000-0000-0000-0000-000000000000');
      const{error}=await supabase.from(table).insert(norm);
      if(error) throw error;
      showToast('ok',`Catálogo actualizado con ${norm.length} registros`); onChanged();
    } catch(e:unknown){ showToast('err',(e as Error)?.message||'Error'); }
    finally{ setBusy(false); }
  }

  const activos=rows.filter(r=>r.activo!==false).length;
  const inactivos=rows.filter(r=>r.activo===false).length;

  return(
    <div className="card p-4">
      <div className="flex items-center justify-between flex-wrap gap-2 mb-3">
        <div>
          <h3 className="font-bold text-[#003b7a]">{title}</h3>
          <div className="text-xs text-slate-500">{activos} activos · {inactivos} inactivos · {rows.length} total</div>
        </div>
        <label className="btn-secondary cursor-pointer text-xs">📥 Cargar Excel
          <input type="file" accept=".xlsx,.xls,.csv" hidden onChange={e=>{const f=e.target.files?.[0];if(f)loadExcel(f);e.currentTarget.value='';}}/>
        </label>
      </div>

      <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-6 gap-2 mb-2">
        {fields.map(f=><div key={f.n}><label className="label">{f.l}</label><input className="input" value={form[f.n]||''} onChange={e=>setForm({...form,[f.n]:e.target.value})}/></div>)}
      </div>
      <button className="btn-primary text-xs mb-3" disabled={busy} onClick={addOne}>+ Agregar registro</button>

      <div className="flex gap-2 mb-3 flex-wrap">
        <input className="input flex-1 min-w-[160px]" placeholder="🔎 Buscar…" value={search} onChange={e=>setSearch(e.target.value)}/>
        {(['activos','inactivos','todos'] as const).map(f=>(
          <button key={f} onClick={()=>setFiltro(f)}
            className={`text-xs px-3 py-1.5 rounded border transition-colors ${filtro===f?'bg-[#003b7a] text-white border-[#003b7a]':'border-slate-300 text-slate-600'}`}>
            {f.charAt(0).toUpperCase()+f.slice(1)}
          </button>
        ))}
      </div>

      {filtered.length>0?(
        <div className="space-y-2 max-h-80 overflow-y-auto">
          {filtered.slice(0,200).map((r,i)=>(
            <div key={i} className={`flex items-center gap-2 p-2 rounded-lg border ${r.activo===false?'bg-slate-50 border-slate-200 opacity-60':'bg-white border-slate-200 hover:border-slate-300'}`}>
              <div className="flex-1 min-w-0">
                <div className="font-medium text-sm truncate">{String(r[nameField]||'')}</div>
                <div className="text-xs text-slate-500 truncate">
                  {fields.filter(f=>f.n!==nameField).map(f=>String(r[f.n]||'')).filter(Boolean).join(' · ')}
                </div>
              </div>
              <div className="flex-shrink-0">
                {r.activo===false?<span className="badge bg-rose-100 text-rose-700 text-xs">Inactivo</span>:<span className="badge bg-emerald-100 text-emerald-700 text-xs">Activo</span>}
              </div>
              <div className="flex gap-1 flex-shrink-0">
                <button
                  className={`text-xs px-2 py-1 rounded font-medium transition-colors ${r.activo===false?'bg-emerald-100 text-emerald-700 hover:bg-emerald-200':'bg-amber-100 text-amber-700 hover:bg-amber-200'}`}
                  onClick={()=>toggleActivo(r)}>
                  {r.activo===false?'Activar':'Desactivar'}
                </button>
                <button
                  className="text-xs px-2 py-1 rounded font-medium bg-rose-100 text-rose-700 hover:bg-rose-200 transition-colors"
                  onClick={()=>eliminar(r)} disabled={busy}>
                  Eliminar
                </button>
              </div>
            </div>
          ))}
          {filtered.length>200&&<div className="text-xs text-slate-500 text-center py-2">Mostrando 200 de {filtered.length}. Usa el buscador para ver más.</div>}
        </div>
      ):<div className="text-sm text-slate-500 text-center py-4">Sin registros con los filtros actuales.</div>}
    </div>
  );
}

// ── MAQUINARIA ────────────────────────────────────────────────────
function MaquinariaModule({maquinaria,onRefresh,showToast}:{maquinaria:Maq[];onRefresh:()=>void;showToast:(k:'ok'|'err'|'info',m:string)=>void}){
  const[form,setForm]=useState({tipo:'motosierra',tipoCustom:'',item_id:'',nombre:'',estado:'activo'});
  const[busy,setBusy]=useState(false);

  async function addOne(){
    if(!form.item_id){showToast('err','Falta ID del equipo');return;}
    if(form.tipo==='otro'&&!form.tipoCustom.trim()){showToast('err','Especifica el tipo de maquinaria');return;}
    const tipoFinal=form.tipo==='otro'?form.tipoCustom.trim():form.tipo;
    setBusy(true);
    try{const{error}=await supabase.from('maquinaria').insert({tipo:tipoFinal,item_id:form.item_id,nombre:form.nombre,estado:form.estado,horas_acum_operativas:0,horas_acum_standby:0});if(error)throw error;showToast('ok','Equipo agregado');setForm({tipo:'motosierra',tipoCustom:'',item_id:'',nombre:'',estado:'activo'});onRefresh();}
    catch(e:unknown){showToast('err',(e as Error)?.message||'Error');}
    finally{setBusy(false);}
  }
  async function cambiarEstado(m:Maq,nuevoEstado:string){
    if(!window.confirm(`¿Cambiar estado de ${m.item_id} a "${nuevoEstado}"?`))return;
    await supabase.from('maquinaria').update({estado:nuevoEstado}).eq('id',m.id);
    showToast('ok','Estado actualizado');onRefresh();
  }
  async function eliminar(m:Maq){
    if(!window.confirm(`⚠️ ¿ELIMINAR ${m.item_id}?`))return;
    setBusy(true);
    try{const{error}=await supabase.from('maquinaria').delete().eq('id',m.id);if(error)throw error;showToast('ok','Eliminado');onRefresh();}
    catch(e:unknown){showToast('err',`No se puede eliminar: ${(e as Error)?.message}`);}
    finally{setBusy(false);}
  }

  return(
    <div className="card p-4 space-y-3">
      <h3 className="font-bold text-[#003b7a]">Maquinaria ({maquinaria.length} equipos)</h3>
      <div className="bg-slate-50 rounded-lg p-3 border border-slate-200 space-y-3">
        <div className="font-semibold text-sm text-slate-700">Agregar equipo</div>
        <div className="grid grid-cols-2 sm:grid-cols-4 gap-2">
          <div><label className="label">Tipo</label>
            <select className="select" value={form.tipo} onChange={e=>setForm({...form,tipo:e.target.value,tipoCustom:''})}>
              <option value="motosierra">Motosierra</option>
              <option value="chipeadora">Chipeadora</option>
              <option value="camion">Camión</option>
              <option value="volqueta">Volqueta</option>
              <option value="excavadora">Excavadora</option>
              <option value="retroexcavadora">Retroexcavadora</option>
              <option value="otro">Otro…</option>
            </select>
          </div>
          {form.tipo==='otro'&&<div><label className="label">¿Cuál tipo?</label><input className="input" value={form.tipoCustom} onChange={e=>setForm({...form,tipoCustom:e.target.value})} placeholder="Ej: Volqueta doble troque"/></div>}
          <div><label className="label">ID único</label><input className="input" value={form.item_id} onChange={e=>setForm({...form,item_id:e.target.value})} placeholder="MS-009"/></div>
          <div><label className="label">Nombre / descripción</label><input className="input" value={form.nombre} onChange={e=>setForm({...form,nombre:e.target.value})}/></div>
          <div><label className="label">Estado inicial</label><select className="select" value={form.estado} onChange={e=>setForm({...form,estado:e.target.value})}><option value="activo">Activo</option><option value="inactivo">Inactivo</option><option value="mantenimiento">Mantenimiento</option></select></div>
        </div>
        <button className="btn-primary text-xs" disabled={busy} onClick={addOne}>+ Agregar equipo</button>
      </div>

      <div className="overflow-x-auto">
        <table className="table w-full text-sm">
          <thead><tr><th>Código</th><th>Tipo</th><th>Nombre</th><th>Horas op.</th><th>Stand-by</th><th>Estado</th><th></th></tr></thead>
          <tbody>
            {maquinaria.map(m=>(
              <tr key={m.id} className={m.estado==='inactivo'?'opacity-50':''}>
                <td className="font-semibold text-[#003b7a]">{m.item_id}</td>
                <td>{m.tipo}</td>
                <td className="text-slate-600">{m.nombre||'—'}</td>
                <td className="text-emerald-600">{(m.horas_acum_operativas||0).toFixed(1)}h</td>
                <td className="text-amber-500">{(m.horas_acum_standby||0).toFixed(1)}h</td>
                <td>
                  <select className="select text-xs w-32" value={m.estado} onChange={e=>cambiarEstado(m,e.target.value)}>
                    <option value="activo">Activo</option>
                    <option value="inactivo">Inactivo</option>
                    <option value="mantenimiento">Mantenimiento</option>
                  </select>
                </td>
                <td>
                  <button className="text-xs px-2 py-1 rounded bg-rose-100 text-rose-700 hover:bg-rose-200" onClick={()=>eliminar(m)} disabled={busy}>Eliminar</button>
                </td>
              </tr>
            ))}
          </tbody>
        </table>
        {!maquinaria.length&&<div className="text-sm text-slate-500 text-center py-4">Sin equipos registrados.</div>}
      </div>
    </div>
  );
}

// ── CONFIG ACTIVIDADES — con nombres visibles y editar/eliminar ────
function ConfigActModule({user,configActs,catalogs,onRefresh,showToast}:{
  user:Profile; configActs:ConfigAct[]; catalogs:Catalogs|null; onRefresh:()=>void;
  showToast:(k:'ok'|'err'|'info',m:string)=>void;
}){
  const[form,setForm]=useState({especialidad_id:'',actividad_id:'',tipo:'A',unidad_es:'',unidad_en:'',meta_total:'',acumulado_previo:'',rendimiento_esperado:'',rendimiento_por:'cuadrilla',tiene_meta:true,es_medible:true});
  const[busy,setBusy]=useState(false);
  const[dbForm,setDbForm]=useState({nombre:'',bloqueo_tipo:'permanente' as 'permanente'|'temporal',columnas:[{nombre:'ID',tipo:'text'},{nombre:'Nombre',tipo:'text'}] as {nombre:string;tipo:string}[]});
  const[dbBusy,setDbBusy]=useState(false);
  const[existingDb,setExistingDb]=useState<ItemDatabase|null>(null);
  const[dbItems,setDbItems]=useState<ItemDB[]>([]);
  const[showDbPanel,setShowDbPanel]=useState(false);
  const[tieneDatabase,setTieneDatabase]=useState(false);
  const[dbNombre,setDbNombre]=useState('');
  const[dbColumnas,setDbColumnas]=useState<string[]>(['ID']);
  const[dbBloqueoTipo,setDbBloqueoTipo]=useState<'permanente'|'temporal'>('permanente');
  const[dbItemsNuevos,setDbItemsNuevos]=useState<Record<string,string>[]>([]);
  const[nuevaColumna,setNuevaColumna]=useState('');
  const[itemManual,setItemManual]=useState<Record<string,string>>({});
  const[mostrarFormItem,setMostrarFormItem]=useState(false);
  const[ajusteFor,setAjusteFor]=useState<string|null>(null);
  const[ajusteForm,setAjusteForm]=useState({fecha:today(),cantidad:'',motivo:''});
  const[ajusteBusy,setAjusteBusy]=useState(false);

  async function recalcularAcumulado(actividadId:string){
    const{data}=await supabase.from('avance_diario').select('id,fecha,created_at,cantidad')
      .eq('actividad_id',actividadId).neq('unidad','cualitativo').order('fecha',{ascending:true}).order('created_at',{ascending:true});
    let run=0;
    for(const r of (data||[]) as Record<string,unknown>[]){
      const anterior=run;
      run+=parseFloat(String(r.cantidad||0));
      if(r.acumulado_anterior!==anterior||r.acumulado_total!==run){
        await supabase.from('avance_diario').update({acumulado_anterior:anterior,acumulado_total:run}).eq('id',r.id as string);
      }
    }
  }

  async function guardarAjuste(cfg:ConfigAct){
    const cant=parseFloat(ajusteForm.cantidad);
    if(!ajusteForm.fecha){showToast('err','Indica la fecha del ajuste');return;}
    if(!cant||isNaN(cant)){showToast('err','Indica una cantidad de ajuste distinta de cero (puede ser negativa)');return;}
    if(!ajusteForm.motivo.trim()){showToast('err','Indica el motivo del ajuste');return;}
    setAjusteBusy(true);
    try{
      const{data:rep,error:re}=await supabase.from('reportes_avance').insert({
        fecha:ajusteForm.fecha,usuario_id:user.id,usuario_nombre:`${user.nombre} (ajuste admin)`,
        especialidad_id:cfg.especialidad_id,jornada_horas:0,clima:'despejado',charla_preturno:false,
        estado:'aprobado',aprobado_por:user.id,aprobado_en:new Date().toISOString(),
      }).select().single();
      if(re||!rep) throw new Error(re?.message||'Error creando el reporte de ajuste');
      const{error:avErr}=await supabase.from('avance_diario').insert({
        reporte_id:(rep as Record<string,unknown>).id as string,fecha:ajusteForm.fecha,usuario_id:user.id,
        actividad_id:cfg.actividad_id,especialidad_id:cfg.especialidad_id,area_id:null,
        cantidad:cant,unidad:cfg.unidad_es,acumulado_anterior:0,acumulado_total:0,
        observacion_es:`[AJUSTE] ${ajusteForm.motivo.trim()}`,
      });
      if(avErr) throw new Error(avErr.message);
      await recalcularAcumulado(cfg.actividad_id);
      showToast('ok','✅ Ajuste registrado y acumulado recalculado');
      setAjusteFor(null); setAjusteForm({fecha:today(),cantidad:'',motivo:''});
      onRefresh();
    } catch(e:unknown){ showToast('err',(e as Error)?.message||'Error'); }
    finally{ setAjusteBusy(false); }
  }

  async function loadItemDb(cfgId:string){
    const{data}=await supabase.from('item_databases').select('*').eq('config_actividad_id',cfgId).maybeSingle();
    if(data){
      const db=data as ItemDatabase;
      setExistingDb(db);
      setDbForm({nombre:db.nombre,bloqueo_tipo:db.bloqueo_tipo,columnas:db.columnas as {nombre:string;tipo:string}[]});
      const{data:items}=await supabase.from('items_database').select('*').eq('database_id',db.id).limit(5);
      setDbItems((items||[]) as ItemDB[]);
    } else { setExistingDb(null); }
    setShowDbPanel(true);
  }

  async function saveItemDb(cfgId:string){
    if(!dbForm.nombre.trim()||!dbForm.columnas.length){showToast('err','Nombre y al menos una columna requeridos');return;}
    setDbBusy(true);
    try{
      if(existingDb){
        await supabase.from('item_databases').update({nombre:dbForm.nombre,bloqueo_tipo:dbForm.bloqueo_tipo,columnas:dbForm.columnas}).eq('id',existingDb.id);
      } else {
        await supabase.from('item_databases').insert({config_actividad_id:cfgId,nombre:dbForm.nombre,bloqueo_tipo:dbForm.bloqueo_tipo,columnas:dbForm.columnas,activo:true});
        await supabase.from('config_actividades').update({tiene_items_unicos:true}).eq('id',cfgId);
      }
      showToast('ok','Base de datos guardada'); onRefresh(); loadItemDb(cfgId);
    } catch(e:unknown){ showToast('err',(e as Error)?.message||'Error'); }
    finally{ setDbBusy(false); }
  }

  function uploadItemsExcel(cfgId:string,file:File){
    const cfg=configActs.find(c=>c.id===cfgId);
    if(!cfg) return;
    const reader=new FileReader();
    reader.onload=async(ev)=>{
      try{
        const wb=XLSX.read(ev.target?.result,{type:'binary'});
        const ws=wb.Sheets[wb.SheetNames[0]];
        const rows=XLSX.utils.sheet_to_json(ws) as Record<string,unknown>[];
        if(!rows.length){showToast('err','Excel vacío');return;}
        let dbId=existingDb?.id;
        if(!dbId){
          const{data:newDb}=await supabase.from('item_databases').insert({config_actividad_id:cfgId,nombre:dbForm.nombre||'Base de datos',bloqueo_tipo:dbForm.bloqueo_tipo,columnas:dbForm.columnas,activo:true}).select().single();
          if(newDb) dbId=(newDb as Record<string,unknown>).id as string;
          await supabase.from('config_actividades').update({tiene_items_unicos:true}).eq('id',cfgId);
        }
        if(!dbId) return;
        const payload=rows.map(r=>({database_id:dbId,datos:Object.fromEntries(Object.entries(r).map(([k,v])=>[k,String(v||'')])),bloqueado:false,bloqueado_fecha:null,bloqueado_en_reporte:null}));
        await supabase.from('items_database').insert(payload);
        showToast('ok',`${payload.length} ítems cargados`); onRefresh(); loadItemDb(cfgId);
      } catch(e:unknown){ showToast('err',(e as Error)?.message||'Error al leer Excel'); }
    };
    reader.readAsBinaryString(file);
  }
  const esps=useMemo(()=>catalogs?uniqueEsp(catalogs.especialidades_actividades.filter(e=>e.activo!==false)):[],[catalogs]);
  const acts=useMemo(()=>catalogs&&form.especialidad_id?actsForEsp(catalogs.especialidades_actividades.filter(e=>e.activo!==false),form.especialidad_id):[],[catalogs,form.especialidad_id]);
  const allActs=useMemo(()=>catalogs?.especialidades_actividades||[],[catalogs]);

  function selAct(actId:string){
    const cfg=configActs.find(c=>c.actividad_id===actId);
    if(cfg) setForm({...form,actividad_id:actId,tipo:cfg.tipo||'A',unidad_es:cfg.unidad_es||'',unidad_en:cfg.unidad_en||'',meta_total:String(cfg.meta_total||''),tiene_meta:cfg.tiene_meta!==false,es_medible:cfg.es_medible!==false,acumulado_previo:String(cfg.acumulado_previo||''),rendimiento_esperado:String(cfg.rendimiento_esperado||'')});
    else setForm({...form,actividad_id:actId});
  }

  async function save(){
    if(!form.actividad_id){showToast('err','Selecciona una actividad');return;}
    if(form.tipo!=='D'&&!form.unidad_es){showToast('err','Ingresa la unidad de medida');return;}
    if(form.tipo==='A'&&!form.meta_total){showToast('err','Ingresa la meta total (requerida para tipo A)');return;}
    setBusy(true);
    try{
      const payload=form.tipo==='D'?{
        especialidad_id:form.especialidad_id,actividad_id:form.actividad_id,
        tipo:'D',es_medible:false,tiene_meta:false,meta_total:null,
        unidad_es:'cualitativo',unidad_en:'qualitative',
        tiene_items_unicos:false,rendimiento_esperado:null,rendimiento_por:null,
        acumulado_previo:0,activo:true,
      }:{
        especialidad_id:form.especialidad_id,actividad_id:form.actividad_id,
        tipo:form.tipo,unidad_es:form.unidad_es||'N/A',unidad_en:form.unidad_en||'N/A',
        tiene_meta:!!(form.tiene_meta&&form.es_medible),
        es_medible:!!form.es_medible,
        meta_total:form.tiene_meta&&form.es_medible&&form.meta_total?parseFloat(form.meta_total):null,
        acumulado_previo:parseFloat(form.acumulado_previo||'0'),
        rendimiento_esperado:form.rendimiento_esperado?parseFloat(form.rendimiento_esperado):null,
        rendimiento_por:form.rendimiento_por,activo:true,
      };
      const{error}=await supabase.from('config_actividades').upsert(payload as unknown as Record<string,unknown>,{onConflict:'actividad_id'});
      if(error) throw error;
      if(tieneDatabase&&dbNombre.trim()&&dbItemsNuevos.length>0){
        const{data:savedCfg}=await supabase.from('config_actividades').select('id').eq('actividad_id',form.actividad_id).single();
        if(savedCfg){
          const cfgId=(savedCfg as Record<string,unknown>).id as string;
          const{data:dbSaved}=await supabase.from('item_databases').upsert({config_actividad_id:cfgId,nombre:dbNombre.trim(),columnas:dbColumnas.map(c=>({nombre:c})),bloqueo_tipo:dbBloqueoTipo,activo:true}).select().single();
          if(dbSaved){
            const dbId=(dbSaved as Record<string,unknown>).id as string;
            await supabase.from('items_database').delete().eq('database_id',dbId);
            await supabase.from('items_database').insert(dbItemsNuevos.map(item=>({database_id:dbId,datos:item,bloqueado:false})));
          }
        }
      }
      showToast('ok','✅ Configuración guardada'); onRefresh();
    } catch(e:unknown){ showToast('err',(e as Error)?.message||'Error'); }
    finally{ setBusy(false); }
  }

  async function eliminarConfig(actId:string){
    if(!window.confirm('¿Eliminar la configuración? Los técnicos no podrán reportar esta actividad hasta que la reconfigures.')) return;
    const{error}=await supabase.from('config_actividades').delete().eq('actividad_id',actId);
    if(error){showToast('err',error.message);return;}
    showToast('ok','Configuración eliminada'); onRefresh();
  }

  return(
    <div className="card p-4 space-y-4">
      <div>
        <h3 className="font-bold text-[#003b7a]">Configuración de actividades ({configActs.length} configuradas)</h3>
        <p className="text-xs text-amber-700 bg-amber-50 border border-amber-200 rounded p-2 mt-2">⚠️ Configura cada actividad aquí antes de que los técnicos puedan reportarla.</p>
      </div>

      <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
        <div><label className="label">Especialidad</label>
          <select className="select" value={form.especialidad_id} onChange={e=>setForm({...form,especialidad_id:e.target.value,actividad_id:''})}>
            <option value="">— Seleccionar —</option>
            {esps.map(e=><option key={e.id} value={e.id}>{e.especialidad_es}</option>)}
          </select>
        </div>
        <div><label className="label">Actividad</label>
          <select className="select" value={form.actividad_id} disabled={!form.especialidad_id} onChange={e=>selAct(e.target.value)}>
            <option value="">— Seleccionar —</option>
            {acts.map(a=>{const yaConf=configActs.some(c=>c.actividad_id===a.id);return<option key={a.id} value={a.id}>{a.actividad_es}{yaConf?' ✓':''}</option>;})}
          </select>
        </div>
      </div>

      <div className="bg-slate-50 rounded-lg p-4 space-y-4 border border-slate-200">
        <div><label className="label">Tipo de actividad</label>
          <div className="grid grid-cols-2 sm:grid-cols-4 gap-2">
            {[{v:'A',l:'A — Con meta',d:'Cantidad con meta total definida'},{v:'B',l:'B — Acumulativa',d:'Acumula sin límite'},{v:'C',l:'C — Ítems únicos',d:'Ítems con ID único'},{v:'D',l:'D — Cualitativa',d:'Solo descripción, sin cantidad'}].map(t=>(
              <button key={t.v} onClick={()=>{
                if(t.v==='D') setForm({...form,tipo:'D',es_medible:false,tiene_meta:false,unidad_es:'cualitativo',unidad_en:'qualitative'});
                else if(t.v==='B'||t.v==='C') setForm({...form,tipo:t.v,tiene_meta:false,es_medible:true});
                else setForm({...form,tipo:t.v,tiene_meta:true,es_medible:true});
              }}
                className={`p-2 rounded-lg border text-xs text-left transition-colors ${form.tipo===t.v?'bg-[#003b7a] text-white border-[#003b7a]':'bg-white border-slate-300 text-slate-700 hover:border-[#003b7a]'}`}>
                <div className="font-bold">{t.l}</div>
                <div className={`mt-0.5 ${form.tipo===t.v?'text-blue-200':'text-slate-400'}`}>{t.d}</div>
              </button>
            ))}
          </div>
        </div>

        {form.tipo!=='D'&&(
          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-3">
            <div><label className="label">Unidad de medida (ES)</label><input className="input" value={form.unidad_es} onChange={e=>setForm({...form,unidad_es:e.target.value})} placeholder="árboles, m³, ha…"/></div>
            <div><label className="label">Unidad (EN)</label><input className="input" value={form.unidad_en} onChange={e=>setForm({...form,unidad_en:e.target.value})}/></div>
            {(form.tipo==='A')&&<>
              <div><label className="label">Meta total del proyecto</label><input type="number" className="input" value={form.meta_total} onChange={e=>setForm({...form,meta_total:e.target.value})} placeholder="Ej: 3466"/></div>
            </>}
            {(form.tipo==='A'||form.tipo==='B')&&<>
              <div><label className="label">Acumulado previo (ya ejecutado antes)</label><input type="number" className="input" value={form.acumulado_previo} onChange={e=>setForm({...form,acumulado_previo:e.target.value})} placeholder="0"/></div>
            </>}
            <div><label className="label">Rendimiento esperado</label><input type="number" className="input" value={form.rendimiento_esperado} onChange={e=>setForm({...form,rendimiento_esperado:e.target.value})}/></div>
            <div><label className="label">Rendimiento por</label>
              <select className="select" value={form.rendimiento_por} onChange={e=>setForm({...form,rendimiento_por:e.target.value})}>
                <option value="cuadrilla">cuadrilla/día</option><option value="persona">persona/día</option><option value="equipo">equipo/día</option>
              </select>
            </div>
          </div>
        )}
        {form.tipo==='D'&&<p className="text-xs text-purple-700 bg-purple-50 border border-purple-200 rounded p-2">En el reporte diario, el técnico solo podrá escribir una descripción de lo ejecutado. No se pide cantidad ni área.</p>}
      </div>

      {/* BASE DE DATOS DE ÍTEMS */}
      <div className="border border-slate-200 rounded-xl p-4 mt-4">
        <div className="flex items-center justify-between mb-3">
          <div>
            <div className="font-bold text-sm text-[#003b7a]">📦 Base de datos de ítems</div>
            <div className="text-xs text-slate-500">Ej: inventario de árboles, puntos de muestreo</div>
          </div>
          <div className="flex gap-2">
            <button type="button" className={`text-sm px-4 py-1.5 rounded-lg font-medium ${!tieneDatabase?'bg-[#003b7a] text-white':'bg-slate-100 text-slate-600'}`} onClick={()=>setTieneDatabase(false)}>No</button>
            <button type="button" className={`text-sm px-4 py-1.5 rounded-lg font-medium ${tieneDatabase?'bg-[#003b7a] text-white':'bg-slate-100 text-slate-600'}`} onClick={()=>setTieneDatabase(true)}>Sí</button>
          </div>
        </div>
        {tieneDatabase&&(
          <div className="space-y-4">
            <div>
              <label className="text-xs font-medium text-slate-600 block mb-1">Nombre de la base</label>
              <input className="input w-full" placeholder="Ej: Inventario árboles BSL8" value={dbNombre} onChange={e=>setDbNombre(e.target.value)}/>
            </div>
            <div>
              <label className="text-xs font-medium text-slate-600 block mb-2">Tipo de bloqueo al usar un ítem</label>
              <div className="flex gap-6">
                <label className="flex items-start gap-2 cursor-pointer">
                  <input type="radio" name="bloqueo_nuevo" checked={dbBloqueoTipo==='permanente'} onChange={()=>setDbBloqueoTipo('permanente')} className="mt-0.5"/>
                  <div><div className="text-sm font-medium">Permanente</div><div className="text-xs text-slate-500">Bloqueado para siempre al usarse</div></div>
                </label>
                <label className="flex items-start gap-2 cursor-pointer">
                  <input type="radio" name="bloqueo_nuevo" checked={dbBloqueoTipo==='temporal'} onChange={()=>setDbBloqueoTipo('temporal')} className="mt-0.5"/>
                  <div><div className="text-sm font-medium">Temporal</div><div className="text-xs text-slate-500">Bloqueado solo el día que se usa</div></div>
                </label>
              </div>
            </div>
            <div>
              <label className="text-xs font-medium text-slate-600 block mb-2">Columnas de la base</label>
              <div className="flex flex-wrap gap-2 mb-2">
                {dbColumnas.map((col,i)=>(
                  <div key={i} className="flex items-center gap-1 bg-blue-50 border border-blue-200 rounded-lg px-2 py-1">
                    <span className="text-sm text-blue-800 font-medium">{col}</span>
                    {dbColumnas.length>1&&<button type="button" className="text-blue-400 hover:text-rose-500 ml-1" onClick={()=>setDbColumnas(prev=>prev.filter((_,idx)=>idx!==i))}>✕</button>}
                  </div>
                ))}
              </div>
              <div className="flex gap-2">
                <input className="input flex-1 text-sm" placeholder="Nombre de columna..." value={nuevaColumna} onChange={e=>setNuevaColumna(e.target.value)}
                  onKeyDown={e=>{if(e.key==='Enter'&&nuevaColumna.trim()){setDbColumnas(p=>[...p,nuevaColumna.trim()]);setNuevaColumna('');}}}/>
                <button type="button" className="btn-secondary text-sm" onClick={()=>{if(nuevaColumna.trim()){setDbColumnas(p=>[...p,nuevaColumna.trim()]);setNuevaColumna('');}}}>+ Columna</button>
              </div>
            </div>
            <div>
              <label className="text-xs font-medium text-slate-600 block mb-2">Ítems ({dbItemsNuevos.length} cargados)</label>
              <div className="flex gap-2 mb-3">
                <label className="btn-secondary text-sm cursor-pointer">
                  📥 Cargar Excel
                  <input type="file" accept=".xlsx,.xls" className="hidden" onChange={async e=>{
                    const file=e.target.files?.[0];if(!file)return;
                    const XLSX=await import('xlsx');
                    const buf=await file.arrayBuffer();
                    const wb=XLSX.read(buf);
                    const ws=wb.Sheets[wb.SheetNames[0]];
                    const rows=XLSX.utils.sheet_to_json(ws) as Record<string,unknown>[];
                    if(rows.length>0){
                      const cols=Object.keys(rows[0]);
                      setDbColumnas(cols);
                      setDbItemsNuevos(rows.map(r=>Object.fromEntries(Object.entries(r).map(([k,v])=>[k,String(v)]))));
                    }
                  }}/>
                </label>
                <button type="button" className="btn-secondary text-sm" onClick={()=>{setItemManual(Object.fromEntries(dbColumnas.map(c=>[c,''])));setMostrarFormItem(true);}}>+ Agregar manual</button>
              </div>
              {mostrarFormItem&&(
                <div className="border border-blue-200 bg-blue-50 rounded-xl p-3 mb-3">
                  <div className="grid grid-cols-2 gap-2 mb-2">
                    {dbColumnas.map(col=>(
                      <div key={col}>
                        <label className="text-xs text-slate-500 block mb-0.5">{col}</label>
                        <input className="input w-full text-sm" placeholder={col} value={itemManual[col]||''} onChange={e=>setItemManual(p=>({...p,[col]:e.target.value}))}/>
                      </div>
                    ))}
                  </div>
                  <div className="flex gap-2">
                    <button type="button" className="btn-primary text-sm" onClick={()=>{setDbItemsNuevos(p=>[...p,{...itemManual}]);setMostrarFormItem(false);}}>Agregar</button>
                    <button type="button" className="btn-secondary text-sm" onClick={()=>setMostrarFormItem(false)}>Cancelar</button>
                  </div>
                </div>
              )}
              {dbItemsNuevos.length>0&&(
                <div className="border border-slate-200 rounded-lg overflow-hidden">
                  <div className="bg-slate-50 px-3 py-1.5 text-xs font-medium text-slate-600 border-b">Vista previa — {dbItemsNuevos.length} ítems</div>
                  <div className="overflow-x-auto max-h-40">
                    <table className="w-full text-xs">
                      <thead><tr className="bg-slate-100">
                        {dbColumnas.map(col=><th key={col} className="text-left px-3 py-2 font-medium text-slate-600">{col}</th>)}
                        <th className="px-2 py-2"></th>
                      </tr></thead>
                      <tbody>
                        {dbItemsNuevos.slice(0,8).map((item,i)=>(
                          <tr key={i} className="border-t border-slate-100">
                            {dbColumnas.map(col=><td key={col} className="px-3 py-1.5 text-slate-700">{item[col]||'—'}</td>)}
                            <td className="px-2 py-1.5"><button type="button" className="text-rose-400 text-xs" onClick={()=>setDbItemsNuevos(p=>p.filter((_,idx)=>idx!==i))}>✕</button></td>
                          </tr>
                        ))}
                        {dbItemsNuevos.length>8&&<tr><td colSpan={dbColumnas.length+1} className="px-3 py-2 text-slate-400 text-center italic">... y {dbItemsNuevos.length-8} ítems más</td></tr>}
                      </tbody>
                    </table>
                  </div>
                </div>
              )}
            </div>
          </div>
        )}
      </div>

      <button className="btn-primary" disabled={busy||!form.actividad_id} onClick={save}>{busy?'Guardando…':'💾 Guardar configuración'}</button>

      {(()=>{
        const cfg=configActs.find(c=>c.actividad_id===form.actividad_id);
        if(!cfg) return null;
        return(
          <div className="border border-indigo-200 rounded-lg p-4 bg-indigo-50 space-y-3">
            <div className="flex items-center justify-between">
              <span className="font-semibold text-sm text-indigo-800">🗂️ Base de datos de ítems</span>
              {!showDbPanel&&<button className="text-xs text-indigo-600 underline" onClick={()=>loadItemDb(cfg.id)}>Configurar</button>}
            </div>
            {cfg.tiene_items_unicos&&!showDbPanel&&<p className="text-xs text-emerald-700">✅ Esta actividad tiene base de datos configurada. ({dbItems.length} ítems cargados)</p>}
            {showDbPanel&&(
              <div className="space-y-3">
                <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
                  <div><label className="label">Nombre de la base</label><input className="input" value={dbForm.nombre} onChange={e=>setDbForm(f=>({...f,nombre:e.target.value}))} placeholder="Inventario de árboles BSL8"/></div>
                  <div><label className="label">Tipo de bloqueo</label>
                    <select className="select" value={dbForm.bloqueo_tipo} onChange={e=>setDbForm(f=>({...f,bloqueo_tipo:e.target.value as 'permanente'|'temporal'}))}>
                      <option value="permanente">Permanente (se bloquea para siempre)</option>
                      <option value="temporal">Temporal (se bloquea solo el día que se usa)</option>
                    </select>
                  </div>
                </div>
                <div>
                  <label className="label">Columnas de la base de datos</label>
                  <div className="space-y-1.5">
                    {dbForm.columnas.map((col,ci)=>(
                      <div key={ci} className="flex gap-2 items-center">
                        <input className="input flex-1" value={col.nombre} onChange={e=>setDbForm(f=>({...f,columnas:f.columnas.map((c2,k)=>k===ci?{...c2,nombre:e.target.value}:c2)}))} placeholder="Nombre columna"/>
                        <select className="select w-28" value={col.tipo} onChange={e=>setDbForm(f=>({...f,columnas:f.columnas.map((c2,k)=>k===ci?{...c2,tipo:e.target.value}:c2)}))}>
                          <option value="text">Texto</option><option value="number">Número</option>
                        </select>
                        {dbForm.columnas.length>1&&<button className="text-rose-500 text-xs px-2" onClick={()=>setDbForm(f=>({...f,columnas:f.columnas.filter((_,k)=>k!==ci)}))}>✕</button>}
                      </div>
                    ))}
                    <button className="btn-secondary text-xs" onClick={()=>setDbForm(f=>({...f,columnas:[...f.columnas,{nombre:'',tipo:'text'}]}))}>+ Columna</button>
                  </div>
                </div>
                <div className="flex gap-2 flex-wrap items-center">
                  <button className="btn-primary text-sm" disabled={dbBusy} onClick={()=>saveItemDb(cfg.id)}>{dbBusy?'Guardando…':'💾 Guardar base de datos'}</button>
                  <label className="btn-secondary text-sm cursor-pointer">
                    📥 Cargar ítems desde Excel
                    <input type="file" accept=".xlsx,.xls" className="hidden" onChange={e=>{if(e.target.files?.[0]) uploadItemsExcel(cfg.id,e.target.files[0]);}}/>
                  </label>
                </div>
                {existingDb&&<p className="text-xs text-slate-500">Base existente: <strong>{existingDb.nombre}</strong> · {dbItems.length} ítems precargados (muestra). Bloqueo: {existingDb.bloqueo_tipo}.</p>}
              </div>
            )}
          </div>
        );
      })()}

      {configActs.length>0&&(
        <div>
          <div className="font-semibold text-sm text-slate-700 mb-2">Actividades configuradas ({configActs.length})</div>
          <div className="space-y-2 max-h-72 overflow-y-auto">
            {configActs.map((c,i)=>{
              const actRow=allActs.find(e=>e.id===c.actividad_id);
              return(
                <div key={i} className="rounded-lg border border-slate-200 bg-white">
                  <div className="flex items-center gap-2 p-3">
                    <div className="flex-1 min-w-0">
                      <div className="font-medium text-sm truncate text-[#003b7a]">{actRow?.actividad_es||'—'}</div>
                      <div className="text-xs text-slate-500">{actRow?.especialidad_es||'—'}</div>
                      <div className="text-xs text-slate-400 mt-0.5">
                        {c.tipo==='D'?'Cualitativa':c.tipo==='A'&&c.meta_total?`Meta: ${c.meta_total} ${c.unidad_es}`:c.tipo==='B'?`Acumulativo · ${c.unidad_es}`:`Tipo ${c.tipo} · ${c.unidad_es}`}
                      </div>
                    </div>
                    <span className={`badge flex-shrink-0 ${c.tipo==='A'?'bg-blue-100 text-blue-800':c.tipo==='B'?'bg-green-100 text-green-800':c.tipo==='C'?'bg-purple-100 text-purple-800':'bg-rose-100 text-rose-800'}`}>{c.tipo}</span>
                    <div className="flex gap-1 flex-shrink-0">
                      {c.tipo!=='D'&&<button className="text-xs px-2 py-1 rounded bg-slate-100 text-slate-700 hover:bg-slate-200" onClick={()=>{setAjusteFor(ajusteFor===c.actividad_id?null:c.actividad_id);setAjusteForm({fecha:today(),cantidad:'',motivo:''});}}>⚖️ Ajustar</button>}
                      <button className="text-xs px-2 py-1 rounded bg-amber-100 text-amber-700 hover:bg-amber-200" onClick={()=>{const e=allActs.find(x=>x.id===c.actividad_id);if(e){const espRow=allActs.find(x=>x.especialidad_es===e.especialidad_es);setForm({...form,especialidad_id:espRow?.id||'',actividad_id:c.actividad_id});selAct(c.actividad_id);}}} >Editar</button>
                      <button className="text-xs px-2 py-1 rounded bg-rose-100 text-rose-700 hover:bg-rose-200" onClick={()=>eliminarConfig(c.actividad_id)}>Eliminar</button>
                    </div>
                  </div>
                  {ajusteFor===c.actividad_id&&(
                    <div className="border-t border-slate-100 p-3 bg-slate-50 space-y-2">
                      <p className="text-xs text-slate-500">Registra una corrección al acumulado de <strong>{actRow?.actividad_es}</strong>. Queda como un movimiento auditable (fecha, cantidad y motivo) — no se sobrescribe nada, se suma o resta al histórico.</p>
                      <div className="grid grid-cols-1 sm:grid-cols-3 gap-2">
                        <div><label className="label">Fecha del ajuste</label><input type="date" className="input" value={ajusteForm.fecha} onChange={e=>setAjusteForm(f=>({...f,fecha:e.target.value}))}/></div>
                        <div><label className="label">Cantidad ({c.unidad_es}) — negativa para restar</label><input type="number" className="input" value={ajusteForm.cantidad} onChange={e=>setAjusteForm(f=>({...f,cantidad:e.target.value}))} placeholder="Ej: 250 o -80"/></div>
                        <div className="sm:col-span-1"><label className="label">Motivo</label><input className="input" value={ajusteForm.motivo} onChange={e=>setAjusteForm(f=>({...f,motivo:e.target.value}))} placeholder="Ej: reportes en papel de junio no digitalizados"/></div>
                      </div>
                      <div className="flex gap-2">
                        <button className="btn-primary text-xs" disabled={ajusteBusy} onClick={()=>guardarAjuste(c)}>{ajusteBusy?'Guardando…':'💾 Guardar ajuste'}</button>
                        <button className="btn-secondary text-xs" onClick={()=>setAjusteFor(null)}>Cancelar</button>
                      </div>
                    </div>
                  )}
                </div>
              );
            })}
          </div>
        </div>
      )}
    </div>
  );
}

// ── USUARIOS ──────────────────────────────────────────────────────
function UsuariosModule({showToast}:{showToast:(k:'ok'|'err'|'info',m:string)=>void}){
  const[users,setUsers]=useState<Profile[]>([]);
  const[form,setForm]=useState({nombre:'',correo:'',clave:'',rol:'tecnico' as UserRole,especialidad_id:''});
  const[busy,setBusy]=useState(false);
  const[loading,setLoading]=useState(true);
  const[search,setSearch]=useState('');

  async function loadUsers(){setLoading(true);const{data}=await supabase.from('profiles').select('*').order('nombre');setUsers((data||[]) as Profile[]);setLoading(false);}
  useEffect(()=>{loadUsers();},[]);

  const filtered=useMemo(()=>{const q=search.trim().toLowerCase();return q?users.filter(u=>u.nombre.toLowerCase().includes(q)||u.correo.toLowerCase().includes(q)):users;},[users,search]);

  async function createUser(){
    if(!form.nombre||!form.correo||!form.clave){showToast('err','Nombre, correo y clave son obligatorios');return;}
    if(form.clave.length<8){showToast('err','La clave debe tener mínimo 8 caracteres');return;}
    setBusy(true);
    try{
      const r=await fetch('/api/admin',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({action:'createUser',email:form.correo,password:form.clave,nombre:form.nombre,rol:form.rol,especialidad_id:form.especialidad_id})});
      const d=await r.json() as {ok:boolean;error?:string};
      if(!d.ok) throw new Error(d.error||'Error al crear usuario');
      showToast('ok','Usuario creado ✓');
      setForm({nombre:'',correo:'',clave:'',rol:'tecnico',especialidad_id:''});
      loadUsers();
    } catch(e:unknown){ showToast('err',(e as Error)?.message||'Error'); }
    finally{ setBusy(false); }
  }

  async function toggleActivo(u:Profile){
    if(!window.confirm(`¿${!u.activo?'Activar':'Desactivar'} a ${u.nombre}?\n${u.activo?'El usuario no podrá ingresar a la app.':''}`)) return;
    await supabase.from('profiles').update({activo:!u.activo,updated_at:new Date().toISOString()}).eq('id',u.id);
    showToast('ok',!u.activo?`${u.nombre} activado`:`${u.nombre} desactivado`);
    loadUsers();
  }

  async function cambiarRol(u:Profile,rol:string){
    if(!window.confirm(`¿Cambiar el rol de ${u.nombre} a "${rol}"?`)) return;
    await supabase.from('profiles').update({rol,updated_at:new Date().toISOString()}).eq('id',u.id);
    showToast('ok','Rol actualizado'); loadUsers();
  }

  async function eliminar(u:Profile){
    if(!window.confirm(`⚠️ ¿ELIMINAR al usuario ${u.nombre}?\n\nEsto eliminará su acceso. Sus reportes se mantienen.`)) return;
    setBusy(true);
    try{
      const r=await fetch('/api/admin',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({action:'deleteUser',user_id:u.id})});
      const d=await r.json() as {ok:boolean;error?:string};
      if(!d.ok) throw new Error(d.error||'Error');
      showToast('ok','Usuario eliminado'); loadUsers();
    } catch(e:unknown){ showToast('err',(e as Error)?.message||'Error'); }
    finally{ setBusy(false); }
  }

  const ROLES=['admin','lider','tecnico','gerencia','cliente','visualizador'];

  return(
    <div className="card p-4 space-y-4">
      <h3 className="font-bold text-[#003b7a]">Usuarios ({users.length})</h3>
      <div className="bg-slate-50 rounded-lg p-3 space-y-2 border border-slate-200">
        <div className="font-semibold text-sm text-slate-700 mb-2">Crear nuevo usuario</div>
        <div className="grid grid-cols-2 sm:grid-cols-3 gap-2">
          <div><label className="label">Nombre completo</label><input className="input" value={form.nombre} onChange={e=>setForm({...form,nombre:e.target.value})}/></div>
          <div><label className="label">Correo</label><input className="input" type="email" value={form.correo} onChange={e=>setForm({...form,correo:e.target.value})}/></div>
          <div><label className="label">Clave inicial (min. 8)</label><input className="input" type="password" value={form.clave} onChange={e=>setForm({...form,clave:e.target.value})}/></div>
          <div><label className="label">Rol</label>
            <select className="select" value={form.rol} onChange={e=>setForm({...form,rol:e.target.value as UserRole})}>
              {ROLES.map(r=><option key={r} value={r}>{r.charAt(0).toUpperCase()+r.slice(1)}</option>)}
            </select>
          </div>
        </div>
        <button className="btn-primary text-xs mt-1" disabled={busy} onClick={createUser}>{busy?'Creando…':'+ Crear usuario'}</button>
      </div>

      <input className="input" placeholder="🔎 Buscar por nombre o correo…" value={search} onChange={e=>setSearch(e.target.value)}/>

      {loading?<div className="text-sm text-slate-500 text-center py-4">Cargando usuarios…</div>:(
        <div className="space-y-2">
          {filtered.map(u=>(
            <div key={u.id} className="border border-slate-200 rounded-xl p-4 flex flex-col gap-2 sm:flex-row sm:items-center sm:justify-between">
              <div className="flex items-center gap-3 min-w-0">
                <div className={`w-10 h-10 rounded-full flex items-center justify-center text-white font-bold text-sm flex-shrink-0 ${u.activo?'bg-[#003b7a]':'bg-slate-400'}`}>
                  {u.nombre?.charAt(0)?.toUpperCase()||'?'}
                </div>
                <div className="min-w-0">
                  <div className="font-semibold text-sm text-[#003b7a]">{u.nombre||'Sin nombre'}</div>
                  <div className="text-xs text-slate-500">{u.correo}</div>
                  <div className="flex items-center gap-2 mt-1">
                    <span className={`text-xs font-bold uppercase px-2 py-0.5 rounded-full ${
                      u.rol==='admin'?'bg-red-100 text-red-700':
                      u.rol==='lider'?'bg-purple-100 text-purple-700':
                      u.rol==='tecnico'?'bg-blue-100 text-blue-700':
                      u.rol==='gerencia'?'bg-amber-100 text-amber-700':
                      u.rol==='cliente'?'bg-green-100 text-green-700':
                      'bg-slate-100 text-slate-600'
                    }`}>{u.rol}</span>
                    {!u.activo&&<span className="text-xs text-rose-500 font-medium">● Inactivo</span>}
                  </div>
                </div>
              </div>
              <div className="flex items-center gap-2 flex-wrap">
                <select className="select text-xs" value={u.rol} onChange={e=>cambiarRol(u,e.target.value)}>
                  {ROLES.map(r=><option key={r} value={r}>{r}</option>)}
                </select>
                <button className={`text-xs px-3 py-1.5 rounded-lg font-medium ${u.activo?'bg-amber-100 text-amber-700 hover:bg-amber-200':'bg-emerald-100 text-emerald-700 hover:bg-emerald-200'}`} onClick={()=>toggleActivo(u)}>
                  {u.activo?'Desactivar':'Activar'}
                </button>
                <button className="text-xs px-3 py-1.5 rounded-lg font-medium bg-rose-100 text-rose-700 hover:bg-rose-200" onClick={()=>eliminar(u)} disabled={busy}>Eliminar</button>
              </div>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

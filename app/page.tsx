'use client';
// Powerchina PDS 360 v2.2
import { useCallback, useEffect, useMemo, useState } from 'react';
import { supabase, type Profile, type UserRole } from '@/lib/supabase';
import * as XLSX from 'xlsx';

// ── TIPOS ─────────────────────────────────────────────────────────
type AppView = 'home'|'planear'|'reporte'|'aprobacion'|'solicitudes'|'dashboard'|'informes'|'catalogos'|'maquinaria'|'config_act'|'usuarios';

interface EspAct  { id:string; especialidad_es:string; especialidad_en:string; actividad_es:string; actividad_en:string; activo?:boolean; }
interface Area     { id:string; area_es:string; area_en:string; activo?:boolean; }
interface Lider    { id:string; nombre:string; documento:string; cargo_es:string; activo?:boolean; }
interface Personal { id:string; nombre:string; documento:string; cargo_es:string; tipo?:string; empresa?:string; activo?:boolean; }
interface Maq      { id:string; tipo:string; item_id:string; nombre:string; estado:string; horas_acum_operativas?:number; horas_acum_standby?:number; }
interface ConfigAct{ id:string; actividad_id:string; especialidad_id:string; tipo:string; unidad_es:string; unidad_en?:string; meta_total?:number; tiene_meta?:boolean; es_medible?:boolean; rendimiento_esperado?:number; rendimiento_por?:string; acumulado_previo?:number; }
interface Catalogs { especialidades_actividades:EspAct[]; areas:Area[]; lideres:Lider[]; personal:Personal[]; }
interface Notif    { id:string; titulo:string; mensaje:string; leida:boolean; created_at:string; }
interface SuspItem { uid:string; tipo_susp:string; otro_desc:string; hora_inicio:string; hora_fin:string; descripcion:string; }
interface AsistItem{ personal_id:string; documento_personal:string; nombre:string; cargo_es:string; asistio:boolean; motivo_ausencia:string; }
interface AreaRep  { uid:string; area_id:string; cantidad:string; }
interface ActRep   { uid:string; actividad_id:string; areas:AreaRep[]; descripcion_cualitativa:string; observacion_es:string; }
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
function horasLost(ss:SuspItem[]):number {
  return ss.reduce((a,s)=>{
    if(!s.hora_inicio||!s.hora_fin) return a;
    const[ih,im]=s.hora_inicio.split(':').map(Number);
    const[fh,fm]=s.hora_fin.split(':').map(Number);
    return a+Math.max(0,(fh+fm/60)-(ih+im/60));
  },0);
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
  const canEdit=u.rol==='admin'||u.rol==='lider'||u.rol==='tecnico';
  const canApprove=u.rol==='admin'||u.rol==='lider';

  return(
    <div className="min-h-screen flex flex-col">
      <Header user={u} onLogout={handleLogout} setView={setView} currentView={view} notifs={notifs} onReadNotifs={marcarLeidas}/>
      <main className="flex-1 p-3 sm:p-5 max-w-7xl mx-auto w-full">
        {view==='home'       &&<HomeScreen user={u} setView={setView} notifs={notifs}/>}
        {view==='planear'    &&canEdit&&<PlaneacionModule user={u} catalogs={catalogs} maquinaria={maquinaria} showToast={showToast}/>}
        {view==='reporte'    &&canEdit&&<ReporteModule user={u} catalogs={catalogs} maquinaria={maquinaria} configActs={configActs} showToast={showToast}/>}
        {view==='aprobacion' &&canApprove&&<AprobacionModule user={u} catalogs={catalogs} configActs={configActs} showToast={showToast} onRefreshNotifs={()=>loadNotifs(u.id)}/>}
        {view==='solicitudes'&&<SolicitudesModule user={u} catalogs={catalogs} showToast={showToast}/>}
        {view==='dashboard'  &&<DashboardModule catalogs={catalogs} configActs={configActs} showToast={showToast}/>}
        {view==='informes'   &&<InformesModule user={u} catalogs={catalogs} configActs={configActs} showToast={showToast}/>}
        {view==='catalogos'  &&u.rol==='admin'&&<CatalogosModule catalogs={catalogs} onRefresh={loadCatalogs} showToast={showToast}/>}
        {view==='maquinaria' &&u.rol==='admin'&&<MaquinariaModule maquinaria={maquinaria} onRefresh={loadCatalogs} showToast={showToast}/>}
        {view==='config_act' &&u.rol==='admin'&&<ConfigActModule configActs={configActs} catalogs={catalogs} onRefresh={loadCatalogs} showToast={showToast}/>}
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
  const canEdit=user.rol==='admin'||user.rol==='lider'||user.rol==='tecnico';
  const canApprove=user.rol==='admin'||user.rol==='lider';
  const tabs=[
    {key:'home' as AppView,label:'Inicio',show:true},
    {key:'planear' as AppView,label:'Planear',show:canEdit},
    {key:'reporte' as AppView,label:'Reporte',show:canEdit},
    {key:'aprobacion' as AppView,label:'✅ Aprobar',show:canApprove},
    {key:'solicitudes' as AppView,label:'Solicitudes',show:true},
    {key:'dashboard' as AppView,label:'Dashboard',show:true},
    {key:'informes' as AppView,label:'Informes',show:true},
    {key:'catalogos' as AppView,label:'Catálogos',show:user.rol==='admin'},
    {key:'maquinaria' as AppView,label:'Maquinaria',show:user.rol==='admin'},
    {key:'config_act' as AppView,label:'Config. Act.',show:user.rol==='admin'},
    {key:'usuarios' as AppView,label:'Usuarios',show:user.rol==='admin'},
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
          <button className="btn-secondary text-xs" onClick={onLogout}>Salir</button>
        </div>
      </div>
      <nav className="max-w-7xl mx-auto px-2 overflow-x-auto">
        <div className="flex gap-0.5">{tabs.filter(t=>t.show).map(t=>(
          <button key={t.key} onClick={()=>setView(t.key)}
            className={`px-3 py-2 text-xs font-medium border-b-2 whitespace-nowrap transition-colors ${currentView===t.key?'border-white text-white':'border-transparent text-blue-200 hover:text-white'}`}>
            {t.label}
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
  const canEdit=user.rol==='admin'||user.rol==='lider'||user.rol==='tecnico';
  const canApprove=user.rol==='admin'||user.rol==='lider';
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

// ── PLANEACIÓN ────────────────────────────────────────────────────
function PlaneacionModule({user,catalogs,maquinaria,showToast}:{
  user:Profile; catalogs:Catalogs|null; maquinaria:Maq[];
  showToast:(k:'ok'|'err'|'info',m:string)=>void;
}){
  const[fecha,setFecha]=useState(today());
  const[actividades,setActividades]=useState<ActForm[]>([emptyAct()]);
  const[blocked,setBlocked]=useState<{documento_personal:string;usuario_nombre:string}[]>([]);
  const[estado,setEstado]=useState<'nuevo'|'borrador'|'enviado'>('nuevo');
  const[saving,setSaving]=useState(false);
  const[progId,setProgId]=useState<string|null>(null);

  const loadFecha=useCallback(async()=>{
    if(!fecha) return;
    const{data:bl}=await supabase.from('personal_asignado')
      .select('documento_personal,usuario_id,programaciones!inner(usuario_nombre)')
      .eq('fecha',fecha).neq('usuario_id',user.id);
    setBlocked((bl||[]).map((b:Record<string,unknown>)=>({
      documento_personal:b.documento_personal as string,
      usuario_nombre:((b.programaciones as Record<string,unknown>)?.usuario_nombre as string)||'otro',
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

  const isRO=estado==='enviado';

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
      const maqNombres=(a.maquinaria_ids||[]).map(id=>maquinaria.find(m=>m.id===id)?.item_id||id).join(", ");
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
      const maqNombres=(a.maquinaria_ids||[]).map(id=>maquinaria.find(m=>m.id===id)?.item_id||id).join(", ");
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
      <div className="card p-4 flex flex-col sm:flex-row gap-3 items-end justify-between flex-wrap">
        <div className="flex-1 min-w-[160px]"><label className="label">Fecha</label><input type="date" className="input" value={fecha} onChange={e=>setFecha(e.target.value)}/></div>
        <div className="flex items-center gap-3 flex-wrap">
          {estado==='nuevo'&&<span className="badge bg-slate-200 text-slate-700">Nuevo</span>}
          {estado==='borrador'&&<span className="badge-borrador">Borrador</span>}
          {estado==='enviado'&&<span className="badge-enviado">Enviado</span>}
          <button className="btn-secondary text-xs" onClick={exportarExcel}>📥 Excel</button>
          <button className="btn-secondary text-xs" onClick={imprimirPlan}>🖨️ PDF / Imprimir</button>
        </div>
      </div>

      {actividades.map((a,idx)=>(
        <ActCard key={a.uid} index={idx} act={a} catalogs={catalogs} maquinaria={maquinaria}
          blockedMap={blockedMap} readOnly={isRO}
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
function ActCard({index,act,catalogs,maquinaria,blockedMap,readOnly,onChange,onRemove}:{
  index:number; act:ActForm; catalogs:Catalogs|null; maquinaria:Maq[];
  blockedMap:Record<string,string>; readOnly:boolean;
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
                  {m.item_id}
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
                return(
                  <div key={p.id} className="relative group">
                    <label className={`flex items-center gap-2 px-3 py-2 border-b border-slate-100 cursor-pointer text-sm ${bl?'bg-rose-50 cursor-not-allowed':sel?'bg-blue-50':'hover:bg-slate-50'}`}>
                      <input type="checkbox" checked={sel} disabled={readOnly||!!bl}
                        onChange={()=>onChange({personal:sel?act.personal.filter(x=>x.documento_personal!==p.documento):[...act.personal,{personal_id:p.id,documento_personal:p.documento}]})}/>
                      <div className="flex-1 min-w-0">
                        <div className={`font-medium truncate ${bl?'text-slate-400':''}`}>{p.nombre}</div>
                        <div className="text-xs text-slate-500">{p.documento} · {p.cargo_es}</div>
                      </div>
                      {bl&&<span className="text-xs text-rose-500 font-medium flex-shrink-0">🔒 {bl}</span>}
                      {sel&&!bl&&<span className="badge bg-blue-100 text-blue-800 flex-shrink-0">✓</span>}
                    </label>
                    {bl&&<div className="absolute left-0 bottom-full mb-1 hidden group-hover:block z-20 bg-slate-800 text-white text-xs rounded p-2 shadow-lg whitespace-nowrap">Asignado por: <strong>{bl}</strong></div>}
                  </div>
                );
              })}
            </div>
          </div>
        </div>

        {/* PANEL RESUMEN — derecho */}
        <div className="lg:w-64 border-t lg:border-t-0 lg:border-l border-slate-200 bg-slate-50 p-4 flex-shrink-0">
          <div className="text-xs font-bold text-[#003b7a] uppercase mb-3">📋 Resumen</div>

          <div className="space-y-2 text-xs">
            <div><span className="text-slate-400">Área:</span> <span className="font-medium">{areaSel?.area_es||'—'}</span></div>
            <div><span className="text-slate-400">Líder:</span> <span className="font-medium">{liderSel?.nombre||'—'}</span>{liderSel&&<div className="text-slate-500">{liderSel.cargo_es}</div>}</div>

            {maqSel.length>0&&(
              <div><span className="text-slate-400">Maquinaria:</span>
                <div className="flex flex-wrap gap-1 mt-1">{maqSel.map(m=><span key={m.id} className="badge bg-orange-100 text-orange-700">{m.item_id}</span>)}</div>
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
  const[maqDia,setMaqDia]=useState<{maquinaria_id:string;nombre:string;novedad:boolean;descripcion:string;hora_inicio:string;hora_fin:string;horas_standby:number}[]>([]);
  const[actReps,setActReps]=useState<ActRep[]>([]);
  const[incidente,setIncidente]=useState({tipo:'sin_novedad',descripcion:'',medidas:''});
  const[notaBit,setNotaBit]=useState('');
  const[saving,setSaving]=useState(false);
  const[solicitudOk,setSolicitudOk]=useState(false);

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
    supabase.from('personal_asignado').select('personal_id,documento_personal,personal!inner(nombre,cargo_es)').eq('fecha',fecha)
      .then(({data})=>{
        const seen=new Set<string>();
        setAsistencia((data||[]).filter((r:Record<string,unknown>)=>{
          const doc=r.documento_personal as string;
          if(seen.has(doc)) return false; seen.add(doc); return true;
        }).map((r:Record<string,unknown>)=>{
          const pers=r.personal as Record<string,unknown>;
          return { personal_id:r.personal_id as string, documento_personal:r.documento_personal as string, nombre:(pers?.nombre as string)||r.documento_personal as string, cargo_es:(pers?.cargo_es as string)||'', asistio:true, motivo_ausencia:'' };
        }));
      });
  },[fecha,espId,user.id,esDiaAnt]);

  useEffect(()=>{
    setMaqDia(maquinaria.filter(m=>m.estado==='activo').map(m=>({maquinaria_id:m.id,nombre:`${m.item_id} – ${m.tipo}`,novedad:false,descripcion:'',hora_inicio:'',hora_fin:'',horas_standby:0})));
  },[maquinaria]);

  useEffect(()=>{
    if(!catalogs||!espId) return;
    const espRow=catalogs.especialidades_actividades.find(e=>e.id===espId);
    if(!espRow) return;
    const t=(espRow.especialidad_es||'').toLowerCase();
    const acts=catalogs.especialidades_actividades.filter(e=>(e.especialidad_es||'').toLowerCase()===t&&e.activo!==false);
    setActReps(acts.map(a=>({uid:gid(),actividad_id:a.id,areas:[{uid:gid(),area_id:'',cantidad:''}],descripcion_cualitativa:'',observacion_es:''})));
  },[catalogs,espId]);

  async function submit(){
    if(!fecha||!espId){showToast('err','Falta fecha o especialidad');return;}
    if(esDiaAnt&&!solicitudOk){showToast('err','Necesitas aprobación para reportar día anterior. Ve a Solicitudes.');return;}
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
          horas_trabajadas:a.asistio?horasReal:0,
        }))):null,
        maqDia.filter(m=>m.novedad).length?supabase.from('novedades_maquinaria').insert(
          maqDia.filter(m=>m.novedad).map(m=>({
            reporte_id:rid,fecha,usuario_id:user.id,maquinaria_id:m.maquinaria_id,
            descripcion:m.descripcion,hora_inicio:m.hora_inicio||null,hora_fin:m.hora_fin||null,horas_standby:m.horas_standby,
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
        if(cfg?.tipo==='D'){
          // Cualitativa — solo observación
          if(ar.descripcion_cualitativa.trim()){
            await supabase.from('avance_diario').insert({
              reporte_id:rid,fecha,usuario_id:user.id,actividad_id:ar.actividad_id,
              especialidad_id:espId,area_id:null,cantidad:0,unidad:'cualitativo',
              acumulado_anterior:0,acumulado_total:0,observacion_es:ar.descripcion_cualitativa,
            });
          }
          continue;
        }
        if(cfg?.es_medible===false) continue;
        for(const area of ar.areas.filter(a=>a.area_id&&parseFloat(a.cantidad)>0)){
          const{data:prev}=await supabase.from('avance_diario').select('cantidad')
            .eq('actividad_id',ar.actividad_id).eq('area_id',area.area_id).eq('usuario_id',user.id);
          const acumPrev=(prev||[]).reduce((s:number,r:Record<string,unknown>)=>s+parseFloat(String(r.cantidad||0)),0);
          const cantidad=parseFloat(area.cantidad);
          await supabase.from('avance_diario').insert({
            reporte_id:rid,fecha,usuario_id:user.id,actividad_id:ar.actividad_id,
            especialidad_id:espId,area_id:area.area_id,cantidad,
            unidad:cfg?.unidad_es||'',acumulado_anterior:acumPrev,acumulado_total:acumPrev+cantidad,
            observacion_es:ar.observacion_es,
          });
        }
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

  if(esDiaAnt&&!solicitudOk&&step===1){
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
              <div key={a.personal_id} className="flex items-center gap-3 p-2 border border-slate-200 rounded-lg flex-wrap">
                <input type="checkbox" checked={a.asistio} onChange={e=>setAsistencia(arr=>arr.map((x,j)=>j===i?{...x,asistio:e.target.checked,motivo_ausencia:''}:x))}/>
                <div className="flex-1"><div className={`text-sm font-medium ${a.asistio?'':'text-slate-400 line-through'}`}>{a.nombre}</div><div className="text-xs text-slate-400">{a.cargo_es}</div></div>
                {!a.asistio&&<select className="select text-xs w-auto" value={a.motivo_ausencia} onChange={e=>setAsistencia(arr=>arr.map((x,j)=>j===i?{...x,motivo_ausencia:e.target.value}:x))}>
                  <option value="">— Motivo —</option>
                  <option value="injustificada">Injustificada</option>
                  <option value="incapacidad">Incapacidad</option>
                  <option value="permiso">Permiso</option>
                </select>}
                {a.asistio&&<span className="text-xs text-emerald-600">{horasReal.toFixed(1)}h</span>}
              </div>
            ))}
          </div>
          {!asistencia.length&&<p className="text-sm text-slate-500">Sin personal planeado para este día.</p>}
          <div className="flex gap-2"><button className="btn-secondary" onClick={()=>setStep(3)}>← Anterior</button><button className="btn-primary" onClick={()=>setStep(5)}>Siguiente →</button></div>
        </div>
      )}

      {step===5&&(
        <div className="card p-4 space-y-3">
          <h3 className="font-bold text-[#003b7a]">Paso 5 — Maquinaria</h3>
          {maqDia.map((m,i)=>(
            <div key={m.maquinaria_id} className="border border-slate-200 rounded-lg p-3">
              <div className="flex items-center justify-between mb-2">
                <span className="font-medium text-sm">{m.nombre}</span>
                <span className="text-xs text-slate-500">Op: {Math.max(0,horasReal-m.horas_standby).toFixed(1)}h | SB: {m.horas_standby}h</span>
              </div>
              <label className="flex items-center gap-2 text-sm mb-2"><input type="checkbox" checked={m.novedad} onChange={e=>setMaqDia(a=>a.map((x,j)=>j===i?{...x,novedad:e.target.checked}:x))}/> ¿Tuvo novedad?</label>
              {m.novedad&&(
                <div className="grid grid-cols-1 sm:grid-cols-3 gap-2">
                  <div><label className="label">Hora inicio</label><input type="time" className="input" value={m.hora_inicio} onChange={e=>setMaqDia(a=>a.map((x,j)=>j===i?{...x,hora_inicio:e.target.value}:x))}/></div>
                  <div><label className="label">Hora fin</label><input type="time" className="input" value={m.hora_fin} onChange={e=>{
                    const[ih,im]=(m.hora_inicio||'0:0').split(':').map(Number);
                    const[fh,fm]=e.target.value.split(':').map(Number);
                    const diff=Math.max(0,(fh+fm/60)-(ih+im/60));
                    setMaqDia(a=>a.map((x,j)=>j===i?{...x,hora_fin:e.target.value,horas_standby:parseFloat(diff.toFixed(2))}:x));
                  }}/></div>
                  <div><label className="label">Descripción</label><input className="input" value={m.descripcion} onChange={e=>setMaqDia(a=>a.map((x,j)=>j===i?{...x,descripcion:e.target.value}:x))}/></div>
                </div>
              )}
            </div>
          ))}
          {!maqDia.length&&<p className="text-sm text-slate-500">Sin maquinaria activa.</p>}
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
                <div className="text-sm text-amber-700">⚠️ {actRow?.actividad_es} — El admin debe configurar esta actividad en <strong>Config. Act.</strong></div>
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
              </div>
            );
          })}
          {!actReps.length&&<p className="text-sm text-slate-500">Selecciona una especialidad en el Paso 1.</p>}
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
  const[aprobadas,setAprobadas]=useState<Set<string>>(new Set());
  const[rechazadas,setRechazadas]=useState<Set<string>>(new Set());
  const[motivos,setMotivos]=useState<Record<string,string>>({});
  const[saving,setSaving]=useState(false);
  const[loading,setLoading]=useState(false);
  const espList=useMemo(()=>catalogs?uniqueEsp(catalogs.especialidades_actividades):[],[catalogs]);

  async function cargar(){
    if(!fecha||!espId){showToast('err','Selecciona fecha y especialidad');return;}
    setLoading(true);
    try{
      const{data:reps}=await supabase.from('reportes_avance').select('*').eq('fecha',fecha).eq('especialidad_id',espId).eq('estado','borrador');
      const repIds=(reps||[]).map((r:Record<string,unknown>)=>r.id as string);
      const avs=repIds.length?(await supabase.from('avance_diario').select('*').in('reporte_id',repIds)).data||[]:[];
      setReportes(reps||[]); setAvances(avs as Record<string,unknown>[]);
      setAprobadas(new Set()); setRechazadas(new Set()); setMotivos({});
    } catch(e:unknown){ showToast('err',(e as Error)?.message||'Error'); }
    finally{ setLoading(false); }
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
                        {isRe&&<input className="input text-xs" placeholder="Motivo del rechazo (obligatorio)…" value={motivos[c.reporte_id]||''} onChange={e=>setMotivos(m=>({...m,[c.reporte_id]:e.target.value}))}/>}
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
                {s.comentario&&<div className="text-xs text-slate-500">Resp: {s.comentario as string}</div>}
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
  const[espId,setEspId]=useState('');
  const[data,setData]=useState<Record<string,unknown>|null>(null);
  const[loading,setLoading]=useState(false);
  const espList=useMemo(()=>catalogs?uniqueEsp(catalogs.especialidades_actividades):[],[catalogs]);

  async function load(){
    setLoading(true);
    try{
      let qR=supabase.from('reportes_avance').select('*').gte('fecha',fechaIni).lte('fecha',fechaFin);
      if(espId) qR=qR.eq('especialidad_id',espId);
      const[reps,asist,avances,incid,maqD]=await Promise.all([
        qR,
        supabase.from('asistencia_real').select('*').gte('fecha',fechaIni).lte('fecha',fechaFin),
        supabase.from('avance_diario').select('*').gte('fecha',fechaIni).lte('fecha',fechaFin),
        supabase.from('incidentes_seg').select('*').gte('fecha',fechaIni).lte('fecha',fechaFin).neq('tipo','sin_novedad'),
        supabase.from('maquinaria').select('*'),
      ]);
      const aD=(asist.data||[]) as Record<string,unknown>[];
      const avD=(avances.data||[]) as Record<string,unknown>[];
      const horasH=aD.filter(a=>a.asistio).reduce((s,a)=>s+parseFloat(String(a.horas_trabajadas||0)),0);
      const pl=aD.length,re=aD.filter(a=>a.asistio).length;

      // Avance por actividad
      const avPorAct:Record<string,number>={};
      avD.forEach(av=>{
        const id=av.actividad_id as string;
        avPorAct[id]=(avPorAct[id]||0)+parseFloat(String(av.cantidad||0));
      });

      // Horas maquinaria por novedades
      const{data:novedades}=await supabase.from('novedades_maquinaria').select('*').gte('fecha',fechaIni).lte('fecha',fechaFin);
      const horasSB=(novedades||[]).reduce((s:number,n:Record<string,unknown>)=>s+parseFloat(String(n.horas_standby||0)),0);

      setData({reportes:reps.data||[],horas_hombre:Math.round(horasH),eficiencia_personal:pl>0?Math.round(re/pl*100):100,incidentes:incid.data||[],maquinaria:maqD.data||[],avance_por_actividad:avPorAct,total_personal_dias:aD.length,horas_standby_total:horasSB.toFixed(1)});
    } catch(e:unknown){ showToast('err',(e as Error)?.message||'Error'); }
    finally{ setLoading(false); }
  }

  useEffect(()=>{load();},[]);

  const actividadesConConfig=useMemo(()=>{
    if(!catalogs||!data) return [];
    const avPorAct=data.avance_por_actividad as Record<string,number>;
    return configActs
      .filter(c=>!espId||(catalogs.especialidades_actividades.find(e=>e.id===c.actividad_id)?.especialidad_es||'').toLowerCase()===(espList.find(e=>e.id===espId)?.especialidad_es||'').toLowerCase())
      .map(c=>{
        const actRow=catalogs.especialidades_actividades.find(e=>e.id===c.actividad_id);
        const avanceHoy=avPorAct[c.actividad_id]||0;
        const acumPrevio=c.acumulado_previo||0;
        const total=avanceHoy+acumPrevio;
        const pct=c.meta_total&&c.tiene_meta?Math.min(100,Math.round(total/c.meta_total*100)):null;
        return{...c,actividad_nombre:actRow?.actividad_es||c.actividad_id,avance_periodo:avanceHoy,total_acumulado:total,pct};
      });
  },[catalogs,configActs,data,espId,espList]);

  return(
    <div className="space-y-4">
      <div className="card p-4 flex gap-3 items-end flex-wrap no-print">
        <div className="flex-1 min-w-[130px]"><label className="label">Desde</label><input type="date" className="input" value={fechaIni} onChange={e=>setFechaIni(e.target.value)}/></div>
        <div className="flex-1 min-w-[130px]"><label className="label">Hasta</label><input type="date" className="input" value={fechaFin} onChange={e=>setFechaFin(e.target.value)}/></div>
        <div className="flex-1 min-w-[160px]"><label className="label">Especialidad</label><select className="select" value={espId} onChange={e=>setEspId(e.target.value)}><option value="">Todas</option>{espList.map(e=><option key={e.id} value={e.id}>{e.especialidad_es}</option>)}</select></div>
        <button className="btn-primary" onClick={load} disabled={loading}>{loading?'Cargando…':'Actualizar'}</button>
      </div>

      {data&&(
        <>
          <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">
            <MC label="Eficiencia personal" value={`${data.eficiencia_personal as number}%`} sub="asistencia" color={(data.eficiencia_personal as number)>=90?'text-emerald-600':(data.eficiencia_personal as number)>=70?'text-amber-500':'text-rose-600'}/>
            <MC label="Horas-hombre" value={`${data.horas_hombre as number}h`} sub="productivas"/>
            <MC label="Horas stand-by" value={`${data.horas_standby_total as string}h`} sub="maquinaria"/>
            <MC label="Incidentes" value={(data.incidentes as unknown[]).length} sub="seguridad" color={(data.incidentes as unknown[]).length>0?'text-rose-600':'text-emerald-600'}/>
          </div>

          {/* AVANCE POR ACTIVIDAD */}
          {actividadesConConfig.length>0&&(
            <div className="card p-4">
              <h3 className="font-bold text-[#003b7a] mb-4">📊 Avance por actividad</h3>
              <div className="space-y-4">
                {actividadesConConfig.map((c,i)=>(
                  <div key={i} className="border border-slate-200 rounded-lg p-3">
                    <div className="flex items-start justify-between flex-wrap gap-2 mb-2">
                      <div>
                        <div className="font-semibold text-sm text-[#003b7a]">{c.actividad_nombre}</div>
                        <div className="text-xs text-slate-500">{c.unidad_es}</div>
                      </div>
                      <div className="text-right">
                        <div className="text-lg font-bold text-[#003b7a]">{c.total_acumulado}</div>
                        {c.meta_total&&c.tiene_meta&&<div className="text-xs text-slate-500">de {c.meta_total} · {c.pct}%</div>}
                      </div>
                    </div>
                    {c.meta_total&&c.tiene_meta&&c.pct!==null&&(
                      <div>
                        <div className="flex justify-between text-xs text-slate-500 mb-1">
                          <span>Acumulado previo: {c.acumulado_previo||0}</span>
                          <span>Este período: +{c.avance_periodo}</span>
                          <span>{c.pct}%</span>
                        </div>
                        <div className="w-full bg-slate-200 rounded-full h-4 overflow-hidden">
                          <div className={`h-4 rounded-full transition-all duration-500 ${c.pct>=90?'bg-emerald-500':c.pct>=50?'bg-blue-500':'bg-amber-500'}`} style={{width:`${c.pct}%`}}/>
                        </div>
                      </div>
                    )}
                    {(!c.meta_total||!c.tiene_meta)&&c.tipo!=='D'&&(
                      <div className="text-xs text-slate-500 italic">Acumulativo sin meta definida · Total: {c.total_acumulado} {c.unidad_es}</div>
                    )}
                  </div>
                ))}
              </div>
            </div>
          )}

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

// ── INFORMES ──────────────────────────────────────────────────────
function InformesModule({user,catalogs,configActs,showToast}:{
  user:Profile; catalogs:Catalogs|null; configActs:ConfigAct[];
  showToast:(k:'ok'|'err'|'info',m:string)=>void;
}){
  const[fechaIni,setFechaIni]=useState(today());
  const[fechaFin,setFechaFin]=useState(today());
  const[espIds,setEspIds]=useState<string[]>([]);
  const[areaIds,setAreaIds]=useState<string[]>([]);
  const[soloAp,setSoloAp]=useState(false);
  const[data,setData]=useState<Record<string,unknown>|null>(null);
  const[loading,setLoading]=useState(false);
  const espList=useMemo(()=>catalogs?uniqueEsp(catalogs.especialidades_actividades):[],[catalogs]);

  async function fetchData(){
    setLoading(true);
    try{
      let qR=supabase.from('reportes_avance').select('*').gte('fecha',fechaIni).lte('fecha',fechaFin);
      if(espIds.length) qR=qR.in('especialidad_id',espIds);
      if(soloAp) qR=qR.eq('estado','aprobado');
      if(user.rol==='tecnico') qR=qR.eq('usuario_id',user.id);
      if(user.rol==='cliente') qR=qR.eq('estado','aprobado');
      const{data:reps}=await qR;
      const repIds=(reps||[]).map((r:Record<string,unknown>)=>r.id as string);
      const[av,as2,sc]=await Promise.all([
        repIds.length?supabase.from('avance_diario').select('*').in('reporte_id',repIds):Promise.resolve({data:[]}),
        repIds.length?supabase.from('asistencia_real').select('*').in('reporte_id',repIds):Promise.resolve({data:[]}),
        repIds.length?supabase.from('suspensiones_clima').select('*').in('reporte_id',repIds):Promise.resolve({data:[]}),
      ]);
      let avD=(av.data||[]) as Record<string,unknown>[];
      if(areaIds.length) avD=avD.filter(a=>areaIds.includes(a.area_id as string));
      const aD=(as2.data||[]) as Record<string,unknown>[];
      const horasH=aD.filter(a=>a.asistio).reduce((s,a)=>s+parseFloat(String(a.horas_trabajadas||0)),0);
      const horasC=((sc.data||[]) as Record<string,unknown>[]).reduce((s,a)=>s+parseFloat(String(a.horas_perdidas||0)),0);

      // Calcular acumulado real incluyendo acumulado previo de config
      const avConMeta=avD.map(av=>{
        const cfg=configActs.find(c=>c.actividad_id===(av.actividad_id as string));
        const acumPrevio=cfg?.acumulado_previo||0;
        return{...av,acumulado_total_real:(av.acumulado_total as number)+acumPrevio};
      });

      setData({reportes:reps||[],avances:avConMeta,asistencia:aD,totales:{horas_hombre:Math.round(horasH),horas_perdidas_clima:Math.round(horasC),dias:repIds.length}});
    } catch(e:unknown){ showToast('err',(e as Error)?.message||'Error'); }
    finally{ setLoading(false); }
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
    const wb=XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(rows),'Avance');
    if(!isC){
      const aRows=asistencia.map(a=>({Fecha:a.fecha,Documento:a.documento_personal,Asistió:a.asistio?'Sí':'No',Motivo:a.motivo_ausencia||'',Horas:a.horas_trabajadas}));
      XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(aRows),'Asistencia');
    }
    XLSX.writeFile(wb,`PDS360_${fechaIni}_${fechaFin}.xlsx`);
    showToast('ok','Excel descargado');
  }

  function toggle(arr:string[],setArr:(v:string[])=>void,val:string){setArr(arr.includes(val)?arr.filter(x=>x!==val):[...arr,val]);}

  return(
    <div className="space-y-4">
      <div className="card p-4 space-y-3 no-print">
        <div className="grid grid-cols-1 sm:grid-cols-3 gap-3 items-end">
          <div><label className="label">Desde</label><input type="date" className="input" value={fechaIni} onChange={e=>setFechaIni(e.target.value)}/></div>
          <div><label className="label">Hasta</label><input type="date" className="input" value={fechaFin} onChange={e=>setFechaFin(e.target.value)}/></div>
          <button className="btn-primary" onClick={fetchData} disabled={loading}>{loading?'Cargando…':'Consultar'}</button>
        </div>
        <div><label className="label">Especialidades</label><div className="flex flex-wrap gap-2">{espList.map(e=><button key={e.id} onClick={()=>toggle(espIds,setEspIds,e.id)} className={`text-xs px-2 py-1 rounded border transition-colors ${espIds.includes(e.id)?'bg-[#003b7a] text-white border-[#003b7a]':'border-slate-300 text-slate-600 hover:border-[#003b7a]'}`}>{e.especialidad_es}</button>)}</div></div>
        <div><label className="label">Áreas</label><div className="flex flex-wrap gap-2">{(catalogs?.areas||[]).map(a=><button key={a.id} onClick={()=>toggle(areaIds,setAreaIds,a.id)} className={`text-xs px-2 py-1 rounded border transition-colors ${areaIds.includes(a.id)?'bg-[#003b7a] text-white border-[#003b7a]':'border-slate-300 text-slate-600 hover:border-[#003b7a]'}`}>{a.area_es}</button>)}</div></div>
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

            {(data.avances as unknown[]).length>0&&(
              <div className="card p-4">
                <h3 className="font-bold text-[#003b7a] mb-3">Avance de actividades</h3>
                <div className="overflow-x-auto">
                  <table className="table w-full">
                    <thead>
                      <tr>
                        <th>Fecha</th><th>Actividad</th><th>Área</th>
                        <th>Hoy</th><th>Acumulado</th><th>Meta</th><th>Avance</th><th>Unidad</th>
                        {user.rol!=='cliente'&&<th>Técnico</th>}
                      </tr>
                    </thead>
                    <tbody>
                      {(data.avances as Record<string,unknown>[]).map((av,i)=>{
                        const aR=catalogs?.especialidades_actividades.find(e=>e.id===(av.actividad_id as string));
                        const arR=catalogs?.areas.find(a=>a.id===(av.area_id as string));
                        const rep=(data.reportes as Record<string,unknown>[]).find(r=>r.id===av.reporte_id);
                        const cfg=configActs.find(c=>c.actividad_id===(av.actividad_id as string));
                        const pct=cfg?.meta_total&&cfg.tiene_meta?Math.round((av.acumulado_total_real as number)/cfg.meta_total*100):null;
                        return(
                          <tr key={i}>
                            <td>{av.fecha as string}</td>
                            <td>{aR?.actividad_es||av.actividad_id as string}</td>
                            <td>{arR?.area_es||av.area_id as string}</td>
                            <td className="font-semibold">{av.cantidad as number}</td>
                            <td className="font-semibold text-[#003b7a]">{av.acumulado_total_real as number}</td>
                            <td className="text-slate-500">{cfg?.meta_total||'—'}</td>
                            <td>
                              {pct!==null?(
                                <div className="flex items-center gap-1">
                                  <div className="w-16 bg-slate-200 rounded-full h-2"><div className={`h-2 rounded-full ${pct>=90?'bg-emerald-500':pct>=50?'bg-blue-500':'bg-amber-500'}`} style={{width:`${pct}%`}}/></div>
                                  <span className="text-xs font-medium">{pct}%</span>
                                </div>
                              ):<span className="text-xs text-slate-400">—</span>}
                            </td>
                            <td>{av.unidad as string}</td>
                            {user.rol!=='cliente'&&<td className="text-xs text-slate-500">{rep?.usuario_nombre as string||''}</td>}
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              </div>
            )}
            {!(data.avances as unknown[]).length&&!loading&&<div className="card p-6 text-center text-slate-500">Sin datos. Consulta primero.</div>}
          </>
        )}
        {!data&&<div className="card p-6 text-center text-slate-500">Selecciona el período y haz clic en Consultar.</div>}
      </div>
    </div>
  );
}

// ── CATÁLOGOS — con botones siempre visibles ───────────────────────
function CatalogosModule({catalogs,onRefresh,showToast}:{catalogs:Catalogs|null;onRefresh:()=>void;showToast:(k:'ok'|'err'|'info',m:string)=>void}){
  return(
    <div className="space-y-4">
      <CatMgr title="Especialidades y Actividades" table="especialidades_actividades" nameField="actividad_es"
        fields={[{n:'especialidad_es',l:'Especialidad (ES)'},{n:'especialidad_en',l:'Especialidad (EN)'},{n:'actividad_es',l:'Actividad (ES)'},{n:'actividad_en',l:'Actividad (EN)'}]}
        rows={catalogs?.especialidades_actividades||[]} onChanged={onRefresh} showToast={showToast}/>
      <CatMgr title="Áreas" table="areas" nameField="area_es"
        fields={[{n:'area_es',l:'Área (ES)'},{n:'area_en',l:'Área (EN)'}]}
        rows={catalogs?.areas||[]} onChanged={onRefresh} showToast={showToast}/>
      <CatMgr title="Líderes" table="lideres" nameField="nombre"
        fields={[{n:'nombre',l:'Nombre'},{n:'documento',l:'Documento'},{n:'cargo_es',l:'Cargo (ES)'},{n:'cargo_en',l:'Cargo (EN)'}]}
        rows={catalogs?.lideres||[]} onChanged={onRefresh} showToast={showToast}/>
      <CatMgr title="Personal" table="personal" nameField="nombre"
        fields={[{n:'nombre',l:'Nombre'},{n:'documento',l:'Documento'},{n:'cargo_es',l:'Cargo (ES)'},{n:'cargo_en',l:'Cargo (EN)'},{n:'tipo',l:'Tipo'},{n:'empresa',l:'Empresa'}]}
        rows={catalogs?.personal||[]} onChanged={onRefresh} showToast={showToast}/>
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
  const[form,setForm]=useState({tipo:'motosierra',item_id:'',nombre:'',estado:'activo'});
  const[busy,setBusy]=useState(false);

  async function addOne(){
    if(!form.item_id){showToast('err','Falta ID del equipo');return;}
    setBusy(true);
    try{const{error}=await supabase.from('maquinaria').insert({...form,horas_acum_operativas:0,horas_acum_standby:0});if(error)throw error;showToast('ok','Equipo agregado');setForm({tipo:'motosierra',item_id:'',nombre:'',estado:'activo'});onRefresh();}
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
      <div className="grid grid-cols-2 sm:grid-cols-4 gap-2">
        <div><label className="label">Tipo</label><select className="select" value={form.tipo} onChange={e=>setForm({...form,tipo:e.target.value})}><option value="motosierra">Motosierra</option><option value="chipeadora">Chipeadora</option><option value="camion">Camión</option><option value="otro">Otro</option></select></div>
        <div><label className="label">ID único</label><input className="input" value={form.item_id} onChange={e=>setForm({...form,item_id:e.target.value})} placeholder="MS-009"/></div>
        <div><label className="label">Nombre</label><input className="input" value={form.nombre} onChange={e=>setForm({...form,nombre:e.target.value})}/></div>
        <div><label className="label">Estado</label><select className="select" value={form.estado} onChange={e=>setForm({...form,estado:e.target.value})}><option value="activo">Activo</option><option value="inactivo">Inactivo</option><option value="mantenimiento">Mantenimiento</option></select></div>
      </div>
      <button className="btn-primary text-xs" disabled={busy} onClick={addOne}>+ Agregar equipo</button>
      <div className="space-y-2">
        {maquinaria.map(m=>(
          <div key={m.id} className={`flex items-center gap-2 p-2 rounded-lg border ${m.estado==='inactivo'?'bg-slate-50 border-slate-200 opacity-60':'bg-white border-slate-200'}`}>
            <div className="flex-1 min-w-0">
              <div className="font-medium text-sm">{m.item_id} — {m.nombre}</div>
              <div className="text-xs text-slate-500">{m.tipo} · Op: {(m.horas_acum_operativas||0).toFixed(1)}h · SB: {(m.horas_acum_standby||0).toFixed(1)}h</div>
            </div>
            <select className="select text-xs w-32 flex-shrink-0" value={m.estado} onChange={e=>cambiarEstado(m,e.target.value)}>
              <option value="activo">Activo</option><option value="inactivo">Inactivo</option><option value="mantenimiento">Mantenimiento</option>
            </select>
            <button className="text-xs px-2 py-1 rounded bg-rose-100 text-rose-700 hover:bg-rose-200 flex-shrink-0" onClick={()=>eliminar(m)} disabled={busy}>Eliminar</button>
          </div>
        ))}
      </div>
    </div>
  );
}

// ── CONFIG ACTIVIDADES — con nombres visibles y editar/eliminar ────
function ConfigActModule({configActs,catalogs,onRefresh,showToast}:{
  configActs:ConfigAct[]; catalogs:Catalogs|null; onRefresh:()=>void;
  showToast:(k:'ok'|'err'|'info',m:string)=>void;
}){
  const[form,setForm]=useState({especialidad_id:'',actividad_id:'',tipo:'A',unidad_es:'',unidad_en:'',meta_total:'',acumulado_previo:'',rendimiento_esperado:'',rendimiento_por:'cuadrilla',tiene_meta:true,es_medible:true});
  const[busy,setBusy]=useState(false);
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
    if(form.es_medible&&form.tipo!=='D'&&!form.unidad_es){showToast('err','Ingresa la unidad de medida');return;}
    if(form.es_medible&&form.tiene_meta&&form.tipo!=='D'&&!form.meta_total){showToast('err','Ingresa la meta total');return;}
    setBusy(true);
    try{
      const payload={
        especialidad_id:form.especialidad_id, actividad_id:form.actividad_id,
        tipo:form.tipo, unidad_es:form.unidad_es||'N/A', unidad_en:form.unidad_en||'N/A',
        tiene_meta:!!(form.tiene_meta&&form.es_medible&&form.tipo!=='D'),
        es_medible:form.tipo==='D'?false:!!form.es_medible,
        meta_total:form.tiene_meta&&form.es_medible&&form.meta_total&&form.tipo!=='D'?parseFloat(form.meta_total):null,
        acumulado_previo:parseFloat(form.acumulado_previo||'0'),
        rendimiento_esperado:form.rendimiento_esperado?parseFloat(form.rendimiento_esperado):null,
        rendimiento_por:form.rendimiento_por, activo:true,
      };
      const{error}=await supabase.from('config_actividades').upsert(payload,{onConflict:'actividad_id'});
      if(error) throw error;
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
              <button key={t.v} onClick={()=>setForm({...form,tipo:t.v})}
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

      <button className="btn-primary" disabled={busy||!form.actividad_id} onClick={save}>{busy?'Guardando…':'💾 Guardar configuración'}</button>

      {configActs.length>0&&(
        <div>
          <div className="font-semibold text-sm text-slate-700 mb-2">Actividades configuradas ({configActs.length})</div>
          <div className="space-y-2 max-h-72 overflow-y-auto">
            {configActs.map((c,i)=>{
              const actRow=allActs.find(e=>e.id===c.actividad_id);
              return(
                <div key={i} className="flex items-center gap-2 p-3 rounded-lg border border-slate-200 bg-white">
                  <div className="flex-1 min-w-0">
                    <div className="font-medium text-sm truncate text-[#003b7a]">{actRow?.actividad_es||'—'}</div>
                    <div className="text-xs text-slate-500">{actRow?.especialidad_es||'—'}</div>
                    <div className="text-xs text-slate-400 mt-0.5">
                      {c.tipo==='D'?'Cualitativa':c.tipo==='A'&&c.meta_total?`Meta: ${c.meta_total} ${c.unidad_es}`:c.tipo==='B'?`Acumulativo · ${c.unidad_es}`:`Tipo ${c.tipo} · ${c.unidad_es}`}
                    </div>
                  </div>
                  <span className={`badge flex-shrink-0 ${c.tipo==='A'?'bg-blue-100 text-blue-800':c.tipo==='B'?'bg-green-100 text-green-800':c.tipo==='C'?'bg-purple-100 text-purple-800':'bg-rose-100 text-rose-800'}`}>{c.tipo}</span>
                  <div className="flex gap-1 flex-shrink-0">
                    <button className="text-xs px-2 py-1 rounded bg-amber-100 text-amber-700 hover:bg-amber-200" onClick={()=>{const e=allActs.find(x=>x.id===c.actividad_id);if(e){const espRow=allActs.find(x=>x.especialidad_es===e.especialidad_es);setForm({...form,especialidad_id:espRow?.id||'',actividad_id:c.actividad_id});selAct(c.actividad_id);}}} >Editar</button>
                    <button className="text-xs px-2 py-1 rounded bg-rose-100 text-rose-700 hover:bg-rose-200" onClick={()=>eliminarConfig(c.actividad_id)}>Eliminar</button>
                  </div>
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
            <div key={u.id} className={`flex items-center gap-2 p-2 rounded-lg border ${!u.activo?'bg-slate-50 border-slate-200 opacity-60':'bg-white border-slate-200'}`}>
              <div className="flex-1 min-w-0">
                <div className="font-medium text-sm truncate">{u.nombre}</div>
                <div className="text-xs text-slate-500 truncate">{u.correo}</div>
              </div>
              <span className="badge bg-slate-100 text-slate-700 uppercase text-xs flex-shrink-0">{u.rol}</span>
              <select className="select text-xs w-28 flex-shrink-0" value={u.rol} onChange={e=>cambiarRol(u,e.target.value)}>
                {ROLES.map(r=><option key={r} value={r}>{r}</option>)}
              </select>
              <div className="flex gap-1 flex-shrink-0">
                <button className={`text-xs px-2 py-1 rounded font-medium transition-colors ${u.activo?'bg-amber-100 text-amber-700 hover:bg-amber-200':'bg-emerald-100 text-emerald-700 hover:bg-emerald-200'}`} onClick={()=>toggleActivo(u)}>{u.activo?'Desactivar':'Activar'}</button>
                <button className="text-xs px-2 py-1 rounded font-medium bg-rose-100 text-rose-700 hover:bg-rose-200" onClick={()=>eliminar(u)} disabled={busy}>Eliminar</button>
              </div>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

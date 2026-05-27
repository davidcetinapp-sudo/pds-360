'use client';
// app/page.tsx  ·  Powerchina PDS 360  v2.0  ·  Supabase edition
// FIX ERROR 1: Concurrencia resuelta con unique constraints en PostgreSQL
// FIX ERROR 2: Sesión nunca expira (autoRefreshToken en Supabase)
// FIX ERROR 3: Cada actividad tiene su propio líder

import { useCallback, useEffect, useMemo, useRef, useState } from 'react';
import { supabase, type Profile, type UserRole } from '@/lib/supabase';
import * as XLSX from 'xlsx';

/* ═══════════════════════════════════════════════════════════════
 *  TIPOS
 * ═══════════════════════════════════════════════════════════════ */
type AppView = 'home'|'planear'|'reporte'|'dashboard'|'informes'|'catalogos'|'maquinaria'|'config_act'|'usuarios';

interface EspAct  { id:string; especialidad_es:string; especialidad_en:string; actividad_es:string; actividad_en:string; activo?:boolean; }
interface Area     { id:string; area_es:string; area_en:string; }
interface Lider    { id:string; nombre:string; documento:string; cargo_es:string; especialidad_id?:string; }
interface Personal { id:string; nombre:string; documento:string; cargo_es:string; tipo?:string; empresa?:string; }
interface Maq      { id:string; tipo:string; item_id:string; nombre:string; estado:string; horas_acum_operativas?:number; horas_acum_standby?:number; }
interface ConfigAct{ id:string; actividad_id:string; tipo:'A'|'B'|'C'; unidad_es:string; meta_total?:number; rendimiento_esperado?:number; rendimiento_por?:string; acumulado_previo?:number; }
interface Catalogs { especialidades_actividades:EspAct[]; areas:Area[]; lideres:Lider[]; personal:Personal[]; }

interface ActForm {
  uid:string; especialidad_id:string; actividad_id:string;
  area_id:string; areas_adicionales:string[]; lider_id:string;
  maquinaria_ids:string[]; rendimiento_esperado:string;
  observacion_es:string; observacion_en:string;
  personal:{personal_id:string;documento_personal:string}[];
}
interface SuspensionItem { uid:string; hora_inicio:string; hora_fin:string; descripcion:string; }
interface MaqDiaItem { maquinaria_id:string; nombre:string; novedad:boolean; descripcion:string; hora_inicio:string; hora_fin:string; horas_standby:number; }
interface AsistItem { personal_id:string; documento_personal:string; nombre:string; cargo_es:string; asistio:boolean; motivo_ausencia:string; }
interface AvanceForm { uid:string; actividad_id:string; area_id:string; cantidad:string; unidad:string; observacion_es:string; observacion_en:string; }

/* ═══════════════════════════════════════════════════════════════
 *  HELPERS
 * ═══════════════════════════════════════════════════════════════ */
const JORNADA = 9;
function today() { const d=new Date(); return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`; }
function uid()   { return Math.random().toString(36).slice(2)+Date.now().toString(36); }

function uniqueEsp(rows:EspAct[]) {
  const seen=new Set<string>(); return rows.filter(r=>{ const k=(r.especialidad_es||'').toLowerCase(); if(seen.has(k))return false; seen.add(k); return true; });
}
function actsForEsp(rows:EspAct[], espId:string) {
  const esp=rows.find(r=>r.id===espId); if(!esp) return [];
  const t=(esp.especialidad_es||'').toLowerCase();
  return rows.filter(r=>(r.especialidad_es||'').toLowerCase()===t);
}
function semaforo(pct:number) { return pct>=70?'🟢':pct>=50?'🟡':'🔴'; }
function horasLostClima(susps:SuspensionItem[]) {
  return susps.reduce((acc,s)=>{
    if(!s.hora_inicio||!s.hora_fin) return acc;
    const [ih,im]=s.hora_inicio.split(':').map(Number);
    const [fh,fm]=s.hora_fin.split(':').map(Number);
    return acc+Math.max(0,(fh+fm/60)-(ih+im/60));
  },0);
}

/* ═══════════════════════════════════════════════════════════════
 *  ROOT
 * ═══════════════════════════════════════════════════════════════ */
export default function App() {
  const [profile, setProfile] = useState<Profile|null>(null);
  const [view,    setView]    = useState<AppView>('home');
  const [catalogs,setCatalogs]= useState<Catalogs|null>(null);
  const [maquinaria,setMaq]   = useState<Maq[]>([]);
  const [configActs,setCA]    = useState<ConfigAct[]>([]);
  const [toast,   setToast]   = useState<{k:'ok'|'err'|'info';m:string}|null>(null);
  const [loading, setLoading] = useState(true);

  useEffect(()=>{
    supabase.auth.getSession().then(async({data:{session}})=>{
      if(session?.user){
        const{data:p}=await supabase.from('profiles').select('*').eq('id',session.user.id).single();
        if(p) setProfile(p as Profile);
      }
      setLoading(false);
    });
    const{data:{subscription}}=supabase.auth.onAuthStateChange((_,session)=>{
      if(!session) { setProfile(null); setView('home'); }
    });
    return()=>subscription.unsubscribe();
  },[]);

  useEffect(()=>{ if(toast){ const t=setTimeout(()=>setToast(null),4500); return()=>clearTimeout(t); } },[toast]);

  const showToast=useCallback((k:'ok'|'err'|'info',m:string)=>setToast({k,m}),[]);

  const loadCatalogs=useCallback(async()=>{
    const[ea,ar,li,pe,ma,ca]=await Promise.all([
      supabase.from('especialidades_actividades').select('*').eq('activo',true).order('especialidad_es'),
      supabase.from('areas').select('*').eq('activo',true).order('area_es'),
      supabase.from('lideres').select('*').eq('activo',true).order('nombre'),
      supabase.from('personal').select('*').eq('activo',true).order('nombre'),
      supabase.from('maquinaria').select('*').order('item_id'),
      supabase.from('config_actividades').select('*').eq('activo',true),
    ]);
    setCatalogs({ especialidades_actividades:(ea.data||[]) as EspAct[], areas:(ar.data||[]) as Area[], lideres:(li.data||[]) as Lider[], personal:(pe.data||[]) as Personal[] });
    setMaq((ma.data||[]) as Maq[]);
    setCA((ca.data||[]) as ConfigAct[]);
  },[]);

  useEffect(()=>{ if(profile) loadCatalogs(); },[profile,loadCatalogs]);

  async function handleLogin(p:Profile){ setProfile(p); }
  async function handleLogout(){
    await supabase.auth.signOut();
    setProfile(null); setCatalogs(null); setView('home');
  }

  if(loading) return <div className="min-h-screen flex items-center justify-center text-slate-500">Cargando…</div>;
  if(!profile) return <LoginScreen onLogin={handleLogin} showToast={showToast} toast={toast}/>;

  const u=profile;
  const canEdit=u.rol==='admin'||u.rol==='lider'||u.rol==='tecnico';

  return (
    <div className="min-h-screen flex flex-col">
      <Header user={u} onLogout={handleLogout} setView={setView} currentView={view}/>
      <main className="flex-1 p-3 sm:p-5 max-w-7xl mx-auto w-full">
        {view==='home'        && <HomeScreen user={u} setView={setView}/>}
        {view==='planear'     && canEdit && <PlaneacionModule user={u} catalogs={catalogs} maquinaria={maquinaria} showToast={showToast}/>}
        {view==='reporte'     && canEdit && <ReporteModule user={u} catalogs={catalogs} maquinaria={maquinaria} configActs={configActs} showToast={showToast}/>}
        {view==='dashboard'   && <DashboardModule user={u} showToast={showToast}/>}
        {view==='informes'    && <InformesModule user={u} catalogs={catalogs} showToast={showToast}/>}
        {view==='catalogos'   && u.rol==='admin' && <CatalogosModule catalogs={catalogs} onRefresh={loadCatalogs} showToast={showToast}/>}
        {view==='maquinaria'  && u.rol==='admin' && <MaquinariaModule maquinaria={maquinaria} onRefresh={loadCatalogs} showToast={showToast}/>}
        {view==='config_act'  && u.rol==='admin' && <ConfigActModule configActs={configActs} catalogs={catalogs} onRefresh={loadCatalogs} showToast={showToast}/>}
        {view==='usuarios'    && u.rol==='admin' && <UsuariosModule showToast={showToast}/>}
      </main>
      <Toast toast={toast}/>
      <footer className="bg-slate-100 border-t text-center text-xs text-slate-500 py-2 no-print">
        Powerchina · PDS 360 · {new Date().getFullYear()}
      </footer>
    </div>
  );
}

/* ─── HEADER ─────────────────────────────────────────────────── */
function Header({user,onLogout,setView,currentView}:{user:Profile;onLogout:()=>void;setView:(v:AppView)=>void;currentView:AppView}){
  const canEdit=user.rol==='admin'||user.rol==='lider'||user.rol==='tecnico';
  const tabs:{key:AppView;label:string;show:boolean}[]=[
    {key:'home',label:'Inicio',show:true},
    {key:'planear',label:'Planear',show:canEdit},
    {key:'reporte',label:'Reporte diario',show:canEdit},
    {key:'dashboard',label:'Dashboard',show:true},
    {key:'informes',label:'Informes',show:true},
    {key:'catalogos',label:'Catálogos',show:user.rol==='admin'},
    {key:'maquinaria',label:'Maquinaria',show:user.rol==='admin'},
    {key:'config_act',label:'Config. Act.',show:user.rol==='admin'},
    {key:'usuarios',label:'Usuarios',show:user.rol==='admin'},
  ];
  return(
    <header className="bg-[#003b7a] text-white shadow no-print">
      <div className="max-w-7xl mx-auto px-4 py-3 flex items-center justify-between gap-2 flex-wrap">
        <div className="flex items-center gap-3">
          <img src="/icons/icon-192.png" alt="PC" className="h-8 w-auto bg-white rounded p-0.5"/>
          <div><div className="font-bold text-base leading-tight">Powerchina · PDS 360</div>
            <div className="text-xs text-blue-200">Gestión integral de obra</div></div>
        </div>
        <div className="flex items-center gap-3 text-sm">
          <div className="hidden sm:block text-right">
            <div className="font-medium">{user.nombre}</div>
            <div className="text-xs text-blue-200 uppercase">{user.rol}</div>
          </div>
          <button className="btn-secondary text-xs" onClick={onLogout}>Salir</button>
        </div>
      </div>
      <nav className="max-w-7xl mx-auto px-2 overflow-x-auto">
        <div className="flex gap-0.5 pb-0">
          {tabs.filter(t=>t.show).map(t=>(
            <button key={t.key} onClick={()=>setView(t.key)}
              className={`px-3 py-2 text-xs font-medium border-b-2 whitespace-nowrap transition-colors ${currentView===t.key?'border-white text-white':'border-transparent text-blue-200 hover:text-white'}`}>
              {t.label}
            </button>
          ))}
        </div>
      </nav>
    </header>
  );
}

/* ─── TOAST ──────────────────────────────────────────────────── */
function Toast({toast}:{toast:{k:string;m:string}|null}){
  if(!toast) return null;
  const c=toast.k==='ok'?'bg-emerald-600':toast.k==='err'?'bg-rose-600':'bg-slate-700';
  return <div className={`fixed bottom-4 left-4 right-4 sm:left-auto sm:right-4 sm:max-w-sm ${c} text-white px-4 py-3 rounded-lg shadow-lg z-50 text-sm no-print`}>{toast.m}</div>;
}

/* ─── LOGIN ──────────────────────────────────────────────────── */
function LoginScreen({onLogin,showToast,toast}:{onLogin:(p:Profile)=>void;showToast:any;toast:any}){
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
      const{data:p,error:pe}=await supabase.from('profiles').select('*').eq('id',data.user.id).single();
      if(pe||!p) throw new Error('Perfil no encontrado. Contacta al administrador.');
      if(!(p as any).activo) throw new Error('Usuario inactivo. Contacta al administrador.');
      onLogin(p as Profile);
    }catch(e:any){showToast('err',e?.message||'Error de red');}
    finally{setLoading(false);}
  }

  return(
    <div className="min-h-screen flex items-center justify-center p-4 bg-gradient-to-br from-[#003b7a] to-[#002752]">
      <div className="card p-6 sm:p-8 w-full max-w-md">
        <div className="flex flex-col items-center mb-6">
          <img src="/icons/icon-192.png" alt="PC" className="h-20 w-auto mb-3"/>
          <h1 className="text-2xl font-bold text-[#003b7a]">Powerchina · PDS 360</h1>
          <p className="text-sm text-slate-500">Gestión integral de obra</p>
        </div>
        <form onSubmit={submit} className="space-y-4">
          <div><label className="label">Correo</label>
            <input className="input" type="email" autoComplete="username" value={correo} onChange={e=>setCorreo(e.target.value)} placeholder="usuario@correo.com"/></div>
          <div><label className="label">Clave</label>
            <input className="input" type="password" autoComplete="current-password" value={clave} onChange={e=>setClave(e.target.value)}/></div>
          <button className="btn-primary w-full py-3" disabled={loading}>{loading?'Ingresando…':'Ingresar'}</button>
        </form>
      </div>
      <Toast toast={toast}/>
    </div>
  );
}

/* ─── HOME ───────────────────────────────────────────────────── */
function HomeScreen({user,setView}:{user:Profile;setView:(v:AppView)=>void}){
  const canEdit=user.rol==='admin'||user.rol==='lider'||user.rol==='tecnico';
  return(
    <div className="space-y-6">
      <div className="text-center py-4">
        <h2 className="text-2xl font-bold text-[#003b7a]">Bienvenido, {user.nombre.split(' ')[0]}</h2>
        <p className="text-slate-500 text-sm mt-1">{today()} · <span className="font-semibold uppercase">{user.rol}</span></p>
      </div>
      {canEdit&&(
        <div className="grid grid-cols-1 sm:grid-cols-2 gap-4 max-w-2xl mx-auto">
          <button onClick={()=>setView('planear')} className="card p-6 text-left hover:shadow-md transition-shadow border-2 border-transparent hover:border-[#003b7a] group">
            <div className="text-3xl mb-3">📋</div>
            <div className="font-bold text-[#003b7a] text-lg group-hover:underline">Planear actividades</div>
            <p className="text-sm text-slate-500 mt-1">Programa actividades con personal y maquinaria. Cada actividad tiene su propio líder.</p>
          </button>
          <button onClick={()=>setView('reporte')} className="card p-6 text-left hover:shadow-md transition-shadow border-2 border-transparent hover:border-emerald-600 group">
            <div className="text-3xl mb-3">📊</div>
            <div className="font-bold text-emerald-700 text-lg group-hover:underline">Reporte de avance</div>
            <p className="text-sm text-slate-500 mt-1">Registra el avance real del día: asistencia, maquinaria y progreso por actividad.</p>
          </button>
        </div>
      )}
      <div className="grid grid-cols-2 sm:grid-cols-4 gap-3 max-w-2xl mx-auto">
        <QuickBtn icon="📈" label="Dashboard"  onClick={()=>setView('dashboard')}/>
        <QuickBtn icon="📁" label="Informes"   onClick={()=>setView('informes')}/>
        {user.rol==='admin'&&<QuickBtn icon="⚙️" label="Catálogos" onClick={()=>setView('catalogos')}/>}
        {user.rol==='admin'&&<QuickBtn icon="👤" label="Usuarios"  onClick={()=>setView('usuarios')}/>}
      </div>
    </div>
  );
}
function QuickBtn({icon,label,onClick}:{icon:string;label:string;onClick:()=>void}){
  return <button onClick={onClick} className="card p-3 text-center hover:shadow-md transition-shadow"><div className="text-2xl">{icon}</div><div className="text-xs font-medium text-slate-700 mt-1">{label}</div></button>;
}

/* ═══════════════════════════════════════════════════════════════
 *  PLANEACIÓN
 * ═══════════════════════════════════════════════════════════════ */
function PlaneacionModule({user,catalogs,maquinaria,showToast}:{user:Profile;catalogs:Catalogs|null;maquinaria:Maq[];showToast:any}){
  const[fecha,setFecha]=useState(today());
  const[actividades,setActividades]=useState<ActForm[]>([emptyAct()]);
  const[blocked,setBlocked]=useState<{documento_personal:string;usuario_nombre:string}[]>([]);
  const[estado,setEstado]=useState<'nuevo'|'borrador'|'enviado'>('nuevo');
  const[saving,setSaving]=useState(false);
  const[progId,setProgId]=useState<string|null>(null);
  const[contexto,setContexto]=useState<string>('');

  const loadFecha=useCallback(async()=>{
    if(!fecha) return;
    // Cargar personal ocupado
    const{data:bl}=await supabase.from('personal_asignado').select('documento_personal, usuario_id, programaciones!inner(usuario_nombre)').eq('fecha',fecha).neq('usuario_id',user.id);
    setBlocked((bl||[]).map((b:any)=>({ documento_personal:b.documento_personal, usuario_nombre:(b.programaciones as any)?.usuario_nombre||'otro' })));

    // Cargar planeación existente
    const{data:prog}=await supabase.from('programaciones').select('*').eq('fecha',fecha).eq('usuario_id',user.id).maybeSingle();
    if(prog){
      setEstado(prog.estado);
      setProgId(prog.id);
      const{data:acts}=await supabase.from('actividades_programadas').select('*').eq('programacion_id',prog.id);
      if(acts?.length){
        const{data:persAll}=await supabase.from('personal_asignado').select('*').eq('programacion_id',prog.id);
        setActividades((acts||[]).map((a:any)=>({
          uid:uid(), especialidad_id:a.especialidad_id||'', actividad_id:a.actividad_id||'',
          area_id:a.area_id||'', areas_adicionales:a.areas_adicionales||[], lider_id:a.lider_id||'',
          maquinaria_ids:a.maquinaria_ids||[], rendimiento_esperado:a.rendimiento_esperado||'',
          observacion_es:a.observacion_es||'', observacion_en:a.observacion_en||'',
          personal:(persAll||[]).filter((p:any)=>p.actividad_programada_id===a.id).map((p:any)=>({ personal_id:p.personal_id, documento_personal:p.documento_personal })),
        })));
        return;
      }
    }
    setEstado('nuevo'); setProgId(null); setActividades([emptyAct()]);

    // Contexto del día anterior
    const ayer=new Date(fecha); ayer.setDate(ayer.getDate()-1);
    const ayerStr=`${ayer.getFullYear()}-${String(ayer.getMonth()+1).padStart(2,'0')}-${String(ayer.getDate()).padStart(2,'0')}`;
    const{data:avAyer}=await supabase.from('avance_diario').select('cantidad, actividad_id, area_id').eq('fecha',ayerStr).eq('usuario_id',user.id);
    if(avAyer?.length){
      const resumen=avAyer.map((a:any)=>`Actividad ${a.actividad_id?.slice(-6)}: ${a.cantidad} en área ${a.area_id?.slice(-6)}`).join(' | ');
      setContexto(`Ayer: ${resumen}`);
    } else setContexto('');
  },[fecha,user.id]);

  useEffect(()=>{ loadFecha(); },[loadFecha]);

  const blockedDocs=useMemo(()=>new Set(blocked.map(b=>b.documento_personal)),[blocked]);
  const isReadOnly=estado==='enviado';

  async function save(est:'borrador'|'enviado'){
    if(!fecha){showToast('err','Falta fecha');return;}
    for(let i=0;i<actividades.length;i++){
      const a=actividades[i];
      if(!a.especialidad_id||!a.actividad_id||!a.area_id||!a.lider_id){showToast('err',`Actividad ${i+1}: completa todos los campos obligatorios`);return;}
      if(!a.personal.length){showToast('err',`Actividad ${i+1}: agrega al menos una persona`);return;}
    }
    if(est==='enviado'&&!window.confirm('¿Enviar planeación? Quedará bloqueada.')) return;
    setSaving(true);
    try{
      // Upsert programación
      const{data:prog,error:pe}=await supabase.from('programaciones').upsert({ id:progId||undefined, fecha, usuario_id:user.id, usuario_nombre:user.nombre, estado:est, updated_at:new Date().toISOString() },{ onConflict:'fecha,usuario_id' }).select().single();
      if(pe||!prog) throw new Error(pe?.message||'Error guardando programación');

      // Borrar actividades y personal anteriores
      await supabase.from('actividades_programadas').delete().eq('programacion_id',prog.id);

      // Insertar actividades con sus líderes individuales
      for(const act of actividades){
        const{data:actR,error:ae}=await supabase.from('actividades_programadas').insert({
          programacion_id:prog.id, fecha, usuario_id:user.id,
          especialidad_id:act.especialidad_id, actividad_id:act.actividad_id,
          area_id:act.area_id, areas_adicionales:act.areas_adicionales,
          lider_id:act.lider_id,  // ← líder propio de esta actividad (FIX ERROR 3)
          maquinaria_ids:act.maquinaria_ids,
          rendimiento_esperado:act.rendimiento_esperado,
          observacion_es:act.observacion_es, observacion_en:act.observacion_en,
        }).select().single();
        if(ae||!actR) throw new Error(ae?.message||'Error en actividad');

        if(act.personal.length){
          const{error:paE}=await supabase.from('personal_asignado').insert(
            act.personal.map(p=>({ programacion_id:prog.id, actividad_programada_id:actR.id, fecha, usuario_id:user.id, personal_id:p.personal_id, documento_personal:p.documento_personal }))
          );
          // FIX ERROR 1: error 23505 = personal duplicado (unique constraint)
          if(paE?.code==='23505') throw new Error('Personal duplicado: alguien ya asignó esa persona hoy');
          if(paE) throw new Error(paE.message);
        }
      }
      setProgId(prog.id); setEstado(est);
      showToast('ok',est==='enviado'?'Planeación enviada ✓':'Borrador guardado ✓');
      await loadFecha();
    }catch(e:any){ showToast('err',e?.message||'Error guardando'); }
    finally{ setSaving(false); }
  }

  return(
    <div className="space-y-4">
      {contexto&&<div className="card p-3 border-l-4 border-amber-400 bg-amber-50 text-sm text-amber-800">📋 {contexto}</div>}
      <div className="card p-4 flex flex-col sm:flex-row gap-3 items-end justify-between flex-wrap">
        <div className="flex-1 min-w-[160px]">
          <label className="label">Fecha</label>
          <input type="date" className="input" value={fecha} onChange={e=>setFecha(e.target.value)}/>
        </div>
        <div className="flex items-center gap-3 flex-wrap">
          {estado==='nuevo'&&<span className="badge bg-slate-200 text-slate-700">Nuevo</span>}
          {estado==='borrador'&&<span className="badge-borrador">Borrador</span>}
          {estado==='enviado'&&<span className="badge-enviado">Enviado</span>}
        </div>
      </div>

      {actividades.map((a,idx)=>(
        <ActCard key={a.uid} index={idx} act={a} catalogs={catalogs} maquinaria={maquinaria}
          blockedDocs={blockedDocs} blockedInfo={blocked} readOnly={isReadOnly}
          onChange={p=>setActividades(arr=>arr.map(x=>x.uid===a.uid?{...x,...p}:x))}
          onRemove={()=>setActividades(arr=>arr.length<=1?arr:arr.filter(x=>x.uid!==a.uid))}/>
      ))}

      <div className="flex flex-wrap gap-2">
        {!isReadOnly&&<button className="btn-secondary" onClick={()=>setActividades(a=>[...a,emptyAct()])}>+ Actividad</button>}
        {!isReadOnly&&<>
          <button className="btn-primary" disabled={saving} onClick={()=>save('borrador')}>{saving?'Guardando…':'Guardar borrador'}</button>
          <button className="btn-success" disabled={saving} onClick={()=>save('enviado')}>{saving?'Enviando…':'Enviar planeación'}</button>
        </>}
        {isReadOnly&&<button className="btn-secondary" disabled={saving} onClick={async()=>{
          if(!window.confirm('¿Reabrir como borrador?'))return;
          setSaving(true);
          await supabase.from('programaciones').update({estado:'borrador',updated_at:new Date().toISOString()}).eq('id',progId!);
          setEstado('borrador'); setSaving(false); showToast('ok','Reabierta');
        }}>Reabrir</button>}
      </div>
    </div>
  );
}

function emptyAct():ActForm{ return{uid:uid(),especialidad_id:'',actividad_id:'',area_id:'',areas_adicionales:[],lider_id:'',maquinaria_ids:[],rendimiento_esperado:'',observacion_es:'',observacion_en:'',personal:[]}; }

function ActCard({index,act,catalogs,maquinaria,blockedDocs,blockedInfo,readOnly,onChange,onRemove}:{
  index:number;act:ActForm;catalogs:Catalogs|null;maquinaria:Maq[];
  blockedDocs:Set<string>;blockedInfo:{documento_personal:string;usuario_nombre:string}[];
  readOnly:boolean;onChange:(p:Partial<ActForm>)=>void;onRemove:()=>void;
}){
  const[search,setSearch]=useState('');
  const especialidades=useMemo(()=>catalogs?uniqueEsp(catalogs.especialidades_actividades):[],[catalogs]);
  const activities=useMemo(()=>catalogs&&act.especialidad_id?actsForEsp(catalogs.especialidades_actividades,act.especialidad_id):[],[catalogs,act.especialidad_id]);
  const personalFiltrado=useMemo(()=>{
    const list=catalogs?.personal||[];
    const q=search.trim().toLowerCase();
    return q?list.filter(p=>p.nombre.toLowerCase().includes(q)||p.documento.toLowerCase().includes(q)||p.cargo_es.toLowerCase().includes(q)):list;
  },[catalogs,search]);

  return(
    <div className="card p-4 space-y-3">
      <div className="flex items-center justify-between">
        <div className="font-semibold text-[#003b7a]">Actividad {index+1}</div>
        {!readOnly&&<button className="btn-ghost text-rose-600 text-xs" onClick={onRemove}>✕ Eliminar</button>}
      </div>
      <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-3">
        <div><label className="label">Especialidad</label>
          <select className="select" disabled={readOnly} value={act.especialidad_id} onChange={e=>onChange({especialidad_id:e.target.value,actividad_id:''})}>
            <option value="">— Seleccionar —</option>
            {especialidades.map(e=><option key={e.id} value={e.id}>{e.especialidad_es}</option>)}
          </select></div>
        <div><label className="label">Actividad</label>
          <select className="select" disabled={readOnly||!act.especialidad_id} value={act.actividad_id} onChange={e=>onChange({actividad_id:e.target.value})}>
            <option value="">— Seleccionar —</option>
            {activities.map(a=><option key={a.id} value={a.id}>{a.actividad_es}</option>)}
          </select></div>
        <div><label className="label">Área principal</label>
          <select className="select" disabled={readOnly} value={act.area_id} onChange={e=>onChange({area_id:e.target.value})}>
            <option value="">— Seleccionar —</option>
            {(catalogs?.areas||[]).map(a=><option key={a.id} value={a.id}>{a.area_es}</option>)}
          </select></div>
        {/* FIX ERROR 3: líder es por actividad, no por especialidad */}
        <div><label className="label">Líder de esta actividad</label>
          <select className="select" disabled={readOnly} value={act.lider_id} onChange={e=>onChange({lider_id:e.target.value})}>
            <option value="">— Seleccionar —</option>
            {(catalogs?.lideres||[]).map(l=><option key={l.id} value={l.id}>{l.nombre} — {l.cargo_es}</option>)}
          </select></div>
      </div>

      <div>
        <label className="label">Áreas adicionales (si la cuadrilla se mueve durante el día)</label>
        <div className="flex flex-wrap gap-2">
          {(catalogs?.areas||[]).filter(a=>a.id!==act.area_id).map(a=>{
            const sel=act.areas_adicionales.includes(a.id);
            return <button key={a.id} disabled={readOnly} onClick={()=>onChange({areas_adicionales:sel?act.areas_adicionales.filter(x=>x!==a.id):[...act.areas_adicionales,a.id]})}
              className={`text-xs px-2 py-1 rounded border transition-colors ${sel?'bg-[#003b7a] text-white border-[#003b7a]':'border-slate-300 text-slate-600 hover:border-[#003b7a]'}`}>{a.area_es}</button>;
          })}
        </div>
      </div>

      <div>
        <label className="label">Maquinaria asignada</label>
        <div className="flex flex-wrap gap-2">
          {maquinaria.filter(m=>m.estado==='activo').map(m=>{
            const sel=act.maquinaria_ids.includes(m.id);
            return <button key={m.id} disabled={readOnly} onClick={()=>onChange({maquinaria_ids:sel?act.maquinaria_ids.filter(x=>x!==m.id):[...act.maquinaria_ids,m.id]})}
              className={`text-xs px-2 py-1 rounded border transition-colors ${sel?'bg-orange-500 text-white border-orange-500':'border-slate-300 text-slate-600 hover:border-orange-400'}`}>{m.item_id}</button>;
          })}
        </div>
      </div>

      <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
        <div><label className="label">Rendimiento esperado</label>
          <input className="input" disabled={readOnly} value={act.rendimiento_esperado} onChange={e=>onChange({rendimiento_esperado:e.target.value})} placeholder="ej: 10 árboles/cuadrilla"/></div>
        <div><label className="label">Observación (ES)</label>
          <input className="input" disabled={readOnly} value={act.observacion_es} onChange={e=>onChange({observacion_es:e.target.value})}/></div>
      </div>

      <div>
        <div className="flex items-center justify-between mb-2">
          <label className="label !mb-0">Personal ({act.personal.length} sel.)</label>
          {!readOnly&&act.personal.length>0&&<button className="text-xs text-rose-500 underline" onClick={()=>onChange({personal:[]})}>Limpiar</button>}
        </div>
        <input className="input mb-2" placeholder="🔎 Buscar nombre, documento o cargo…" value={search} onChange={e=>setSearch(e.target.value)}/>
        <div className="border border-slate-200 rounded-md max-h-52 overflow-y-auto">
          {personalFiltrado.map(p=>{
            const sel=!!act.personal.find(x=>x.documento_personal===p.documento);
            const bl=blockedDocs.has(p.documento);
            const blInfo=blockedInfo.find(b=>b.documento_personal===p.documento);
            return(
              <label key={p.id} className={`flex items-center gap-2 px-3 py-2 border-b border-slate-100 cursor-pointer text-sm ${bl?'bg-rose-50 opacity-70 cursor-not-allowed':sel?'bg-blue-50':'hover:bg-slate-50'}`}>
                <input type="checkbox" checked={sel} disabled={readOnly||bl} onChange={()=>{
                  if(bl) return;
                  onChange({personal:sel?act.personal.filter(x=>x.documento_personal!==p.documento):[...act.personal,{personal_id:p.id,documento_personal:p.documento}]});
                }}/>
                <div className="flex-1 min-w-0">
                  <div className="font-medium truncate">{p.nombre}</div>
                  <div className="text-xs text-slate-500">{p.documento} · {p.cargo_es}</div>
                </div>
                {bl&&<span className="badge-ocupado text-xs">Ocupado por {blInfo?.usuario_nombre}</span>}
                {sel&&!bl&&<span className="badge bg-blue-100 text-blue-800">✓</span>}
              </label>
            );
          })}
        </div>
      </div>
    </div>
  );
}

/* ═══════════════════════════════════════════════════════════════
 *  REPORTE DIARIO — 8 PASOS
 * ═══════════════════════════════════════════════════════════════ */
function ReporteModule({user,catalogs,maquinaria,configActs,showToast}:{user:Profile;catalogs:Catalogs|null;maquinaria:Maq[];configActs:ConfigAct[];showToast:any}){
  const[step,setStep]=useState(1);
  const[fecha,setFecha]=useState(today());
  const[espId,setEspId]=useState(user.especialidad_id||'');
  const[jornadaHrs,setJornadaHrs]=useState(JORNADA);
  const[clima,setClima]=useState('despejado');
  const[suspensiones,setSuspensiones]=useState<SuspensionItem[]>([]);
  const[charla,setCharla]=useState(true);
  const[charlaTema,setCharlaTema]=useState('');
  const[asistencia,setAsistencia]=useState<AsistItem[]>([]);
  const[maqDia,setMaqDia]=useState<MaqDiaItem[]>([]);
  const[avances,setAvances]=useState<AvanceForm[]>([]);
  const[incidente,setIncidente]=useState({tipo:'sin_novedad',descripcion:'',medidas:'',area_id:''});
  const[notaBit,setNotaBit]=useState('');
  const[saving,setSaving]=useState(false);

  const horasClima=useMemo(()=>horasLostClima(suspensiones),[suspensiones]);
  const horasReal=Math.max(0,jornadaHrs-horasClima);
  const STEPS=['Encabezado','Condiciones','Charla','Asistencia','Maquinaria','Avance','Seguridad','Enviar'];

  useEffect(()=>{
    if(!fecha||!espId) return;
    supabase.from('personal_asignado').select('personal_id, documento_personal, personal!inner(nombre,cargo_es)').eq('fecha',fecha)
      .then(({data})=>{
        const seen=new Set<string>();
        const list=(data||[]).filter((r:any)=>{
          if(seen.has(r.documento_personal)) return false;
          seen.add(r.documento_personal); return true;
        });
        setAsistencia(list.map((r:any)=>({ personal_id:r.personal_id, documento_personal:r.documento_personal, nombre:(r.personal as any)?.nombre||r.documento_personal, cargo_es:(r.personal as any)?.cargo_es||'', asistio:true, motivo_ausencia:'' })));
      });
  },[fecha,espId]);

  useEffect(()=>{
    const maqFilt=maquinaria.filter(m=>m.estado==='activo');
    setMaqDia(maqFilt.map(m=>({maquinaria_id:m.id,nombre:`${m.item_id} – ${m.tipo}`,novedad:false,descripcion:'',hora_inicio:'',hora_fin:'',horas_standby:0})));
  },[maquinaria]);

  useEffect(()=>{
    if(!catalogs||!espId) return;
    const espRow=catalogs.especialidades_actividades.find(e=>e.id===espId);
    if(!espRow) return;
    const t=(espRow.especialidad_es||'').toLowerCase();
    const acts=catalogs.especialidades_actividades.filter(e=>(e.especialidad_es||'').toLowerCase()===t);
    setAvances(acts.map(a=>({uid:uid(),actividad_id:a.id,area_id:'',cantidad:'',unidad:configActs.find(c=>c.actividad_id===a.id)?.unidad_es||'',observacion_es:'',observacion_en:''})));
  },[catalogs,espId,configActs]);

  async function submit(){
    if(!fecha||!espId){showToast('err','Falta fecha o especialidad');return;}
    setSaving(true);
    try{
      const{data:rep,error:re}=await supabase.from('reportes_avance').insert({
        fecha, usuario_id:user.id, usuario_nombre:user.nombre, especialidad_id:espId,
        jornada_horas:jornadaHrs, clima, charla_preturno:charla, charla_tema:charlaTema, estado:'borrador',
      }).select().single();
      if(re||!rep) throw new Error(re?.message||'Error creando reporte');

      const rid=rep.id;
      await Promise.all([
        suspensiones.length?supabase.from('suspensiones_clima').insert(suspensiones.map(s=>({reporte_id:rid,fecha,usuario_id:user.id,hora_inicio:s.hora_inicio||null,hora_fin:s.hora_fin||null,horas_perdidas:horasLostClima([s]),descripcion:s.descripcion}))):null,
        asistencia.length?supabase.from('asistencia_real').insert(asistencia.map(a=>({reporte_id:rid,fecha,usuario_id:user.id,personal_id:a.personal_id,documento_personal:a.documento_personal,asistio:a.asistio,motivo_ausencia:a.motivo_ausencia||null,horas_trabajadas:a.asistio?horasReal:0}))):null,
        maqDia.filter(m=>m.novedad).length?supabase.from('novedades_maquinaria').insert(maqDia.filter(m=>m.novedad).map(m=>({reporte_id:rid,fecha,usuario_id:user.id,maquinaria_id:m.maquinaria_id,descripcion:m.descripcion,hora_inicio:m.hora_inicio||null,hora_fin:m.hora_fin||null,horas_standby:m.horas_standby}))):null,
        incidente.tipo!=='sin_novedad'?supabase.from('incidentes_seg').insert({reporte_id:rid,fecha,usuario_id:user.id,tipo:incidente.tipo,descripcion:incidente.descripcion,medidas_tomadas:incidente.medidas,area_id:incidente.area_id||null}):null,
        notaBit?supabase.from('bitacora_decisiones').insert({fecha,usuario_id:user.id,descripcion:notaBit,especialidad_id:espId}):null,
      ]);

      // Avances con acumulado
      for(const av of avances.filter(a=>parseFloat(a.cantidad)>0)){
        const{data:prev}=await supabase.from('avance_diario').select('cantidad').eq('actividad_id',av.actividad_id).eq('area_id',av.area_id).eq('usuario_id',user.id);
        const acumPrev=(prev||[]).reduce((s:number,r:any)=>s+parseFloat(r.cantidad||0),0);
        const cantidad=parseFloat(av.cantidad);
        await supabase.from('avance_diario').insert({reporte_id:rid,fecha,usuario_id:user.id,actividad_id:av.actividad_id,especialidad_id:espId,area_id:av.area_id,cantidad,unidad:av.unidad,acumulado_anterior:acumPrev,acumulado_total:acumPrev+cantidad,observacion_es:av.observacion_es,observacion_en:av.observacion_en});
      }

      showToast('ok','Reporte guardado. Espera aprobación del líder.');
      setStep(1);
    }catch(e:any){ showToast('err',e?.message||'Error'); }
    finally{ setSaving(false); }
  }

  const asistio=asistencia.filter(a=>a.asistio).length;
  const efic=asistencia.length>0?Math.round((asistio/asistencia.length)*100):100;

  return(
    <div className="space-y-4">
      <div className="card p-3">
        <div className="flex items-center gap-1 overflow-x-auto">
          {STEPS.map((s,i)=>{const n=i+1; return(
            <div key={n} className="flex items-center gap-1 flex-shrink-0">
              <button onClick={()=>setStep(n)} className={n<step?'step-done':n===step?'step-active':'step-pending'}>{n<step?'✓':n}</button>
              <span className={`text-xs whitespace-nowrap hidden sm:inline ${n===step?'font-semibold text-[#003b7a]':'text-slate-500'}`}>{s}</span>
              {i<STEPS.length-1&&<span className="text-slate-300 mx-1">›</span>}
            </div>
          );})}
        </div>
      </div>

      {step===1&&<div className="card p-4 space-y-3"><h3 className="font-bold text-[#003b7a]">Paso 1 — Encabezado</h3>
        <div className="grid grid-cols-1 sm:grid-cols-3 gap-3">
          <div><label className="label">Fecha</label><input type="date" className="input" value={fecha} onChange={e=>setFecha(e.target.value)}/></div>
          <div><label className="label">Especialidad</label>
            <select className="select" value={espId} onChange={e=>setEspId(e.target.value)}>
              <option value="">— Seleccionar —</option>
              {catalogs&&uniqueEsp(catalogs.especialidades_actividades).map(e=><option key={e.id} value={e.id}>{e.especialidad_es}</option>)}
            </select></div>
          <div><label className="label">Horas de jornada</label><input type="number" className="input" min={1} max={24} value={jornadaHrs} onChange={e=>setJornadaHrs(parseFloat(e.target.value)||JORNADA)}/></div>
        </div>
        <button className="btn-primary" onClick={()=>{if(!fecha||!espId){showToast('err','Completa fecha y especialidad');return;}setStep(2);}}>Siguiente →</button>
      </div>}

      {step===2&&<div className="card p-4 space-y-3"><h3 className="font-bold text-[#003b7a]">Paso 2 — Condiciones del día</h3>
        <div><label className="label">Clima</label>
          <div className="flex gap-2 flex-wrap">
            {['despejado','nublado','lluvia','tormenta','suspendido'].map(c=>(
              <button key={c} onClick={()=>setClima(c)} className={`text-sm px-3 py-1.5 rounded-full border transition-colors ${clima===c?'bg-[#003b7a] text-white border-[#003b7a]':'border-slate-300 text-slate-600 hover:border-[#003b7a]'}`}>
                {c==='despejado'?'☀️ Despejado':c==='nublado'?'☁️ Nublado':c==='lluvia'?'🌧️ Lluvia':c==='tormenta'?'⛈️ Tormenta':'🚫 Suspendido'}
              </button>
            ))}
          </div>
        </div>
        <div><label className="label">Suspensiones por clima</label>
          {suspensiones.map((s,i)=>(
            <div key={s.uid} className="flex gap-2 items-center mt-2 flex-wrap">
              <input type="time" className="input w-32" value={s.hora_inicio} onChange={e=>setSuspensiones(a=>a.map((x,j)=>j===i?{...x,hora_inicio:e.target.value}:x))}/>
              <span className="text-sm text-slate-500">a</span>
              <input type="time" className="input w-32" value={s.hora_fin} onChange={e=>setSuspensiones(a=>a.map((x,j)=>j===i?{...x,hora_fin:e.target.value}:x))}/>
              <input className="input flex-1" value={s.descripcion} onChange={e=>setSuspensiones(a=>a.map((x,j)=>j===i?{...x,descripcion:e.target.value}:x))} placeholder="Descripción"/>
              <button className="btn-ghost text-rose-500 text-xs" onClick={()=>setSuspensiones(a=>a.filter(x=>x.uid!==s.uid))}>✕</button>
            </div>
          ))}
          <button className="btn-secondary text-xs mt-2" onClick={()=>setSuspensiones(a=>[...a,{uid:uid(),hora_inicio:'',hora_fin:'',descripcion:''}])}>+ Agregar suspensión</button>
          {horasClima>0&&<p className="text-sm text-amber-600 mt-2">⏱ {horasClima.toFixed(1)}h perdidas → {horasReal.toFixed(1)}h operativas</p>}
        </div>
        <div className="flex gap-2"><button className="btn-secondary" onClick={()=>setStep(1)}>← Anterior</button><button className="btn-primary" onClick={()=>setStep(3)}>Siguiente →</button></div>
      </div>}

      {step===3&&<div className="card p-4 space-y-3"><h3 className="font-bold text-[#003b7a]">Paso 3 — Charla preturno</h3>
        <label className="flex items-center gap-3 cursor-pointer"><input type="checkbox" checked={charla} onChange={e=>setCharla(e.target.checked)} className="w-5 h-5"/><span className="text-sm font-medium">Se realizó la charla preturno</span></label>
        {charla&&<div><label className="label">Tema</label><input className="input" value={charlaTema} onChange={e=>setCharlaTema(e.target.value)} placeholder="Ej: Uso seguro de motosierra…"/></div>}
        <div className="flex gap-2"><button className="btn-secondary" onClick={()=>setStep(2)}>← Anterior</button><button className="btn-primary" onClick={()=>setStep(4)}>Siguiente →</button></div>
      </div>}

      {step===4&&<div className="card p-4 space-y-3"><h3 className="font-bold text-[#003b7a]">Paso 4 — Asistencia real</h3>
        <div className="flex gap-4 text-sm flex-wrap">
          <span>Planeado: <strong>{asistencia.length}</strong></span>
          <span className="text-emerald-600">Asistió: <strong>{asistio}</strong></span>
          <span>Cumplimiento: <strong>{efic}%</strong> {semaforo(efic)}</span>
        </div>
        <div className="space-y-2">
          {asistencia.map((a,i)=>(
            <div key={a.personal_id} className="flex items-center gap-3 p-2 border border-slate-200 rounded-lg flex-wrap">
              <input type="checkbox" checked={a.asistio} onChange={e=>setAsistencia(arr=>arr.map((x,j)=>j===i?{...x,asistio:e.target.checked,motivo_ausencia:e.target.checked?'':x.motivo_ausencia}:x))}/>
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
          {!asistencia.length&&<p className="text-sm text-slate-500">Sin personal planeado para este día/especialidad.</p>}
        </div>
        <div className="flex gap-2"><button className="btn-secondary" onClick={()=>setStep(3)}>← Anterior</button><button className="btn-primary" onClick={()=>setStep(5)}>Siguiente →</button></div>
      </div>}

      {step===5&&<div className="card p-4 space-y-3"><h3 className="font-bold text-[#003b7a]">Paso 5 — Maquinaria del día</h3>
        {maqDia.map((m,i)=>(
          <div key={m.maquinaria_id} className="border border-slate-200 rounded-lg p-3">
            <div className="flex items-center justify-between mb-2">
              <span className="font-medium text-sm">{m.nombre}</span>
              <span className="text-xs text-slate-500">Op: {Math.max(0,horasReal-m.horas_standby).toFixed(1)}h | SB: {m.horas_standby}h</span>
            </div>
            <label className="flex items-center gap-2 text-sm mb-2"><input type="checkbox" checked={m.novedad} onChange={e=>setMaqDia(a=>a.map((x,j)=>j===i?{...x,novedad:e.target.checked}:x))}/> ¿Tuvo novedad hoy?</label>
            {m.novedad&&<div className="grid grid-cols-1 sm:grid-cols-3 gap-2">
              <div><label className="label">Hora inicio</label><input type="time" className="input" value={m.hora_inicio} onChange={e=>setMaqDia(a=>a.map((x,j)=>j===i?{...x,hora_inicio:e.target.value}:x))}/></div>
              <div><label className="label">Hora fin</label><input type="time" className="input" value={m.hora_fin} onChange={e=>{ const[ih,im]=(m.hora_inicio||'0:0').split(':').map(Number);const[fh,fm]=e.target.value.split(':').map(Number);const diff=Math.max(0,(fh+fm/60)-(ih+im/60)); setMaqDia(a=>a.map((x,j)=>j===i?{...x,hora_fin:e.target.value,horas_standby:parseFloat(diff.toFixed(2))}:x));  }}/></div>
              <div><label className="label">Descripción</label><input className="input" value={m.descripcion} onChange={e=>setMaqDia(a=>a.map((x,j)=>j===i?{...x,descripcion:e.target.value}:x))}/></div>
            </div>}
          </div>
        ))}
        <div className="flex gap-2"><button className="btn-secondary" onClick={()=>setStep(4)}>← Anterior</button><button className="btn-primary" onClick={()=>setStep(6)}>Siguiente →</button></div>
      </div>}

      {step===6&&<div className="card p-4 space-y-3"><h3 className="font-bold text-[#003b7a]">Paso 6 — Avance de actividades</h3>
        <p className="text-xs text-slate-500">Solo llena las actividades que ejecutaste hoy.</p>
        {avances.map((av,i)=>{
          const actRow=catalogs?.especialidades_actividades.find(e=>e.id===av.actividad_id);
          const cfg=configActs.find(c=>c.actividad_id===av.actividad_id);
          return(
            <div key={av.uid} className="border border-slate-200 rounded-lg p-3 space-y-2">
              <div className="font-medium text-sm text-[#003b7a]">{actRow?.actividad_es||av.actividad_id} {cfg&&<span className="text-xs text-slate-400 ml-1">Tipo {cfg.tipo} · {cfg.unidad_es}</span>}</div>
              <div className="grid grid-cols-1 sm:grid-cols-3 gap-2">
                <div><label className="label">Área</label>
                  <select className="select" value={av.area_id} onChange={e=>setAvances(a=>a.map((x,j)=>j===i?{...x,area_id:e.target.value}:x))}>
                    <option value="">— Área —</option>
                    {(catalogs?.areas||[]).map(a=><option key={a.id} value={a.id}>{a.area_es}</option>)}
                  </select></div>
                <div><label className="label">Cantidad hoy</label><input type="number" className="input" min={0} value={av.cantidad} onChange={e=>setAvances(a=>a.map((x,j)=>j===i?{...x,cantidad:e.target.value}:x))}/></div>
                <div><label className="label">Unidad</label><input className="input" value={av.unidad} onChange={e=>setAvances(a=>a.map((x,j)=>j===i?{...x,unidad:e.target.value}:x))} placeholder={cfg?.unidad_es||'unidad'}/></div>
              </div>
            </div>
          );
        })}
        <div className="flex gap-2"><button className="btn-secondary" onClick={()=>setStep(5)}>← Anterior</button><button className="btn-primary" onClick={()=>setStep(7)}>Siguiente →</button></div>
      </div>}

      {step===7&&<div className="card p-4 space-y-3"><h3 className="font-bold text-[#003b7a]">Paso 7 — Seguridad y novedades</h3>
        <div><label className="label">Incidente de seguridad</label>
          <div className="flex gap-2 flex-wrap">
            {['sin_novedad','casi_accidente','incidente','accidente'].map(t=>(
              <button key={t} onClick={()=>setIncidente(i=>({...i,tipo:t}))} className={`text-xs px-3 py-1.5 rounded-full border transition-colors ${incidente.tipo===t?t==='sin_novedad'?'bg-emerald-500 text-white border-emerald-500':t==='accidente'?'bg-rose-600 text-white border-rose-600':'bg-amber-500 text-white border-amber-500':'border-slate-300 text-slate-600 hover:border-slate-400'}`}>
                {t==='sin_novedad'?'✅ Sin novedad':t==='casi_accidente'?'⚠️ Casi accidente':t==='incidente'?'🔶 Incidente':'🚨 Accidente'}
              </button>
            ))}
          </div>
        </div>
        {incidente.tipo!=='sin_novedad'&&<>
          <div><label className="label">Descripción</label><textarea className="textarea" rows={2} value={incidente.descripcion} onChange={e=>setIncidente(i=>({...i,descripcion:e.target.value}))}/></div>
          <div><label className="label">Medidas tomadas</label><textarea className="textarea" rows={2} value={incidente.medidas} onChange={e=>setIncidente(i=>({...i,medidas:e.target.value}))}/></div>
        </>}
        <div><label className="label">Nota de bitácora (opcional)</label>
          <textarea className="textarea" rows={2} value={notaBit} onChange={e=>setNotaBit(e.target.value)} placeholder="Decisiones importantes del día…"/></div>
        <div className="flex gap-2"><button className="btn-secondary" onClick={()=>setStep(6)}>← Anterior</button><button className="btn-primary" onClick={()=>setStep(8)}>Siguiente →</button></div>
      </div>}

      {step===8&&<div className="card p-4 space-y-4"><h3 className="font-bold text-[#003b7a]">Paso 8 — Resumen y envío</h3>
        <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">
          <div className="card p-3 text-center"><div className="text-2xl font-bold text-[#003b7a]">{asistio}</div><div className="text-xs text-slate-500">Personal activo</div></div>
          <div className="card p-3 text-center"><div className="text-2xl font-bold text-emerald-600">{(asistio*horasReal).toFixed(0)}h</div><div className="text-xs text-slate-500">Horas-hombre</div></div>
          <div className="card p-3 text-center"><div className="text-2xl font-bold text-amber-500">{horasClima.toFixed(1)}h</div><div className="text-xs text-slate-500">Perdidas clima</div></div>
          <div className="card p-3 text-center"><div className="text-2xl font-bold text-[#003b7a]">{avances.filter(a=>parseFloat(a.cantidad)>0).length}</div><div className="text-xs text-slate-500">Actividades</div></div>
        </div>
        <div className="bg-slate-50 rounded-lg p-3 text-sm space-y-1">
          <div><strong>Fecha:</strong> {fecha}</div>
          <div><strong>Clima:</strong> {clima} {horasClima>0?`· ${horasClima.toFixed(1)}h perdidas`:''}</div>
          <div><strong>Asistencia:</strong> {asistio}/{asistencia.length} ({efic}%) {semaforo(efic)}</div>
          <div><strong>Incidente:</strong> {incidente.tipo}</div>
        </div>
        <p className="text-xs text-slate-500">Quedará en <strong>Borrador</strong> hasta que el líder o admin lo apruebe.</p>
        <div className="flex gap-2 flex-wrap">
          <button className="btn-secondary" onClick={()=>setStep(7)}>← Anterior</button>
          <button className="btn-success" disabled={saving} onClick={submit}>{saving?'Enviando…':'✅ Enviar reporte del día'}</button>
        </div>
      </div>}
    </div>
  );
}

/* ═══════════════════════════════════════════════════════════════
 *  DASHBOARD
 * ═══════════════════════════════════════════════════════════════ */
function DashboardModule({user,showToast}:{user:Profile;showToast:any}){
  const[fechaIni,setFechaIni]=useState(today());
  const[fechaFin,setFechaFin]=useState(today());
  const[data,setData]=useState<any>(null);
  const[loading,setLoading]=useState(false);

  const load=useCallback(async()=>{
    setLoading(true);
    try{
      const[reps,asist,avances,novMaq,incid]=await Promise.all([
        supabase.from('reportes_avance').select('*').gte('fecha',fechaIni).lte('fecha',fechaFin),
        supabase.from('asistencia_real').select('*').gte('fecha',fechaIni).lte('fecha',fechaFin),
        supabase.from('avance_diario').select('*').gte('fecha',fechaIni).lte('fecha',fechaFin),
        supabase.from('novedades_maquinaria').select('*,maquinaria(item_id,tipo)').gte('fecha',fechaIni).lte('fecha',fechaFin),
        supabase.from('incidentes_seg').select('*').gte('fecha',fechaIni).lte('fecha',fechaFin).neq('tipo','sin_novedad'),
      ]);
      const asistData=asist.data||[];
      const horasH=asistData.filter((a:any)=>a.asistio).reduce((s:number,a:any)=>s+parseFloat(a.horas_trabajadas||0),0);
      const planeados=asistData.length, reales=asistData.filter((a:any)=>a.asistio).length;
      const maqData=await supabase.from('maquinaria').select('*');
      setData({ reportes:reps.data||[], horas_hombre:Math.round(horasH), eficiencia_personal:planeados>0?Math.round(reales/planeados*100):100, avances:avances.data||[], novedades_maquinaria:novMaq.data||[], incidentes:incid.data||[], maquinaria:maqData.data||[] });
    }catch(e:any){showToast('err',e?.message||'Error');}
    finally{setLoading(false);}
  },[fechaIni,fechaFin,showToast]);

  useEffect(()=>{load();},[load]);

  return(
    <div className="space-y-4">
      <div className="card p-4 flex gap-3 items-end flex-wrap no-print">
        <div className="flex-1 min-w-[140px]"><label className="label">Desde</label><input type="date" className="input" value={fechaIni} onChange={e=>setFechaIni(e.target.value)}/></div>
        <div className="flex-1 min-w-[140px]"><label className="label">Hasta</label><input type="date" className="input" value={fechaFin} onChange={e=>setFechaFin(e.target.value)}/></div>
        <button className="btn-primary" onClick={load} disabled={loading}>{loading?'Cargando…':'Actualizar'}</button>
      </div>
      {data&&<>
        <div className="grid grid-cols-2 sm:grid-cols-4 gap-3">
          <MetricCard label="Eficiencia personal" value={`${data.eficiencia_personal}%`} sub="asistencia" color={data.eficiencia_personal>=90?'text-emerald-600':data.eficiencia_personal>=70?'text-amber-500':'text-rose-600'}/>
          <MetricCard label="Horas-hombre" value={`${data.horas_hombre}h`} sub="productivas"/>
          <MetricCard label="Reportes" value={data.reportes.length} sub="en el período"/>
          <MetricCard label="Incidentes" value={data.incidentes.length} sub="de seguridad" color={data.incidentes.length>0?'text-rose-600':'text-emerald-600'}/>
        </div>
        {data.maquinaria.length>0&&(
          <div className="card p-4">
            <h3 className="font-bold text-[#003b7a] mb-3">🔧 Maquinaria — horas acumuladas</h3>
            <div className="table-wrap">
              <table className="table"><thead><tr><th>Equipo</th><th>Tipo</th><th>Horas op.</th><th>Horas SB</th><th>Eficiencia</th></tr></thead>
                <tbody>{data.maquinaria.map((m:any)=>{
                  const t=(m.horas_acum_operativas||0)+(m.horas_acum_standby||0);
                  const ef=t>0?Math.round((m.horas_acum_operativas||0)/t*100):100;
                  return <tr key={m.id}><td className="font-medium">{m.item_id}</td><td>{m.tipo}</td><td className="text-emerald-600">{(m.horas_acum_operativas||0).toFixed(1)}h</td><td className="text-amber-500">{(m.horas_acum_standby||0).toFixed(1)}h</td><td>{semaforo(ef)} {ef}%</td></tr>;
                })}</tbody>
              </table>
            </div>
          </div>
        )}
        {data.incidentes.length>0&&(
          <div className="card p-4 border-l-4 border-rose-400">
            <h3 className="font-bold text-rose-700 mb-2">⚠️ Incidentes de seguridad</h3>
            {data.incidentes.map((i:any)=><div key={i.id} className="text-sm mb-1"><strong>{i.fecha} · {i.tipo}</strong>: {i.descripcion}</div>)}
          </div>
        )}
      </>}
    </div>
  );
}
function MetricCard({label,value,sub,color='text-[#003b7a]'}:{label:string;value:any;sub:string;color?:string}){
  return <div className="card p-4 text-center"><div className={`text-2xl sm:text-3xl font-bold ${color}`}>{value}</div><div className="text-xs font-semibold text-slate-700 mt-1">{label}</div><div className="text-xs text-slate-400">{sub}</div></div>;
}

/* ═══════════════════════════════════════════════════════════════
 *  INFORMES
 * ═══════════════════════════════════════════════════════════════ */
function InformesModule({user,catalogs,showToast}:{user:Profile;catalogs:Catalogs|null;showToast:any}){
  const[fechaIni,setFechaIni]=useState(today());
  const[fechaFin,setFechaFin]=useState(today());
  const[espIds,setEspIds]=useState<string[]>([]);
  const[areaIds,setAreaIds]=useState<string[]>([]);
  const[soloAprobados,setSoloAprobados]=useState(false);
  const[data,setData]=useState<any>(null);
  const[loading,setLoading]=useState(false);

  const espList=useMemo(()=>catalogs?uniqueEsp(catalogs.especialidades_actividades):[],[catalogs]);

  async function fetchData(){
    setLoading(true);
    try{
      let qRep=supabase.from('reportes_avance').select('*').gte('fecha',fechaIni).lte('fecha',fechaFin);
      if(espIds.length) qRep=qRep.in('especialidad_id',espIds);
      if(soloAprobados) qRep=qRep.eq('estado','aprobado');
      if(user.rol==='tecnico') qRep=qRep.eq('usuario_id',user.id);
      const{data:reps}=await qRep;
      const repIds=(reps||[]).map((r:any)=>r.id);
      let[av,as2,nm,sc,inc]=await Promise.all([
        repIds.length?supabase.from('avance_diario').select('*').in('reporte_id',repIds):Promise.resolve({data:[]}),
        repIds.length?supabase.from('asistencia_real').select('*').in('reporte_id',repIds):Promise.resolve({data:[]}),
        repIds.length?supabase.from('novedades_maquinaria').select('*').in('reporte_id',repIds):Promise.resolve({data:[]}),
        repIds.length?supabase.from('suspensiones_clima').select('*').in('reporte_id',repIds):Promise.resolve({data:[]}),
        repIds.length?supabase.from('incidentes_seg').select('*').in('reporte_id',repIds).neq('tipo','sin_novedad'):Promise.resolve({data:[]}),
      ]);
      let avData=(av.data||[]) as any[];
      if(areaIds.length) avData=avData.filter((a:any)=>areaIds.includes(a.area_id));
      const asistData=as2.data||[] as any[];
      const horasH=(asistData as any[]).filter((a:any)=>a.asistio).reduce((s:number,a:any)=>s+parseFloat(a.horas_trabajadas||0),0);
      const horasClima=(sc.data||[]).reduce((s:number,a:any)=>s+parseFloat(a.horas_perdidas||0),0);
      setData({ reportes:reps||[], avances:avData, asistencia:asistData, novedades_maquinaria:nm.data||[], suspensiones_clima:sc.data||[], incidentes:inc.data||[], totales:{ horas_hombre:Math.round(horasH), horas_perdidas_clima:Math.round(horasClima), dias:repIds.length } });
    }catch(e:any){showToast('err',e?.message||'Error');}
    finally{setLoading(false);}
  }

  function exportExcel(){
    if(!data) return;
    const rows=data.avances.map((av:any)=>{
      const actRow=catalogs?.especialidades_actividades.find(e=>e.id===av.actividad_id);
      const areaRow=catalogs?.areas.find(a=>a.id===av.area_id);
      const rep=data.reportes.find((r:any)=>r.id===av.reporte_id);
      return{'Fecha':av.fecha,'Especialidad':actRow?.especialidad_es||av.especialidad_id,'Actividad':actRow?.actividad_es||av.actividad_id,'Área':areaRow?.area_es||av.area_id,'Cantidad':av.cantidad,'Unidad':av.unidad,'Acumulado':av.acumulado_total,'Usuario':rep?.usuario_nombre||''};
    });
    const wb=XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(rows),'Avance');
    XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(data.asistencia.map((a:any)=>({'Fecha':a.fecha,'Doc':a.documento_personal,'Asistió':a.asistio?'Sí':'No','Motivo':a.motivo_ausencia||'','Horas':a.horas_trabajadas}))),'Asistencia');
    XLSX.writeFile(wb,`PDS360_${fechaIni}_${fechaFin}.xlsx`);
    showToast('ok','Excel descargado');
  }

  function toggleFilter(arr:string[],setArr:any,val:string){ setArr(arr.includes(val)?arr.filter(x=>x!==val):[...arr,val]); }

  return(
    <div className="space-y-4">
      <div className="card p-4 space-y-3 no-print">
        <div className="grid grid-cols-1 sm:grid-cols-3 gap-3 items-end">
          <div><label className="label">Desde</label><input type="date" className="input" value={fechaIni} onChange={e=>setFechaIni(e.target.value)}/></div>
          <div><label className="label">Hasta</label><input type="date" className="input" value={fechaFin} onChange={e=>setFechaFin(e.target.value)}/></div>
          <button className="btn-primary" onClick={fetchData} disabled={loading}>{loading?'Cargando…':'Consultar'}</button>
        </div>
        <div><label className="label">Especialidades (vacío = todas)</label>
          <div className="flex flex-wrap gap-2">{espList.map(e=><button key={e.id} onClick={()=>toggleFilter(espIds,setEspIds,e.id)} className={`text-xs px-2 py-1 rounded border transition-colors ${espIds.includes(e.id)?'bg-[#003b7a] text-white border-[#003b7a]':'border-slate-300 text-slate-600 hover:border-[#003b7a]'}`}>{e.especialidad_es}</button>)}</div>
        </div>
        <div><label className="label">Áreas (vacío = todas)</label>
          <div className="flex flex-wrap gap-2">{(catalogs?.areas||[]).map(a=><button key={a.id} onClick={()=>toggleFilter(areaIds,setAreaIds,a.id)} className={`text-xs px-2 py-1 rounded border transition-colors ${areaIds.includes(a.id)?'bg-[#003b7a] text-white border-[#003b7a]':'border-slate-300 text-slate-600 hover:border-[#003b7a]'}`}>{a.area_es}</button>)}</div>
        </div>
        <div className="flex items-center gap-3 flex-wrap">
          <label className="flex items-center gap-2 text-sm cursor-pointer"><input type="checkbox" checked={soloAprobados} onChange={e=>setSoloAprobados(e.target.checked)}/> Solo aprobados</label>
          <button className="btn-secondary text-xs" onClick={exportExcel}>📥 Excel</button>
          <button className="btn-secondary text-xs" onClick={()=>window.print()}>🖨️ PDF</button>
        </div>
      </div>

      <div className="print-area space-y-4">
        {data?.totales&&<div className="card p-4"><h3 className="font-bold text-[#003b7a] mb-3">Totales del período</h3>
          <div className="grid grid-cols-3 gap-3 text-center">
            <div><div className="text-xl font-bold text-[#003b7a]">{data.totales.dias}</div><div className="text-xs text-slate-500">Reportes</div></div>
            <div><div className="text-xl font-bold text-emerald-600">{data.totales.horas_hombre}h</div><div className="text-xs text-slate-500">Horas-hombre</div></div>
            <div><div className="text-xl font-bold text-amber-500">{data.totales.horas_perdidas_clima}h</div><div className="text-xs text-slate-500">Perdidas clima</div></div>
          </div>
        </div>}

        {data?.avances?.length>0&&<div className="card p-4"><h3 className="font-bold text-[#003b7a] mb-3">Avance de actividades</h3>
          <div className="table-wrap"><table className="table">
            <thead><tr><th>Fecha</th><th>Actividad</th><th>Área</th><th>Hoy</th><th>Acumulado</th><th>Unidad</th></tr></thead>
            <tbody>{data.avances.map((av:any,i:number)=>{
              const actRow=catalogs?.especialidades_actividades.find(e=>e.id===av.actividad_id);
              const areaRow=catalogs?.areas.find(a=>a.id===av.area_id);
              return <tr key={i}><td>{av.fecha}</td><td>{actRow?.actividad_es||av.actividad_id}</td><td>{areaRow?.area_es||av.area_id}</td><td className="font-semibold">{av.cantidad}</td><td className="font-semibold text-[#003b7a]">{av.acumulado_total}</td><td>{av.unidad}</td></tr>;
            })}</tbody>
          </table></div>
        </div>}

        {!data?.avances?.length&&!loading&&<div className="card p-6 text-center text-slate-500">Sin datos para el período seleccionado.</div>}
      </div>
    </div>
  );
}

/* ═══════════════════════════════════════════════════════════════
 *  CATÁLOGOS
 * ═══════════════════════════════════════════════════════════════ */
function CatalogosModule({catalogs,onRefresh,showToast}:{catalogs:Catalogs|null;onRefresh:()=>void;showToast:any}){
  return(
    <div className="space-y-4">
      <CatMgr title="Especialidades y Actividades" table="especialidades_actividades"
        fields={[{n:'especialidad_es',l:'Especialidad (ES)'},{n:'especialidad_en',l:'Especialidad (EN)'},{n:'actividad_es',l:'Actividad (ES)'},{n:'actividad_en',l:'Actividad (EN)'}]}
        rows={catalogs?.especialidades_actividades||[]} onChanged={onRefresh} showToast={showToast}/>
      <CatMgr title="Áreas" table="areas"
        fields={[{n:'area_es',l:'Área (ES)'},{n:'area_en',l:'Área (EN)'}]}
        rows={catalogs?.areas||[]} onChanged={onRefresh} showToast={showToast}/>
      <CatMgr title="Líderes" table="lideres"
        fields={[{n:'nombre',l:'Nombre'},{n:'documento',l:'Documento'},{n:'cargo_es',l:'Cargo (ES)'},{n:'cargo_en',l:'Cargo (EN)'}]}
        rows={catalogs?.lideres||[]} onChanged={onRefresh} showToast={showToast}/>
      <CatMgr title="Personal" table="personal"
        fields={[{n:'nombre',l:'Nombre'},{n:'documento',l:'Documento'},{n:'cargo_es',l:'Cargo (ES)'},{n:'cargo_en',l:'Cargo (EN)'},{n:'tipo',l:'Tipo (directo/subcont.)'},{n:'empresa',l:'Empresa'}]}
        rows={catalogs?.personal||[]} onChanged={onRefresh} showToast={showToast}/>
    </div>
  );
}

function CatMgr({title,table,fields,rows,onChanged,showToast}:{title:string;table:string;fields:{n:string;l:string}[];rows:any[];onChanged:()=>void;showToast:any}){
  const[form,setForm]=useState<Record<string,string>>({});
  const[search,setSearch]=useState('');
  const[busy,setBusy]=useState(false);
  const filtered=useMemo(()=>{const q=search.trim().toLowerCase();return q?rows.filter(r=>fields.some(f=>String(r[f.n]||'').toLowerCase().includes(q))):rows;},[rows,search,fields]);

  async function addOne(){
    if(!form[fields[0].n]){showToast('err',`Falta ${fields[0].l}`);return;}
    setBusy(true);
    try{
      const{error}=await supabase.from(table).insert({...form,activo:true});
      if(error) throw error;
      showToast('ok','Agregado'); setForm({}); onChanged();
    }catch(e:any){showToast('err',e?.message||'Error');}
    finally{setBusy(false);}
  }

  async function loadExcel(file:File){
    setBusy(true);
    try{
      const buf=await file.arrayBuffer();
      const wb=XLSX.read(buf);
      const data:any[]=XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{defval:''});
      const norm=data.map(row=>{const o:any={activo:true};for(const k in row){o[String(k).trim().toLowerCase().replace(/\s+/g,'_')]=row[k];}return o;});
      if(!norm.length){showToast('err','Excel vacío');return;}
      if(!window.confirm(`¿Reemplazar "${title}" con ${norm.length} filas?`)) return;
      await supabase.from(table).delete().neq('id','00000000-0000-0000-0000-000000000000');
      const{error}=await supabase.from(table).insert(norm);
      if(error) throw error;
      showToast('ok','Catálogo reemplazado'); onChanged();
    }catch(e:any){showToast('err',e?.message||'Error');}
    finally{setBusy(false);}
  }

  return(
    <div className="card p-4">
      <div className="flex items-center justify-between flex-wrap gap-2 mb-3">
        <h3 className="font-bold text-[#003b7a]">{title} ({rows.length})</h3>
        <label className="btn-secondary cursor-pointer text-xs">📥 Cargar Excel<input type="file" accept=".xlsx,.xls,.csv" hidden onChange={e=>{const f=e.target.files?.[0];if(f)loadExcel(f);e.currentTarget.value='';}} /></label>
      </div>
      <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-6 gap-2 mb-2">
        {fields.map(f=><div key={f.n}><label className="label">{f.l}</label><input className="input" value={form[f.n]||''} onChange={e=>setForm({...form,[f.n]:e.target.value})}/></div>)}
      </div>
      <button className="btn-primary text-xs" disabled={busy} onClick={addOne}>+ Agregar</button>
      {rows.length>0&&<>
        <input className="input mt-3" placeholder="🔎 Buscar…" value={search} onChange={e=>setSearch(e.target.value)}/>
        <div className="table-wrap mt-2 max-h-64 overflow-auto">
          <table className="table"><thead><tr>{fields.map(f=><th key={f.n}>{f.l}</th>)}</tr></thead>
            <tbody>{filtered.slice(0,200).map((r,i)=><tr key={i}>{fields.map(f=><td key={f.n}>{String(r[f.n]??'')}</td>)}</tr>)}</tbody>
          </table>
          {filtered.length>200&&<div className="text-xs text-slate-500 p-2">Mostrando 200 de {filtered.length}</div>}
        </div>
      </>}
    </div>
  );
}

/* ═══════════════════════════════════════════════════════════════
 *  MAQUINARIA
 * ═══════════════════════════════════════════════════════════════ */
function MaquinariaModule({maquinaria,onRefresh,showToast}:{maquinaria:Maq[];onRefresh:()=>void;showToast:any}){
  const[form,setForm]=useState({tipo:'motosierra',item_id:'',nombre:'',estado:'activo'});
  const[busy,setBusy]=useState(false);

  async function addOne(){
    if(!form.item_id){showToast('err','Falta ID');return;}
    setBusy(true);
    try{
      const{error}=await supabase.from('maquinaria').insert({...form,horas_acum_operativas:0,horas_acum_standby:0});
      if(error) throw error;
      showToast('ok','Equipo agregado'); setForm({tipo:'motosierra',item_id:'',nombre:'',estado:'activo'}); onRefresh();
    }catch(e:any){showToast('err',e?.message||'Error');}
    finally{setBusy(false);}
  }

  return(
    <div className="card p-4 space-y-3">
      <h3 className="font-bold text-[#003b7a]">Maquinaria ({maquinaria.length} equipos)</h3>
      <div className="grid grid-cols-2 sm:grid-cols-4 gap-2">
        <div><label className="label">Tipo</label>
          <select className="select" value={form.tipo} onChange={e=>setForm({...form,tipo:e.target.value})}>
            <option value="motosierra">Motosierra</option><option value="chipeadora">Chipeadora</option><option value="camion">Camión</option><option value="otro">Otro</option>
          </select></div>
        <div><label className="label">ID único</label><input className="input" value={form.item_id} onChange={e=>setForm({...form,item_id:e.target.value})} placeholder="MS-009"/></div>
        <div><label className="label">Nombre</label><input className="input" value={form.nombre} onChange={e=>setForm({...form,nombre:e.target.value})}/></div>
        <div><label className="label">Estado</label>
          <select className="select" value={form.estado} onChange={e=>setForm({...form,estado:e.target.value})}>
            <option value="activo">Activo</option><option value="inactivo">Inactivo</option><option value="mantenimiento">Mantenimiento</option>
          </select></div>
      </div>
      <button className="btn-primary text-xs" disabled={busy} onClick={addOne}>+ Agregar equipo</button>
      <div className="table-wrap"><table className="table">
        <thead><tr><th>ID</th><th>Tipo</th><th>Nombre</th><th>Estado</th><th>Horas op.</th><th>Horas SB</th></tr></thead>
        <tbody>{maquinaria.map(m=><tr key={m.id}><td className="font-medium">{m.item_id}</td><td>{m.tipo}</td><td>{m.nombre}</td><td><span className={m.estado==='activo'?'text-emerald-600':m.estado==='mantenimiento'?'text-amber-500':'text-slate-400'}>{m.estado}</span></td><td>{m.horas_acum_operativas||0}h</td><td>{m.horas_acum_standby||0}h</td></tr>)}</tbody>
      </table></div>
    </div>
  );
}

/* ═══════════════════════════════════════════════════════════════
 *  CONFIG ACTIVIDADES
 * ═══════════════════════════════════════════════════════════════ */
function ConfigActModule({configActs,catalogs,onRefresh,showToast}:{configActs:ConfigAct[];catalogs:Catalogs|null;onRefresh:()=>void;showToast:any}){
  const[form,setForm]=useState<any>({especialidad_id:'',actividad_id:'',tipo:'A',unidad_es:'',unidad_en:'',meta_total:'',acumulado_previo:'',rendimiento_esperado:'',rendimiento_por:'cuadrilla',activo:true});
  const[busy,setBusy]=useState(false);
  const esps=useMemo(()=>catalogs?uniqueEsp(catalogs.especialidades_actividades):[],[catalogs]);
  const acts=useMemo(()=>catalogs&&form.especialidad_id?actsForEsp(catalogs.especialidades_actividades,form.especialidad_id):[],[catalogs,form.especialidad_id]);

  async function save(){
    if(!form.actividad_id){showToast('err','Selecciona actividad');return;}
    setBusy(true);
    try{
      const{error}=await supabase.from('config_actividades').upsert({...form,meta_total:form.meta_total?parseFloat(form.meta_total):null,acumulado_previo:parseFloat(form.acumulado_previo||'0'),rendimiento_esperado:form.rendimiento_esperado?parseFloat(form.rendimiento_esperado):null},{onConflict:'actividad_id'});
      if(error) throw error;
      showToast('ok','Guardado'); onRefresh();
    }catch(e:any){showToast('err',e?.message||'Error');}
    finally{setBusy(false);}
  }

  return(
    <div className="card p-4 space-y-3">
      <h3 className="font-bold text-[#003b7a]">Configuración de actividades ({configActs.length})</h3>
      <div className="grid grid-cols-1 sm:grid-cols-3 gap-2">
        <div><label className="label">Especialidad</label>
          <select className="select" value={form.especialidad_id} onChange={e=>setForm({...form,especialidad_id:e.target.value,actividad_id:''})}>
            <option value="">— Seleccionar —</option>{esps.map(e=><option key={e.id} value={e.id}>{e.especialidad_es}</option>)}
          </select></div>
        <div><label className="label">Actividad</label>
          <select className="select" value={form.actividad_id} onChange={e=>setForm({...form,actividad_id:e.target.value})} disabled={!form.especialidad_id}>
            <option value="">— Seleccionar —</option>{acts.map(a=><option key={a.id} value={a.id}>{a.actividad_es}</option>)}
          </select></div>
        <div><label className="label">Tipo</label>
          <select className="select" value={form.tipo} onChange={e=>setForm({...form,tipo:e.target.value})}>
            <option value="A">A — Con meta numérica</option><option value="B">B — Acumulativa sin meta</option><option value="C">C — Ítems únicos</option>
          </select></div>
        <div><label className="label">Unidad (ES)</label><input className="input" value={form.unidad_es} onChange={e=>setForm({...form,unidad_es:e.target.value})} placeholder="árboles, m³…"/></div>
        <div><label className="label">Unidad (EN)</label><input className="input" value={form.unidad_en} onChange={e=>setForm({...form,unidad_en:e.target.value})}/></div>
        {form.tipo==='A'&&<div><label className="label">Meta total</label><input type="number" className="input" value={form.meta_total} onChange={e=>setForm({...form,meta_total:e.target.value})}/></div>}
        {form.tipo==='A'&&<div><label className="label">Acumulado previo (día 0)</label><input type="number" className="input" value={form.acumulado_previo} onChange={e=>setForm({...form,acumulado_previo:e.target.value})}/></div>}
        <div><label className="label">Rendimiento esperado</label><input type="number" className="input" value={form.rendimiento_esperado} onChange={e=>setForm({...form,rendimiento_esperado:e.target.value})}/></div>
        <div><label className="label">Por</label>
          <select className="select" value={form.rendimiento_por} onChange={e=>setForm({...form,rendimiento_por:e.target.value})}>
            <option value="cuadrilla">cuadrilla/día</option><option value="persona">persona/día</option><option value="equipo">equipo/día</option>
          </select></div>
      </div>
      <button className="btn-primary text-xs" disabled={busy} onClick={save}>+ Guardar configuración</button>
      {configActs.length>0&&<div className="table-wrap mt-3"><table className="table">
        <thead><tr><th>Actividad ID</th><th>Tipo</th><th>Unidad</th><th>Meta</th><th>Rendimiento</th></tr></thead>
        <tbody>{configActs.map((c,i)=><tr key={i}><td className="font-medium text-xs">{c.actividad_id.slice(-8)}</td><td><span className={`badge ${c.tipo==='A'?'bg-blue-100 text-blue-800':c.tipo==='B'?'bg-green-100 text-green-800':'bg-purple-100 text-purple-800'}`}>{c.tipo}</span></td><td>{c.unidad_es}</td><td>{c.meta_total||'—'}</td><td>{c.rendimiento_esperado?`${c.rendimiento_esperado}/${c.rendimiento_por}`:'—'}</td></tr>)}
        </tbody>
      </table></div>}
    </div>
  );
}

/* ═══════════════════════════════════════════════════════════════
 *  GESTIÓN DE USUARIOS (Admin)
 * ═══════════════════════════════════════════════════════════════ */
function UsuariosModule({showToast}:{showToast:any}){
  const[users,setUsers]=useState<Profile[]>([]);
  const[form,setForm]=useState({nombre:'',correo:'',clave:'',rol:'tecnico' as UserRole,especialidad_id:''});
  const[busy,setBusy]=useState(false);
  const[loading,setLoading]=useState(true);

  async function loadUsers(){
    setLoading(true);
    const{data}=await supabase.from('profiles').select('*').order('nombre');
    setUsers((data||[]) as Profile[]);
    setLoading(false);
  }
  useEffect(()=>{loadUsers();},[]);

  async function createUser(){
    if(!form.nombre||!form.correo||!form.clave){showToast('err','Nombre, correo y clave son obligatorios');return;}
    setBusy(true);
    try{
      const r=await fetch('/api/admin',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({action:'createUser',...form})});
      const d=await r.json();
      if(!d.ok) throw new Error(d.error||'Error');
      showToast('ok','Usuario creado'); setForm({nombre:'',correo:'',clave:'',rol:'tecnico',especialidad_id:''}); loadUsers();
    }catch(e:any){showToast('err',e?.message||'Error');}
    finally{setBusy(false);}
  }

  async function toggleActivo(u:Profile){
    await supabase.from('profiles').update({activo:!u.activo}).eq('id',u.id);
    showToast('ok',u.activo?'Usuario desactivado':'Usuario activado'); loadUsers();
  }

  return(
    <div className="card p-4 space-y-4">
      <h3 className="font-bold text-[#003b7a]">Usuarios ({users.length})</h3>
      <div className="grid grid-cols-2 sm:grid-cols-3 gap-2">
        <div><label className="label">Nombre</label><input className="input" value={form.nombre} onChange={e=>setForm({...form,nombre:e.target.value})}/></div>
        <div><label className="label">Correo</label><input className="input" type="email" value={form.correo} onChange={e=>setForm({...form,correo:e.target.value})}/></div>
        <div><label className="label">Clave inicial</label><input className="input" type="password" value={form.clave} onChange={e=>setForm({...form,clave:e.target.value})}/></div>
        <div><label className="label">Rol</label>
          <select className="select" value={form.rol} onChange={e=>setForm({...form,rol:e.target.value as UserRole})}>
            <option value="admin">Admin</option><option value="lider">Líder</option><option value="tecnico">Técnico</option>
            <option value="gerencia">Gerencia</option><option value="cliente">Cliente</option><option value="visualizador">Visualizador</option>
          </select></div>
      </div>
      <button className="btn-primary text-xs" disabled={busy} onClick={createUser}>{busy?'Creando…':'+ Crear usuario'}</button>
      {loading?<div className="text-sm text-slate-500">Cargando…</div>:(
        <div className="table-wrap"><table className="table">
          <thead><tr><th>Nombre</th><th>Correo</th><th>Rol</th><th>Estado</th><th>Acción</th></tr></thead>
          <tbody>{users.map(u=>(
            <tr key={u.id}>
              <td className="font-medium">{u.nombre}</td>
              <td className="text-xs text-slate-500">{u.correo}</td>
              <td><span className="badge bg-slate-100 text-slate-700 uppercase">{u.rol}</span></td>
              <td>{u.activo?<span className="text-emerald-600 text-xs">✓ Activo</span>:<span className="text-rose-500 text-xs">✗ Inactivo</span>}</td>
              <td><button className="text-xs underline text-slate-500 hover:text-rose-600" onClick={()=>toggleActivo(u)}>{u.activo?'Desactivar':'Activar'}</button></td>
            </tr>
          ))}</tbody>
        </table></div>
      )}
    </div>
  );
}

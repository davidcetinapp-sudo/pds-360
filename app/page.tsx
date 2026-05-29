'use client';
// Powerchina PDS 360 v2.1
import { useCallback, useEffect, useMemo, useState } from 'react';
import { supabase, type Profile, type UserRole } from '@/lib/supabase';
import * as XLSX from 'xlsx';

type AppView = 'home'|'planear'|'reporte'|'aprobacion'|'solicitudes'|'dashboard'|'informes'|'catalogos'|'maquinaria'|'config_act'|'usuarios';
interface EspAct { id:string; especialidad_es:string; especialidad_en:string; actividad_es:string; actividad_en:string; activo?:boolean; }
interface Area { id:string; area_es:string; area_en:string; activo?:boolean; }
interface Lider { id:string; nombre:string; documento:string; cargo_es:string; activo?:boolean; }
interface Personal { id:string; nombre:string; documento:string; cargo_es:string; tipo?:string; empresa?:string; activo?:boolean; }
interface Maq { id:string; tipo:string; item_id:string; nombre:string; estado:string; horas_acum_operativas?:number; horas_acum_standby?:number; }
interface ConfigAct { id:string; actividad_id:string; especialidad_id:string; tipo:'A'|'B'|'C'; unidad_es:string; unidad_en?:string; meta_total?:number; tiene_meta?:boolean; es_medible?:boolean; rendimiento_esperado?:number; rendimiento_por?:string; acumulado_previo?:number; }
interface Catalogs { especialidades_actividades:EspAct[]; areas:Area[]; lideres:Lider[]; personal:Personal[]; }
interface Notif { id:string; titulo:string; mensaje:string; leida:boolean; created_at:string; }
interface SuspItem { uid:string; hora_inicio:string; hora_fin:string; descripcion:string; }
interface AsistItem { personal_id:string; documento_personal:string; nombre:string; cargo_es:string; asistio:boolean; motivo_ausencia:string; }
interface ActForm { uid:string; especialidad_id:string; actividad_id:string; area_id:string; areas_adicionales:string[]; lider_id:string; maquinaria_ids:string[]; rendimiento_esperado:string; observacion_es:string; observacion_en:string; personal:{personal_id:string;documento_personal:string}[]; }
interface AreaRep { uid:string; area_id:string; cantidad:string; }
interface ActRep { uid:string; actividad_id:string; areas:AreaRep[]; observacion_es:string; }

function today():string { const d=new Date(); return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`; }
function gid():string { return Math.random().toString(36).slice(2)+Date.now().toString(36); }
function sem(p:number):string { return p>=70?'🟢':p>=50?'🟡':'🔴'; }
function uniqueEsp(rows:EspAct[]):EspAct[] { const s=new Set<string>(); return rows.filter(r=>{ const k=(r.especialidad_es||'').toLowerCase(); if(s.has(k))return false; s.add(k); return true; }); }
function actsForEsp(rows:EspAct[],espId:string):EspAct[] { const e=rows.find(r=>r.id===espId); if(!e)return[]; const t=(e.especialidad_es||'').toLowerCase(); return rows.filter(r=>(r.especialidad_es||'').toLowerCase()===t); }
function horasLost(ss:SuspItem[]):number { return ss.reduce((a,s)=>{ if(!s.hora_inicio||!s.hora_fin)return a; const[ih,im]=s.hora_inicio.split(':').map(Number); const[fh,fm]=s.hora_fin.split(':').map(Number); return a+Math.max(0,(fh+fm/60)-(ih+im/60)); },0); }

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
      try{
        const{data:{session}}=await supabase.auth.getSession();
        if(session?.user){
          const{data:p}=await supabase.from('profiles').select('*').eq('id',session.user.id).single();
          if(p)setProfile(p as Profile);
        }
      }catch(e){ console.error('Auth error:',e); }
      finally{ setLoading(false); }
    })();
    const{data:{subscription}}=supabase.auth.onAuthStateChange((_,session)=>{ if(!session){setProfile(null);setView('home');} });
    return()=>subscription.unsubscribe();
  },[]);

  useEffect(()=>{ if(toast){const t=setTimeout(()=>setToast(null),4500);return()=>clearTimeout(t);} },[toast]);
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
    }catch(e){ console.error('Error cargando catálogos:',e); }
  },[]);

  const loadNotifs=useCallback(async(uid:string)=>{
    try{
      const{data}=await supabase.from('notificaciones').select('*').eq('usuario_id',uid).eq('leida',false).order('created_at',{ascending:false}).limit(20);
      setNotifs((data||[]) as Notif[]);
    }catch{ setNotifs([]); }
  },[]);

  useEffect(()=>{ if(profile){loadCatalogs();loadNotifs(profile.id);} },[profile,loadCatalogs,loadNotifs]);

  async function marcarLeidas(){ if(!profile)return; try{await supabase.from('notificaciones').update({leida:true}).eq('usuario_id',profile.id).eq('leida',false);}catch{} setNotifs([]); }
  async function handleLogout(){ await supabase.auth.signOut(); setProfile(null);setCatalogs(null);setView('home'); }

  if(loading)return(
    <div className="min-h-screen flex flex-col items-center justify-center gap-3 text-slate-500">
      <div className="w-8 h-8 border-4 border-[#003b7a] border-t-transparent rounded-full animate-spin"/>
      <span className="text-sm">Iniciando PDS 360…</span>
    </div>
  );
  if(!profile)return <LoginScreen onLogin={setProfile} showToast={showToast} toast={toast}/>;

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
        {view==='dashboard'  &&<DashboardModule showToast={showToast}/>}
        {view==='informes'   &&<InformesModule user={u} catalogs={catalogs} showToast={showToast}/>}
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

function Header({user,onLogout,setView,currentView,notifs,onReadNotifs}:{user:Profile;onLogout:()=>void;setView:(v:AppView)=>void;currentView:AppView;notifs:Notif[];onReadNotifs:()=>void}) {
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
                  notifs.map(n=><div key={n.id} className="p-3 border-b hover:bg-slate-50"><div className="font-medium text-sm text-slate-800">{n.titulo}</div><div className="text-xs text-slate-500 mt-0.5">{n.mensaje}</div><div className="text-xs text-slate-400 mt-1">{new Date(n.created_at).toLocaleString('es-CO')}</div></div>)
                }
              </div>
            )}
          </div>
          <div className="hidden sm:block text-right"><div className="font-medium text-sm">{user.nombre}</div><div className="text-xs text-blue-200 uppercase">{user.rol}</div></div>
          <button className="btn-secondary text-xs" onClick={onLogout}>Salir</button>
        </div>
      </div>
      <nav className="max-w-7xl mx-auto px-2 overflow-x-auto">
        <div className="flex gap-0.5">{tabs.filter(t=>t.show).map(t=><button key={t.key} onClick={()=>setView(t.key)} className={`px-3 py-2 text-xs font-medium border-b-2 whitespace-nowrap transition-colors ${currentView===t.key?'border-white text-white':'border-transparent text-blue-200 hover:text-white'}`}>{t.label}</button>)}</div>
      </nav>
    </header>
  );
}

function Toast({toast}:{toast:{k:string;m:string}|null}) {
  if(!toast)return null;
  const c=toast.k==='ok'?'bg-emerald-600':toast.k==='err'?'bg-rose-600':'bg-slate-700';
  return <div className={`fixed bottom-4 left-4 right-4 sm:left-auto sm:right-4 sm:max-w-sm ${c} text-white px-4 py-3 rounded-lg shadow-lg z-50 text-sm no-print`}>{toast.m}</div>;
}

function LoginScreen({onLogin,showToast,toast}:{onLogin:(p:Profile)=>void;showToast:(k:'ok'|'err'|'info',m:string)=>void;toast:any}) {
  const[correo,setCorreo]=useState('');
  const[clave,setClave]=useState('');
  const[loading,setLoading]=useState(false);
  async function submit(e:React.FormEvent){
    e.preventDefault();
    if(!correo||!clave){showToast('err','Ingresa correo y clave');return;}
    setLoading(true);
    try{
      const{data,error}=await supabase.auth.signInWithPassword({email:correo,password:clave});
      if(error||!data.user)throw new Error(error?.message||'Credenciales inválidas');
      const{data:p}=await supabase.from('profiles').select('*').eq('id',data.user.id).single();
      if(!p)throw new Error('Perfil no encontrado. Contacta al administrador.');
      if(!(p as any).activo)throw new Error('Usuario inactivo. Contacta al administrador.');
      onLogin(p as Profile);
    }catch(e:any){showToast('err',e?.message||'Error de red');}
    finally{setLoading(false);}
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

function HomeScreen({user,setView,notifs}:{user:Profile;setView:(v:AppView)=>void;notifs:Notif[]}) {
  const canEdit=user.rol==='admin'||user.rol==='lider'||user.rol==='tecnico';
  const canApprove=user.rol==='admin'||user.rol==='lider';
  return(
    <div className="space-y-6">
      <div className="text-center py-4">
        <h2 className="text-2xl font-bold text-[#003b7a]">Bienvenido, {user.nombre.split(' ')[0]}</h2>
        <p className="text-slate-500 text-sm mt-1">{today()} · <span className="font-semibold uppercase">{user.rol}</span></p>
      </div>
      {notifs.length>0&&<div className="card p-3 border-l-4 border-amber-400 bg-amber-50"><div className="font-semibold text-amber-800 text-sm mb-1">🔔 {notifs.length} notificaciones pendientes</div>{notifs.slice(0,3).map(n=><div key={n.id} className="text-xs text-amber-700">• {n.titulo}: {n.mensaje}</div>)}</div>}
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
function QBtn({icon,label,onClick}:{icon:string;label:string;onClick:()=>void}) {
  return <button onClick={onClick} className="card p-3 text-center hover:shadow-md transition-shadow"><div className="text-2xl">{icon}</div><div className="text-xs font-medium text-slate-700 mt-1">{label}</div></button>;
}

// ── PLANEACIÓN ────────────────────────────────────────────────────
function PlaneacionModule({user,catalogs,maquinaria,showToast}:{user:Profile;catalogs:Catalogs|null;maquinaria:Maq[];showToast:(k:'ok'|'err'|'info',m:string)=>void}) {
  const[fecha,setFecha]=useState(today());
  const[actividades,setActividades]=useState<ActForm[]>([emptyAct()]);
  const[blocked,setBlocked]=useState<{documento_personal:string;usuario_nombre:string}[]>([]);
  const[estado,setEstado]=useState<'nuevo'|'borrador'|'enviado'>('nuevo');
  const[saving,setSaving]=useState(false);
  const[progId,setProgId]=useState<string|null>(null);

  const loadFecha=useCallback(async()=>{
    if(!fecha)return;
    const{data:bl}=await supabase.from('personal_asignado').select('documento_personal,usuario_id,programaciones!inner(usuario_nombre)').eq('fecha',fecha).neq('usuario_id',user.id);
    setBlocked((bl||[]).map((b:any)=>({documento_personal:b.documento_personal,usuario_nombre:(b.programaciones as any)?.usuario_nombre||'otro'})));
    const{data:prog}=await supabase.from('programaciones').select('*').eq('fecha',fecha).eq('usuario_id',user.id).maybeSingle();
    if(prog){
      setEstado(prog.estado);setProgId(prog.id);
      const{data:acts}=await supabase.from('actividades_programadas').select('*').eq('programacion_id',prog.id);
      if(acts?.length){
        const{data:pa}=await supabase.from('personal_asignado').select('*').eq('programacion_id',prog.id);
        setActividades((acts||[]).map((a:any)=>({uid:gid(),especialidad_id:a.especialidad_id||'',actividad_id:a.actividad_id||'',area_id:a.area_id||'',areas_adicionales:a.areas_adicionales||[],lider_id:a.lider_id||'',maquinaria_ids:a.maquinaria_ids||[],rendimiento_esperado:a.rendimiento_esperado||'',observacion_es:a.observacion_es||'',observacion_en:a.observacion_en||'',personal:(pa||[]).filter((p:any)=>p.actividad_programada_id===a.id).map((p:any)=>({personal_id:p.personal_id,documento_personal:p.documento_personal}))})));
        return;
      }
    }
    setEstado('nuevo');setProgId(null);setActividades([emptyAct()]);
  },[fecha,user.id]);

  useEffect(()=>{loadFecha();},[loadFecha]);
  const blockedMap=useMemo(()=>{const m:Record<string,string>={};blocked.forEach(b=>{m[b.documento_personal]=b.usuario_nombre;});return m;},[blocked]);
  const isRO=estado==='enviado';

  async function save(est:'borrador'|'enviado'){
    for(let i=0;i<actividades.length;i++){const a=actividades[i];if(!a.especialidad_id||!a.actividad_id||!a.area_id||!a.lider_id){showToast('err',`Actividad ${i+1}: completa todos los campos`);return;}if(!a.personal.length){showToast('err',`Actividad ${i+1}: agrega al menos una persona`);return;}}
    if(est==='enviado'&&!window.confirm('¿Enviar planeación?'))return;
    setSaving(true);
    try{
      const{data:prog,error:pe}=await supabase.from('programaciones').upsert({id:progId||undefined,fecha,usuario_id:user.id,usuario_nombre:user.nombre,estado:est,updated_at:new Date().toISOString()},{onConflict:'fecha,usuario_id'}).select().single();
      if(pe||!prog)throw new Error(pe?.message||'Error');
      await supabase.from('actividades_programadas').delete().eq('programacion_id',prog.id);
      for(const act of actividades){
        const{data:aR,error:ae}=await supabase.from('actividades_programadas').insert({programacion_id:prog.id,fecha,usuario_id:user.id,especialidad_id:act.especialidad_id,actividad_id:act.actividad_id,area_id:act.area_id,areas_adicionales:act.areas_adicionales,lider_id:act.lider_id,maquinaria_ids:act.maquinaria_ids,rendimiento_esperado:act.rendimiento_esperado,observacion_es:act.observacion_es,observacion_en:act.observacion_en}).select().single();
        if(ae||!aR)throw new Error(ae?.message||'Error en actividad');
        if(act.personal.length){const{error:paE}=await supabase.from('personal_asignado').insert(act.personal.map(p=>({programacion_id:prog.id,actividad_programada_id:aR.id,fecha,usuario_id:user.id,personal_id:p.personal_id,documento_personal:p.documento_personal})));if(paE?.code==='23505')throw new Error('Personal duplicado');if(paE)throw new Error(paE.message);}
      }
      setProgId(prog.id);setEstado(est);showToast('ok',est==='enviado'?'Planeación enviada ✓':'Borrador guardado ✓');await loadFecha();
    }catch(e:any){showToast('err',e?.message||'Error');}
    finally{setSaving(false);}
  }

  function exportarExcel(){
    const rows=actividades.map((a,i)=>({'N°':i+1,Fecha:fecha,Actividad:a.actividad_id,Área:a.area_id,Líder:a.lider_id,Personal:a.personal.length}));
    const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(rows),'Planeación');XLSX.writeFile(wb,`Planeacion_${fecha}.xlsx`);
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
          <button className="btn-secondary text-xs" onClick={()=>window.print()}>🖨️ PDF</button>
        </div>
      </div>
      {actividades.map((a,idx)=><ActCard key={a.uid} index={idx} act={a} catalogs={catalogs} maquinaria={maquinaria} blockedMap={blockedMap} readOnly={isRO} onChange={p=>setActividades(arr=>arr.map(x=>x.uid===a.uid?{...x,...p}:x))} onRemove={()=>setActividades(arr=>arr.length<=1?arr:arr.filter(x=>x.uid!==a.uid))}/>)}
      <div className="flex flex-wrap gap-2">
        {!isRO&&<button className="btn-secondary" onClick={()=>setActividades(a=>[...a,emptyAct()])}>+ Actividad</button>}
        {!isRO&&<><button className="btn-primary" disabled={saving} onClick={()=>save('borrador')}>{saving?'Guardando…':'Guardar borrador'}</button><button className="btn-success" disabled={saving} onClick={()=>save('enviado')}>{saving?'Enviando…':'Enviar planeación'}</button></>}
        {isRO&&<button className="btn-secondary" disabled={saving} onClick={async()=>{if(!window.confirm('¿Reabrir?'))return;setSaving(true);await supabase.from('programaciones').update({estado:'borrador'}).eq('id',progId!);setEstado('borrador');setSaving(false);}}>Reabrir</button>}
      </div>
    </div>
  );
}

function emptyAct():ActForm{return{uid:gid(),especialidad_id:'',actividad_id:'',area_id:'',areas_adicionales:[],lider_id:'',maquinaria_ids:[],rendimiento_esperado:'',observacion_es:'',observacion_en:'',personal:[]};}

function ActCard({index,act,catalogs,maquinaria,blockedMap,readOnly,onChange,onRemove}:{index:number;act:ActForm;catalogs:Catalogs|null;maquinaria:Maq[];blockedMap:Record<string,string>;readOnly:boolean;onChange:(p:Partial<ActForm>)=>void;onRemove:()=>void;}) {
  const[search,setSearch]=useState('');
  const esps=useMemo(()=>catalogs?uniqueEsp(catalogs.especialidades_actividades.filter(e=>e.activo!==false)):[],[catalogs]);
  const acts=useMemo(()=>catalogs&&act.especialidad_id?actsForEsp(catalogs.especialidades_actividades.filter(e=>e.activo!==false),act.especialidad_id):[],[catalogs,act.especialidad_id]);
  const persActivos=useMemo(()=>(catalogs?.personal||[]).filter(p=>p.activo!==false),[catalogs]);
  const persFilt=useMemo(()=>{const q=search.trim().toLowerCase();return q?persActivos.filter(p=>p.nombre.toLowerCase().includes(q)||p.documento.toLowerCase().includes(q)):persActivos;},[persActivos,search]);
  return(
    <div className="card p-4 space-y-3">
      <div className="flex items-center justify-between"><div className="font-semibold text-[#003b7a]">Actividad {index+1}</div>{!readOnly&&<button className="btn-ghost text-rose-600 text-xs" onClick={onRemove}>✕ Eliminar</button>}</div>
      <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-4 gap-3">
        <div><label className="label">Especialidad</label><select className="select" disabled={readOnly} value={act.especialidad_id} onChange={e=>onChange({especialidad_id:e.target.value,actividad_id:''})}><option value="">— Seleccionar —</option>{esps.map(e=><option key={e.id} value={e.id}>{e.especialidad_es}</option>)}</select></div>
        <div><label className="label">Actividad</label><select className="select" disabled={readOnly||!act.especialidad_id} value={act.actividad_id} onChange={e=>onChange({actividad_id:e.target.value})}><option value="">— Seleccionar —</option>{acts.map(a=><option key={a.id} value={a.id}>{a.actividad_es}</option>)}</select></div>
        <div><label className="label">Área principal</label><select className="select" disabled={readOnly} value={act.area_id} onChange={e=>onChange({area_id:e.target.value})}><option value="">— Seleccionar —</option>{(catalogs?.areas||[]).map(a=><option key={a.id} value={a.id}>{a.area_es}</option>)}</select></div>
        <div><label className="label">Líder</label><select className="select" disabled={readOnly} value={act.lider_id} onChange={e=>onChange({lider_id:e.target.value})}><option value="">— Seleccionar —</option>{(catalogs?.lideres||[]).filter(l=>l.activo!==false).map(l=><option key={l.id} value={l.id}>{l.nombre}</option>)}</select></div>
      </div>
      <div><label className="label">Áreas adicionales</label><div className="flex flex-wrap gap-2">{(catalogs?.areas||[]).filter(a=>a.id!==act.area_id).map(a=>{const sel=act.areas_adicionales.includes(a.id);return<button key={a.id} disabled={readOnly} onClick={()=>onChange({areas_adicionales:sel?act.areas_adicionales.filter(x=>x!==a.id):[...act.areas_adicionales,a.id]})} className={`text-xs px-2 py-1 rounded border transition-colors ${sel?'bg-[#003b7a] text-white border-[#003b7a]':'border-slate-300 text-slate-600 hover:border-[#003b7a]'}`}>{a.area_es}</button>;})}</div></div>
      <div><label className="label">Maquinaria</label><div className="flex flex-wrap gap-2">{maquinaria.filter(m=>m.estado==='activo').map(m=>{const sel=act.maquinaria_ids.includes(m.id);return<button key={m.id} disabled={readOnly} onClick={()=>onChange({maquinaria_ids:sel?act.maquinaria_ids.filter(x=>x!==m.id):[...act.maquinaria_ids,m.id]})} className={`text-xs px-2 py-1 rounded border transition-colors ${sel?'bg-orange-500 text-white border-orange-500':'border-slate-300 text-slate-600 hover:border-orange-400'}`}>{m.item_id}</button>;})}</div></div>
      <div>
        <div className="flex items-center justify-between mb-2"><label className="label !mb-0">Personal ({act.personal.length} sel.)</label>{!readOnly&&act.personal.length>0&&<button className="text-xs text-rose-500 underline" onClick={()=>onChange({personal:[]})}>Limpiar</button>}</div>
        <input className="input mb-2" placeholder="🔎 Buscar…" value={search} onChange={e=>setSearch(e.target.value)}/>
        <div className="border border-slate-200 rounded-md max-h-52 overflow-y-auto">
          {persFilt.map(p=>{
            const sel=!!act.personal.find(x=>x.documento_personal===p.documento);
            const bl=blockedMap[p.documento];
            return(
              <div key={p.id} className="relative group">
                <label className={`flex items-center gap-2 px-3 py-2 border-b border-slate-100 cursor-pointer text-sm ${bl?'bg-rose-50 cursor-not-allowed':sel?'bg-blue-50':'hover:bg-slate-50'}`}>
                  <input type="checkbox" checked={sel} disabled={readOnly||!!bl} onChange={()=>onChange({personal:sel?act.personal.filter(x=>x.documento_personal!==p.documento):[...act.personal,{personal_id:p.id,documento_personal:p.documento}]})}/>
                  <div className="flex-1 min-w-0"><div className={`font-medium truncate ${bl?'text-slate-400':''}`}>{p.nombre}</div><div className="text-xs text-slate-500">{p.documento} · {p.cargo_es}</div></div>
                  {bl&&<span className="text-xs text-rose-500 font-medium">🔒 {bl}</span>}
                  {sel&&!bl&&<span className="badge bg-blue-100 text-blue-800">✓</span>}
                </label>
                {bl&&<div className="absolute left-0 bottom-full mb-1 hidden group-hover:block z-20 bg-slate-800 text-white text-xs rounded p-2 shadow-lg whitespace-nowrap">Asignado por: <strong>{bl}</strong></div>}
              </div>
            );
          })}
        </div>
      </div>
    </div>
  );
}

// ── REPORTE DIARIO ────────────────────────────────────────────────
function ReporteModule({user,catalogs,maquinaria,configActs,showToast}:{user:Profile;catalogs:Catalogs|null;maquinaria:Maq[];configActs:ConfigAct[];showToast:(k:'ok'|'err'|'info',m:string)=>void}) {
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
    if(!fecha||!espId)return;
    if(esDiaAnt){supabase.from('solicitudes_reporte_pasado').select('*').eq('tecnico_id',user.id).eq('fecha_reporte',fecha).eq('estado','aprobado').maybeSingle().then(({data})=>setSolicitudOk(!!data));}
    else{setSolicitudOk(false);}
    supabase.from('personal_asignado').select('personal_id,documento_personal,personal!inner(nombre,cargo_es)').eq('fecha',fecha).then(({data})=>{
      const seen=new Set<string>();
      setAsistencia((data||[]).filter((r:any)=>{if(seen.has(r.documento_personal))return false;seen.add(r.documento_personal);return true;}).map((r:any)=>({personal_id:r.personal_id,documento_personal:r.documento_personal,nombre:(r.personal as any)?.nombre||r.documento_personal,cargo_es:(r.personal as any)?.cargo_es||'',asistio:true,motivo_ausencia:''})));
    });
  },[fecha,espId,user.id,esDiaAnt]);

  useEffect(()=>{setMaqDia(maquinaria.filter(m=>m.estado==='activo').map(m=>({maquinaria_id:m.id,nombre:`${m.item_id} – ${m.tipo}`,novedad:false,descripcion:'',hora_inicio:'',hora_fin:'',horas_standby:0})));},[maquinaria]);

  useEffect(()=>{
    if(!catalogs||!espId)return;
    const espRow=catalogs.especialidades_actividades.find(e=>e.id===espId);
    if(!espRow)return;
    const t=(espRow.especialidad_es||'').toLowerCase();
    const acts=catalogs.especialidades_actividades.filter(e=>(e.especialidad_es||'').toLowerCase()===t&&e.activo!==false);
    setActReps(acts.map(a=>({uid:gid(),actividad_id:a.id,areas:[{uid:gid(),area_id:'',cantidad:''}],observacion_es:''})));
  },[catalogs,espId]);

  async function submit(){
    if(!fecha||!espId){showToast('err','Falta fecha o especialidad');return;}
    if(esDiaAnt&&!solicitudOk){showToast('err','Necesitas aprobación para reportar día anterior. Ve a Solicitudes.');return;}
    setSaving(true);
    try{
      const{data:rep,error:re}=await supabase.from('reportes_avance').insert({fecha,usuario_id:user.id,usuario_nombre:user.nombre,especialidad_id:espId,jornada_horas:jornadaHrs,clima,charla_preturno:charla,charla_tema:charlaTema,estado:'borrador'}).select().single();
      if(re||!rep)throw new Error(re?.message||'Error');
      const rid=rep.id;
      await Promise.all([
        susps.length?supabase.from('suspensiones_clima').insert(susps.map(s=>({reporte_id:rid,fecha,usuario_id:user.id,hora_inicio:s.hora_inicio||null,hora_fin:s.hora_fin||null,horas_perdidas:horasLost([s]),descripcion:s.descripcion}))):null,
        asistencia.length?supabase.from('asistencia_real').insert(asistencia.map(a=>({reporte_id:rid,fecha,usuario_id:user.id,personal_id:a.personal_id,documento_personal:a.documento_personal,asistio:a.asistio,motivo_ausencia:a.motivo_ausencia||null,horas_trabajadas:a.asistio?horasReal:0}))):null,
        maqDia.filter(m=>m.novedad).length?supabase.from('novedades_maquinaria').insert(maqDia.filter(m=>m.novedad).map(m=>({reporte_id:rid,fecha,usuario_id:user.id,maquinaria_id:m.maquinaria_id,descripcion:m.descripcion,hora_inicio:m.hora_inicio||null,hora_fin:m.hora_fin||null,horas_standby:m.horas_standby}))):null,
        incidente.tipo!=='sin_novedad'?supabase.from('incidentes_seg').insert({reporte_id:rid,fecha,usuario_id:user.id,tipo:incidente.tipo,descripcion:incidente.descripcion,medidas_tomadas:incidente.medidas}):null,
        notaBit?supabase.from('bitacora_decisiones').insert({fecha,usuario_id:user.id,descripcion:notaBit,especialidad_id:espId}):null,
      ]);
      for(const ar of actReps){
        const cfg=configActs.find(c=>c.actividad_id===ar.actividad_id);
        if(cfg?.es_medible===false)continue;
        for(const area of ar.areas.filter(a=>a.area_id&&parseFloat(a.cantidad)>0)){
          const{data:prev}=await supabase.from('avance_diario').select('cantidad').eq('actividad_id',ar.actividad_id).eq('area_id',area.area_id).eq('usuario_id',user.id);
          const acumPrev=(prev||[]).reduce((s:number,r:any)=>s+parseFloat(r.cantidad||0),0);
          const cantidad=parseFloat(area.cantidad);
          await supabase.from('avance_diario').insert({reporte_id:rid,fecha,usuario_id:user.id,actividad_id:ar.actividad_id,especialidad_id:espId,area_id:area.area_id,cantidad,unidad:cfg?.unidad_es||'',acumulado_anterior:acumPrev,acumulado_total:acumPrev+cantidad,observacion_es:ar.observacion_es});
        }
      }
      try{const{data:la}=await supabase.from('profiles').select('id').in('rol',['admin','lider']);if(la?.length)await supabase.from('notificaciones').insert(la.map((x:any)=>({usuario_id:x.id,tipo:'reporte_enviado',titulo:'Nuevo reporte por aprobar',mensaje:`${user.nombre} envió su reporte del ${fecha}`,data:{reporte_id:rid}})));}catch{}
      showToast('ok','Reporte enviado ✓ Espera aprobación.');setStep(1);
    }catch(e:any){showToast('err',e?.message||'Error');}
    finally{setSaving(false);}
  }

  const asistio=asistencia.filter(a=>a.asistio).length;
  const efic=asistencia.length>0?Math.round((asistio/asistencia.length)*100):100;

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
      <div className="card p-3"><div className="flex items-center gap-1 overflow-x-auto">{STEPS.map((s,i)=>{const n=i+1;return(<div key={n} className="flex items-center gap-1 flex-shrink-0"><button onClick={()=>setStep(n)} className={n<step?'step-done':n===step?'step-active':'step-pending'}>{n<step?'✓':n}</button><span className={`text-xs whitespace-nowrap hidden sm:inline ${n===step?'font-semibold text-[#003b7a]':'text-slate-500'}`}>{s}</span>{i<STEPS.length-1&&<span className="text-slate-300 mx-1">›</span>}</div>);})}</div></div>
      {step===1&&<div className="card p-4 space-y-3"><h3 className="font-bold text-[#003b7a]">Paso 1 — Encabezado</h3><div className="grid grid-cols-1 sm:grid-cols-3 gap-3"><div><label className="label">Fecha</label><input type="date" className="input" value={fecha} onChange={e=>setFecha(e.target.value)}/></div><div><label className="label">Especialidad</label><select className="select" value={espId} onChange={e=>setEspId(e.target.value)}><option value="">— Seleccionar —</option>{catalogs&&uniqueEsp(catalogs.especialidades_actividades.filter(e=>e.activo!==false)).map(e=><option key={e.id} value={e.id}>{e.especialidad_es}</option>)}</select></div><div><label className="label">Horas de jornada</label><input type="number" className="input" min={1} max={24} value={jornadaHrs} onChange={e=>setJornadaHrs(parseFloat(e.target.value)||9)}/></div></div><button className="btn-primary" onClick={()=>{if(!fecha||!espId){showToast('err','Completa fecha y especialidad');return;}setStep(2);}}>Siguiente →</button></div>}
      {step===2&&<div className="card p-4 space-y-3"><h3 className="font-bold text-[#003b7a]">Paso 2 — Condiciones del día</h3><div className="flex gap-2 flex-wrap">{['despejado','nublado','lluvia','tormenta','suspendido'].map(c=><button key={c} onClick={()=>setClima(c)} className={`text-sm px-3 py-1.5 rounded-full border transition-colors ${clima===c?'bg-[#003b7a] text-white border-[#003b7a]':'border-slate-300 text-slate-600'}`}>{c==='despejado'?'☀️ Despejado':c==='nublado'?'☁️ Nublado':c==='lluvia'?'🌧️ Lluvia':c==='tormenta'?'⛈️ Tormenta':'🚫 Suspendido'}</button>)}</div>{susps.map((s,i)=><div key={s.uid} className="flex gap-2 items-center mt-2 flex-wrap"><input type="time" className="input w-32" value={s.hora_inicio} onChange={e=>setSusps(a=>a.map((x,j)=>j===i?{...x,hora_inicio:e.target.value}:x))}/><span className="text-sm text-slate-500">a</span><input type="time" className="input w-32" value={s.hora_fin} onChange={e=>setSusps(a=>a.map((x,j)=>j===i?{...x,hora_fin:e.target.value}:x))}/><input className="input flex-1" value={s.descripcion} onChange={e=>setSusps(a=>a.map((x,j)=>j===i?{...x,descripcion:e.target.value}:x))} placeholder="Descripción"/><button className="btn-ghost text-rose-500 text-xs" onClick={()=>setSusps(a=>a.filter(x=>x.uid!==s.uid))}>✕</button></div>)}<button className="btn-secondary text-xs mt-2" onClick={()=>setSusps(a=>[...a,{uid:gid(),hora_inicio:'',hora_fin:'',descripcion:''}])}>+ Agregar suspensión</button>{horasClima>0&&<p className="text-sm text-amber-600 mt-2">⏱ {horasClima.toFixed(1)}h perdidas → {horasReal.toFixed(1)}h operativas</p>}<div className="flex gap-2"><button className="btn-secondary" onClick={()=>setStep(1)}>← Anterior</button><button className="btn-primary" onClick={()=>setStep(3)}>Siguiente →</button></div></div>}
      {step===3&&<div className="card p-4 space-y-3"><h3 className="font-bold text-[#003b7a]">Paso 3 — Charla preturno</h3><label className="flex items-center gap-3 cursor-pointer"><input type="checkbox" checked={charla} onChange={e=>setCharla(e.target.checked)} className="w-5 h-5"/><span className="text-sm font-medium">Se realizó la charla preturno</span></label>{charla&&<div><label className="label">Tema</label><input className="input" value={charlaTema} onChange={e=>setCharlaTema(e.target.value)} placeholder="Tema de la charla…"/></div>}<div className="flex gap-2"><button className="btn-secondary" onClick={()=>setStep(2)}>← Anterior</button><button className="btn-primary" onClick={()=>setStep(4)}>Siguiente →</button></div></div>}
      {step===4&&<div className="card p-4 space-y-3"><h3 className="font-bold text-[#003b7a]">Paso 4 — Asistencia real</h3><div className="flex gap-4 text-sm flex-wrap"><span>Planeado: <strong>{asistencia.length}</strong></span><span className="text-emerald-600">Asistió: <strong>{asistio}</strong></span><span>Cumplimiento: <strong>{efic}%</strong> {sem(efic)}</span></div><div className="space-y-2">{asistencia.map((a,i)=><div key={a.personal_id} className="flex items-center gap-3 p-2 border border-slate-200 rounded-lg flex-wrap"><input type="checkbox" checked={a.asistio} onChange={e=>setAsistencia(arr=>arr.map((x,j)=>j===i?{...x,asistio:e.target.checked,motivo_ausencia:''}:x))}/><div className="flex-1"><div className={`text-sm font-medium ${a.asistio?'':'text-slate-400 line-through'}`}>{a.nombre}</div><div className="text-xs text-slate-400">{a.cargo_es}</div></div>{!a.asistio&&<select className="select text-xs w-auto" value={a.motivo_ausencia} onChange={e=>setAsistencia(arr=>arr.map((x,j)=>j===i?{...x,motivo_ausencia:e.target.value}:x))}><option value="">— Motivo —</option><option value="injustificada">Injustificada</option><option value="incapacidad">Incapacidad</option><option value="permiso">Permiso</option></select>}{a.asistio&&<span className="text-xs text-emerald-600">{horasReal.toFixed(1)}h</span>}</div>)}</div>{!asistencia.length&&<p className="text-sm text-slate-500">Sin personal planeado para este día.</p>}<div className="flex gap-2"><button className="btn-secondary" onClick={()=>setStep(3)}>← Anterior</button><button className="btn-primary" onClick={()=>setStep(5)}>Siguiente →</button></div></div>}
      {step===5&&<div className="card p-4 space-y-3"><h3 className="font-bold text-[#003b7a]">Paso 5 — Maquinaria</h3>{maqDia.map((m,i)=><div key={m.maquinaria_id} className="border border-slate-200 rounded-lg p-3"><div className="flex items-center justify-between mb-2"><span className="font-medium text-sm">{m.nombre}</span><span className="text-xs text-slate-500">Op: {Math.max(0,horasReal-m.horas_standby).toFixed(1)}h | SB: {m.horas_standby}h</span></div><label className="flex items-center gap-2 text-sm mb-2"><input type="checkbox" checked={m.novedad} onChange={e=>setMaqDia(a=>a.map((x,j)=>j===i?{...x,novedad:e.target.checked}:x))}/> ¿Tuvo novedad?</label>{m.novedad&&<div className="grid grid-cols-1 sm:grid-cols-3 gap-2"><div><label className="label">Hora inicio</label><input type="time" className="input" value={m.hora_inicio} onChange={e=>setMaqDia(a=>a.map((x,j)=>j===i?{...x,hora_inicio:e.target.value}:x))}/></div><div><label className="label">Hora fin</label><input type="time" className="input" value={m.hora_fin} onChange={e=>{const[ih,im]=(m.hora_inicio||'0:0').split(':').map(Number);const[fh,fm]=e.target.value.split(':').map(Number);const diff=Math.max(0,(fh+fm/60)-(ih+im/60));setMaqDia(a=>a.map((x,j)=>j===i?{...x,hora_fin:e.target.value,horas_standby:parseFloat(diff.toFixed(2))}:x));}}/></div><div><label className="label">Descripción</label><input className="input" value={m.descripcion} onChange={e=>setMaqDia(a=>a.map((x,j)=>j===i?{...x,descripcion:e.target.value}:x))}/></div></div>}</div>)}{!maqDia.length&&<p className="text-sm text-slate-500">Sin maquinaria activa.</p>}<div className="flex gap-2"><button className="btn-secondary" onClick={()=>setStep(4)}>← Anterior</button><button className="btn-primary" onClick={()=>setStep(6)}>Siguiente →</button></div></div>}
      {step===6&&<div className="card p-4 space-y-4"><h3 className="font-bold text-[#003b7a]">Paso 6 — Avance de actividades</h3><p className="text-xs text-slate-500 bg-blue-50 border border-blue-200 rounded p-2">💡 Puedes reportar múltiples áreas por actividad. El personal no se bloquea entre técnicos en el reporte diario.</p>
        {actReps.map((ar,ai)=>{
          const actRow=catalogs?.especialidades_actividades.find(e=>e.id===ar.actividad_id);
          const cfg=configActs.find(c=>c.actividad_id===ar.actividad_id);
          if(cfg?.es_medible===false)return<div key={ar.uid} className="border border-slate-200 rounded-lg p-3 bg-slate-50"><div className="text-sm text-slate-500">{actRow?.actividad_es} — <em>Actividad no medible</em></div></div>;
          if(!cfg)return<div key={ar.uid} className="border border-amber-200 rounded-lg p-3 bg-amber-50"><div className="text-sm text-amber-700">⚠️ {actRow?.actividad_es} — El admin debe configurar esta actividad en <strong>Config. Act.</strong> primero.</div></div>;
          return(
            <div key={ar.uid} className="border border-slate-200 rounded-lg p-3 space-y-3">
              <div className="font-medium text-sm text-[#003b7a]">{actRow?.actividad_es}<span className="ml-2 text-xs text-slate-400">Tipo {cfg.tipo} · {cfg.unidad_es}{cfg.tiene_meta&&cfg.meta_total?` · Meta: ${cfg.meta_total}`:''}</span></div>
              {ar.areas.map((area,areai)=>(
                <div key={area.uid} className="flex gap-2 items-end flex-wrap bg-slate-50 p-2 rounded border border-slate-100">
                  <div className="flex-1 min-w-[120px]"><label className="label">Área</label><select className="select" value={area.area_id} onChange={e=>setActReps(arr=>arr.map((x,j)=>j===ai?{...x,areas:x.areas.map((a2,k)=>k===areai?{...a2,area_id:e.target.value}:a2)}:x))}><option value="">— Área —</option>{(catalogs?.areas||[]).map(a=><option key={a.id} value={a.id}>{a.area_es}</option>)}</select></div>
                  <div className="w-32"><label className="label">Cantidad</label><input type="number" className="input" min={0} value={area.cantidad} onChange={e=>setActReps(arr=>arr.map((x,j)=>j===ai?{...x,areas:x.areas.map((a2,k)=>k===areai?{...a2,cantidad:e.target.value}:a2)}:x))}/></div>
                  {ar.areas.length>1&&<button className="btn-ghost text-rose-500 text-xs pb-2" onClick={()=>setActReps(arr=>arr.map((x,j)=>j===ai?{...x,areas:x.areas.filter((_,k)=>k!==areai)}:x))}>✕</button>}
                </div>
              ))}
              <button className="btn-secondary text-xs" onClick={()=>setActReps(arr=>arr.map((x,j)=>j===ai?{...x,areas:[...x.areas,{uid:gid(),area_id:'',cantidad:''}]}:x))}>+ Otra área</button>
              <div><label className="label">Observación</label><textarea className="textarea" rows={1} value={ar.observacion_es} onChange={e=>setActReps(arr=>arr.map((x,j)=>j===ai?{...x,observacion_es:e.target.value}:x))}/></div>
            </div>
          );
        })}
        {!actReps.length&&<p className="text-sm text-slate-500">Selecciona una especialidad en el Paso 1.</p>}
        <div className="flex gap-2"><button className="btn-secondary" onClick={()=>setStep(5)}>← Anterior</button><button className="btn-primary" onClick={()=>setStep(7)}>Siguiente →</button></div>
      </div>}
      {step===7&&<div className="card p-4 space-y-3"><h3 className="font-bold text-[#003b7a]">Paso 7 — Seguridad</h3><div className="flex gap-2 flex-wrap">{['sin_novedad','casi_accidente','incidente','accidente'].map(t=><button key={t} onClick={()=>setIncidente(i=>({...i,tipo:t}))} className={`text-xs px-3 py-1.5 rounded-full border transition-colors ${incidente.tipo===t?t==='sin_novedad'?'bg-emerald-500 text-white border-emerald-500':t==='accidente'?'bg-rose-600 text-white border-rose-600':'bg-amber-500 text-white border-amber-500':'border-slate-300 text-slate-600'}`}>{t==='sin_novedad'?'✅ Sin novedad':t==='casi_accidente'?'⚠️ Casi accidente':t==='incidente'?'🔶 Incidente':'🚨 Accidente'}</button>)}</div>{incidente.tipo!=='sin_novedad'&&<><div><label className="label">Descripción</label><textarea className="textarea" rows={2} value={incidente.descripcion} onChange={e=>setIncidente(i=>({...i,descripcion:e.target.value}))}/></div><div><label className="label">Medidas tomadas</label><textarea className="textarea" rows={2} value={incidente.medidas} onChange={e=>setIncidente(i=>({...i,medidas:e.target.value}))}/></div></>}<div><label className="label">Nota de bitácora (opcional)</label><textarea className="textarea" rows={2} value={notaBit} onChange={e=>setNotaBit(e.target.value)} placeholder="Decisiones importantes del día…"/></div><div className="flex gap-2"><button className="btn-secondary" onClick={()=>setStep(6)}>← Anterior</button><button className="btn-primary" onClick={()=>setStep(8)}>Siguiente →</button></div></div>}
      {step===8&&<div className="card p-4 space-y-4"><h3 className="font-bold text-[#003b7a]">Paso 8 — Resumen y envío</h3><div className="grid grid-cols-2 sm:grid-cols-4 gap-3"><div className="card p-3 text-center"><div className="text-2xl font-bold text-[#003b7a]">{asistio}</div><div className="text-xs text-slate-500">Personal activo</div></div><div className="card p-3 text-center"><div className="text-2xl font-bold text-emerald-600">{(asistio*horasReal).toFixed(0)}h</div><div className="text-xs text-slate-500">Horas-hombre</div></div><div className="card p-3 text-center"><div className="text-2xl font-bold text-amber-500">{horasClima.toFixed(1)}h</div><div className="text-xs text-slate-500">Perdidas clima</div></div><div className="card p-3 text-center"><div className="text-2xl font-bold text-[#003b7a]">{actReps.filter(a=>a.areas.some(ar=>parseFloat(ar.cantidad)>0)).length}</div><div className="text-xs text-slate-500">Actividades</div></div></div><div className="bg-slate-50 rounded-lg p-3 text-sm space-y-1"><div><strong>Fecha:</strong> {fecha}</div><div><strong>Clima:</strong> {clima}</div><div><strong>Asistencia:</strong> {asistio}/{asistencia.length} ({efic}%) {sem(efic)}</div><div><strong>Incidente:</strong> {incidente.tipo}</div></div><p className="text-xs text-slate-500">Quedará en <strong>Borrador</strong> hasta que el líder o admin lo apruebe.</p><div className="flex gap-2"><button className="btn-secondary" onClick={()=>setStep(7)}>← Anterior</button><button className="btn-success" disabled={saving} onClick={submit}>{saving?'Enviando…':'✅ Enviar reporte'}</button></div></div>}
    </div>
  );
}

// ── APROBACIÓN ────────────────────────────────────────────────────
function AprobacionModule({user,catalogs,configActs,showToast,onRefreshNotifs}:{user:Profile;catalogs:Catalogs|null;configActs:ConfigAct[];showToast:(k:'ok'|'err'|'info',m:string)=>void;onRefreshNotifs:()=>void}) {
  const[fecha,setFecha]=useState(today());
  const[espId,setEspId]=useState('');
  const[reportes,setReportes]=useState<any[]>([]);
  const[avances,setAvances]=useState<any[]>([]);
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
      const repIds=(reps||[]).map((r:any)=>r.id);
      const avs=repIds.length?(await supabase.from('avance_diario').select('*').in('reporte_id',repIds)).data||[]:[];
      setReportes(reps||[]);setAvances(avs);setAprobadas(new Set());setRechazadas(new Set());setMotivos({});
    }catch(e:any){showToast('err',e?.message||'Error');}
    finally{setLoading(false);}
  }

  const actividadesConReps=useMemo(()=>{
    const acts=new Map<string,{actividad_id:string;cuadrillas:{reporte_id:string;usuario_nombre:string;areas:{area_id:string;cantidad:number;unidad:string}[];total:number}[]}>();
    avances.forEach((av:any)=>{
      if(!acts.has(av.actividad_id))acts.set(av.actividad_id,{actividad_id:av.actividad_id,cuadrillas:[]});
      const entry=acts.get(av.actividad_id)!;
      const rep=reportes.find(r=>r.id===av.reporte_id);
      let cuad=entry.cuadrillas.find(c=>c.reporte_id===av.reporte_id);
      if(!cuad){cuad={reporte_id:av.reporte_id,usuario_nombre:rep?.usuario_nombre||'Técnico',areas:[],total:0};entry.cuadrillas.push(cuad);}
      cuad.areas.push({area_id:av.area_id,cantidad:parseFloat(av.cantidad||0),unidad:av.unidad});
      cuad.total+=parseFloat(av.cantidad||0);
    });
    reportes.forEach(r=>{if(!avances.some((av:any)=>av.reporte_id===r.id)){if(!acts.has('sin_avance'))acts.set('sin_avance',{actividad_id:'',cuadrillas:[]});acts.get('sin_avance')!.cuadrillas.push({reporte_id:r.id,usuario_nombre:r.usuario_nombre,areas:[],total:0});}});
    return Array.from(acts.values());
  },[avances,reportes]);

  const consolidado=useMemo(()=>{const c:Record<string,number>={};actividadesConReps.forEach(act=>{act.cuadrillas.filter(cu=>aprobadas.has(cu.reporte_id)).forEach(cu=>{c[act.actividad_id]=(c[act.actividad_id]||0)+cu.total;});});return c;},[actividadesConReps,aprobadas]);

  function toggleAp(id:string){setAprobadas(p=>{const n=new Set(p);if(n.has(id))n.delete(id);else{n.add(id);setRechazadas(r=>{const nr=new Set(r);nr.delete(id);return nr;});}return n;});}
  function toggleRe(id:string){setRechazadas(p=>{const n=new Set(p);if(n.has(id))n.delete(id);else{n.add(id);setAprobadas(a=>{const na=new Set(a);na.delete(id);return na;});}return n;});}

  async function enviar(){
    const toAp=Array.from(aprobadas);const toRe=Array.from(rechazadas);
    if(!toAp.length&&!toRe.length){showToast('err','Sin decisiones');return;}
    for(const id of toRe){if(!motivos[id]){showToast('err','Indica el motivo de rechazo para todos los rechazados');return;}}
    if(!window.confirm(`¿Confirmar? ${toAp.length} aprobados · ${toRe.length} rechazados`))return;
    setSaving(true);
    try{
      if(toAp.length){
        await supabase.from('reportes_avance').update({estado:'aprobado',aprobado_por:user.id,aprobado_en:new Date().toISOString()}).in('id',toAp);
        await supabase.from('aprobacion_informes').insert(toAp.map(id=>({reporte_id:id,aprobado_por:user.id,estado:'aprobado',version:1})));
        const rA=reportes.filter(r=>toAp.includes(r.id));
        if(rA.length)try{await supabase.from('notificaciones').insert(rA.map(r=>({usuario_id:r.usuario_id,tipo:'aprobado',titulo:'Reporte aprobado ✅',mensaje:`Tu reporte del ${fecha} fue aprobado por ${user.nombre}`,data:{reporte_id:r.id}})));}catch{}
      }
      if(toRe.length){
        await supabase.from('aprobacion_informes').insert(toRe.map(id=>({reporte_id:id,aprobado_por:user.id,estado:'rechazado',version:1,comentarios:motivos[id]||''})));
        const rR=reportes.filter(r=>toRe.includes(r.id));
        if(rR.length)try{await supabase.from('notificaciones').insert(rR.map(r=>({usuario_id:r.usuario_id,tipo:'rechazado',titulo:'Reporte rechazado ❌',mensaje:`Tu reporte del ${fecha} fue rechazado. Motivo: ${motivos[r.id]||'Ver detalles'}`,data:{reporte_id:r.id}})));}catch{}
      }
      showToast('ok',`✅ ${toAp.length} aprobados · ❌ ${toRe.length} rechazados`);onRefreshNotifs();await cargar();
    }catch(e:any){showToast('err',e?.message||'Error');}
    finally{setSaving(false);}
  }

  const areaMap=useMemo(()=>{const m:Record<string,string>={};(catalogs?.areas||[]).forEach(a=>{m[a.id]=a.area_es;});return m;},[catalogs]);

  return(
    <div className="space-y-4">
      <div className="card p-4 flex gap-3 items-end flex-wrap">
        <div className="flex-1 min-w-[140px]"><label className="label">Fecha</label><input type="date" className="input" value={fecha} onChange={e=>setFecha(e.target.value)}/></div>
        <div className="flex-1 min-w-[180px]"><label className="label">Especialidad</label><select className="select" value={espId} onChange={e=>setEspId(e.target.value)}><option value="">— Seleccionar —</option>{espList.map(e=><option key={e.id} value={e.id}>{e.especialidad_es}</option>)}</select></div>
        <button className="btn-primary" onClick={cargar} disabled={loading}>{loading?'Cargando…':'Cargar reportes'}</button>
      </div>
      {reportes.length>0&&<>
        <div className="card p-4 border-2 border-[#003b7a]">
          <h3 className="font-bold text-[#003b7a] mb-3">📊 Consolidado del día — se actualiza al marcar</h3>
          <div className="flex items-center gap-4 mb-4"><span className="text-sm font-medium"><strong>{aprobadas.size}</strong> de <strong>{reportes.length}</strong> aprobados</span><div className="flex-1 progress-bar"><div className="progress-fill-blue transition-all duration-300" style={{width:`${reportes.length>0?Math.round(aprobadas.size/reportes.length*100):0}%`}}/></div><span className="text-sm font-bold text-[#003b7a]">{reportes.length>0?Math.round(aprobadas.size/reportes.length*100):0}%</span></div>
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
                  {pct!==null&&<><div className="progress-bar"><div className="progress-fill-blue transition-all duration-300" style={{width:`${pct}%`}}/></div><div className="text-xs text-slate-500 mt-1">{pct}% de {cfg?.meta_total} {cfg?.unidad_es}</div></>}
                  <div className="text-xs text-emerald-600 mt-1 font-medium">{act.cuadrillas.filter(c=>aprobadas.has(c.reporte_id)).length}/{act.cuadrillas.length} cuadrillas aprobadas</div>
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
                    <div key={c.reporte_id} className={`border rounded-lg p-3 transition-all ${isAp?'border-emerald-400 bg-emerald-50':isRe?'border-rose-400 bg-rose-50':'border-slate-200 bg-white'}`}>
                      <div className="flex items-start justify-between gap-2 flex-wrap mb-2">
                        <div><div className="font-medium text-sm">{c.usuario_nombre}</div>{c.areas.length>0&&<div className="text-xs text-slate-500 mt-0.5">{c.areas.map(a=>`${areaMap[a.area_id]||a.area_id}: ${a.cantidad} ${a.unidad}`).join(' · ')}</div>}{c.total>0&&<div className="text-base font-bold text-[#003b7a] mt-1">Total: {c.total} {cfg?.unidad_es||'und'}</div>}</div>
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
          <div className="text-sm text-slate-600"><strong>{aprobadas.size}</strong> por aprobar · <strong>{rechazadas.size}</strong> por rechazar</div>
          <button className="btn-success" disabled={saving||(!aprobadas.size&&!rechazadas.size)} onClick={enviar}>{saving?'Procesando…':'📤 Enviar decisiones'}</button>
        </div>
      </>}
      {!loading&&!reportes.length&&espId&&<div className="card p-6 text-center text-slate-500">Sin reportes pendientes para este día y especialidad.</div>}
      {!espId&&<div className="card p-6 text-center text-slate-500">Selecciona una especialidad y fecha para cargar los reportes.</div>}
    </div>
  );
}

// ── SOLICITUDES ───────────────────────────────────────────────────
function SolicitudesModule({user,catalogs,showToast}:{user:Profile;catalogs:Catalogs|null;showToast:(k:'ok'|'err'|'info',m:string)=>void}) {
  const[solicitudes,setSolicitudes]=useState<any[]>([]);
  const[form,setForm]=useState({fecha_reporte:'',especialidad_id:'',motivo:''});
  const[saving,setSaving]=useState(false);
  const canApprove=user.rol==='admin'||user.rol==='lider';
  const espList=useMemo(()=>catalogs?uniqueEsp(catalogs.especialidades_actividades):[],[catalogs]);

  async function load(){try{let q=supabase.from('solicitudes_reporte_pasado').select('*').order('created_at',{ascending:false});if(!canApprove)q=q.eq('tecnico_id',user.id);const{data}=await q;setSolicitudes(data||[]);}catch{setSolicitudes([]);}}
  useEffect(()=>{load();},[]);

  async function crear(){
    if(!form.fecha_reporte||!form.especialidad_id||!form.motivo){showToast('err','Completa todos los campos');return;}
    if(form.fecha_reporte>=today()){showToast('err','Solo para días anteriores');return;}
    setSaving(true);
    try{
      const{error}=await supabase.from('solicitudes_reporte_pasado').insert({tecnico_id:user.id,tecnico_nombre:user.nombre,fecha_reporte:form.fecha_reporte,especialidad_id:form.especialidad_id,motivo:form.motivo});
      if(error)throw error;
      try{const{data:admins}=await supabase.from('profiles').select('id').in('rol',['admin','lider']);if(admins?.length)await supabase.from('notificaciones').insert(admins.map((a:any)=>({usuario_id:a.id,tipo:'solicitud',titulo:'Solicitud de reporte pasado',mensaje:`${user.nombre} solicita reportar el ${form.fecha_reporte}`,data:{}})));}catch{}
      showToast('ok','Solicitud enviada.');setForm({fecha_reporte:'',especialidad_id:'',motivo:''});await load();
    }catch(e:any){showToast('err',e?.message||'Error');}
    finally{setSaving(false);}
  }

  async function decidir(id:string,estado:'aprobado'|'rechazado'){
    const mot=estado==='rechazado'?window.prompt('Motivo del rechazo:'):null;
    if(estado==='rechazado'&&!mot)return;
    await supabase.from('solicitudes_reporte_pasado').update({estado,aprobado_por:user.id,aprobado_en:new Date().toISOString(),comentario:mot||''}).eq('id',id);
    const sol=solicitudes.find(s=>s.id===id);
    if(sol)try{await supabase.from('notificaciones').insert({usuario_id:sol.tecnico_id,tipo:`solicitud_${estado}`,titulo:`Solicitud ${estado==='aprobado'?'aprobada ✅':'rechazada ❌'}`,mensaje:`Tu solicitud para reportar el ${sol.fecha_reporte} fue ${estado}`,data:{}});}catch{}
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
          <div key={s.id} className={`border rounded-lg p-3 ${s.estado==='aprobado'?'border-emerald-400 bg-emerald-50':s.estado==='rechazado'?'border-rose-400 bg-rose-50':'border-slate-200'}`}>
            <div className="flex items-start justify-between gap-2 flex-wrap">
              <div><div className="font-medium text-sm">{s.tecnico_nombre} — <strong>{s.fecha_reporte}</strong></div><div className="text-xs text-slate-500 mt-0.5">Motivo: {s.motivo}</div>{s.comentario&&<div className="text-xs text-slate-500">Resp: {s.comentario}</div>}</div>
              <div className="flex items-center gap-2 flex-wrap">
                <span className={`badge ${s.estado==='aprobado'?'bg-emerald-100 text-emerald-800':s.estado==='rechazado'?'bg-rose-100 text-rose-800':'badge-borrador'}`}>{s.estado}</span>
                {canApprove&&s.estado==='pendiente'&&<><button className="btn-success text-xs" onClick={()=>decidir(s.id,'aprobado')}>✓ Aprobar</button><button className="btn-danger text-xs" onClick={()=>decidir(s.id,'rechazado')}>✗ Rechazar</button></>}
              </div>
            </div>
          </div>
        ))}</div>
      </div>
    </div>
  );
}

// ── DASHBOARD ─────────────────────────────────────────────────────
function DashboardModule({showToast}:{showToast:(k:'ok'|'err'|'info',m:string)=>void}) {
  const[fechaIni,setFechaIni]=useState(today());const[fechaFin,setFechaFin]=useState(today());const[data,setData]=useState<any>(null);const[loading,setLoading]=useState(false);
  const load=useCallback(async()=>{
    setLoading(true);
    try{
      const[reps,asist,incid,maqD]=await Promise.all([supabase.from('reportes_avance').select('*').gte('fecha',fechaIni).lte('fecha',fechaFin),supabase.from('asistencia_real').select('*').gte('fecha',fechaIni).lte('fecha',fechaFin),supabase.from('incidentes_seg').select('*').gte('fecha',fechaIni).lte('fecha',fechaFin).neq('tipo','sin_novedad'),supabase.from('maquinaria').select('*')]);
      const aD=asist.data||[];const horasH=aD.filter((a:any)=>a.asistio).reduce((s:number,a:any)=>s+parseFloat(a.horas_trabajadas||0),0);const pl=aD.length,re=aD.filter((a:any)=>a.asistio).length;
      setData({reportes:reps.data||[],horas_hombre:Math.round(horasH),eficiencia_personal:pl>0?Math.round(re/pl*100):100,incidentes:incid.data||[],maquinaria:maqD.data||[]});
    }catch{showToast('err','Error cargando dashboard');}finally{setLoading(false);}
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
          <MC label="Eficiencia personal" value={`${data.eficiencia_personal}%`} sub="asistencia" color={data.eficiencia_personal>=90?'text-emerald-600':data.eficiencia_personal>=70?'text-amber-500':'text-rose-600'}/>
          <MC label="Horas-hombre" value={`${data.horas_hombre}h`} sub="productivas"/>
          <MC label="Reportes" value={data.reportes.length} sub="en el período"/>
          <MC label="Incidentes" value={data.incidentes.length} sub="seguridad" color={data.incidentes.length>0?'text-rose-600':'text-emerald-600'}/>
        </div>
        {data.maquinaria.length>0&&<div className="card p-4"><h3 className="font-bold text-[#003b7a] mb-3">🔧 Maquinaria acumulada</h3><div className="overflow-x-auto"><table className="table w-full"><thead><tr><th>Equipo</th><th>Tipo</th><th>Horas op.</th><th>Stand-by</th><th>Eficiencia</th></tr></thead><tbody>{data.maquinaria.map((m:any)=>{const t=(m.horas_acum_operativas||0)+(m.horas_acum_standby||0);const ef=t>0?Math.round((m.horas_acum_operativas||0)/t*100):100;return<tr key={m.id}><td className="font-medium">{m.item_id}</td><td>{m.tipo}</td><td className="text-emerald-600">{(m.horas_acum_operativas||0).toFixed(1)}h</td><td className="text-amber-500">{(m.horas_acum_standby||0).toFixed(1)}h</td><td>{sem(ef)} {ef}%</td></tr>;})}</tbody></table></div></div>}
      </>}
    </div>
  );
}
function MC({label,value,sub,color='text-[#003b7a]'}:{label:string;value:any;sub:string;color?:string}) {
  return<div className="card p-4 text-center"><div className={`text-2xl sm:text-3xl font-bold ${color}`}>{value}</div><div className="text-xs font-semibold text-slate-700 mt-1">{label}</div><div className="text-xs text-slate-400">{sub}</div></div>;
}

// ── INFORMES ──────────────────────────────────────────────────────
function InformesModule({user,catalogs,showToast}:{user:Profile;catalogs:Catalogs|null;showToast:(k:'ok'|'err'|'info',m:string)=>void}) {
  const[fechaIni,setFechaIni]=useState(today());const[fechaFin,setFechaFin]=useState(today());const[espIds,setEspIds]=useState<string[]>([]);const[areaIds,setAreaIds]=useState<string[]>([]);const[soloAp,setSoloAp]=useState(false);const[data,setData]=useState<any>(null);const[loading,setLoading]=useState(false);
  const espList=useMemo(()=>catalogs?uniqueEsp(catalogs.especialidades_actividades):[],[catalogs]);
  async function fetchData(){
    setLoading(true);
    try{
      let qR=supabase.from('reportes_avance').select('*').gte('fecha',fechaIni).lte('fecha',fechaFin);
      if(espIds.length)qR=qR.in('especialidad_id',espIds);if(soloAp)qR=qR.eq('estado','aprobado');if(user.rol==='tecnico')qR=qR.eq('usuario_id',user.id);if(user.rol==='cliente')qR=qR.eq('estado','aprobado');
      const{data:reps}=await qR;const repIds=(reps||[]).map((r:any)=>r.id);
      const[av,as2,sc]=await Promise.all([repIds.length?supabase.from('avance_diario').select('*').in('reporte_id',repIds):Promise.resolve({data:[]}),repIds.length?supabase.from('asistencia_real').select('*').in('reporte_id',repIds):Promise.resolve({data:[]}),repIds.length?supabase.from('suspensiones_clima').select('*').in('reporte_id',repIds):Promise.resolve({data:[]})]);
      let avD=(av.data||[]) as any[];if(areaIds.length)avD=avD.filter((a:any)=>areaIds.includes(a.area_id));
      const aD=(as2.data||[]) as any[];const horasH=aD.filter((a:any)=>a.asistio).reduce((s:number,a:any)=>s+parseFloat(a.horas_trabajadas||0),0);const horasC=((sc.data||[]) as any[]).reduce((s:number,a:any)=>s+parseFloat(a.horas_perdidas||0),0);
      setData({reportes:reps||[],avances:avD,asistencia:aD,totales:{horas_hombre:Math.round(horasH),horas_perdidas_clima:Math.round(horasC),dias:repIds.length}});
    }catch{showToast('err','Error');}finally{setLoading(false);}
  }
  function exportExcel(){
    if(!data)return;
    const isC=user.rol==='cliente';
    const rows=data.avances.map((av:any)=>{const aR=catalogs?.especialidades_actividades.find(e=>e.id===av.actividad_id);const arR=catalogs?.areas.find(a=>a.id===av.area_id);const rep=data.reportes.find((r:any)=>r.id===av.reporte_id);const o:any={Fecha:av.fecha,Especialidad:aR?.especialidad_es||'',Actividad:aR?.actividad_es||'',Área:arR?.area_es||'',Cantidad:av.cantidad,Unidad:av.unidad,Acumulado:av.acumulado_total};if(!isC)o.Técnico=rep?.usuario_nombre||'';return o;});
    const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(rows),'Avance');
    if(!isC){const aRows=data.asistencia.map((a:any)=>({Fecha:a.fecha,Documento:a.documento_personal,Asistió:a.asistio?'Sí':'No',Motivo:a.motivo_ausencia||'',Horas:a.horas_trabajadas}));XLSX.utils.book_append_sheet(wb,XLSX.utils.json_to_sheet(aRows),'Asistencia');}
    XLSX.writeFile(wb,`PDS360_${fechaIni}_${fechaFin}.xlsx`);showToast('ok','Excel descargado');
  }
  function toggle(arr:string[],setArr:(v:string[])=>void,val:string){setArr(arr.includes(val)?arr.filter(x=>x!==val):[...arr,val]);}
  return(
    <div className="space-y-4">
      <div className="card p-4 space-y-3 no-print">
        <div className="grid grid-cols-1 sm:grid-cols-3 gap-3 items-end"><div><label className="label">Desde</label><input type="date" className="input" value={fechaIni} onChange={e=>setFechaIni(e.target.value)}/></div><div><label className="label">Hasta</label><input type="date" className="input" value={fechaFin} onChange={e=>setFechaFin(e.target.value)}/></div><button className="btn-primary" onClick={fetchData} disabled={loading}>{loading?'Cargando…':'Consultar'}</button></div>
        <div><label className="label">Especialidades</label><div className="flex flex-wrap gap-2">{espList.map(e=><button key={e.id} onClick={()=>toggle(espIds,setEspIds,e.id)} className={`text-xs px-2 py-1 rounded border transition-colors ${espIds.includes(e.id)?'bg-[#003b7a] text-white border-[#003b7a]':'border-slate-300 text-slate-600 hover:border-[#003b7a]'}`}>{e.especialidad_es}</button>)}</div></div>
        <div><label className="label">Áreas</label><div className="flex flex-wrap gap-2">{(catalogs?.areas||[]).map(a=><button key={a.id} onClick={()=>toggle(areaIds,setAreaIds,a.id)} className={`text-xs px-2 py-1 rounded border transition-colors ${areaIds.includes(a.id)?'bg-[#003b7a] text-white border-[#003b7a]':'border-slate-300 text-slate-600 hover:border-[#003b7a]'}`}>{a.area_es}</button>)}</div></div>
        <div className="flex items-center gap-3 flex-wrap"><label className="flex items-center gap-2 text-sm cursor-pointer"><input type="checkbox" checked={soloAp} onChange={e=>setSoloAp(e.target.checked)}/> Solo aprobados</label><button className="btn-secondary text-xs" onClick={exportExcel} disabled={!data}>📥 Excel</button><button className="btn-secondary text-xs" onClick={()=>window.print()}>🖨️ PDF</button></div>
      </div>
      <div className="print-area space-y-4">
        {data?.totales&&<div className="card p-4"><h3 className="font-bold text-[#003b7a] mb-3">Totales del período</h3><div className="grid grid-cols-3 gap-3 text-center"><div><div className="text-xl font-bold text-[#003b7a]">{data.totales.dias}</div><div className="text-xs text-slate-500">Reportes</div></div><div><div className="text-xl font-bold text-emerald-600">{data.totales.horas_hombre}h</div><div className="text-xs text-slate-500">Horas-hombre</div></div><div><div className="text-xl font-bold text-amber-500">{data.totales.horas_perdidas_clima}h</div><div className="text-xs text-slate-500">Perdidas clima</div></div></div></div>}
        {data?.avances?.length>0&&<div className="card p-4"><h3 className="font-bold text-[#003b7a] mb-3">Avance de actividades</h3><div className="overflow-x-auto"><table className="table w-full"><thead><tr><th>Fecha</th><th>Actividad</th><th>Área</th><th>Cantidad</th><th>Acumulado</th><th>Unidad</th>{user.rol!=='cliente'&&<th>Técnico</th>}</tr></thead><tbody>{data.avances.map((av:any,i:number)=>{const aR=catalogs?.especialidades_actividades.find(e=>e.id===av.actividad_id);const arR=catalogs?.areas.find(a=>a.id===av.area_id);const rep=data.reportes.find((r:any)=>r.id===av.reporte_id);return<tr key={i}><td>{av.fecha}</td><td>{aR?.actividad_es||av.actividad_id}</td><td>{arR?.area_es||av.area_id}</td><td className="font-semibold">{av.cantidad}</td><td className="font-semibold text-[#003b7a]">{av.acumulado_total}</td><td>{av.unidad}</td>{user.rol!=='cliente'&&<td className="text-xs text-slate-500">{rep?.usuario_nombre||''}</td>}</tr>;})} </tbody></table></div></div>}
        {!data?.avances?.length&&!loading&&<div className="card p-6 text-center text-slate-500">Sin datos. Consulta primero.</div>}
      </div>
    </div>
  );
}

// ── CATÁLOGOS — con botones Activar/Desactivar/Eliminar ────────────
function CatalogosModule({catalogs,onRefresh,showToast}:{catalogs:Catalogs|null;onRefresh:()=>void;showToast:(k:'ok'|'err'|'info',m:string)=>void}) {
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

function CatMgr({title,table,nameField,fields,rows,onChanged,showToast}:{title:string;table:string;nameField:string;fields:{n:string;l:string}[];rows:any[];onChanged:()=>void;showToast:(k:'ok'|'err'|'info',m:string)=>void}) {
  const[form,setForm]=useState<Record<string,string>>({});
  const[search,setSearch]=useState('');
  const[filtro,setFiltro]=useState<'activos'|'inactivos'|'todos'>('activos');
  const[busy,setBusy]=useState(false);

  const filtered=useMemo(()=>{
    let list=rows;
    if(filtro==='activos')list=list.filter(r=>r.activo!==false);
    else if(filtro==='inactivos')list=list.filter(r=>r.activo===false);
    const q=search.trim().toLowerCase();
    return q?list.filter(r=>fields.some(f=>String(r[f.n]||'').toLowerCase().includes(q))):list;
  },[rows,search,fields,filtro]);

  async function addOne(){
    if(!form[fields[0].n]){showToast('err',`Falta ${fields[0].l}`);return;}setBusy(true);
    try{const{error}=await supabase.from(table).insert({...form,activo:true});if(error)throw error;showToast('ok','Agregado correctamente');setForm({});onChanged();}
    catch(e:any){showToast('err',e?.message||'Error');}finally{setBusy(false);}
  }

  async function toggleActivo(row:any){
    const nuevo=row.activo===false;
    if(!window.confirm(`¿${nuevo?'Activar':'Desactivar'} "${row[nameField]}"?\n${!nuevo?'Ya no estará disponible para nuevas asignaciones. Los registros históricos se mantienen.':''}`))return;
    const{error}=await supabase.from(table).update({activo:nuevo}).eq('id',row.id);
    if(error){showToast('err',error.message);return;}
    showToast('ok',nuevo?'Activado ✓':'Desactivado ✓');onChanged();
  }

  async function eliminar(row:any){
    if(!window.confirm(`⚠️ ¿ELIMINAR PERMANENTEMENTE "${row[nameField]}"?\n\nSe recomienda DESACTIVAR en lugar de eliminar para mantener el historial.`))return;
    setBusy(true);
    try{const{error}=await supabase.from(table).delete().eq('id',row.id);if(error)throw error;showToast('ok','Eliminado');onChanged();}
    catch(e:any){showToast('err',`No se puede eliminar: ${e?.message}. Intenta desactivarlo.`);}
    finally{setBusy(false);}
  }

  async function loadExcel(file:File){
    setBusy(true);
    try{
      const buf=await file.arrayBuffer();const wb=XLSX.read(buf);const data:any[]=XLSX.utils.sheet_to_json(wb.Sheets[wb.SheetNames[0]],{defval:''});
      const norm=data.map(row=>{const o:any={activo:true};for(const k in row){o[String(k).trim().toLowerCase().replace(/\s+/g,'_')]=row[k];}return o;});
      if(!norm.length){showToast('err','Excel vacío');return;}
      if(!window.confirm(`¿Reemplazar "${title}" con ${norm.length} filas del Excel?`))return;
      await supabase.from(table).delete().neq('id','00000000-0000-0000-0000-000000000000');
      const{error}=await supabase.from(table).insert(norm);
      if(error)throw error;showToast('ok',`Catálogo actualizado con ${norm.length} registros`);onChanged();
    }catch(e:any){showToast('err',e?.message||'Error');}finally{setBusy(false);}
  }

  const activos=rows.filter(r=>r.activo!==false).length;
  const inactivos=rows.filter(r=>r.activo===false).length;

  return(
    <div className="card p-4">
      <div className="flex items-center justify-between flex-wrap gap-2 mb-3">
        <div><h3 className="font-bold text-[#003b7a]">{title}</h3><div className="text-xs text-slate-500 mt-0.5">{activos} activos · {inactivos} inactivos · {rows.length} total</div></div>
        <label className="btn-secondary cursor-pointer text-xs">📥 Cargar Excel<input type="file" accept=".xlsx,.xls,.csv" hidden onChange={e=>{const f=e.target.files?.[0];if(f)loadExcel(f);e.currentTarget.value='';}} /></label>
      </div>

      {/* Formulario agregar */}
      <div className="grid grid-cols-2 sm:grid-cols-3 lg:grid-cols-6 gap-2 mb-2">{fields.map(f=><div key={f.n}><label className="label">{f.l}</label><input className="input" value={form[f.n]||''} onChange={e=>setForm({...form,[f.n]:e.target.value})}/></div>)}</div>
      <button className="btn-primary text-xs mb-3" disabled={busy} onClick={addOne}>+ Agregar registro</button>

      {/* Búsqueda y filtros */}
      <div className="flex gap-2 mb-3 flex-wrap">
        <input className="input flex-1 min-w-[160px]" placeholder="🔎 Buscar…" value={search} onChange={e=>setSearch(e.target.value)}/>
        {(['activos','inactivos','todos'] as const).map(f=>(
          <button key={f} onClick={()=>setFiltro(f)} className={`text-xs px-3 py-1.5 rounded border transition-colors ${filtro===f?'bg-[#003b7a] text-white border-[#003b7a]':'border-slate-300 text-slate-600'}`}>
            {f.charAt(0).toUpperCase()+f.slice(1)}
          </button>
        ))}
      </div>

      {/* LISTA con botones visibles — diseño de tarjetas en lugar de tabla para evitar scroll horizontal */}
      {filtered.length>0&&(
        <div className="space-y-2 max-h-80 overflow-y-auto">
          {filtered.slice(0,200).map((r,i)=>(
            <div key={i} className={`flex items-center gap-2 p-2 rounded-lg border ${r.activo===false?'bg-slate-50 border-slate-200 opacity-60':'bg-white border-slate-200 hover:border-slate-300'}`}>
              {/* Info principal */}
              <div className="flex-1 min-w-0">
                <div className="font-medium text-sm truncate">{String(r[nameField]??'')}</div>
                <div className="text-xs text-slate-500 truncate">
                  {fields.filter(f=>f.n!==nameField).map(f=>String(r[f.n]??'')).filter(Boolean).join(' · ')}
                </div>
              </div>
              {/* Estado */}
              <div className="flex-shrink-0">
                {r.activo===false
                  ?<span className="badge bg-rose-100 text-rose-700 text-xs">Inactivo</span>
                  :<span className="badge bg-emerald-100 text-emerald-700 text-xs">Activo</span>
                }
              </div>
              {/* Botones — siempre visibles */}
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
      )}
      {!filtered.length&&<div className="text-sm text-slate-500 text-center py-4">Sin registros con los filtros actuales.</div>}
    </div>
  );
}

// ── MAQUINARIA ────────────────────────────────────────────────────
function MaquinariaModule({maquinaria,onRefresh,showToast}:{maquinaria:Maq[];onRefresh:()=>void;showToast:(k:'ok'|'err'|'info',m:string)=>void}) {
  const[form,setForm]=useState({tipo:'motosierra',item_id:'',nombre:'',estado:'activo'});const[busy,setBusy]=useState(false);
  async function addOne(){if(!form.item_id){showToast('err','Falta ID del equipo');return;}setBusy(true);try{const{error}=await supabase.from('maquinaria').insert({...form,horas_acum_operativas:0,horas_acum_standby:0});if(error)throw error;showToast('ok','Equipo agregado');setForm({tipo:'motosierra',item_id:'',nombre:'',estado:'activo'});onRefresh();}catch(e:any){showToast('err',e?.message||'Error');}finally{setBusy(false);}  }
  async function cambiarEstado(m:Maq,nuevoEstado:string){if(!window.confirm(`¿Cambiar estado de ${m.item_id} a "${nuevoEstado}"?`))return;await supabase.from('maquinaria').update({estado:nuevoEstado}).eq('id',m.id);showToast('ok','Estado actualizado');onRefresh();}
  async function eliminar(m:Maq){if(!window.confirm(`⚠️ ¿ELIMINAR ${m.item_id} permanentemente?`))return;setBusy(true);try{const{error}=await supabase.from('maquinaria').delete().eq('id',m.id);if(error)throw error;showToast('ok','Eliminado');onRefresh();}catch(e:any){showToast('err',`No se puede eliminar: ${e?.message}`);}finally{setBusy(false);}  }
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

// ── CONFIG ACTIVIDADES ────────────────────────────────────────────
function ConfigActModule({configActs,catalogs,onRefresh,showToast}:{configActs:ConfigAct[];catalogs:Catalogs|null;onRefresh:()=>void;showToast:(k:'ok'|'err'|'info',m:string)=>void}) {
  const[form,setForm]=useState({especialidad_id:'',actividad_id:'',tipo:'A',unidad_es:'',unidad_en:'',meta_total:'',acumulado_previo:'',rendimiento_esperado:'',rendimiento_por:'cuadrilla',tiene_meta:true,es_medible:true});
  const[busy,setBusy]=useState(false);
  const esps=useMemo(()=>catalogs?uniqueEsp(catalogs.especialidades_actividades.filter(e=>e.activo!==false)):[],[catalogs]);
  const acts=useMemo(()=>catalogs&&form.especialidad_id?actsForEsp(catalogs.especialidades_actividades.filter(e=>e.activo!==false),form.especialidad_id):[],[catalogs,form.especialidad_id]);

  function selAct(actId:string){
    const cfg=configActs.find(c=>c.actividad_id===actId);
    if(cfg)setForm({...form,actividad_id:actId,tipo:cfg.tipo||'A',unidad_es:cfg.unidad_es||'',unidad_en:cfg.unidad_en||'',meta_total:String(cfg.meta_total||''),tiene_meta:cfg.tiene_meta!==false,es_medible:cfg.es_medible!==false,acumulado_previo:String(cfg.acumulado_previo||''),rendimiento_esperado:String(cfg.rendimiento_esperado||'')});
    else setForm({...form,actividad_id:actId});
  }

  async function save(){
    if(!form.actividad_id){showToast('err','Selecciona una actividad');return;}
    if(form.es_medible&&form.tiene_meta&&!form.meta_total){showToast('err','Ingresa la meta total');return;}
    if(form.es_medible&&!form.unidad_es){showToast('err','Ingresa la unidad de medida');return;}
    setBusy(true);
    try{
      const payload={especialidad_id:form.especialidad_id,actividad_id:form.actividad_id,tipo:form.es_medible?form.tipo:'B',unidad_es:form.unidad_es||'N/A',unidad_en:form.unidad_en||'N/A',tiene_meta:!!(form.tiene_meta&&form.es_medible),es_medible:!!form.es_medible,meta_total:form.tiene_meta&&form.es_medible&&form.meta_total?parseFloat(form.meta_total):null,acumulado_previo:parseFloat(form.acumulado_previo||'0'),rendimiento_esperado:form.rendimiento_esperado?parseFloat(form.rendimiento_esperado):null,rendimiento_por:form.rendimiento_por,activo:true};
      const{error}=await supabase.from('config_actividades').upsert(payload,{onConflict:'actividad_id'});
      if(error)throw error;showToast('ok','✅ Configuración guardada');onRefresh();
    }catch(e:any){showToast('err',e?.message||'Error');}finally{setBusy(false);}
  }

  async function eliminarConfig(actId:string){
    if(!window.confirm('¿Eliminar la configuración? Los técnicos no podrán reportar esta actividad hasta que la reconfigures.'))return;
    await supabase.from('config_actividades').delete().eq('actividad_id',actId);
    showToast('ok','Configuración eliminada');onRefresh();
  }

  return(
    <div className="card p-4 space-y-4">
      <div><h3 className="font-bold text-[#003b7a]">Configuración de actividades ({configActs.length} configuradas)</h3><p className="text-xs text-amber-700 bg-amber-50 border border-amber-200 rounded p-2 mt-2">⚠️ Cada actividad debe configurarse aquí antes de que los técnicos puedan reportarla.</p></div>
      <div className="grid grid-cols-1 sm:grid-cols-2 gap-3">
        <div><label className="label">Especialidad</label><select className="select" value={form.especialidad_id} onChange={e=>setForm({...form,especialidad_id:e.target.value,actividad_id:''})}><option value="">— Seleccionar —</option>{esps.map(e=><option key={e.id} value={e.id}>{e.especialidad_es}</option>)}</select></div>
        <div><label className="label">Actividad</label><select className="select" value={form.actividad_id} disabled={!form.especialidad_id} onChange={e=>selAct(e.target.value)}><option value="">— Seleccionar —</option>{acts.map(a=>{const yaConf=configActs.some(c=>c.actividad_id===a.id);return<option key={a.id} value={a.id}>{a.actividad_es}{yaConf?' ✓':''}</option>;})}</select></div>
      </div>
      <div className="bg-slate-50 rounded-lg p-4 space-y-4 border border-slate-200">
        <label className="flex items-start gap-3 cursor-pointer p-3 rounded-lg border border-slate-200 bg-white hover:border-[#003b7a] transition-colors"><input type="checkbox" checked={form.es_medible} onChange={e=>setForm({...form,es_medible:e.target.checked,tiene_meta:e.target.checked?form.tiene_meta:false})} className="w-5 h-5 mt-0.5"/><div><div className="font-medium text-sm">¿Esta actividad es medible?</div><div className="text-xs text-slate-500">Desmarca si es trabajo administrativo, gestión, capacitaciones, etc.</div></div></label>
        {form.es_medible&&<label className="flex items-start gap-3 cursor-pointer p-3 rounded-lg border border-slate-200 bg-white hover:border-[#003b7a] transition-colors"><input type="checkbox" checked={form.tiene_meta} onChange={e=>setForm({...form,tiene_meta:e.target.checked})} className="w-5 h-5 mt-0.5"/><div><div className="font-medium text-sm">¿Tiene meta numérica definida?</div><div className="text-xs text-slate-500">Desmarca si acumula sin límite (conteo de fauna, monitoreo continuo, etc.)</div></div></label>}
        {form.es_medible&&(
          <div className="grid grid-cols-1 sm:grid-cols-2 lg:grid-cols-3 gap-3 pt-2">
            <div><label className="label">Tipo</label><select className="select" value={form.tipo} onChange={e=>setForm({...form,tipo:e.target.value})}><option value="A">A — Con meta</option><option value="B">B — Acumulativa</option><option value="C">C — Ítems únicos</option></select></div>
            <div><label className="label">Unidad (ES)</label><input className="input" value={form.unidad_es} onChange={e=>setForm({...form,unidad_es:e.target.value})} placeholder="árboles, m³, ha…"/></div>
            <div><label className="label">Unidad (EN)</label><input className="input" value={form.unidad_en} onChange={e=>setForm({...form,unidad_en:e.target.value})}/></div>
            {form.tiene_meta&&<div><label className="label">Meta total</label><input type="number" className="input" value={form.meta_total} onChange={e=>setForm({...form,meta_total:e.target.value})} placeholder="Ej: 3466"/></div>}
            {form.tiene_meta&&<div><label className="label">Acumulado previo</label><input type="number" className="input" value={form.acumulado_previo} onChange={e=>setForm({...form,acumulado_previo:e.target.value})} placeholder="0"/></div>}
            <div><label className="label">Rendimiento esperado</label><input type="number" className="input" value={form.rendimiento_esperado} onChange={e=>setForm({...form,rendimiento_esperado:e.target.value})}/></div>
            <div><label className="label">Rendimiento por</label><select className="select" value={form.rendimiento_por} onChange={e=>setForm({...form,rendimiento_por:e.target.value})}><option value="cuadrilla">cuadrilla/día</option><option value="persona">persona/día</option><option value="equipo">equipo/día</option></select></div>
          </div>
        )}
      </div>
      <button className="btn-primary" disabled={busy||!form.actividad_id} onClick={save}>{busy?'Guardando…':'💾 Guardar configuración'}</button>
      {configActs.length>0&&(
        <div className="space-y-2 max-h-72 overflow-y-auto">
          <div className="font-semibold text-sm text-slate-700 mb-2">Actividades configuradas</div>
          {configActs.map((c,i)=>{
            const actRow=catalogs?.especialidades_actividades.find(e=>e.id===c.actividad_id);
            return(
              <div key={i} className="flex items-center gap-2 p-2 rounded-lg border border-slate-200 bg-white">
                <div className="flex-1 min-w-0">
                  <div className="font-medium text-sm truncate">{actRow?.actividad_es||c.actividad_id.slice(-8)}</div>
                  <div className="text-xs text-slate-500">
                    {c.es_medible===false?'No medible':c.tiene_meta?`Meta: ${c.meta_total} ${c.unidad_es}`:`Acumulativo · ${c.unidad_es}`}
                    {' · Tipo '}{c.tipo}
                  </div>
                </div>
                <div className="flex gap-1 flex-shrink-0">
                  <button className="text-xs px-2 py-1 rounded bg-amber-100 text-amber-700 hover:bg-amber-200" onClick={()=>selAct(c.actividad_id)}>Editar</button>
                  <button className="text-xs px-2 py-1 rounded bg-rose-100 text-rose-700 hover:bg-rose-200" onClick={()=>eliminarConfig(c.actividad_id)}>Eliminar</button>
                </div>
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
}

// ── USUARIOS ──────────────────────────────────────────────────────
function UsuariosModule({showToast}:{showToast:(k:'ok'|'err'|'info',m:string)=>void}) {
  const[users,setUsers]=useState<Profile[]>([]);
  const[form,setForm]=useState({nombre:'',correo:'',clave:'',rol:'tecnico' as UserRole,especialidad_id:''});
  const[busy,setBusy]=useState(false);const[loading,setLoading]=useState(true);const[search,setSearch]=useState('');
  async function loadUsers(){setLoading(true);const{data}=await supabase.from('profiles').select('*').order('nombre');setUsers((data||[]) as Profile[]);setLoading(false);}
  useEffect(()=>{loadUsers();},[]);
  const filtered=useMemo(()=>{const q=search.trim().toLowerCase();return q?users.filter(u=>u.nombre.toLowerCase().includes(q)||u.correo.toLowerCase().includes(q)):users;},[users,search]);
  async function createUser(){
    if(!form.nombre||!form.correo||!form.clave){showToast('err','Nombre, correo y clave son obligatorios');return;}
    if(form.clave.length<8){showToast('err','Clave mínimo 8 caracteres');return;}setBusy(true);
    try{const r=await fetch('/api/admin',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({action:'createUser',...form})});const d=await r.json();if(!d.ok)throw new Error(d.error||'Error');showToast('ok','Usuario creado ✓');setForm({nombre:'',correo:'',clave:'',rol:'tecnico',especialidad_id:''});loadUsers();}
    catch(e:any){showToast('err',e?.message||'Error');}finally{setBusy(false);}
  }
  async function toggleActivo(u:Profile){
    if(!window.confirm(`¿${!u.activo?'Activar':'Desactivar'} a ${u.nombre}?\n${u.activo?'El usuario no podrá ingresar a la app.':''}`))return;
    await supabase.from('profiles').update({activo:!u.activo,updated_at:new Date().toISOString()}).eq('id',u.id);
    showToast('ok',!u.activo?`${u.nombre} activado`:`${u.nombre} desactivado`);loadUsers();
  }
  async function cambiarRol(u:Profile,rol:string){
    if(!window.confirm(`¿Cambiar el rol de ${u.nombre} a "${rol}"?`))return;
    await supabase.from('profiles').update({rol,updated_at:new Date().toISOString()}).eq('id',u.id);
    showToast('ok','Rol actualizado');loadUsers();
  }
  async function eliminar(u:Profile){
    if(!window.confirm(`⚠️ ¿ELIMINAR al usuario ${u.nombre}?\n\nEsto eliminará su acceso. Sus reportes se mantienen.`))return;setBusy(true);
    try{const r=await fetch('/api/admin',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({action:'deleteUser',user_id:u.id})});const d=await r.json();if(!d.ok)throw new Error(d.error||'Error');showToast('ok','Usuario eliminado');loadUsers();}
    catch(e:any){showToast('err',e?.message||'Error');}finally{setBusy(false);}
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
          <div><label className="label">Clave (min. 8)</label><input className="input" type="password" value={form.clave} onChange={e=>setForm({...form,clave:e.target.value})}/></div>
          <div><label className="label">Rol</label><select className="select" value={form.rol} onChange={e=>setForm({...form,rol:e.target.value as UserRole})}>{ROLES.map(r=><option key={r} value={r}>{r.charAt(0).toUpperCase()+r.slice(1)}</option>)}</select></div>
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
                <button className="text-xs px-2 py-1 rounded font-medium bg-rose-100 text-rose-700 hover:bg-rose-200 transition-colors" onClick={()=>eliminar(u)} disabled={busy}>Eliminar</button>
              </div>
            </div>
          ))}
        </div>
      )}
    </div>
  );
}

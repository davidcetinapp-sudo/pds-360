-- ============================================================
-- Powerchina · PDS 360  ·  Supabase Schema v2.0
-- Ejecutar completo en: Supabase → SQL Editor → New query
-- ============================================================

-- ── Extensiones ──────────────────────────────────────────────
create extension if not exists "uuid-ossp";

-- ── Función helper: obtener rol del usuario actual ───────────
create or replace function public.get_my_role()
returns text
language sql
security definer stable
as $$
  select rol from public.profiles where id = auth.uid()
$$;

-- ============================================================
-- 1. PROFILES  (extiende auth.users)
-- ============================================================
create table public.profiles (
  id              uuid references auth.users on delete cascade primary key,
  nombre          text not null,
  correo          text unique not null,
  rol             text not null default 'tecnico'
                  check (rol in ('admin','lider','tecnico','gerencia','cliente','visualizador')),
  especialidad_id uuid,
  activo          boolean default true,
  created_at      timestamptz default now(),
  updated_at      timestamptz default now()
);

-- Crear perfil automáticamente al registrar usuario
create or replace function public.handle_new_user()
returns trigger language plpgsql security definer as $$
begin
  insert into public.profiles (id, nombre, correo, rol)
  values (
    new.id,
    coalesce(new.raw_user_meta_data->>'nombre', split_part(new.email,'@',1)),
    new.email,
    coalesce(new.raw_user_meta_data->>'rol','tecnico')
  )
  on conflict (id) do nothing;
  return new;
end;
$$;

create or replace trigger on_auth_user_created
  after insert on auth.users
  for each row execute procedure public.handle_new_user();

-- ============================================================
-- 2. CATÁLOGOS
-- ============================================================
create table public.especialidades_actividades (
  id              uuid primary key default gen_random_uuid(),
  especialidad_es text not null,
  especialidad_en text not null default '',
  actividad_es    text not null,
  actividad_en    text not null default '',
  activo          boolean default true,
  created_at      timestamptz default now()
);

create table public.areas (
  id         uuid primary key default gen_random_uuid(),
  area_es    text not null,
  area_en    text not null default '',
  activo     boolean default true,
  created_at timestamptz default now()
);

create table public.lideres (
  id              uuid primary key default gen_random_uuid(),
  nombre          text not null,
  documento       text not null,
  cargo_es        text not null default '',
  cargo_en        text not null default '',
  especialidad_id uuid,
  activo          boolean default true,
  created_at      timestamptz default now()
);

create table public.personal (
  id                uuid primary key default gen_random_uuid(),
  nombre            text not null,
  documento         text not null unique,
  cargo_es          text not null default '',
  cargo_en          text not null default '',
  tipo              text not null default 'directo'
                    check (tipo in ('directo','subcontratista')),
  empresa           text,
  subcontratista_id uuid,
  activo            boolean default true,
  created_at        timestamptz default now()
);

create table public.subcontratistas (
  id         uuid primary key default gen_random_uuid(),
  empresa    text not null,
  nit        text,
  contacto   text,
  telefono   text,
  email      text,
  activo     boolean default true,
  created_at timestamptz default now()
);

-- ============================================================
-- 3. MAQUINARIA
-- ============================================================
create table public.maquinaria (
  id                    uuid primary key default gen_random_uuid(),
  tipo                  text not null,
  item_id               text not null unique,
  nombre                text not null default '',
  especialidad_id       uuid,
  estado                text not null default 'activo'
                        check (estado in ('activo','inactivo','mantenimiento')),
  fecha_ingreso         date,
  horas_acum_operativas numeric default 0,
  horas_acum_standby    numeric default 0,
  created_at            timestamptz default now()
);

-- ============================================================
-- 4. CONFIG DE ACTIVIDADES  (Tipo A / B / C)
-- ============================================================
create table public.config_actividades (
  id                   uuid primary key default gen_random_uuid(),
  especialidad_id      uuid not null,
  actividad_id         uuid not null references public.especialidades_actividades(id),
  tipo                 text not null check (tipo in ('A','B','C')),
  unidad_es            text not null default '',
  unidad_en            text not null default '',
  meta_total           numeric,
  tiene_items_unicos   boolean default false,
  rendimiento_esperado numeric,
  rendimiento_por      text default 'cuadrilla',
  acumulado_previo     numeric default 0,
  activo               boolean default true,
  created_at           timestamptz default now(),
  unique (actividad_id)
);

-- ============================================================
-- 5. INVENTARIO DE ÍTEMS ÚNICOS  (Tipo C)
-- ============================================================
create table public.inventario_items (
  id              uuid primary key default gen_random_uuid(),
  actividad_id    uuid not null references public.especialidades_actividades(id),
  item_id         text not null,
  descripcion     text,
  area_id         uuid references public.areas(id),
  estado          text not null default 'pendiente'
                  check (estado in ('pendiente','ejecutado')),
  fecha_ejecucion date,
  tecnico_id      uuid references public.profiles(id),
  reporte_id      uuid,
  created_at      timestamptz default now()
);

-- ============================================================
-- 6. PLANEACIÓN
-- ============================================================
create table public.programaciones (
  id             uuid primary key default gen_random_uuid(),
  fecha          date not null,
  usuario_id     uuid not null references public.profiles(id),
  usuario_nombre text not null,
  estado         text not null default 'borrador'
                 check (estado in ('borrador','enviado')),
  created_at     timestamptz default now(),
  updated_at     timestamptz default now(),
  unique (fecha, usuario_id)   -- un usuario, una planeación por fecha
);

-- ── FIX ERROR 3: cada actividad tiene su PROPIO líder ────────
create table public.actividades_programadas (
  id                  uuid primary key default gen_random_uuid(),
  programacion_id     uuid not null references public.programaciones(id) on delete cascade,
  fecha               date not null,
  usuario_id          uuid not null references public.profiles(id),
  especialidad_id     uuid not null,
  actividad_id        uuid not null references public.especialidades_actividades(id),
  area_id             uuid not null references public.areas(id),
  areas_adicionales   uuid[] default '{}',
  lider_id            uuid references public.lideres(id),  -- ← líder propio de esta actividad
  maquinaria_ids      uuid[] default '{}',
  rendimiento_esperado text,
  observacion_es      text,
  observacion_en      text,
  created_at          timestamptz default now()
);

-- ── FIX ERROR 1: unique constraint = sin bloqueos de concurrencia ──
create table public.personal_asignado (
  id                      uuid primary key default gen_random_uuid(),
  programacion_id         uuid not null references public.programaciones(id) on delete cascade,
  actividad_programada_id uuid not null references public.actividades_programadas(id) on delete cascade,
  fecha                   date not null,
  usuario_id              uuid not null references public.profiles(id),
  personal_id             uuid not null references public.personal(id),
  documento_personal      text not null,
  created_at              timestamptz default now(),
  -- Si alguien ya está asignado esa fecha, PostgreSQL lanza error 23505 inmediatamente
  -- sin bloquear a nadie más
  unique (fecha, documento_personal, usuario_id)
);

-- ============================================================
-- 7. REPORTE DIARIO DE AVANCE
-- ============================================================
create table public.reportes_avance (
  id              uuid primary key default gen_random_uuid(),
  fecha           date not null,
  usuario_id      uuid not null references public.profiles(id),
  usuario_nombre  text not null,
  especialidad_id uuid not null,
  jornada_horas   numeric not null default 9,
  clima           text not null default 'despejado',
  charla_preturno boolean default true,
  charla_tema     text,
  estado          text not null default 'borrador'
                  check (estado in ('borrador','aprobado')),
  aprobado_por    uuid references public.profiles(id),
  aprobado_en     timestamptz,
  version         int default 1,
  created_at      timestamptz default now(),
  updated_at      timestamptz default now()
);

create table public.asistencia_real (
  id                 uuid primary key default gen_random_uuid(),
  reporte_id         uuid not null references public.reportes_avance(id) on delete cascade,
  fecha              date not null,
  usuario_id         uuid not null references public.profiles(id),
  personal_id        uuid not null references public.personal(id),
  documento_personal text not null,
  asistio            boolean not null default true,
  motivo_ausencia    text,
  horas_trabajadas   numeric default 0,
  created_at         timestamptz default now()
);

create table public.avance_diario (
  id                 uuid primary key default gen_random_uuid(),
  reporte_id         uuid not null references public.reportes_avance(id) on delete cascade,
  fecha              date not null,
  usuario_id         uuid not null references public.profiles(id),
  actividad_id       uuid not null references public.especialidades_actividades(id),
  especialidad_id    uuid not null,
  area_id            uuid not null references public.areas(id),
  cantidad           numeric not null default 0,
  unidad             text not null default '',
  acumulado_anterior numeric default 0,
  acumulado_total    numeric default 0,
  observacion_es     text,
  observacion_en     text,
  created_at         timestamptz default now()
);

create table public.ejecucion_items (
  id           uuid primary key default gen_random_uuid(),
  reporte_id   uuid not null references public.reportes_avance(id) on delete cascade,
  fecha        date not null,
  usuario_id   uuid not null references public.profiles(id),
  actividad_id uuid not null,
  item_id      text not null,
  area_id      uuid references public.areas(id),
  tecnico_id   uuid references public.profiles(id),
  created_at   timestamptz default now()
);

create table public.novedades_maquinaria (
  id            uuid primary key default gen_random_uuid(),
  reporte_id    uuid not null references public.reportes_avance(id) on delete cascade,
  fecha         date not null,
  usuario_id    uuid not null references public.profiles(id),
  maquinaria_id uuid not null references public.maquinaria(id),
  descripcion   text,
  hora_inicio   time,
  hora_fin      time,
  horas_standby numeric default 0,
  created_at    timestamptz default now()
);

create table public.suspensiones_clima (
  id             uuid primary key default gen_random_uuid(),
  reporte_id     uuid not null references public.reportes_avance(id) on delete cascade,
  fecha          date not null,
  usuario_id     uuid not null references public.profiles(id),
  hora_inicio    time,
  hora_fin       time,
  horas_perdidas numeric default 0,
  descripcion    text,
  created_at     timestamptz default now()
);

create table public.incidentes_seg (
  id              uuid primary key default gen_random_uuid(),
  reporte_id      uuid not null references public.reportes_avance(id) on delete cascade,
  fecha           date not null,
  usuario_id      uuid not null references public.profiles(id),
  tipo            text not null default 'sin_novedad'
                  check (tipo in ('sin_novedad','casi_accidente','incidente','accidente')),
  descripcion     text,
  medidas_tomadas text,
  area_id         uuid references public.areas(id),
  created_at      timestamptz default now()
);

create table public.aprobacion_informes (
  id           uuid primary key default gen_random_uuid(),
  reporte_id   uuid not null references public.reportes_avance(id) on delete cascade,
  aprobado_por uuid not null references public.profiles(id),
  estado       text not null check (estado in ('aprobado','rechazado')),
  version      int default 1,
  comentarios  text,
  created_at   timestamptz default now()
);

create table public.fotos_reporte (
  id           uuid primary key default gen_random_uuid(),
  reporte_id   uuid not null references public.reportes_avance(id) on delete cascade,
  fecha        date not null,
  usuario_id   uuid not null references public.profiles(id),
  actividad_id uuid,
  area_id      uuid references public.areas(id),
  storage_path text not null,
  descripcion  text,
  created_at   timestamptz default now()
);

create table public.calendario (
  id            uuid primary key default gen_random_uuid(),
  fecha         date not null unique,
  tipo          text not null check (tipo in ('festivo','no_laborable','laborable_especial')),
  descripcion   text,
  horas_jornada numeric,
  created_at    timestamptz default now()
);

create table public.bitacora_decisiones (
  id              uuid primary key default gen_random_uuid(),
  fecha           date not null,
  usuario_id      uuid not null references public.profiles(id),
  descripcion     text not null,
  impacto         text,
  especialidad_id uuid,
  area_id         uuid references public.areas(id),
  created_at      timestamptz default now()
);

create table public.log_cambios (
  id             uuid primary key default gen_random_uuid(),
  tabla_afectada text not null,
  registro_id    uuid,
  campo_cambiado text,
  valor_anterior text,
  valor_nuevo    text,
  usuario_id     uuid references public.profiles(id),
  motivo         text,
  created_at     timestamptz default now()
);

-- ============================================================
-- ÍNDICES  (rendimiento en consultas frecuentes)
-- ============================================================
create index idx_prog_fecha        on public.programaciones(fecha);
create index idx_prog_usuario      on public.programaciones(usuario_id);
create index idx_actprog_prog      on public.actividades_programadas(programacion_id);
create index idx_actprog_fecha     on public.actividades_programadas(fecha);
create index idx_persasig_fecha    on public.personal_asignado(fecha, documento_personal);
create index idx_rep_fecha         on public.reportes_avance(fecha);
create index idx_rep_usuario       on public.reportes_avance(usuario_id);
create index idx_rep_esp           on public.reportes_avance(especialidad_id);
create index idx_avance_fecha      on public.avance_diario(fecha);
create index idx_avance_act        on public.avance_diario(actividad_id);
create index idx_avance_area       on public.avance_diario(area_id);

-- ============================================================
-- RLS  (Row Level Security)
-- ============================================================
alter table public.profiles                 enable row level security;
alter table public.especialidades_actividades enable row level security;
alter table public.areas                    enable row level security;
alter table public.lideres                  enable row level security;
alter table public.personal                 enable row level security;
alter table public.subcontratistas          enable row level security;
alter table public.maquinaria               enable row level security;
alter table public.config_actividades       enable row level security;
alter table public.inventario_items         enable row level security;
alter table public.programaciones           enable row level security;
alter table public.actividades_programadas  enable row level security;
alter table public.personal_asignado        enable row level security;
alter table public.reportes_avance          enable row level security;
alter table public.asistencia_real          enable row level security;
alter table public.avance_diario            enable row level security;
alter table public.ejecucion_items          enable row level security;
alter table public.novedades_maquinaria     enable row level security;
alter table public.suspensiones_clima       enable row level security;
alter table public.incidentes_seg           enable row level security;
alter table public.aprobacion_informes      enable row level security;
alter table public.fotos_reporte            enable row level security;
alter table public.calendario               enable row level security;
alter table public.bitacora_decisiones      enable row level security;
alter table public.log_cambios              enable row level security;

-- Profiles
create policy "profiles_select" on public.profiles for select to authenticated using (true);
create policy "profiles_update_self" on public.profiles for update to authenticated using (auth.uid() = id);
create policy "profiles_admin" on public.profiles for all to authenticated using (get_my_role() = 'admin');

-- Catálogos (lectura = todos los autenticados, escritura = admin)
create policy "ea_select" on public.especialidades_actividades for select to authenticated using (true);
create policy "ea_admin"  on public.especialidades_actividades for all   to authenticated using (get_my_role() = 'admin');
create policy "ar_select" on public.areas        for select to authenticated using (true);
create policy "ar_admin"  on public.areas        for all   to authenticated using (get_my_role() = 'admin');
create policy "li_select" on public.lideres      for select to authenticated using (true);
create policy "li_admin"  on public.lideres      for all   to authenticated using (get_my_role() = 'admin');
create policy "pe_select" on public.personal     for select to authenticated using (true);
create policy "pe_admin"  on public.personal     for all   to authenticated using (get_my_role() = 'admin');
create policy "su_select" on public.subcontratistas for select to authenticated using (true);
create policy "su_admin"  on public.subcontratistas for all   to authenticated using (get_my_role() = 'admin');
create policy "ma_select" on public.maquinaria       for select to authenticated using (true);
create policy "ma_admin"  on public.maquinaria       for all   to authenticated using (get_my_role() = 'admin');
create policy "ca_select" on public.config_actividades for select to authenticated using (true);
create policy "ca_admin"  on public.config_actividades for all   to authenticated using (get_my_role() = 'admin');
create policy "ii_select" on public.inventario_items   for select to authenticated using (true);
create policy "ii_admin"  on public.inventario_items   for all   to authenticated using (get_my_role() = 'admin');
create policy "cal_select" on public.calendario        for select to authenticated using (true);
create policy "cal_admin"  on public.calendario        for all   to authenticated using (get_my_role() = 'admin');

-- Planeación
create policy "prog_select" on public.programaciones for select to authenticated using (true);
create policy "prog_write"  on public.programaciones for all to authenticated using (
  (auth.uid() = usuario_id or get_my_role() = 'admin')
  and get_my_role() in ('admin','lider','tecnico')
);
create policy "ap_select" on public.actividades_programadas for select to authenticated using (true);
create policy "ap_write"  on public.actividades_programadas for all to authenticated using (
  auth.uid() = usuario_id or get_my_role() = 'admin'
);
create policy "pa_select" on public.personal_asignado for select to authenticated using (true);
create policy "pa_write"  on public.personal_asignado for all to authenticated using (
  auth.uid() = usuario_id or get_my_role() = 'admin'
);

-- Reporte diario
create policy "rep_select" on public.reportes_avance for select to authenticated using (true);
create policy "rep_write"  on public.reportes_avance for all to authenticated using (
  auth.uid() = usuario_id or get_my_role() in ('admin','lider')
);
create policy "as_all"  on public.asistencia_real      for all to authenticated using (auth.uid() = usuario_id or get_my_role() in ('admin','lider'));
create policy "av_all"  on public.avance_diario        for all to authenticated using (auth.uid() = usuario_id or get_my_role() in ('admin','lider'));
create policy "ei_all"  on public.ejecucion_items      for all to authenticated using (auth.uid() = usuario_id or get_my_role() in ('admin','lider'));
create policy "nm_all"  on public.novedades_maquinaria for all to authenticated using (auth.uid() = usuario_id or get_my_role() in ('admin','lider'));
create policy "sc_all"  on public.suspensiones_clima   for all to authenticated using (auth.uid() = usuario_id or get_my_role() in ('admin','lider'));
create policy "is_all"  on public.incidentes_seg       for all to authenticated using (auth.uid() = usuario_id or get_my_role() in ('admin','lider'));
create policy "fr_all"  on public.fotos_reporte        for all to authenticated using (auth.uid() = usuario_id or get_my_role() in ('admin','lider'));
create policy "apro_select" on public.aprobacion_informes for select to authenticated using (true);
create policy "apro_write"  on public.aprobacion_informes for insert to authenticated with check (get_my_role() in ('admin','lider'));
create policy "bit_select"  on public.bitacora_decisiones for select to authenticated using (true);
create policy "bit_write"   on public.bitacora_decisiones for insert to authenticated with check (get_my_role() in ('admin','lider','tecnico'));
create policy "log_select"  on public.log_cambios for select to authenticated using (get_my_role() in ('admin','gerencia'));
create policy "log_insert"  on public.log_cambios for insert to authenticated with check (true);

-- ============================================================
-- DATOS INICIALES — Maquinaria
-- ============================================================
insert into public.maquinaria (tipo, item_id, nombre, estado) values
  ('motosierra','MS-001','Motosierra 001','activo'),
  ('motosierra','MS-002','Motosierra 002','activo'),
  ('motosierra','MS-003','Motosierra 003','activo'),
  ('motosierra','MS-004','Motosierra 004','activo'),
  ('motosierra','MS-005','Motosierra 005','activo'),
  ('motosierra','MS-006','Motosierra 006','activo'),
  ('motosierra','MS-007','Motosierra 007','activo'),
  ('motosierra','MS-008','Motosierra 008','activo'),
  ('chipeadora','CH-001','Chipeadora 001','activo');

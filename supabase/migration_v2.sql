-- Powerchina · PDS 360 · Migración v2.1
-- Ejecutar en Supabase → SQL Editor DESPUÉS del schema principal

-- 1. Columnas nuevas en config_actividades
alter table public.config_actividades
  add column if not exists tiene_meta boolean default true,
  add column if not exists es_medible  boolean default true;

-- 2. Especialidades asignadas a líderes
create table if not exists public.lider_especialidades (
  id              uuid primary key default gen_random_uuid(),
  lider_id        uuid not null references public.lideres(id) on delete cascade,
  especialidad_id uuid not null references public.especialidades_actividades(id) on delete cascade,
  created_at      timestamptz default now(),
  unique (lider_id, especialidad_id)
);
alter table public.lider_especialidades enable row level security;
create policy "le_select" on public.lider_especialidades for select to authenticated using (true);
create policy "le_admin"  on public.lider_especialidades for all    to authenticated using (get_my_role() = 'admin');

-- 3. Solicitudes de reporte de día anterior
create table if not exists public.solicitudes_reporte_pasado (
  id              uuid primary key default gen_random_uuid(),
  tecnico_id      uuid not null references public.profiles(id),
  tecnico_nombre  text not null,
  especialidad_id uuid not null,
  fecha_reporte   date not null,
  motivo          text not null,
  estado          text not null default 'pendiente'
                  check (estado in ('pendiente','aprobado','rechazado')),
  aprobado_por    uuid references public.profiles(id),
  aprobado_en     timestamptz,
  comentario      text,
  created_at      timestamptz default now()
);
alter table public.solicitudes_reporte_pasado enable row level security;
create policy "srp_select" on public.solicitudes_reporte_pasado for select to authenticated using (auth.uid() = tecnico_id or get_my_role() in ('admin','lider'));
create policy "srp_insert" on public.solicitudes_reporte_pasado for insert to authenticated with check (auth.uid() = tecnico_id);
create policy "srp_update" on public.solicitudes_reporte_pasado for update to authenticated using (get_my_role() in ('admin','lider'));

-- 4. Notificaciones en app
create table if not exists public.notificaciones (
  id          uuid primary key default gen_random_uuid(),
  usuario_id  uuid not null references public.profiles(id) on delete cascade,
  tipo        text not null,
  titulo      text not null,
  mensaje     text not null,
  leida       boolean default false,
  data        jsonb,
  created_at  timestamptz default now()
);
alter table public.notificaciones enable row level security;
create policy "notif_select" on public.notificaciones for select to authenticated using (auth.uid() = usuario_id);
create policy "notif_update" on public.notificaciones for update to authenticated using (auth.uid() = usuario_id);
create policy "notif_insert" on public.notificaciones for insert to authenticated with check (true);

-- 5. Índices
create index if not exists idx_notif_usuario on public.notificaciones(usuario_id, leida);
create index if not exists idx_srp_estado    on public.solicitudes_reporte_pasado(estado);

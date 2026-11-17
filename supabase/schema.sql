-- =====================================================================
--  Mi Día de Calidad — esquema de base de datos (Supabase / PostgreSQL)
--  Cómo usarlo: Supabase → SQL Editor → New query → pegar todo → Run.
--  Se puede ejecutar más de una vez sin dañar datos.
-- =====================================================================
--  Diseño: cada tabla guarda el registro completo en "data" (jsonb) y
--  expone columnas calculadas para consultas y reportes. Cada usuario
--  solo puede ver y modificar sus propios registros (Row Level Security).
-- =====================================================================

-- ---------- Tareas ----------
create table if not exists public.tasks (
  user_id     uuid not null default auth.uid() references auth.users(id) on delete cascade,
  id          text not null,
  updated_at  timestamptz not null default now(),
  deleted     boolean not null default false,
  data        jsonb not null default '{}'::jsonb,
  title          text generated always as (data->>'title') stored,
  status         text generated always as (data->>'status') stored,           -- pendiente | programada | en_curso | bloqueada | completada | pospuesta | cancelada
  clasificacion  text generated always as (data->>'clasificacion') stored,    -- inmediata | programable | espera | delegable | seguimiento | referencia
  priority       int  generated always as (case when data->>'priority' ~ '^[0-9]+$' then (data->>'priority')::int end) stored,  -- 1 crítica … 4 baja
  category       text generated always as (data->>'category') stored,
  area           text generated always as (data->>'area') stored,
  due_date       text generated always as (data->>'due_date') stored,         -- AAAA-MM-DD
  scheduled_date text generated always as (data->>'scheduled_date') stored,
  primary key (user_id, id)
);

-- ---------- Intramed: requerimientos, incidencias y soporte ----------
create table if not exists public.intramed (
  user_id     uuid not null default auth.uid() references auth.users(id) on delete cascade,
  id          text not null,
  updated_at  timestamptz not null default now(),
  deleted     boolean not null default false,
  data        jsonb not null default '{}'::jsonb,
  title     text generated always as (data->>'title') stored,
  tipo      text generated always as (data->>'tipo') stored,      -- reporte | necesidad | soporte  (origen: hospital | calidad | otra)
  status    text generated always as (data->>'status') stored,    -- recibido | en_analisis | enviado | en_desarrollo | en_validacion | implementado | cerrado
  hospital  text generated always as (data->>'hospital') stored,
  due_date  text generated always as (data->>'due_date') stored,
  primary key (user_id, id)
);

-- ---------- Notas de reunión con Intramed ----------
create table if not exists public.intramed_meetings (
  user_id uuid not null default auth.uid() references auth.users(id) on delete cascade,
  id text not null, updated_at timestamptz not null default now(), deleted boolean not null default false,
  data jsonb not null default '{}'::jsonb,
  date text generated always as (data->>'date') stored,
  primary key (user_id, id)
);

-- ---------- Reuniones imprevistas / bloqueos ----------
create table if not exists public.events (
  user_id uuid not null default auth.uid() references auth.users(id) on delete cascade,
  id text not null, updated_at timestamptz not null default now(), deleted boolean not null default false,
  data jsonb not null default '{}'::jsonb,
  date text generated always as (data->>'date') stored,
  primary key (user_id, id)
);

-- ---------- Configuración por día (sábado de descanso, jornada especial, visita, comida reducida…) ----------
create table if not exists public.days (
  user_id uuid not null default auth.uid() references auth.users(id) on delete cascade,
  id text not null, updated_at timestamptz not null default now(), deleted boolean not null default false,
  data jsonb not null default '{}'::jsonb,
  date text generated always as (data->>'date') stored,
  tipo text generated always as (data->>'tipo') stored,
  primary key (user_id, id)
);

-- ---------- Cierres de jornada ----------
create table if not exists public.closures (
  user_id uuid not null default auth.uid() references auth.users(id) on delete cascade,
  id text not null, updated_at timestamptz not null default now(), deleted boolean not null default false,
  data jsonb not null default '{}'::jsonb,
  date text generated always as (data->>'date') stored,
  primary key (user_id, id)
);

-- ---------- Imprevistos (tabla 'bomberazos'; guarda los cambios para poder revertir) ----------
create table if not exists public.bomberazos (
  user_id uuid not null default auth.uid() references auth.users(id) on delete cascade,
  id text not null, updated_at timestamptz not null default now(), deleted boolean not null default false,
  data jsonb not null default '{}'::jsonb,
  date  text generated always as (data->>'date') stored,
  state text generated always as (data->>'state') stored,   -- aplicado | aceptado | revertido
  primary key (user_id, id)
);

-- ---------- Ajustes (un registro por usuario, id = 'settings') ----------
create table if not exists public.settings (
  user_id uuid not null default auth.uid() references auth.users(id) on delete cascade,
  id text not null, updated_at timestamptz not null default now(), deleted boolean not null default false,
  data jsonb not null default '{}'::jsonb,
  primary key (user_id, id)
);


-- ---------- Seguimientos (macro actividades con etapas y tiempo diario reservado) ----------
create table if not exists public.projects (
  user_id uuid not null default auth.uid() references auth.users(id) on delete cascade,
  id text not null, updated_at timestamptz not null default now(), deleted boolean not null default false,
  data jsonb not null default '{}'::jsonb,
  name   text generated always as (data->>'name') stored,
  status text generated always as (data->>'status') stored,     -- activo | pausa | terminado
  primary key (user_id, id)
);

-- ---------- Avance documental (documentos y formatos con flujo de revisiones) ----------
create table if not exists public.docs (
  user_id uuid not null default auth.uid() references auth.users(id) on delete cascade,
  id text not null, updated_at timestamptz not null default now(), deleted boolean not null default false,
  data jsonb not null default '{}'::jsonb,
  code       text generated always as (data->>'code') stored,
  name       text generated always as (data->>'name') stored,
  flow       text generated always as (data->>'flow') stored,     -- documento_nuevo | documento_actualizacion | formato_ece
  stage      text generated always as (data->>'stage') stored,    -- etapa actual del flujo
  owner      text generated always as (data->>'owner') stored,
  project_id text generated always as (data->>'project_id') stored,
  due_date   text generated always as (data->>'due_date') stored,
  primary key (user_id, id)
);

-- ---------- Seguridad, índices y tiempo real ----------
do $$
declare t text;
begin
  foreach t in array array['tasks','intramed','intramed_meetings','events','days','closures','bomberazos','settings','projects','docs'] loop
    execute format('alter table public.%I enable row level security', t);
    execute format('grant select, insert, update on public.%I to authenticated', t);
    execute format('revoke all on public.%I from anon', t);
    execute format('drop policy if exists "propios_select" on public.%I', t);
    execute format('drop policy if exists "propios_insert" on public.%I', t);
    execute format('drop policy if exists "propios_update" on public.%I', t);
    execute format('create policy "propios_select" on public.%I for select to authenticated using (user_id = auth.uid())', t);
    execute format('create policy "propios_insert" on public.%I for insert to authenticated with check (user_id = auth.uid())', t);
    execute format('create policy "propios_update" on public.%I for update to authenticated using (user_id = auth.uid()) with check (user_id = auth.uid())', t);
    -- Sin política de DELETE: los borrados son lógicos (deleted = true) para que se sincronicen entre dispositivos.
    execute format('create index if not exists %I on public.%I (user_id, updated_at)', t || '_upd_idx', t);
    begin
      execute format('alter publication supabase_realtime add table public.%I', t);
    exception when duplicate_object then null;
    end;
  end loop;
end $$;

-- Vista de apoyo para reportes (solo lectura, respeta RLS del usuario)
create or replace view public.v_tareas_abiertas with (security_invoker = true) as
  select user_id, id, title, status, clasificacion, priority, category, area, due_date, scheduled_date, updated_at
  from public.tasks
  where not deleted and coalesce(status, '') not in ('completada','cancelada');
grant select on public.v_tareas_abiertas to authenticated;

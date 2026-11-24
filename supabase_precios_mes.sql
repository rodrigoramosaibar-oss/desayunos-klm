-- Precios por mes (Escandallo Desayunos)
-- Ejecutar UNA vez en Supabase > SQL Editor. Después, pulsar ↻ en la app.
create table if not exists public.precios_mes (
  mes            text    not null,            -- 'YYYY-MM'
  codigo         text    not null,
  precio_compra  numeric not null default 0,
  factor         numeric not null default 1,
  unidad_conteo  text,
  updated_at     timestamptz not null default now(),
  primary key (mes, codigo)
);
alter table public.precios_mes enable row level security;
-- Mismo criterio que el resto de tablas: usuarios autenticados leen y escriben.
-- (Si tus otras tablas restringen la escritura por rol, replica aquí esa política.)
drop policy if exists precios_mes_sel on public.precios_mes;
drop policy if exists precios_mes_ins on public.precios_mes;
drop policy if exists precios_mes_upd on public.precios_mes;
create policy precios_mes_sel on public.precios_mes for select to authenticated using (true);
create policy precios_mes_ins on public.precios_mes for insert to authenticated with check (true);
create policy precios_mes_upd on public.precios_mes for update to authenticated using (true) with check (true);

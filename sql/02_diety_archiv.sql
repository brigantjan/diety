-- =====================================================================
-- SUPTel – Diéty: archív mesiacov (súčty po ľuďoch), schéma public
-- Spúšťa sa raz. Jeden riadok = jeden mesiac (pár kB), záťaž je zanedbateľná.
-- =====================================================================
create table if not exists public.diety_archiv (
  rok         int  not null check (rok between 2020 and 2100),
  mesiac      int  not null check (mesiac between 1 and 12),
  data        jsonb not null,
  updated_at  timestamptz not null default now(),
  updated_by  uuid default auth.uid(),
  primary key (rok, mesiac)
);

create or replace function public.diety_archiv_pred_zmenou()
returns trigger language plpgsql security definer set search_path = public as $$
begin
  new.updated_at := now();
  new.updated_by := auth.uid();
  return new;
end $$;

drop trigger if exists diety_archiv_stamp on public.diety_archiv;
create trigger diety_archiv_stamp before insert or update on public.diety_archiv
  for each row execute function public.diety_archiv_pred_zmenou();

alter table public.diety_archiv enable row level security;

drop policy if exists diety_archiv_citanie on public.diety_archiv;
create policy diety_archiv_citanie on public.diety_archiv
  for select to authenticated using (public.diety_ma_pristup());
drop policy if exists diety_archiv_vlozenie on public.diety_archiv;
create policy diety_archiv_vlozenie on public.diety_archiv
  for insert to authenticated with check (public.diety_ma_pristup());
drop policy if exists diety_archiv_zmena on public.diety_archiv;
create policy diety_archiv_zmena on public.diety_archiv
  for update to authenticated using (public.diety_ma_pristup()) with check (public.diety_ma_pristup());

revoke all on public.diety_archiv from anon, authenticated;
grant select, insert (rok, mesiac, data), update (rok, mesiac, data) on public.diety_archiv to authenticated;
revoke execute on function public.diety_archiv_pred_zmenou() from anon, authenticated, public;

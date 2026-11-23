-- =====================================================================
-- SUPTel – Diéty: tabuľky v Supabase projekte "dispecing-dev" (schéma public)
-- Spúšťa sa raz (SQL Editor v Supabase). Nič existujúce nemení.
-- =====================================================================

-- Kto má prístup k Diétam (používatelia z Supabase Auth – rovnaké účty ako Objednávky)
create table if not exists public.diety_pristup (
  user_id   uuid primary key references auth.users(id) on delete cascade,
  poznamka  text,
  created_at timestamptz not null default now()
);

-- Nastavenia: zamestnanci, sadzby, časy (jeden JSON dokument)
create table if not exists public.diety_nastavenia (
  id          text primary key default 'main',
  data        jsonb not null,
  updated_at  timestamptz not null default now(),
  updated_by  uuid default auth.uid()
);

-- História – každá zmena nastavení sa sem automaticky odloží (záloha)
create table if not exists public.diety_historia (
  id          bigserial primary key,
  data        jsonb not null,
  platne_do   timestamptz not null default now(),
  zmenil      uuid
);

create or replace function public.diety_ma_pristup()
returns boolean language sql stable security definer set search_path = public as $$
  select exists (select 1 from public.diety_pristup where user_id = auth.uid());
$$;

create or replace function public.diety_pred_zmenou()
returns trigger language plpgsql security definer set search_path = public as $$
begin
  insert into public.diety_historia(data, zmenil) values (old.data, old.updated_by);
  new.updated_at := now();
  new.updated_by := auth.uid();
  return new;
end $$;

drop trigger if exists diety_nastavenia_historia on public.diety_nastavenia;
create trigger diety_nastavenia_historia before update on public.diety_nastavenia
  for each row execute function public.diety_pred_zmenou();

alter table public.diety_pristup    enable row level security;
alter table public.diety_nastavenia enable row level security;
alter table public.diety_historia   enable row level security;

drop policy if exists diety_pristup_vlastny on public.diety_pristup;
create policy diety_pristup_vlastny on public.diety_pristup
  for select to authenticated using (user_id = auth.uid());

drop policy if exists diety_nastavenia_citanie on public.diety_nastavenia;
create policy diety_nastavenia_citanie on public.diety_nastavenia
  for select to authenticated using (public.diety_ma_pristup());

drop policy if exists diety_nastavenia_zmena on public.diety_nastavenia;
create policy diety_nastavenia_zmena on public.diety_nastavenia
  for update to authenticated using (public.diety_ma_pristup()) with check (public.diety_ma_pristup());

drop policy if exists diety_historia_citanie on public.diety_historia;
create policy diety_historia_citanie on public.diety_historia
  for select to authenticated using (public.diety_ma_pristup());

revoke all on public.diety_pristup, public.diety_nastavenia, public.diety_historia from anon;
revoke all on public.diety_pristup, public.diety_nastavenia, public.diety_historia from authenticated;
grant select on public.diety_pristup to authenticated;
grant select, update (data) on public.diety_nastavenia to authenticated;
grant select on public.diety_historia to authenticated;
revoke execute on function public.diety_ma_pristup() from anon;
revoke execute on function public.diety_pred_zmenou() from anon, authenticated;

-- ---------------------------------------------------------------------
-- Pridelenie prístupu (doplň e-mail, ktorým sa prihlasuješ do Objednávok):
--   insert into public.diety_pristup(user_id, poznamka)
--   select id, 'Ján Brigant' from auth.users where email = 'TVOJ@EMAIL.sk'
--   on conflict do nothing;
-- Počiatočné nastavenia (zamestnanci, sadzby) sa nahrajú zvlášť – nie sú v repozitári,
-- pretože obsahujú osobné údaje.
-- ---------------------------------------------------------------------

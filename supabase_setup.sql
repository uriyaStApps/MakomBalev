-- MakomBalev: schema makom_balev in the shared uriyaStApps Supabase project.
-- Sign-in with Google + allow-list (managed by admins from the site). Safe to re-run.
-- Tables from Excel: every uploaded sheet becomes a "dataset" (its columns) + rows (one JSON per row).

create schema if not exists makom_balev;
grant usage on schema makom_balev to anon, authenticated;

-- ---------- usage log (insert-only from the browser) ----------
create table if not exists makom_balev.usage_logs (
  id bigint generated always as identity primary key,
  visited_at timestamptz not null default now(),
  visitor_name text,
  action text not null default 'page_view'
);
alter table makom_balev.usage_logs enable row level security;
drop policy if exists "anyone can insert logs" on makom_balev.usage_logs;
create policy "anyone can insert logs" on makom_balev.usage_logs for insert to anon, authenticated with check (true);
grant insert on makom_balev.usage_logs to anon, authenticated;

-- ---------- allowed users (managed from the site's "משתמשים" tab by admins) ----------
create table if not exists makom_balev.allowed_users (
  email text primary key,
  full_name text,
  is_admin boolean not null default false,
  added_by text default (auth.jwt() ->> 'email'),
  added_at timestamptz not null default now()
);
insert into makom_balev.allowed_users (email, full_name, is_admin) values
  ('uriya19@gmail.com', 'אוריה', true),
  ('siegelyoel@gmail.com', 'יואל זיגל', true)
  on conflict (email) do update set is_admin = true;

create or replace function makom_balev.is_allowed() returns boolean
  language sql stable security definer set search_path = '' as $$
  select exists (select 1 from makom_balev.allowed_users
                 where lower(email) = lower(coalesce(auth.jwt() ->> 'email', '')))
$$;
create or replace function makom_balev.is_admin() returns boolean
  language sql stable security definer set search_path = '' as $$
  select exists (select 1 from makom_balev.allowed_users
                 where is_admin and lower(email) = lower(coalesce(auth.jwt() ->> 'email', '')))
$$;
revoke all on function makom_balev.is_allowed() from public;
revoke all on function makom_balev.is_admin() from public;
grant execute on function makom_balev.is_allowed() to authenticated;
grant execute on function makom_balev.is_admin() to authenticated;

alter table makom_balev.allowed_users enable row level security;
drop policy if exists "see own row or all if admin" on makom_balev.allowed_users;
drop policy if exists "admins add users" on makom_balev.allowed_users;
drop policy if exists "admins edit users" on makom_balev.allowed_users;
drop policy if exists "admins remove users" on makom_balev.allowed_users;
create policy "see own row or all if admin" on makom_balev.allowed_users for select to authenticated
  using (lower(email) = lower(auth.jwt() ->> 'email') or makom_balev.is_admin());
create policy "admins add users" on makom_balev.allowed_users for insert to authenticated with check (makom_balev.is_admin());
create policy "admins edit users" on makom_balev.allowed_users for update to authenticated using (makom_balev.is_admin()) with check (makom_balev.is_admin());
create policy "admins remove users" on makom_balev.allowed_users for delete to authenticated using (makom_balev.is_admin());
grant select, insert, update, delete on makom_balev.allowed_users to authenticated;

-- ---------- datasets (one per uploaded Excel sheet) ----------
-- columns = [{"key":"c1","label":"שם","type":"text|number|date"}, ...]
create table if not exists makom_balev.datasets (
  id bigint generated always as identity primary key,
  name text not null unique,
  columns jsonb not null default '[]'::jsonb,
  source_file text,
  created_by text default (auth.jwt() ->> 'email'),
  created_at timestamptz not null default now(),
  updated_by text default (auth.jwt() ->> 'email'),
  updated_at timestamptz not null default now()
);

-- ---------- rows ----------
create table if not exists makom_balev.dataset_rows (
  id bigint generated always as identity primary key,
  dataset_id bigint not null references makom_balev.datasets(id) on delete cascade,
  data jsonb not null default '{}'::jsonb,
  origin text not null default 'site' check (origin in ('excel', 'site')),
  created_by text default (auth.jwt() ->> 'email'),
  created_at timestamptz not null default now(),
  updated_by text default (auth.jwt() ->> 'email'),
  updated_at timestamptz not null default now()
);
create index if not exists dataset_rows_dataset_idx on makom_balev.dataset_rows (dataset_id, id);

-- ---------- change history (for operational reports later) ----------
create table if not exists makom_balev.row_history (
  id bigint generated always as identity primary key,
  dataset_id bigint,
  row_id bigint,
  action text not null,            -- update | delete
  old_data jsonb,
  new_data jsonb,
  changed_by text default (auth.jwt() ->> 'email'),
  changed_at timestamptz not null default now()
);

create or replace function makom_balev.touch_row() returns trigger
  language plpgsql security definer set search_path = '' as $$
begin
  if tg_op = 'UPDATE' then
    new.updated_at := now();
    new.updated_by := coalesce(auth.jwt() ->> 'email', new.updated_by);
    if new.data is distinct from old.data then
      insert into makom_balev.row_history (dataset_id, row_id, action, old_data, new_data)
        values (old.dataset_id, old.id, 'update', old.data, new.data);
    end if;
    return new;
  else
    insert into makom_balev.row_history (dataset_id, row_id, action, old_data)
      values (old.dataset_id, old.id, 'delete', old.data);
    return old;
  end if;
end $$;
drop trigger if exists dataset_rows_touch on makom_balev.dataset_rows;
create trigger dataset_rows_touch before update or delete on makom_balev.dataset_rows
  for each row execute function makom_balev.touch_row();

create or replace function makom_balev.touch_dataset() returns trigger
  language plpgsql set search_path = '' as $$
begin
  new.updated_at := now();
  new.updated_by := coalesce(auth.jwt() ->> 'email', new.updated_by);
  return new;
end $$;
drop trigger if exists datasets_touch on makom_balev.datasets;
create trigger datasets_touch before update on makom_balev.datasets
  for each row execute function makom_balev.touch_dataset();

-- ---------- access: allowed users only; deleting a whole table is admin-only ----------
alter table makom_balev.datasets enable row level security;
alter table makom_balev.dataset_rows enable row level security;
alter table makom_balev.row_history enable row level security;

drop policy if exists "allowed read datasets" on makom_balev.datasets;
drop policy if exists "allowed add datasets" on makom_balev.datasets;
drop policy if exists "allowed edit datasets" on makom_balev.datasets;
drop policy if exists "admins remove datasets" on makom_balev.datasets;
create policy "allowed read datasets" on makom_balev.datasets for select to authenticated using (makom_balev.is_allowed());
create policy "allowed add datasets" on makom_balev.datasets for insert to authenticated with check (makom_balev.is_allowed());
create policy "allowed edit datasets" on makom_balev.datasets for update to authenticated using (makom_balev.is_allowed()) with check (makom_balev.is_allowed());
create policy "admins remove datasets" on makom_balev.datasets for delete to authenticated using (makom_balev.is_admin());

drop policy if exists "allowed read rows" on makom_balev.dataset_rows;
drop policy if exists "allowed add rows" on makom_balev.dataset_rows;
drop policy if exists "allowed edit rows" on makom_balev.dataset_rows;
drop policy if exists "allowed remove rows" on makom_balev.dataset_rows;
create policy "allowed read rows" on makom_balev.dataset_rows for select to authenticated using (makom_balev.is_allowed());
create policy "allowed add rows" on makom_balev.dataset_rows for insert to authenticated with check (makom_balev.is_allowed());
create policy "allowed edit rows" on makom_balev.dataset_rows for update to authenticated using (makom_balev.is_allowed()) with check (makom_balev.is_allowed());
create policy "allowed remove rows" on makom_balev.dataset_rows for delete to authenticated using (makom_balev.is_allowed());

drop policy if exists "allowed read history" on makom_balev.row_history;
create policy "allowed read history" on makom_balev.row_history for select to authenticated using (makom_balev.is_allowed());

grant select, insert, update, delete on makom_balev.datasets to authenticated;
grant select, insert, update, delete on makom_balev.dataset_rows to authenticated;
grant select on makom_balev.row_history to authenticated;
-- anon gets nothing on the data tables

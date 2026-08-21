-- ============================================================================
-- Delsport UK Booking System — Supabase schema
-- ============================================================================
-- Run this in Supabase: Project → SQL Editor → New query → paste → Run.
--
-- Safe to re-run: this drops any of these tables if they already exist
-- first, then rebuilds them fresh. Only do this if you haven't put real
-- data into these specific tables yet — if you have, stop and tell me
-- before running this, so we handle it differently.
--
-- This replaces the current storage (window.storage) with real Postgres
-- tables, and replaces the custom username/password login with Supabase's
-- own authentication — which checks logins server-side, not in the browser.
--
-- WHO CAN ACCESS WHAT: every table below is locked down so that only
-- someone who has successfully logged in (an "authenticated" Supabase user)
-- can read or write anything. Nobody else — including someone who finds
-- the URL, or opens dev tools — can touch the data, because Postgres
-- itself enforces this, not just the app's interface.
-- ============================================================================

drop table if exists bookings cascade;
drop table if exists blackouts cascade;
drop table if exists facilities cascade;
drop table if exists members cascade;
drop table if exists sites cascade;
drop table if exists profiles cascade;

-- ---------- sites ----------
create table sites (
  id uuid primary key default gen_random_uuid(),
  name text not null,
  address text,
  contact text,
  logo_url text,
  created_at timestamptz not null default now()
);

-- ---------- facilities ----------
create table facilities (
  id uuid primary key default gen_random_uuid(),
  site_id uuid not null references sites(id) on delete cascade,
  name text not null,
  type text,
  rate numeric(10,2) default 0,
  capacity int not null default 1,
  min_notice_hours int not null default 0,
  max_advance_days int not null default 0,
  custom_fields jsonb not null default '[]', -- [{id,label,type,options}]
  created_at timestamptz not null default now()
);

-- ---------- members ----------
create table members (
  id uuid primary key default gen_random_uuid(),
  name text not null,
  company text,
  email text,
  phone text,
  created_at timestamptz not null default now()
);

-- ---------- bookings ----------
create table bookings (
  id uuid primary key default gen_random_uuid(),
  facility_id uuid not null references facilities(id) on delete cascade,
  member_id uuid references members(id) on delete set null,
  date date not null,
  start_time time not null,
  end_time time not null,
  purpose text,
  price numeric(10,2) default 0,
  status text not null default 'confirmed' check (status in ('confirmed','cancelled','pending','declined')), -- pending/declined kept only for legacy data
  notes text,
  spaces int not null default 1,
  recurring_id uuid, -- groups bookings created together via weekly-repeat or pick-specific-dates
  hirer_name text,
  hirer_contact text,
  company text,
  custom_values jsonb not null default '{}',
  created_at timestamptz not null default now()
);
create index bookings_facility_date_idx on bookings(facility_id, date);
create index bookings_date_idx on bookings(date);
create index bookings_member_idx on bookings(member_id);

-- ---------- blackout / closure rules ----------
create table blackouts (
  id uuid primary key default gen_random_uuid(),
  label text not null,
  scope text not null check (scope in ('all','sites','facility')),
  site_ids uuid[] default '{}',       -- used when scope = 'sites'
  facility_id uuid references facilities(id) on delete cascade, -- used when scope = 'facility'
  start_date date not null,
  end_date date not null,
  all_day boolean not null default true,
  days int[] default '{}',            -- 0=Sun..6=Sat, used when all_day = false
  start_time time,
  end_time time,
  created_at timestamptz not null default now()
);

-- ---------- team member profile (display name shown in the sidebar) ----------
-- Supabase Auth handles the actual login (email + password) in its own
-- built-in auth.users table. This table just attaches a friendly display
-- name to each logged-in account, since auth.users doesn't have one.
create table profiles (
  id uuid primary key references auth.users(id) on delete cascade,
  display_name text not null,
  created_at timestamptz not null default now()
);

-- ============================================================================
-- Row Level Security — the actual access control.
-- Every table: only logged-in (authenticated) users can do anything at all.
-- Nobody logged out, and nobody without an account, can read or write.
-- ============================================================================

alter table sites enable row level security;
alter table facilities enable row level security;
alter table members enable row level security;
alter table bookings enable row level security;
alter table blackouts enable row level security;
alter table profiles enable row level security;

create policy "authenticated full access" on sites
  for all using (auth.role() = 'authenticated') with check (auth.role() = 'authenticated');
create policy "authenticated full access" on facilities
  for all using (auth.role() = 'authenticated') with check (auth.role() = 'authenticated');
create policy "authenticated full access" on members
  for all using (auth.role() = 'authenticated') with check (auth.role() = 'authenticated');
create policy "authenticated full access" on bookings
  for all using (auth.role() = 'authenticated') with check (auth.role() = 'authenticated');
create policy "authenticated full access" on blackouts
  for all using (auth.role() = 'authenticated') with check (auth.role() = 'authenticated');

-- profiles: everyone logged in can view names (so the sidebar can show
-- "who's logged in"), but you can only edit your own.
create policy "authenticated can view profiles" on profiles
  for select using (auth.role() = 'authenticated');
create policy "users manage own profile" on profiles
  for insert with check (auth.uid() = id);
create policy "users update own profile" on profiles
  for update using (auth.uid() = id);

-- ============================================================================
-- Next steps (see SUPABASE_MIGRATION_GUIDE.md for the full walkthrough):
-- 1. Create your Supabase project.
-- 2. Run this whole file in the SQL Editor.
-- 3. In Authentication → Users, manually create an account for yourself and
--    your business partner (email + password) — this replaces the old
--    username/password screen with Supabase's own real login.
-- 4. Add a matching row to `profiles` for each account (id = their auth user
--    id, display_name = their name).
-- 5. Set up the free keep-alive ping (see SUPABASE_MIGRATION_GUIDE.md,
--    Step 4.5) so the free tier never auto-pauses — costs nothing, takes
--    about 5 minutes.
-- 6. Hand this project's URL + API key to whoever is doing the app rebuild
--    (Claude Code, or a developer) to connect the frontend to it.
-- ============================================================================

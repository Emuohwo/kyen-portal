-- ============================================================
-- KYEN PRODUCTS — Sales Hub Database Schema
-- Run this entire file in Supabase → SQL Editor → Run
-- ============================================================

-- 1. Config (single row, id always = 1)
create table if not exists config (
  id               integer primary key default 1,
  admin_password   text    not null default 'kyen2024',
  visit_benchmark  integer not null default 5,
  prices           jsonb   not null default '{}'
);
-- Seed one default row so the app always finds config
insert into config (id) values (1) on conflict (id) do nothing;

-- 2. Team members
create table if not exists team (
  id          text primary key,
  name        text not null,
  username    text not null unique,
  password    text not null,
  created_at  timestamptz default now()
);

-- 3. Sales transactions
create table if not exists sales (
  id          text primary key,
  rep_id      text,
  rep_name    text,
  date        date not null,
  channel     text,
  sku         text,
  qty         numeric,
  price       numeric,
  amount      numeric,
  created_at  timestamptz default now()
);
create index if not exists sales_date_idx    on sales (date);
create index if not exists sales_rep_id_idx  on sales (rep_id);

-- 4. Outlet visits
create table if not exists visits (
  id          text primary key,
  rep_id      text,
  rep_name    text,
  date        date not null,
  outlets     jsonb not null default '[]',
  created_at  timestamptz default now()
);
create index if not exists visits_date_idx   on visits (date);
create index if not exists visits_rep_id_idx on visits (rep_id);

-- 5. Monthly targets (one row per rep per month)
create table if not exists targets (
  id              text primary key,   -- format: repId_year_month
  rep_id          text not null,
  year            integer not null,
  month           integer not null,
  value_ngn       numeric default 0,
  volume_units    integer default 0,
  visits_target   integer default 0,
  unique (rep_id, year, month)
);

-- ============================================================
-- Row Level Security — disable for now (simple password auth)
-- Enable later when you add proper Supabase Auth
-- ============================================================
alter table config  disable row level security;
alter table team    disable row level security;
alter table sales   disable row level security;
alter table visits  disable row level security;
alter table targets disable row level security;

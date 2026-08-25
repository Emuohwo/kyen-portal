-- ============================================================
-- KYEN PRODUCTS — Sales Hub Schema v2
-- Run this full file in Supabase → SQL Editor → Run
-- Safe to re-run: uses IF NOT EXISTS and ON CONFLICT
-- ============================================================

-- 1. Config
create table if not exists config (
  id               integer primary key default 1,
  admin_password   text    not null default 'kyen2024',
  visit_benchmark  integer not null default 5,
  prices           jsonb   not null default '{}'
);
insert into config (id) values (1) on conflict (id) do nothing;

-- 2. Team members
create table if not exists team (
  id          text primary key,
  name        text not null,
  username    text not null unique,
  password    text not null,
  created_at  timestamptz default now()
);

-- 3. Dynamic SKUs (seeded from app on first load)
create table if not exists skus (
  sku         text primary key,
  name        text not null,
  cat         text not null,
  is_active   boolean default true,
  sort_order  integer default 99,
  created_at  timestamptz default now()
);

-- 4. Customers / Outlets
create table if not exists customers (
  id            text primary key,
  name          text not null,
  type          text,           -- Supermarket | Retail Store | Wholesaler
  address       text,
  contact_name  text,
  contact_phone text,
  city          text,
  zone          text,
  created_by    text,
  created_at    timestamptz default now()
);
create index if not exists customers_name_idx on customers (name);

-- 5. Sales transactions
create table if not exists sales (
  id            text primary key,
  rep_id        text,
  rep_name      text,
  date          date not null,
  channel       text,
  customer_id   text references customers(id) on delete set null,
  customer_name text,
  sku           text,
  qty           numeric,
  price         numeric,
  amount        numeric,
  created_at    timestamptz default now()
);
create index if not exists sales_date_idx   on sales (date);
create index if not exists sales_rep_idx    on sales (rep_id);
create index if not exists sales_cust_idx   on sales (customer_id);

-- 6. Outlet visits
create table if not exists visits (
  id           text primary key,
  rep_id       text,
  rep_name     text,
  date         date not null,
  customer_id  text references customers(id) on delete set null,
  visit_type   text default 'First Visit',  -- First Visit | Follow-up | Merchandising
  notes        text,
  created_at   timestamptz default now()
);
create index if not exists visits_date_idx on visits (date);
create index if not exists visits_rep_idx  on visits (rep_id);
create index if not exists visits_cust_idx on visits (customer_id);

-- 7. Monthly targets (per-rep, per-SKU volume + overall value + visits)
create table if not exists targets (
  id              text primary key,   -- repId_year_month
  rep_id          text not null,
  year            integer not null,
  month           integer not null,
  value_ngn       numeric default 0,
  visits_target   integer default 0,
  sku_targets     jsonb default '{}', -- { [sku]: targetQty }
  unique (rep_id, year, month)
);

-- 8. Visit plans
create table if not exists visit_plans (
  id            text primary key,
  rep_id        text,
  rep_name      text,
  planned_date  date not null,
  customer_id   text references customers(id) on delete set null,
  customer_name text,
  visit_type    text default 'First Visit',
  notes         text,
  status        text default 'planned',  -- planned | completed | cancelled
  created_at    timestamptz default now()
);
create index if not exists vplans_date_idx on visit_plans (planned_date);
create index if not exists vplans_rep_idx  on visit_plans (rep_id);

-- Disable RLS for now (simple password auth)
alter table config       disable row level security;
alter table team         disable row level security;
alter table skus         disable row level security;
alter table customers    disable row level security;
alter table sales        disable row level security;
alter table visits       disable row level security;
alter table targets      disable row level security;
alter table visit_plans  disable row level security;

-- Enable Realtime on key tables
alter publication supabase_realtime add table sales;
alter publication supabase_realtime add table visits;
alter publication supabase_realtime add table visit_plans;

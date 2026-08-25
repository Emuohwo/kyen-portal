-- ============================================================
-- 2. Config
-- ============================================================

create table config (
    id integer primary key default 1,
    admin_password text not null default 'kyen2024',
    visit_benchmark integer not null default 5,
    prices jsonb not null default '{}'
);

insert into config (id)
values (1)
on conflict (id) do nothing;


-- ============================================================
-- 3. Team members
-- ============================================================

create table team (
    id text primary key,
    name text not null,
    username text not null unique,
    password text not null,
    created_at timestamptz default now()
);


-- ============================================================
-- 4. Dynamic SKUs
-- ============================================================

create table skus (
    sku text primary key,
    name text not null,
    cat text not null,
    is_active boolean default true,
    sort_order integer default 99,
    created_at timestamptz default now()
);


-- ============================================================
-- 5. Customers / Outlets
-- ============================================================

create table customers (
    id text primary key,
    name text not null,
    type text,
    address text,
    "contactName" text,
    "contactPhone" text,
    city text,
    zone text,
    "createdBy" text,
    "createdAt" timestamptz default now()
);

create index customers_name_idx
on customers (name);


-- ============================================================
-- 6. Sales transactions
-- ============================================================

create table sales (
    id text primary key,
    "repId" text,
    "repName" text,
    date date not null,
    channel text,
    "customerId" text references customers(id) on delete set null,
    "customerName" text,
    sku text,
    qty numeric,
    price numeric,
    amount numeric,
    "createdAt" timestamptz default now()
);

create index sales_date_idx
on sales (date);

create index sales_rep_idx
on sales ("repId");

create index sales_cust_idx
on sales ("customerId");


-- ============================================================
-- 7. Outlet visits
-- ============================================================

create table visits (
    id text primary key,
    "repId" text,
    "repName" text,
    date date not null,
    "customerId" text references customers(id) on delete set null,
    "visitType" text default 'First Visit',
    notes text,
    "createdAt" timestamptz default now()
);

create index visits_date_idx
on visits (date);

create index visits_rep_idx
on visits ("repId");

create index visits_cust_idx
on visits ("customerId");


-- ============================================================
-- 8. Monthly targets
-- ============================================================

create table targets (
    id text primary key,
    "repId" text not null,
    year integer not null,
    month integer not null,
    "valueNgn" numeric default 0,
    "visitsTarget" integer default 0,
    "skuTargets" jsonb default '{}',
    unique ("repId", year, month)
);


-- ============================================================
-- 9. Visit plans
-- ============================================================

create table visit_plans (
    id text primary key,
    "repId" text,
    "repName" text,
    "plannedDate" date not null,
    "customerId" text references customers(id) on delete set null,
    "customerName" text,
    "visitType" text default 'First Visit',
    notes text,
    status text default 'planned',
    "createdAt" timestamptz default now()
);

create index vplans_date_idx
on visit_plans ("plannedDate");

create index vplans_rep_idx
on visit_plans ("repId");


-- ============================================================
-- 10. Disable RLS
-- ============================================================

alter table config disable row level security;
alter table team disable row level security;
alter table skus disable row level security;
alter table customers disable row level security;
alter table sales disable row level security;
alter table visits disable row level security;
alter table targets disable row level security;
alter table visit_plans disable row level security;


-- ============================================================
-- 11. Enable Realtime
-- ============================================================

alter publication supabase_realtime add table sales;
alter publication supabase_realtime add table visits;
alter publication supabase_realtime add table visit_plans;
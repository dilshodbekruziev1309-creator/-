-- ============================================================
-- OMBOR HISOBI — TO'LIQ MA'LUMOTLAR BAZASI SXEMASI
-- Yangi firma uchun: bo'sh Supabase loyihasida SQL Editor'ga
-- shu faylning BUTUN matnini joylashtirib, bitta marta Run qiling.
-- ============================================================

create extension if not exists "pgcrypto";

-- ---------- Ombor (tayyor mahsulot) ----------
create table products (
  id uuid primary key default gen_random_uuid(),
  name text not null,
  article text,
  unit text default 'dona',
  price numeric default 0,
  qty numeric default 0,
  cost_price numeric default 0,
  created_at timestamptz default now()
);

create table stock_movements (
  id uuid primary key default gen_random_uuid(),
  product_id uuid references products(id) on delete cascade,
  date date not null,
  qty numeric not null,
  type text,
  note text,
  created_at timestamptz default now()
);

-- ---------- Mijozlar va agentlar ----------
create table customers (
  id uuid primary key default gen_random_uuid(),
  name text not null,
  phone text,
  address text,
  is_agent boolean default false,
  commission_customer_id uuid references customers(id) on delete set null,
  created_at timestamptz default now()
);

-- ---------- Fakturalar va to'lovlar ----------
create table invoices (
  id uuid primary key default gen_random_uuid(),
  number text,
  date date,
  customer_id uuid references customers(id) on delete set null,
  items jsonb,
  total numeric,
  created_by text,
  created_at timestamptz default now()
);

create table payments (
  id uuid primary key default gen_random_uuid(),
  customer_id uuid references customers(id) on delete set null,
  amount numeric,
  date date,
  note text,
  created_by text,
  created_at timestamptz default now()
);

-- ---------- Sozlamalar ----------
create table settings (
  id int primary key,
  company_name text,
  company_phone text,
  usd_rate numeric default 0,
  app_password text,
  telegram_bot_token text,
  telegram_chat_id text,
  last_telegram_sent date
);
insert into settings (id, company_name, company_phone) values (1, 'Mening korxonam', '') on conflict (id) do nothing;

-- ---------- Xom ashyo ----------
create table raw_materials (
  id uuid primary key default gen_random_uuid(),
  name text not null,
  unit text default 'kg',
  reorder_point numeric default 0,
  created_at timestamptz default now()
);

-- ---------- Ta'minotchilar ----------
create table suppliers (
  id uuid primary key default gen_random_uuid(),
  name text not null,
  phone text,
  address text,
  created_at timestamptz default now()
);

create table supplier_payments (
  id uuid primary key default gen_random_uuid(),
  supplier_id uuid references suppliers(id) on delete cascade,
  amount numeric not null,
  date date not null,
  note text,
  created_by text,
  created_at timestamptz default now()
);

create table raw_material_batches (
  id uuid primary key default gen_random_uuid(),
  raw_material_id uuid references raw_materials(id) on delete cascade,
  supplier_id uuid references suppliers(id) on delete set null,
  date date not null,
  qty numeric not null,
  remaining_qty numeric not null,
  unit_cost numeric not null,
  active boolean default true,
  is_adjustment boolean default false,
  note text,
  created_at timestamptz default now()
);

-- ---------- Tayyor mahsulot / Ishlab chiqarish ----------
create table finished_products (
  id uuid primary key default gen_random_uuid(),
  name text not null,
  unit text default 'dona',
  volume_liters numeric default 1,
  linked_product_id uuid references products(id) on delete set null,
  created_at timestamptz default now()
);

create table product_norms (
  id uuid primary key default gen_random_uuid(),
  finished_product_id uuid references finished_products(id) on delete cascade,
  raw_material_id uuid references raw_materials(id) on delete cascade,
  qty_per_unit numeric not null,
  created_at timestamptz default now()
);

create table production_batches (
  id uuid primary key default gen_random_uuid(),
  finished_product_id uuid references finished_products(id) on delete cascade,
  date date not null,
  qty numeric not null,
  unit_cost numeric not null,
  total_cost numeric not null,
  created_by text,
  created_at timestamptz default now()
);

create table production_consumptions (
  id uuid primary key default gen_random_uuid(),
  production_batch_id uuid references production_batches(id) on delete cascade,
  raw_material_batch_id uuid references raw_material_batches(id) on delete set null,
  qty numeric not null,
  unit_cost numeric not null,
  created_at timestamptz default now()
);

-- ---------- Kassa ----------
create table cash_categories (
  id uuid primary key default gen_random_uuid(),
  name text not null unique,
  created_at timestamptz default now()
);

create table cash_transactions (
  id uuid primary key default gen_random_uuid(),
  date date not null,
  type text not null,
  amount numeric not null,
  currency text default 'som',
  rate numeric,
  category text,
  note text,
  payment_id uuid references payments(id) on delete set null,
  supplier_payment_id uuid references supplier_payments(id) on delete set null,
  created_by text,
  created_at timestamptz default now()
);

-- ---------- Asosiy vositalar ----------
create table fixed_assets (
  id uuid primary key default gen_random_uuid(),
  name text not null,
  purchase_price numeric not null default 0,
  purchase_date date,
  depreciation_rate numeric not null default 0,
  include_in_cost boolean default true,
  note text,
  created_at timestamptz default now()
);

-- ============================================================
-- Ruxsatlar (RLS) — ilova bitta umumiy hisob sifatida ishlaydi,
-- shuning uchun barcha amallarga ruxsat beramiz.
-- ============================================================
alter table products enable row level security;
alter table customers enable row level security;
alter table invoices enable row level security;
alter table payments enable row level security;
alter table settings enable row level security;
alter table stock_movements enable row level security;
alter table finished_products enable row level security;
alter table raw_materials enable row level security;
alter table raw_material_batches enable row level security;
alter table product_norms enable row level security;
alter table production_batches enable row level security;
alter table production_consumptions enable row level security;
alter table cash_transactions enable row level security;
alter table cash_categories enable row level security;
alter table suppliers enable row level security;
alter table supplier_payments enable row level security;
alter table fixed_assets enable row level security;

create policy "public full access" on products for all using (true) with check (true);
create policy "public full access" on customers for all using (true) with check (true);
create policy "public full access" on invoices for all using (true) with check (true);
create policy "public full access" on payments for all using (true) with check (true);
create policy "public full access" on settings for all using (true) with check (true);
create policy "public full access" on stock_movements for all using (true) with check (true);
create policy "public full access" on finished_products for all using (true) with check (true);
create policy "public full access" on raw_materials for all using (true) with check (true);
create policy "public full access" on raw_material_batches for all using (true) with check (true);
create policy "public full access" on product_norms for all using (true) with check (true);
create policy "public full access" on production_batches for all using (true) with check (true);
create policy "public full access" on production_consumptions for all using (true) with check (true);
create policy "public full access" on cash_transactions for all using (true) with check (true);
create policy "public full access" on cash_categories for all using (true) with check (true);
create policy "public full access" on suppliers for all using (true) with check (true);
create policy "public full access" on supplier_payments for all using (true) with check (true);
create policy "public full access" on fixed_assets for all using (true) with check (true);

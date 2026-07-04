-- ══════════════════════════════════════════════════════════════════
-- CV Tailor Pro — Supabase Schema
-- Run this in: Supabase Dashboard > SQL Editor > New Query
-- ══════════════════════════════════════════════════════════════════

-- 1. Trials table — tracks one-time free trial per device fingerprint
create table if not exists trials (
  id uuid default gen_random_uuid() primary key,
  fingerprint text unique not null,
  ip text,
  file_name text,
  created_at timestamptz default now()
);

-- 2. Users table — subscription, plan, and monthly usage
create table if not exists users (
  id uuid default gen_random_uuid() primary key,
  clerk_id text unique not null,
  stripe_customer_id text unique,
  stripe_subscription_id text unique,
  plan text,                       -- 'starter' | 'growth' | 'pro' | 'unlimited'
  cv_limit integer default 0,      -- CVs allowed per month
  usage_count integer default 0,   -- CVs used this month
  usage_reset_at timestamptz,      -- when usage resets (monthly)
  subscription_status text default 'inactive',  -- 'active' | 'inactive'
  created_at timestamptz default now(),
  updated_at timestamptz default now()
);

-- 3. Tailor log — one row per CV tailored
create table if not exists tailor_log (
  id uuid default gen_random_uuid() primary key,
  clerk_id text,
  file_name text,
  plan text,
  created_at timestamptz default now()
);

-- 4. Purchase log — one row per subscription event
create table if not exists purchase_log (
  id uuid default gen_random_uuid() primary key,
  clerk_id text,
  plan text,
  stripe_session_id text unique,
  created_at timestamptz default now()
);

-- 5. Row Level Security — service key bypasses this (used by API)
alter table trials enable row level security;
alter table users enable row level security;
alter table tailor_log enable row level security;
alter table purchase_log enable row level security;

create policy "No public access to trials" on trials for all using (false);
create policy "No public access to users" on users for all using (false);
create policy "No public access to tailor_log" on tailor_log for all using (false);
create policy "No public access to purchase_log" on purchase_log for all using (false);

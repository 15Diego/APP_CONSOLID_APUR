-- Consolidador de Apuração: estrutura persistente para Streamlit + Supabase.
-- Execute como uma única migração no projeto Supabase antes de liberar o app.

create table if not exists public.app_users (
    id uuid primary key references auth.users(id) on delete cascade,
    email text not null,
    display_name text,
    role text not null default 'user' check (role in ('admin', 'user')),
    created_at timestamptz not null default now(),
    updated_at timestamptz not null default now()
);

create or replace function public.handle_new_app_user()
returns trigger
language plpgsql
security definer set search_path = public
as $$
begin
    insert into public.app_users (id, email, display_name, role)
    values (
      new.id,
      coalesce(new.email, ''),
      coalesce(new.raw_user_meta_data ->> 'name', split_part(coalesce(new.email, ''), '@', 1)),
      case when not exists (select 1 from public.app_users) then 'admin' else 'user' end
    )
    on conflict (id) do nothing;
    return new;
end;
$$;

drop trigger if exists on_consolidator_user_created on auth.users;
create trigger on_consolidator_user_created
    after insert on auth.users
    for each row execute procedure public.handle_new_app_user();

revoke execute on function public.handle_new_app_user() from public, anon, authenticated;

create table if not exists public.configuration_profiles (
    id uuid primary key default gen_random_uuid(),
    owner_id uuid not null references public.app_users(id) on delete cascade,
    name text not null,
    configuration jsonb not null default '{}'::jsonb,
    is_shared boolean not null default false,
    created_at timestamptz not null default now(),
    updated_at timestamptz not null default now(),
    unique (owner_id, name)
);

create table if not exists public.processing_jobs (
    id uuid primary key default gen_random_uuid(),
    user_id uuid not null references public.app_users(id) on delete restrict,
    name text not null,
    status text not null default 'processing' check (status in ('processing', 'completed', 'completed_with_errors', 'failed')),
    configuration jsonb not null default '{}'::jsonb,
    total_files integer not null default 0,
    valid_files integer not null default 0,
    total_rows integer not null default 0,
    filtered_rows integer not null default 0,
    output_path text,
    csv_path text,
    report_path text,
    error_message text,
    created_at timestamptz not null default now(),
    completed_at timestamptz
);

create table if not exists public.processing_files (
    id uuid primary key default gen_random_uuid(),
    job_id uuid not null references public.processing_jobs(id) on delete cascade,
    original_name text not null,
    storage_path text,
    detected_sheet_name text,
    header_row integer,
    status text not null check (status in ('valid', 'invalid', 'processed', 'failed')),
    read_rows integer not null default 0,
    filtered_rows integer not null default 0,
    source_total numeric(18,2),
    error_message text,
    created_at timestamptz not null default now()
);

create table if not exists public.audit_events (
    id bigint generated always as identity primary key,
    user_id uuid references public.app_users(id) on delete set null,
    job_id uuid references public.processing_jobs(id) on delete set null,
    action text not null,
    details jsonb not null default '{}'::jsonb,
    created_at timestamptz not null default now()
);

create index if not exists idx_processing_jobs_user_created on public.processing_jobs (user_id, created_at desc);
create index if not exists idx_processing_files_job on public.processing_files (job_id);
create index if not exists idx_audit_events_created on public.audit_events (created_at desc);

alter table public.app_users enable row level security;
alter table public.configuration_profiles enable row level security;
alter table public.processing_jobs enable row level security;
alter table public.processing_files enable row level security;
alter table public.audit_events enable row level security;

create policy "users read own profile" on public.app_users for select using (auth.uid() = id);
create policy "users update own profile" on public.app_users for update using (auth.uid() = id) with check (auth.uid() = id);
create policy "users manage own profiles" on public.configuration_profiles for all using (auth.uid() = owner_id) with check (auth.uid() = owner_id);
create policy "users read own or shared profiles" on public.configuration_profiles for select using (auth.uid() = owner_id or is_shared = true);
create policy "users manage own jobs" on public.processing_jobs for all using (auth.uid() = user_id) with check (auth.uid() = user_id);
create policy "users read own processing files" on public.processing_files for select using (exists (select 1 from public.processing_jobs j where j.id = job_id and j.user_id = auth.uid()));
create policy "users read own audit events" on public.audit_events for select using (auth.uid() = user_id);

insert into storage.buckets (id, name, public, file_size_limit, allowed_mime_types)
values ('fiscal-files', 'fiscal-files', false, 104857600, array[
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  'application/vnd.ms-excel',
  'application/vnd.ms-excel.sheet.binary.macroEnabled.12'
])
on conflict (id) do update set public = false, file_size_limit = 104857600;

create policy "authenticated users access own fiscal objects" on storage.objects
for all to authenticated
using (bucket_id = 'fiscal-files' and (storage.foldername(name))[1] = auth.uid()::text)
with check (bucket_id = 'fiscal-files' and (storage.foldername(name))[1] = auth.uid()::text);

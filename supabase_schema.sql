-- ===========================================================================
-- Premion Proposal Builder -- Supabase schema
--
-- Paste this whole file into the Supabase SQL editor and run it. Every
-- statement is idempotent, so it's safe to re-run after a later stage adds
-- a new section.
--
-- Row Level Security is enabled with no policies on every table: the app is
-- server-side and talks to Supabase with the service key (which bypasses
-- RLS), so leaving RLS on with zero policies means an anon/public key can
-- read nothing at all. Do not add permissive policies without a reason.
-- ===========================================================================


-- ---------------------------------------------------------------------------
-- Stage 1: master deck versions
--
-- The .pptx itself lives in the private `decks` storage bucket;
-- storage_path is its object key inside that bucket.
-- ---------------------------------------------------------------------------
create table if not exists public.deck_versions (
    id           bigint generated always as identity primary key,
    storage_path text        not null,
    filename     text        not null,
    uploaded_at  timestamptz not null default now(),
    notes        text,
    active       boolean     not null default false
);

-- "exactly one active": a partial unique index across only the true rows, so
-- any number of inactive versions can sit alongside a single active one.
-- Activation must therefore clear the old active row before setting the new
-- one (see db.activate_deck_version).
create unique index if not exists deck_versions_single_active
    on public.deck_versions (active)
    where active;

alter table public.deck_versions enable row level security;

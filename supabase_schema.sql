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


-- ---------------------------------------------------------------------------
-- Stage 2: products / rate card
--
-- `key` is what the app addresses a product by everywhere (media plan line
-- products, the AI draft schema, Section C's toggles), so it's the real
-- identity of a row -- `id` only exists to be a surrogate key.
--
-- Live Sports packages are products too, keyed "sport:<sport_key>" with
-- display_group 'sports'. Their rates come from the per-sport ratecard rather
-- than one product default, which is exactly why they need rows of their own.
--
-- line_type is 'rate' (billed as impressions x CPM) or 'flat' (a one-time fee),
-- matching the media-plan grid's own Type column. display_group is which part
-- of the form the product belongs to (premion / am / broadcast / sports).
-- ---------------------------------------------------------------------------
create table if not exists public.products (
    id             bigint generated always as identity primary key,
    key            text    not null unique,
    name           text    not null,
    line_type      text    not null default 'rate' check (line_type in ('rate', 'flat')),
    default_cpm    numeric(10, 2),
    targeting_copy text,
    display_group  text,
    active         boolean not null default true,
    sort_order     integer not null default 0
);

alter table public.products enable row level security;


-- ---------------------------------------------------------------------------
-- Stage 3: audience catalog
--
-- `audiences` is the brochure catalog (segment name, grouping, whether it's
-- selectable in a Salesforce RFP) merged with how often each segment has been
-- used. `audience_usage` is the raw year-to-date delivery log, one row per
-- segment *string* as booked -- which is often a combination of segments
-- rather than a single catalog name, hence the separate table and the
-- deliberately different column name.
-- ---------------------------------------------------------------------------
create table if not exists public.audiences (
    id             bigint generated always as identity primary key,
    segment        text    not null unique,
    category       text,
    subcategory    text,
    rfp_selectable boolean not null default false,
    times_used     integer not null default 0,
    active         boolean not null default true
);

create table if not exists public.audience_usage (
    id                    bigint generated always as identity primary key,
    segment_string        text not null,
    delivered_impressions bigint,
    is_custom             boolean
);

create index if not exists audience_usage_segment_idx
    on public.audience_usage (segment_string);

alter table public.audiences enable row level security;
alter table public.audience_usage enable row level security;


-- ---------------------------------------------------------------------------
-- Stage 4: proposal history
--
-- form_json is the whole form state at generate time, every plan option
-- included -- capture only, nothing reads it back yet. deck_version_id is
-- null when the deck came from the local fallback, so a logged proposal never
-- claims a version it didn't actually build from.
-- ---------------------------------------------------------------------------
create table if not exists public.proposals (
    id              uuid primary key default gen_random_uuid(),
    client_name     text,
    vertical        text,
    market          text,
    form_json       jsonb       not null,
    generated_at    timestamptz not null default now(),
    deck_version_id bigint      references public.deck_versions (id),
    output_filename text
);

create index if not exists proposals_generated_at_idx
    on public.proposals (generated_at desc);

alter table public.proposals enable row level security;

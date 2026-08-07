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


-- ---------------------------------------------------------------------------
-- Case study vault
--
-- The .pptx lives in the private `case_studies` bucket; storage_path is its
-- object key. verticals/products are the tags a proposal is matched against
-- at generate time -- arrays because one case study routinely speaks to
-- several verticals (a law firm case study is neither purely one thing nor
-- the other) and almost always demonstrates more than one product.
--
-- added_by is free text on purpose: anyone using the app can contribute one,
-- and there are no user accounts to attribute it to.
-- ---------------------------------------------------------------------------
create table if not exists public.case_studies (
    id           uuid primary key default gen_random_uuid(),
    filename     text        not null,
    storage_path text        not null,
    title        text,
    verticals    text[]      not null default '{}',
    products     text[]      not null default '{}',
    summary      text,
    date_added   timestamptz not null default now(),
    added_by     text,
    active       boolean     not null default true
);

-- The generate-time picker filters on vertical and orders by recency.
create index if not exists case_studies_verticals_idx
    on public.case_studies using gin (verticals);
create index if not exists case_studies_date_added_idx
    on public.case_studies (date_added desc);

alter table public.case_studies enable row level security;


-- ---------------------------------------------------------------------------
-- Stage 5: proposal history (revision + reuse)
--
-- The proposals table stops being capture-only here: the History page reads
-- form_json back to rehydrate the form, rebuild a deck as it was originally
-- presented, or start a new proposal from an old one.
--
-- History is append-only in spirit. Regenerating a loaded proposal inserts a
-- NEW row carrying parent_proposal_id back to the one it came from, rather
-- than overwriting anything -- what a client was actually sent must stay
-- exactly as it was logged. Delete is the only removal, and it's there for
-- junk and test rows.
--
-- parent_proposal_id is ON DELETE SET NULL, not CASCADE, deliberately:
-- deleting a test row that a real proposal happens to descend from must
-- orphan the link, never take the descendant with it.
--
-- The attached-file columns are the exception to storing-the-recipe. A deck
-- is sometimes downloaded, hand-edited in PowerPoint and sent in that state,
-- at which point form_json no longer describes what the client saw; the
-- final file can be attached so the row still tells the truth. It lives in
-- the private `proposal_files` bucket (created by setup_supabase.py, not
-- here -- buckets aren't DDL).
-- ---------------------------------------------------------------------------
alter table public.proposals
    add column if not exists parent_proposal_id uuid
    references public.proposals (id) on delete set null;

alter table public.proposals add column if not exists revision_label     text;
alter table public.proposals add column if not exists file_storage_path  text;
alter table public.proposals add column if not exists file_attached_at   timestamptz;
alter table public.proposals add column if not exists file_note          text;

-- Walking a client's revision thread, and the History page's client search.
create index if not exists proposals_parent_idx
    on public.proposals (parent_proposal_id);
create index if not exists proposals_client_name_idx
    on public.proposals (lower(client_name));

-- The client logo, stored alongside the deck it was used on. An uploaded
-- file is not part of form_json, so without this a rebuild silently fell
-- back to the placeholder and "Rebuild as presented" wasn't faithful. Null
-- means the proposal genuinely had no logo -- which the UI distinguishes
-- from "had one, but it can't be fetched right now". Lives in the same
-- private `proposal_files` bucket as attached finals.
alter table public.proposals add column if not exists logo_storage_path text;


-- ---------------------------------------------------------------------------
-- Stage 6: lightweight identity
--
-- Not authentication -- the shared password is still the gate. This is only
-- "who is using the app right now", so a proposal, a case study upload or an
-- attached file can say who it came from.
--
-- The list builds itself: anyone can add a name from the picker, and it's
-- there for everyone afterwards. No admin UI, because this table is trivially
-- editable in the Supabase dashboard if a typo ever needs cleaning up.
--
-- name_key is the deduplication key -- lower(trim(name)) -- and carries the
-- unique index rather than `name` itself, so "matt", "Matt" and " Matt " are
-- one person while the display name keeps whatever capitalisation was typed
-- first. Doing it as a stored column rather than a unique index on an
-- expression keeps the conflict target nameable from PostgREST's upsert.
-- ---------------------------------------------------------------------------
create table if not exists public.team_members (
    id       uuid        primary key default gen_random_uuid(),
    name     text        not null,
    name_key text        not null,
    active   boolean     not null default true,
    added_at timestamptz not null default now()
);

create unique index if not exists team_members_name_key_idx
    on public.team_members (name_key);

alter table public.team_members enable row level security;

insert into public.team_members (name, name_key)
values ('Matt', 'matt'), ('Andrea', 'andrea')
on conflict (name_key) do nothing;

-- Who generated a proposal. Nullable: every row logged before identity
-- existed stays anonymous rather than being attributed to a guess.
alter table public.proposals add column if not exists created_by text;

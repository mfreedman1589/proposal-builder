-- ===========================================================================
-- Proposal Builder -- Supabase schema
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

-- ---------------------------------------------------------------------------
-- Stage 6: pre-rendered case study slides
--
-- A case study can be grafted into a proposal two ways. Copying its slide XML
-- across decks is the faithful-looking one and has produced five distinct
-- corruption bugs here; it also can't reproduce PowerPoint's live autofit, so
-- flattened text can render a line taller than its box and clip. A slide
-- rendered to an image has no relationships, theme, layout, placeholders,
-- autofit or embedded parts, so none of that can happen -- and it turns out
-- SMALLER, because copying drags in the source deck's full-resolution photos.
--
-- The images are generated locally (see render_case_study_images.py): they
-- need PowerPoint and the Proxima Nova brand font, and Streamlit Cloud has
-- neither -- rendering there would silently substitute fonts.
--
-- slide_images is an ordered array of object keys in the SAME `case_studies`
-- bucket the .pptx lives in. Empty means "no images yet", which is the signal
-- the app falls back to the copy path on, so it needs no separate flag.
-- ---------------------------------------------------------------------------
alter table public.case_studies
    add column if not exists slide_images text[] not null default '{}';
alter table public.case_studies
    add column if not exists image_width int;
alter table public.case_studies
    add column if not exists images_generated_at timestamptz;

-- "Which case studies still need images?" is the only query the admin script
-- and the vault browser's coverage column ask.
create index if not exists case_studies_needs_images_idx
    on public.case_studies (active)
    where cardinality(slide_images) = 0;


-- ---------------------------------------------------------------------------
-- Stage 7: market viewer profiles
--
-- One row per slide of PREMION's "OTT Viewer Profiles By Market" deck: 205
-- Nielsen DMAs plus one national roll-up. Each slide is a single full-bleed
-- image with no text and no speaker notes, so the market name was read off a
-- rendering rather than parsed -- see market_profiles_index.json, which
-- tests/validate_markets.py regenerates and checks against the canonical DMA
-- list. Treat that file as the source of the seed data.
--
-- `key` is the addressable identity, exactly as it is on products: a slug
-- ('new_york', 'total_us'). It exists because `dma` cannot be the key -- the
-- national roll-up has no DMA, and a unique index over a nullable column
-- would admit a second one (nulls are distinct in Postgres).
--
-- The national roll-up IS selectable, deliberately: dma null, kind
-- 'national'. National advertisers are a real case and the deck ships the
-- slide for exactly that reason.
--
-- Five DMAs have no slide in the source deck -- Honolulu, Palm Springs,
-- Anchorage, Fairbanks, Juneau. They still get rows, and are still
-- selectable as target DMAs: a rep selling into Honolulu has a real
-- proposal to build, and refusing the market because Premion never authored
-- a slide for it would be the app inventing a restriction the business
-- doesn't have.
--
-- What they don't get is an image. `image_path is null` is the one and only
-- signal for "no profile slide exists", and it has to stay visible rather
-- than silent -- the picker says so at selection time, so a rep knows before
-- they build why no profile appears, instead of reading its absence as a bug
-- or as their own mistake. The profile toggle is NOT dropped for these
-- markets either; it stays where it is, disabled and explained. Silently
-- removing a control is how a seller ends up thinking the feature is broken.
--
-- So: `slide_number is null` and `image_path is null` mean the profile was
-- never authored, which is different from "not imported yet". Nothing here
-- distinguishes those two, because the seed covers the whole deck in one
-- pass -- if that ever stops being true, it needs a real column, not an
-- inference from a null.
--
-- image_path is an object key in the private `market_profiles` bucket, which
-- is its own bucket rather than a corner of `case_studies`: different
-- lifecycle (this set is replaced wholesale when Premion reissues the deck,
-- where the vault accretes one case study at a time) and cleaner quota
-- accounting against the 1GB free tier. Buckets aren't DDL -- setup_supabase.py
-- creates it, like every other bucket here.
--
-- The images are JPEG at 2560px, the same width and quality the case study
-- image path uses. The renders are 128.2 MiB as PNG and 34.8 MiB as JPEG,
-- and these slides are photographic and fully opaque, which is the case the
-- deck optimizer already converts for.
--
-- `stats` is deliberately present and deliberately empty, and NOTHING READS
-- IT. Every slide states OTT penetration, viewer population, device
-- ownership, view habits and a full demographic block, so the data will
-- obviously be wanted -- but extracting it means reading 206 images, which
-- is a separate task that has not been done and currently has no consumer.
--
-- Do not assume this column is populated. Null means "never extracted", not
-- "this market has no data", and every row is null today. Anything that
-- starts reading it must handle null as the normal case, not the exception,
-- until a backfill exists and this comment says otherwise.
--
-- It ships now rather than being added later only because schema has to land
-- before the code that depends on it, so an extra round trip costs a deploy
-- window for a column that is free to carry.
-- ---------------------------------------------------------------------------
create table if not exists public.market_profiles (
    id                  bigint      generated always as identity primary key,
    key                 text        not null unique,
    label               text        not null,
    dma                 text,
    kind                text        not null default 'dma'
                                    check (kind in ('dma', 'national')),
    slide_number        int,
    image_path          text,
    image_width         int,
    images_generated_at timestamptz,
    stats               jsonb,
    active              boolean     not null default true,
    created_at          timestamptz not null default now()
);

-- A DMA is claimed by at most one profile. Partial, so the national roll-up
-- (dma null) sits outside it rather than needing an exemption.
create unique index if not exists market_profiles_dma_idx
    on public.market_profiles (dma)
    where dma is not null;

-- The picker lists active profiles in slide order, which is alphabetical.
create index if not exists market_profiles_active_idx
    on public.market_profiles (active, slide_number);

alter table public.market_profiles enable row level security;


-- The DMA a proposal is TARGETED at -- a new field, not a rename of anything.
--
-- `proposals.market` above keeps its existing meaning and its existing job:
-- the ORIGINATING market (DC / Harrisburg), which drives station branding,
-- the Total TV slide variants and the broadcast DMA, and records which market
-- produced the proposal. Those two are independent -- a DC-originated
-- proposal can target any DMA in the country -- and collapsing them would
-- put a competitor station's branding in front of a client, which is the
-- worst thing this deck can do.
--
-- Stored as the market_profiles.key slug, and deliberately NOT a foreign key:
-- history is append-only and a logged proposal must not depend on a lookup
-- row still existing, the same reason fetch_case_study ignores `active` and
-- parent_proposal_id is ON DELETE SET NULL. Null means no target DMA was
-- chosen, which is every proposal logged before this existed.
alter table public.proposals add column if not exists target_dma text;


-- ---------------------------------------------------------------------------
-- Stage 8: several target DMAs, not one
--
-- Reps routinely target several markets at once, and the deck was designed
-- for it -- the national roll-up slide is replaced by ONE PROFILE SLIDE PER
-- SELECTED MARKET, in picker order. `target_dma` above could only ever hold
-- the first, which would have silently shipped a one-market deck for a
-- three-market brief.
--
-- ADDITIVE, not a type change. `alter column ... type text[] using
-- array[target_dma]` would rewrite every existing row in place, and history
-- here is append-only: a proposals row is the record of what a client was
-- actually sent, so it is never rewritten -- not even into a shape that
-- happens to mean the same thing. A rewrite is also the one form of this
-- that can't be undone if it goes wrong mid-flight.
--
-- `target_dma` is therefore KEPT and FROZEN. Nothing writes it from here on;
-- it is read only as the backfill source below. It is not dropped, because
-- dropping is the single irreversible option available and it costs nothing
-- to leave in place. Do not start dual-writing it either -- two copies of
-- one fact drifting apart is a bug this project has already paid for more
-- than once (see the canonical DMA list, and the _product_seed_key).
--
-- As it happens the backfill is a no-op today: the scalar column shipped
-- only days before this and no proposal was ever generated with it set. It
-- is written correctly anyway, because "no rows yet" is a fact about right
-- now and this file is re-run later by definition.
alter table public.proposals
    add column if not exists target_dmas text[] not null default '{}';

-- Backfill: one-element array from whatever the scalar held. Guarded on the
-- array still being empty so re-running the file can never overwrite a real
-- multi-market selection with the single legacy value.
update public.proposals
   set target_dmas = array[target_dma]
 where target_dma is not null
   and cardinality(target_dmas) = 0;

-- "Which markets are we selling into" is the question this column exists to
-- answer, and it's a containment query over an array.
create index if not exists proposals_target_dmas_idx
    on public.proposals using gin (target_dmas);


-- ---------------------------------------------------------------------------
-- Stage 9: audience usage workbook versioning
--
-- The YTD usage workbook (Segment Name / Delivered Impressions, one row per
-- booked stack) that Matt refreshes periodically is now an uploadable,
-- versioned source, the same shape as deck_versions: the raw .xlsx lives in
-- the private `audience_usage_workbooks` storage bucket, storage_path is its
-- object key, and activating a version (db.activate_audience_usage_version)
-- re-derives `audience_usage` and merges into `audiences` from whichever
-- workbook is active -- never a silent re-parse on every read. Every
-- previous workbook stays in storage, same "reversible" guarantee the deck
-- has.
-- ---------------------------------------------------------------------------
create table if not exists public.audience_usage_versions (
    id           bigint generated always as identity primary key,
    storage_path text        not null,
    filename     text        not null,
    uploaded_at  timestamptz not null default now(),
    notes        text,
    active       boolean     not null default false
);

create unique index if not exists audience_usage_versions_single_active
    on public.audience_usage_versions (active)
    where active;

alter table public.audience_usage_versions enable row level security;

-- A flexible bag for ranking signals beyond times_used (stack count) --
-- impressions today ("rank by impressions, not by count"), a response-rate
-- metric later -- so a second metric is a new KEY, not a migration. Read as
-- metrics->>'impressions' rather than a dedicated column.
alter table public.audiences
    add column if not exists metrics jsonb not null default '{}'::jsonb;

-- ---------------------------------------------------------------------------
-- Stage 10: app settings
--
-- One generic key/value table for small, admin-editable config records that
-- aren't a catalog (products, audiences, case studies) and aren't per-
-- proposal -- the co-viewing coefficient (multiplier, source, study date,
-- footnote text) is the first. A second setting is a new ROW, not a
-- migration, same jsonb-bag philosophy as audiences.metrics above. Read
-- through db.fetch_setting(key)/db.upsert_setting(key, value); the app falls
-- back to a hardcoded constant (FALLBACK_COVIEWING in app.py) when this
-- table is unreachable or the row doesn't exist yet, same as every other
-- loader in this file's own convention.
-- ---------------------------------------------------------------------------
create table if not exists public.app_settings (
    key        text primary key,
    value      jsonb       not null default '{}'::jsonb,
    updated_at timestamptz not null default now()
);

alter table public.app_settings enable row level security;

-- ---------------------------------------------------------------------------
-- Stage 11: in-app feedback
--
-- A persistent "Report an issue" popover on every page (BACKLOG.md) --
-- category, free-text notes, and a best-effort state snapshot (never a
-- screenshot; that's a deferred fast-follow) so "the avails row
-- disappeared" is reproducible: which page, who reported it, which deck
-- build, and the Build page's own raw source-of-truth state
-- (targeting_groups, plan_options, the scalar Section A/flight fields) --
-- never the derived, markup-applied numbers Generate computes, which are
-- recomputable from these anyway. See app.capture_feedback_state.
--
-- state_json is a few KB, same as proposals.form_json -- no storage bucket
-- needed for this table. status is a plain string rather than a boolean so
-- a third state ("wontfix") is a value, not a migration.
-- ---------------------------------------------------------------------------
create table if not exists public.feedback (
    id           uuid primary key default gen_random_uuid(),
    category     text        not null,
    notes        text        not null,
    page         text,
    created_by   text,
    state_json   jsonb       not null default '{}'::jsonb,
    status       text        not null default 'open',
    created_at   timestamptz not null default now(),
    resolved_at  timestamptz
);

create index if not exists feedback_created_at_idx
    on public.feedback (created_at desc);
create index if not exists feedback_status_idx
    on public.feedback (status);

alter table public.feedback enable row level security;


-- ---------------------------------------------------------------------------
-- Stage 12: slide vault
--
-- The case study vault again, one level down: a colleague uploads a .pptx and
-- picks the individual SLIDES worth keeping, not the whole deck. One row is
-- one picked slide, not one upload -- several rows share one storage_path
-- (slide_index tells them apart) when a colleague picks more than one slide
-- out of the same source deck. The .pptx is stored whole rather than split
-- apart: extracting one slide into its own file is exactly the cross-deck
-- copy problem that produced five distinct corruption bugs on the case-study
-- vault (see CLAUDE.md), so a vault slide is grafted the same way a case
-- study is -- by index, out of the source deck it always lived in.
--
-- placement is the contributor's suggested default anchor (front / the
-- case-study-adjacent slot / appendix); a rep can override it per proposal at
-- generate time, so it is a default, not a constraint enforced here.
--
-- purpose is a free-text label for what KIND of slide this is (capabilities,
-- research, creative example, pricing, ...) -- looser than verticals/products
-- on purpose, since the taxonomy is a guess and this is one text column, so
-- getting it wrong later is a data edit, not a migration.
--
-- slide_image (singular) is the image-vs-copy signal, same convention as
-- case_studies.slide_images: null means "not rendered yet," which is the
-- pending-render queue, so there is nothing separate to keep in sync.
-- ---------------------------------------------------------------------------
create table if not exists public.slide_vault (
    id                  uuid primary key default gen_random_uuid(),
    filename            text        not null,
    storage_path        text        not null,
    slide_index         int         not null,
    title               text,
    verticals           text[]      not null default '{}',
    products            text[]      not null default '{}',
    purpose             text,
    placement           text        not null default 'before_plan',
    summary             text,
    date_added          timestamptz not null default now(),
    added_by            text,
    active              boolean     not null default true,
    slide_image         text,
    image_width         int,
    image_generated_at  timestamptz
);

create index if not exists slide_vault_verticals_idx
    on public.slide_vault using gin (verticals);
create index if not exists slide_vault_date_added_idx
    on public.slide_vault (date_added desc);
-- "Which vault slides still need an image?" -- the admin script's pending
-- queue and the vault browser's coverage column.
create index if not exists slide_vault_needs_image_idx
    on public.slide_vault (active)
    where slide_image is null;

alter table public.slide_vault enable row level security;

-- ---------------------------------------------------------------------------
-- Stage 13: advertisers (the canonical client spine)
--
-- client_name has always been free text per proposal, which is fine until a
-- second table needs to reference "the same client" -- an attribution report
-- links to a proposal by advertiser, and neither Premion's export nor a
-- rep's own typing agrees on spelling with what a proposal used
-- ("Cardinal Plumbing" vs. "Cardinal Plumbing Pixel" vs. a rep's own typo).
-- Built now, before reports start accumulating against free-text names
-- (ATTRIBUTION_REPORT_PLAN.md Phase 2) -- deferring costs more, not less: a
-- later backfill would have to reconcile two independently-spelled
-- free-text sources (proposals.client_name AND attribution_reports'
-- own) instead of one, and the fuzzy-match code this table enables has to
-- be written now regardless.
--
-- Exact shape of team_members: name_key (lower(trim(name))) carries the
-- dedup, not canonical_name itself, so "Cardinal Plumbing" and "cardinal
-- plumbing" resolve to one row while the display name keeps whatever
-- capitalisation was typed first.
-- ---------------------------------------------------------------------------
create table if not exists public.advertisers (
    id              uuid        primary key default gen_random_uuid(),
    canonical_name  text        not null,
    name_key        text        not null,
    created_at      timestamptz not null default now()
);

create unique index if not exists advertisers_name_key_idx
    on public.advertisers (name_key);

alter table public.advertisers enable row level security;

-- Nullable on purpose: every proposal logged before this table existed
-- stays unlinked (still shows its own client_name as typed) rather than
-- being attributed to a guessed advertiser row. A one-time backfill script
-- groups existing proposals by normalized client_name and back-fills this
-- column -- a data migration, not a schema requirement, so it is
-- deliberately not part of this DDL.
alter table public.proposals
    add column if not exists advertiser_id uuid references public.advertisers (id) on delete set null;
create index if not exists proposals_advertiser_id_idx
    on public.proposals (advertiser_id);

-- ---------------------------------------------------------------------------
-- Stage 14: attribution reports
--
-- One row per generated (or, in Phase 2, merely parsed-and-confirmed)
-- attribution report. report_json is a jsonb bag the same way feedback.
-- state_json is -- the parsed export facts today, the full report payload
-- once deck assembly lands in Phase 3+. proposal_id is nullable: a report
-- built in no-proposal mode never gets one, same distinction proposals.
-- parent_proposal_id already draws for a proposal with no ancestor.
-- Both FKs are ON DELETE SET NULL, never CASCADE -- deleting a proposal or
-- advertiser must never take a real report down with it.
-- ---------------------------------------------------------------------------
create table if not exists public.attribution_reports (
    id              uuid        primary key default gen_random_uuid(),
    advertiser_id   uuid        references public.advertisers (id) on delete set null,
    proposal_id     uuid        references public.proposals (id) on delete set null,
    report_json     jsonb       not null default '{}'::jsonb,
    status          text        not null default 'parsed',
    created_by      text,
    created_at      timestamptz not null default now()
);

create index if not exists attribution_reports_advertiser_id_idx
    on public.attribution_reports (advertiser_id);
create index if not exists attribution_reports_created_at_idx
    on public.attribution_reports (created_at desc);

alter table public.attribution_reports enable row level security;

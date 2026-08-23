# Deploying to Streamlit Community Cloud

**Live app: https://proposal-builder.streamlit.app/**

Recorded here because nothing in the repo said where the deployed instance
was, which meant a live outage couldn't be checked from the code — the
diagnosis had to be made entirely from git, and "is it back up?" could only
be answered by asking someone. A deployment doc that doesn't name the
deployment is missing the one fact you need when it's down.

Everything in the repo is ready. This is what you do in the dashboard, and
what you need to decide.

---

## 1. Secrets — paste this into the dashboard

Streamlit Cloud has a **Secrets** box (Advanced settings during deploy, or
**⋮ → Settings → Secrets** afterwards) that expects TOML. Paste exactly this,
filling in the four values:

```toml
APP_PASSWORD = "the shared password sellers type to get in"
ANTHROPIC_API_KEY = "sk-ant-..."
SUPABASE_URL = "https://<your-project-ref>.supabase.co"
SUPABASE_SERVICE_KEY = "eyJ...  (the service_role key, not the anon key)"
```

All four are the same keys already in your local `.streamlit/secrets.toml` —
copy the values straight across. No other secrets are read anywhere in the
codebase.

⚠️ **`SUPABASE_SERVICE_KEY` is the `service_role` key and bypasses Row Level
Security entirely.** That's deliberate — the app is server-side and every
table has RLS on with no policies, so this key is the only thing that can read
them. It must never appear in the repo, in client-side code, or in a
screenshot. Streamlit stores secrets encrypted and doesn't expose them to the
browser.

---

## 2. Deploy, in order

1. Go to **share.streamlit.io** and sign in with GitHub.
2. **Create app** → **Deploy a public app from GitHub**.
3. Fill in:
   - **Repository:** `mfreedman1589/proposal-builder`
   - **Branch:** `main`
   - **Main file path:** `app.py`
   - **App URL:** pick your subdomain (this is public and hard to change
     later — something like `premion-proposal-builder`).
4. Open **Advanced settings** *before* deploying:
   - **Python version:** **3.12** (matches local; `runtime.txt` records the
     same version but the dropdown is what Streamlit Cloud actually uses).
   - **Secrets:** paste the TOML block above.
5. Click **Deploy**. First build takes a few minutes while it installs
   `requirements.txt`.
6. When it comes up, log in with `APP_PASSWORD` and check the top of the page
   shows **no yellow warning banners** — those mean it fell back to built-in
   data instead of reaching Supabase (see troubleshooting below).

If the repo is private, grant Streamlit access to it when GitHub prompts
during sign-in.

---

## 3. Decisions you need to make

**Who can open the app.** Streamlit Community Cloud apps are **public by
default** — anyone with the URL reaches the password gate. The gate is a
single shared password, which is thin protection for something wired to a
service-role key. Under **Settings → Sharing** you can restrict viewing to
specific email addresses. Recommended unless you want sellers to be able to
share the link freely.

**Which Supabase project.** It'll use the same project as local, so a deployed
seller and you are writing to the same `proposals` table and the same vault.
That's probably what you want. If you'd rather keep production separate, make
a second Supabase project, run `supabase_schema.sql` and
`python setup_supabase.py all` against it, and use its URL/key in the cloud
secrets.

**Whether to keep the app awake.** Free-tier apps sleep after a period of
inactivity and take ~30s to wake. Fine for occasional use; annoying if a
seller is demoing live.

---

## 4. What happens to the gitignored files

`.pptx`, `.pdf` and `case_studies_source/` are gitignored, so **none of them
exist on the cloud server.** That's fine by design — the master deck, the
audience catalog and the case studies all come from Supabase. It does mean the
local fallbacks can't engage there, so the app was explicitly tested with
every one of those files absent:

| File | On the cloud | If Supabase is also unreachable |
|---|---|---|
| `TEGNA_MASTER_DECK_v1_1.pptx` | not present; deck comes from the `decks` bucket | Generate stops with a plain-English message — no traceback |
| `PREMION_Audience Targeting (1).pdf` | not present; catalog comes from the `audiences` table | Catalog is empty, avails can still be typed by hand, everything else works |
| `case_studies_source/` | not present; case studies come from the `case_studies` bucket | Case study section warns and offers nothing |
| `audience_segments_derived.csv`, `audience_usage_ytd.csv`, `placeholder_logo.png` | **committed**, so present | n/a |

Verified end-to-end in both states (Supabase up and Supabase unreachable) with
all three files removed: the app starts, all three pages render, and nothing
raises.

---

## 5. Resource use

Measured over three consecutive generates on the real 119-slide deck:

| | |
|---|---|
| Baseline after imports | ~73 MiB |
| Peak during a generate | **~280 MiB** |
| Back down to | ~90 MiB |
| Growth over 3 generates | +18 MiB, flattening — no leak |
| Deck on disk | 43.9 MiB, downloaded **once per version**, not per generate |

Comfortably inside Community Cloud's limit. The one thing to know: peak is
*per concurrent generate*, so several sellers hitting Generate at the same
moment multiply it. A handful of users is fine.

Temp files: uploaded decks land in a scratch directory that sweeps anything
older than 6 hours; the optimizer's working files are deleted on both the
success and failure paths; the deck and case-study download caches keep one
file per version, which is what makes repeat generates cheap.

---

## 6. If something goes wrong

**`AttributeError` on a cross-module name at startup** (e.g. `module
'assembly' has no attribute 'VERTICAL_ATTRIBUTION'` at `app.py:750`) —
**reboot the app first, before touching git.** A redeploy can reuse a warm
Python process, so a new `app.py` ends up running against a module still
cached in `sys.modules` from before the change. The giveaway is that the
failing combination exists in *no commit*: check with
`git log --oneline -S "<name>"` against each file, and if both sides came
from the same commit, the repo is fine and the process is stale. Reboot app
(dashboard → app menu) kills the process and rebuilds from a clean checkout.
Do not revert — that deletes working code and leaves the fault in place.

**Yellow warning banners at the top of the form** — the app is running on
built-in fallbacks because it couldn't reach Supabase. Check the secrets are
present and correct, and that the Supabase project isn't paused (free-tier
projects pause after a week of inactivity).

**"A proposal can't be generated until Supabase is reachable again"** — same
cause. Nothing is lost; the form keeps everything you typed.

**Build fails installing requirements** — check the Python version dropdown is
3.12. The pins in `requirements.txt` are exact and were tested against it.

**Build succeeds, app boots, then every request 500s with a `TypeError` deep in
starlette or uvicorn** — a transitive dependency resolved to a version the
pinned `streamlit` can't drive. This happened once for real: streamlit declares
only `starlette<2,>=0.40.0`, and a starlette release added a required
keyword-only argument to `GZipResponder`, which streamlit's own middleware
subclasses and constructs positionally. `starlette` is pinned in
`requirements.txt` for exactly this reason. Local dev won't reproduce it — the
local install is whatever old version was resolved months ago, which is the
working one. Fix by pinning the offending package to the version local dev has,
after checking it sits inside streamlit's declared range.

**App works but the deck is old** — the deck cache is keyed on the active
version id, so activating a new version through the *Update master deck* page
picks it up immediately. Changing the deck in Supabase by hand won't.

To ship a code change: push to `main`. Streamlit Cloud redeploys
automatically. Use **⋮ → Reboot** to clear caches without redeploying.

---

## Schema changes go before the push, and the push is not the finish line

Moved verbatim from `CLAUDE.md`.

- **DDL is not something the app can do — and the deployed app rebuilds from GitHub on every push, so schema goes first.** The service key can't issue DDL over PostgREST, so schema changes go in `supabase_schema.sql` and are pasted into the Supabase SQL editor by hand; every statement is idempotent so the file can be re-run after a later stage appends to it. **Ordering matters and is not optional:** run the SQL *before* pushing code that depends on it, or Streamlit Cloud redeploys against a database missing the columns and the live app breaks in the window between. **And the rule ends at a confirmed deploy, not at a completed push** — a redeploy can reuse a warm process (see the Streamlit Cloud staleness note below), so a push that "succeeded" is not proof the new code is running. The sequence is: run the SQL → push → *confirm the live app is serving the new code* (https://proposal-builder.streamlit.app/), rebooting it if it isn't. Treating the push as the finish line is how a schema-dependent change gets reported as working when nothing has actually exercised it. Stage 5 (proposal history) is the current example — `log_proposal` writes `parent_proposal_id`, which doesn't exist until that DDL is run. `setup_supabase.py` does everything the client *can* do (buckets, uploads, seeding) with `python setup_supabase.py all|deck|products|audiences|case_studies|proposal_files|settings` — each step checks for what it would create, so re-running is safe. Note `setup_supabase.py products` (and, the same way, `settings`) is a *bootstrap, not a sync*: it upserts on `key`, so running it after editing rates — or the co-viewing coefficient — in Supabase would overwrite those edits with the older hardcoded values. Every table has RLS enabled with no policies — the app is server-side and uses the service key (which bypasses RLS), so an anon key can read nothing. Don't add permissive policies without a reason.

## A redeploy can reuse a warm process

- **The same staleness happens on Streamlit Cloud, and there "pushed" is not the same as "running that code".** A redeploy can reuse a warm Python process, so a new `app.py` runs against a module still in `sys.modules` from before the change. That produced a live outage: `AttributeError: module 'assembly' has no attribute 'VERTICAL_ATTRIBUTION'` at `app.py:750`, on a repo where **both halves had always shipped in the same commit** (`65db40a`, app.py +50 and assembly.py +68). It took down every page rather than one feature, because that reference sits at module scope — `DRAFT_KEY_SECTIONS.update({... for v in assembly.VERTICAL_ATTRIBUTION})` runs at import.

  **Triage an `AttributeError` on a cross-module name by rebooting first, before touching git.** The tell is that *the failing combination exists in no commit*: an `app.py` that has the reference against an `assembly` that lacks the definition. `git log --oneline -S "<name>"` on each file settles it in seconds — same commit on both sides means the repo is fine and the process is stale, and the fix is the dashboard's **Reboot app**, which kills the process and rebuilds from a clean checkout. Reverting is the tempting move and the wrong one: it deletes working code, triggers another rebuild, and leaves the actual fault untouched. Confirm by importing the pushed tree in a clean directory (`git archive origin/main | tar -x -C <tmp>`, then `python -c "import app"`) rather than reasoning about it.

  `tests/test_cross_module_refs.py` covers the *other* failure — the genuinely half-complete push — by resolving every `module.NAME` between first-party modules. It would not have caught this incident, and saying so is the point: a guard whose scope you misremember is worse than none.

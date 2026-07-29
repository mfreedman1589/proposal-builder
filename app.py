"""
app.py -- Phase 1 Streamlit form for the Premion Proposal Builder (build spec
section 5), wired directly to assembly.py's selection -> deletion -> fill
pipeline. No Supabase, no Claude API yet: the product list/CPMs are a
hardcoded dict below, and Campaign Specs bullets are typed in by hand.
"""

import io
from datetime import date

import pandas as pd
import streamlit as st

import assembly

st.set_page_config(page_title="Premion Proposal Builder", layout="wide")

MASTER_DECK_PATH = assembly.MASTER_DECK_PATH

VERTICALS = {
    "None": "none",
    "Education": "education",
    "Healthcare": "healthcare",
    "Retail": "retail",
    "Travel & Tourism": "travel",
    "Home Improvement": "home_improvement",
    "Banking & Finance": "banking",
    "Entertainment": "entertainment",
    "Casual Dining & QSR": "dining_qsr",
    "Automotive": "auto",
}

SPORTS = {
    "NFL - Regular Season": "nfl_reg",
    "NFL - Playoffs": "nfl_playoffs",
    "NFL - Home Team": "nfl_home_team",
    "NBA - Regular Season": "nba_reg",
    "NBA - Playoffs": "nba_playoffs",
    "WNBA - Regular Season": "wnba_reg",
    "WNBA - Playoffs": "wnba_playoffs",
    "NHL - Regular Season": "nhl_reg",
    "NHL - Playoffs": "nhl_playoffs",
    "MLB - Regular Season": "mlb_reg",
    "MLB - Playoffs": "mlb_playoffs",
    "NCAAF - Regular Season": "ncaaf_reg",
    "NCAAF - Conference": "ncaaf_conference",
    "NCAAF - Playoffs": "ncaaf_playoffs",
    "NCAAF - Home Team": "ncaaf_home_team",
    "NCAA Basketball": "ncaa_basketball",
    "Golf / PGA": "golf_pga",
    "Soccer (Professional)": "soccer_pro",
    "Prestige Sports": "prestige_sports",
    "All Live Sports": "all_live_sports",
}

# Hardcoded product list + CPM defaults (spec section 3 "products" table,
# no Supabase yet). line_type mirrors the DB schema's premion/am/broadcast.
PRODUCTS = {
    "premion_streaming_tv": {"label": "Premion Streaming TV", "default_cpm": 35.00, "line_type": "premion"},
    "streaming_retargeting": {"label": "Streaming Retargeting", "default_cpm": 12.00, "line_type": "premion"},
    "audience_targeting": {"label": "Audience Targeting", "default_cpm": 40.00, "line_type": "am"},
    "geofencing": {"label": "Geofencing", "default_cpm": 18.00, "line_type": "am"},
    "site_retargeting_display": {"label": "Site Retargeting - Display", "default_cpm": 8.00, "line_type": "am"},
    "site_retargeting_preroll": {"label": "Site Retargeting - Pre-roll", "default_cpm": 15.00, "line_type": "am"},
    "live_sports": {"label": "Live Sports", "default_cpm": 45.00, "line_type": "premion"},
    "broadcast_tv": {"label": "Broadcast Schedule", "default_cpm": 5.50, "line_type": "broadcast"},
}

STREAMING_RETARGETING_TARGETING = "Retarget Exposed CTV Viewers"
LIVE_SPORTS_TARGETING = "100% Live, 100% In-Game, 100% CTV"
MEDIA_PLAN_FIELDS = ["Tactic", "Flight", "Geo", "Targeting", "Impressions", "CPM"]

PRESETS = {
    "Quick Pitch": "quick_pitch",
    "Standard": "standard",
    "Extended": "extended",
}


def check_password():
    if st.session_state.get("authed"):
        return True

    st.title("Premion Proposal Builder")
    pwd = st.text_input("Password", type="password")
    if st.button("Log in"):
        expected = st.secrets.get("APP_PASSWORD")
        if expected is None:
            st.error("APP_PASSWORD is not set in .streamlit/secrets.toml.")
        elif pwd == expected:
            st.session_state["authed"] = True
            st.rerun()
        else:
            st.error("Incorrect password.")
    return False


def lines_to_bullets(text):
    return [line.strip() for line in text.splitlines() if line.strip()]


def first_line(text):
    bullets = lines_to_bullets(text)
    return bullets[0] if bullets else ""


def month_list(start, end):
    """['Sep 2026', 'Oct 2026', ...] inclusive, for every calendar month
    between start and end (order-swapped defensively if entered backwards).
    """
    if end < start:
        start, end = end, start
    months = []
    cur = date(start.year, start.month, 1)
    last = date(end.year, end.month, 1)
    while cur <= last:
        months.append(cur.strftime("%b %Y"))
        next_month = cur.month + 1
        next_year = cur.year
        if next_month > 12:
            next_month = 1
            next_year += 1
        cur = date(next_year, next_month, 1)
    return months


def format_flight_label(all_months, active_months):
    if not active_months:
        return ""
    if active_months == all_months and len(active_months) > 1:
        return f"{active_months[0]} - {active_months[-1]}"
    return ", ".join(active_months)


def resolve_row_defaults(tactic, default_geo, default_targeting, flight_label):
    """What a row's Flight/Geo/Targeting should be right now, given its
    Tactic name -- used both at initial seed time and to soft-update
    not-yet-edited rows when the shared form fields change."""
    if tactic == PRODUCTS["streaming_retargeting"]["label"]:
        targeting = STREAMING_RETARGETING_TARGETING
    elif tactic.startswith("Live Sports"):
        targeting = LIVE_SPORTS_TARGETING
    else:
        targeting = default_targeting
    return {"Flight": flight_label, "Geo": default_geo, "Targeting": targeting}


def seed_media_plan_rows(selections, market_label, default_targeting, flight_label):
    """Section C -> Section E: each selected product seeds a proposal line
    with its default CPM (spec section 5, Section C description). Targeting
    defaults to the Campaign Specs Audience field, except for Streaming
    Retargeting and Live Sports, which have fixed standard targeting copy.
    Geo/Flight default to the Geography/Timing-derived values."""
    rows = []
    products = selections["products"]

    def _row(label, cpm, targeting):
        return {"Tactic": label, "Flight": flight_label, "Geo": market_label,
                "Targeting": targeting, "Impressions": 0, "CPM": cpm}

    if selections.get("_premion_streaming_tv"):
        p = PRODUCTS["premion_streaming_tv"]
        rows.append(_row(p["label"], p["default_cpm"], default_targeting))

    if products.get("streaming_retargeting"):
        p = PRODUCTS["streaming_retargeting"]
        rows.append(_row(p["label"], p["default_cpm"], STREAMING_RETARGETING_TARGETING))

    am = products.get("audience_marketplace", {})
    if am.get("enabled"):
        if am.get("audience_targeting"):
            p = PRODUCTS["audience_targeting"]
            rows.append(_row(p["label"], p["default_cpm"], default_targeting))
        if am.get("geofencing"):
            p = PRODUCTS["geofencing"]
            rows.append(_row(p["label"], p["default_cpm"], default_targeting))
        if am.get("site_retargeting_display"):
            p = PRODUCTS["site_retargeting_display"]
            rows.append(_row(p["label"], p["default_cpm"], default_targeting))
        if am.get("site_retargeting_preroll"):
            p = PRODUCTS["site_retargeting_preroll"]
            rows.append(_row(p["label"], p["default_cpm"], default_targeting))

    sports = products.get("live_sports", {})
    if sports.get("enabled"):
        for sport_key in sports.get("sports", []):
            label = next((k for k, v in SPORTS.items() if v == sport_key), sport_key)
            rows.append(_row(f"Live Sports - {label}", PRODUCTS["live_sports"]["default_cpm"], LIVE_SPORTS_TARGETING))

    if products.get("total_tv"):
        p = PRODUCTS["broadcast_tv"]
        rows.append(_row(p["label"], p["default_cpm"], default_targeting))

    if not rows:
        rows.append(_row("", 0.0, ""))

    return rows


def build_included_list(targeting, commercial_production):
    included = [
        "Dedicated Account Management Team",
        "Monthly Reporting Calls & Optimizations",
        "Dashboard Access",
        "Web Attribution (Pixel Required)",
    ]
    if commercial_production:
        included.append("Commercial Production")
    if targeting.get("sales_attribution"):
        included.append("Sales Attribution (CRM Upload Required)")
    if targeting.get("brand_lift"):
        included.append("Brand Lift Study")
    return included


def main():
    if not check_password():
        return

    st.title("Premion Proposal Builder")
    st.caption("Phase 1 -- hardcoded product/CPM list, manually-typed Campaign Specs copy. No Supabase or Claude API yet.")

    # ---------------- Section A: Client basics ----------------
    st.header("A. Client basics")
    col1, col2 = st.columns(2)
    with col1:
        client_name = st.text_input("Client name", value="Acme Test Co")
        market_choice = st.radio("Market", ["DC", "Harrisburg"], horizontal=True)
        vertical_choice = st.selectbox("Vertical", list(VERTICALS.keys()), index=0)
        agency_involved = st.toggle("Ad agency involved? (gross markup x1.15)", value=False)
    with col2:
        logo_file = st.file_uploader("Client logo", type=["png", "jpg", "jpeg"])
        discovery_notes = st.text_area("Discovery notes", height=100, help="Reference only in phase 1 -- no Claude API wired up yet.")

    vertical_key = VERTICALS[vertical_choice]
    market_label = "Washington, DC DMA" if market_choice == "DC" else "Harrisburg DMA"

    if vertical_key == "healthcare":
        st.info("Healthcare targeting (Crossix) is automatically included for the Healthcare vertical.")
    if vertical_key == "travel":
        st.info("Arrivalist destination attribution is automatically included for Travel & Tourism.")
    if vertical_key == "auto":
        st.info("Polk Audiences / Polk Signals automotive targeting is automatically included for Automotive.")

    # ---------------- Section B: Deck scope ----------------
    st.header("B. Deck scope")
    col1, col2 = st.columns(2)
    with col1:
        preset = st.radio("Preset", list(PRESETS.keys()), index=1, horizontal=True)
        preset_key = PRESETS[preset]
        if preset_key == "quick_pitch":
            st.caption("Quick Pitch: client title, What You Told Us, Why Premion, and one targeting/vertical slide, plus selected add-ons.")
        elif preset_key == "standard":
            st.caption("Standard: a fixed core slide set (cover, specs, intro, one Premium Content highlight, personalized targeting, core attribution, media plan) plus selected add-ons.")
        else:
            st.caption("Extended: the full core-content deck plus selected add-ons.")
    with col2:
        tegna_positioning = st.toggle("Include TEGNA media positioning slides", value=True)
        include_vertical_slides = True
        include_avails_template = False
        if vertical_key != "none":
            include_vertical_slides = st.toggle(f"Include {vertical_choice} vertical slides", value=True)
            include_avails_template = st.toggle("Include personalized targeting / avails table", value=True)
            if not include_avails_template:
                st.caption(f"The {vertical_choice} vertical's own static Precision Targeting slide will be included instead.")

    # ---------------- Section C: Products ----------------
    st.header("C. Products")
    col1, col2, col3 = st.columns(3)
    with col1:
        premion_streaming_tv = st.checkbox("Premion Streaming TV", value=True)
        streaming_retargeting = st.checkbox("Streaming Retargeting")
    with col2:
        am_enabled = st.checkbox("Audience Marketplace")
        am_audience_targeting = am_geofencing = am_site_display = am_site_preroll = False
        if am_enabled:
            am_audience_targeting = st.checkbox("  Audience Targeting", key="am_at")
            am_geofencing = st.checkbox("  Geofencing", key="am_gf")
            am_site_display = st.checkbox("  Site Retargeting - Display", key="am_srd")
            am_site_preroll = st.checkbox("  Site Retargeting - Pre-roll", key="am_srp")
    with col3:
        total_tv = st.checkbox("Total TV", value=False)
        live_sports_enabled = st.checkbox("Live Sports", value=False)
        selected_sports = []
        if live_sports_enabled:
            selected_sports = st.multiselect("Sports packages", list(SPORTS.keys()))

    # ---------------- Section D: Targeting & attribution ----------------
    st.header("D. Targeting & attribution")
    st.caption("Standard (always included): Dashboard, Reporting, Web Attribution.")
    col1, col2 = st.columns(2)
    with col1:
        spanish_campaign = st.toggle("Spanish-language campaign?")
        first_party_data = st.checkbox("First-party data targeting")
        linear_reach_extension = st.checkbox("Linear reach extension", disabled=not total_tv,
                                              help="Only available when Total TV is selected")
        if not total_tv:
            linear_reach_extension = False
    with col2:
        sales_attribution = st.checkbox("Sales attribution")
        brand_lift = st.checkbox("Brand lift")
        commercial_production = st.checkbox("Commercial production")

    # ---------------- Section D2: Audiences & avails ----------------
    avails_rows = []
    if vertical_key != "none" and include_avails_template:
        st.header("D2. Audiences & avails")
        default_avails = pd.DataFrame([
            {"Audience": "", "Geo": market_label, "Max Monthly Avails": 0},
        ])
        avails_df = st.data_editor(default_avails, num_rows="dynamic", key="avails_editor", use_container_width=True)
        total_avails_val = int(avails_df["Max Monthly Avails"].fillna(0).sum())
        st.caption(f"Total avails: {total_avails_val:,}")
        for _, row in avails_df.iterrows():
            if str(row["Audience"]).strip():
                avails_rows.append({
                    "audience": str(row["Audience"]),
                    "geo": str(row["Geo"]),
                    "avails": f"{int(row['Max Monthly Avails']):,}",
                })

    # ---------------- Section A2: Campaign Specs (manual copy) ----------------
    st.header("Campaign Specs copy")
    st.caption("Manually typed for phase 1 (no Claude API yet). One bullet per line. Audience/Geography feed the media plan's Targeting/Geo defaults below.")
    spec_col1, spec_col2 = st.columns(2)
    with spec_col1:
        goals_text = st.text_area("Goals & Approach", height=90)
        audience_text = st.text_area("Audience", height=90)
        geography_text = st.text_area("Geography", height=90, value=market_label)
    with spec_col2:
        budget_text = st.text_area("Budget & Allocation", height=90)
        placements_text = st.text_area("Placements & Creative", height=90)
        timing_text = st.text_area("Timing", height=90, help="Narrative copy for the Campaign Specs slide. Actual flight dates for the media plan are set below.")

    default_targeting = first_line(audience_text)
    default_geo = first_line(geography_text) or market_label

    # ---------------- Section E: Proposal / media plan ----------------
    st.header("E. Proposal / media plan")
    st.caption("Phase 1 supports a single plan option (Option A). Cost = Impressions/1000 x CPM, gross x1.15 if agency toggle is on.")

    st.subheader("Flight & breakout")
    fcol1, fcol2, fcol3 = st.columns([1, 1, 1])
    with fcol1:
        flight_start = st.date_input("Flight start", value=date(2026, 9, 1))
    with fcol2:
        flight_end = st.date_input("Flight end", value=date(2026, 11, 30))
    with fcol3:
        breakout_mode = st.radio("Breakout", ["Monthly (default)", "Full Flight"], horizontal=True)

    all_months = month_list(flight_start, flight_end)
    active_months = st.multiselect(
        "Active months (uncheck to skip a month -- custom flighting)",
        all_months, default=all_months,
    )
    n_months = max(1, len(active_months))
    flight_label = format_flight_label(all_months, active_months) or "TBD"
    st.caption(f"{n_months} active month(s): {flight_label}")

    products_selection = {
        "streaming_retargeting": streaming_retargeting,
        "audience_marketplace": {
            "enabled": am_enabled,
            "audience_targeting": am_audience_targeting,
            "geofencing": am_geofencing,
            "site_retargeting_display": am_site_display,
            "site_retargeting_preroll": am_site_preroll,
        },
        "live_sports": {
            "enabled": live_sports_enabled,
            "sports": [SPORTS[s] for s in selected_sports],
        },
        "total_tv": total_tv,
    }
    seed_selections = {"products": products_selection, "_premion_streaming_tv": premion_streaming_tv}
    product_seed_key = str(seed_selections)
    # Shared Audience/Geography/Flight fields drive a *soft* update: only
    # rows the user hasn't touched get refreshed. Rows the user has edited
    # (tracked via media_plan_dirty) keep whatever they typed -- editing a
    # shared field never silently wipes a customized line.
    shared_fields_key = default_targeting + "||" + default_geo + "||" + flight_label

    if "media_plan_version" not in st.session_state:
        st.session_state["media_plan_version"] = 0

    if "media_plan_rows" not in st.session_state:
        st.session_state["media_plan_rows"] = seed_media_plan_rows(
            seed_selections, default_geo, default_targeting, flight_label)
        st.session_state["media_plan_dirty"] = [False] * len(st.session_state["media_plan_rows"])
        st.session_state["_product_seed_key"] = product_seed_key
        st.session_state["_shared_fields_key"] = shared_fields_key
    elif st.session_state["_product_seed_key"] != product_seed_key:
        # Product selections changed -- which tactics exist is a structural
        # change, so the row list itself is rebuilt from scratch.
        st.session_state["media_plan_rows"] = seed_media_plan_rows(
            seed_selections, default_geo, default_targeting, flight_label)
        st.session_state["media_plan_dirty"] = [False] * len(st.session_state["media_plan_rows"])
        st.session_state["_product_seed_key"] = product_seed_key
        st.session_state["_shared_fields_key"] = shared_fields_key
        st.session_state["media_plan_version"] += 1
    elif st.session_state["_shared_fields_key"] != shared_fields_key:
        for row, dirty in zip(st.session_state["media_plan_rows"], st.session_state["media_plan_dirty"]):
            if not dirty:
                row.update(resolve_row_defaults(row.get("Tactic", ""), default_geo, default_targeting, flight_label))
        st.session_state["_shared_fields_key"] = shared_fields_key
        st.session_state["media_plan_version"] += 1

    plan_df = pd.DataFrame(st.session_state["media_plan_rows"])
    editor_key = f"media_plan_editor_{st.session_state['media_plan_version']}"
    impressions_label = "Impressions (Monthly)" if breakout_mode.startswith("Monthly") else "Impressions (Full Flight)"
    edited_plan_df = st.data_editor(
        plan_df, num_rows="dynamic", key=editor_key, use_container_width=True,
        column_config={"Impressions": st.column_config.NumberColumn(impressions_label)},
    )
    st.caption("Duplicate a line (e.g. same product, different audience/impressions), then edit the copy. "
               "Rows you've customized won't auto-update when Audience/Geography/flight dates change above.")

    # Reconcile edits: diff against the pre-render snapshot to update dirty
    # flags, then persist both back to session_state as the new baseline.
    prev_rows = st.session_state["media_plan_rows"]
    prev_dirty = st.session_state["media_plan_dirty"]
    new_rows = edited_plan_df.to_dict("records")
    new_dirty = []
    for i, row in enumerate(new_rows):
        if i < len(prev_rows):
            changed = any(str(row.get(f, "")) != str(prev_rows[i].get(f, "")) for f in MEDIA_PLAN_FIELDS)
            new_dirty.append(prev_dirty[i] or changed)
        else:
            new_dirty.append(True)  # a row added via the grid's own "+" is treated as customized
    st.session_state["media_plan_rows"] = new_rows
    st.session_state["media_plan_dirty"] = new_dirty

    dcol1, dcol2 = st.columns([3, 1])
    tactic_labels = [f"{i}: {row.get('Tactic', '') or '(blank)'}" for i, row in enumerate(new_rows)]
    with dcol1:
        dup_pick = st.selectbox("Line to duplicate", tactic_labels, label_visibility="collapsed") if tactic_labels else None
    with dcol2:
        if st.button("Duplicate line", disabled=not tactic_labels):
            idx = int(dup_pick.split(":")[0])
            st.session_state["media_plan_rows"] = new_rows + [dict(new_rows[idx])]
            st.session_state["media_plan_dirty"] = new_dirty + [True]  # a duplicate is immediately customizable
            st.session_state["media_plan_version"] += 1
            st.rerun()

    markup = 1.15 if agency_involved else 1.0
    preview_rows = []
    monthly_total_impressions = 0.0
    monthly_total_cost = 0.0
    for _, row in edited_plan_df.iterrows():
        if not str(row.get("Tactic", "")).strip():
            continue
        entered_impressions = float(row.get("Impressions") or 0)
        cpm = float(row.get("CPM") or 0)
        if breakout_mode.startswith("Full Flight"):
            full_flight_impressions = entered_impressions
            full_flight_cost = (full_flight_impressions / 1000.0) * cpm * markup
            monthly_impressions = full_flight_impressions / n_months
            monthly_cost = full_flight_cost / n_months
        else:
            monthly_impressions = entered_impressions
            monthly_cost = (monthly_impressions / 1000.0) * cpm * markup
            full_flight_impressions = monthly_impressions * n_months
            full_flight_cost = monthly_cost * n_months

        monthly_total_impressions += monthly_impressions
        monthly_total_cost += monthly_cost
        preview_rows.append({
            "tactic": str(row["Tactic"]), "flight": str(row.get("Flight", "")) or flight_label,
            "geo": str(row.get("Geo", "")), "targeting": str(row.get("Targeting", "")),
            "monthly_impressions": monthly_impressions, "monthly_cost": monthly_cost,
            "full_flight_impressions": full_flight_impressions, "full_flight_cost": full_flight_cost,
        })

    full_flight_total_impressions = monthly_total_impressions * n_months
    full_flight_total_cost = monthly_total_cost * n_months

    preview_display = pd.DataFrame([
        {
            "tactic": r["tactic"], "flight": r["flight"], "geo": r["geo"], "targeting": r["targeting"],
            "monthly impressions": f"{int(r['monthly_impressions']):,}",
            "monthly cost": f"${r['monthly_cost']:,.0f}",
            "full flight impressions": f"{int(r['full_flight_impressions']):,}",
            "full flight cost": f"${r['full_flight_cost']:,.0f}",
        }
        for r in preview_rows
    ]) if preview_rows else pd.DataFrame(columns=[
        "tactic", "flight", "geo", "targeting", "monthly impressions", "monthly cost",
        "full flight impressions", "full flight cost"])
    st.dataframe(preview_display, use_container_width=True)

    gross_suffix = " gross" if agency_involved else ""
    st.caption(f"Monthly totals: {int(monthly_total_impressions):,} impressions / ${monthly_total_cost:,.0f}{gross_suffix}")
    st.caption(f"**Full Flight Total ({n_months} month{'s' if n_months != 1 else ''}): "
               f"{int(full_flight_total_impressions):,} impressions / ${full_flight_total_cost:,.0f}{gross_suffix}**")

    included_list = build_included_list(
        {"sales_attribution": sales_attribution, "brand_lift": brand_lift},
        commercial_production,
    )
    st.caption("Included with Campaign: " + ", ".join(included_list))

    # ---------------- Generate ----------------
    st.header("Generate")
    proposal_title = st.text_input("Proposal title (appears on cover + media plan)",
                                    value="Total TV Strategy" if total_tv else "CTV Strategy")

    if st.button("Generate proposal", type="primary"):
        selections = {
            "preset": preset_key,
            "market": market_choice,
            "vertical": vertical_key if include_vertical_slides else "none",
            "agency_involved": agency_involved,
            "spanish_campaign": spanish_campaign,
            "tegna_positioning": tegna_positioning,
            "include_avails_template": include_avails_template and bool(avails_rows),
            "products": products_selection,
        }
        selections["targeting_attribution"] = {
            "first_party_data": first_party_data,
            "linear_reach_extension": linear_reach_extension,
            "sales_attribution": sales_attribution,
            "brand_lift": brand_lift,
        }

        if not avails_rows:
            avails_rows_final = [{"audience": "", "geo": market_label, "avails": "0"}]
            total_avails_str = "0"
        else:
            avails_rows_final = avails_rows
            total_avails_str = f"{sum(int(r['avails'].replace(',', '')) for r in avails_rows):,}"

        media_plan_rows_final = [
            {"tactic": r["tactic"], "flight": r["flight"], "geo": r["geo"], "targeting": r["targeting"],
             "impressions": f"{int(r['monthly_impressions']):,}", "cost": f"${r['monthly_cost']:,.0f}" + (" (Gross)" if agency_involved else "")}
            for r in preview_rows
        ] or [{"tactic": "", "flight": flight_label, "geo": market_label, "targeting": "", "impressions": "0", "cost": "$0"}]

        full_flight_total = None
        if n_months > 1 and preview_rows:
            full_flight_total = {
                "label": f"Full Flight Total ({n_months} months)",
                "impressions": f"{int(full_flight_total_impressions):,}",
                "cost": f"${full_flight_total_cost:,.0f}" + (" (Gross)" if agency_involved else ""),
            }

        fill_data = {
            "client_name": client_name or "Client",
            "proposal_title": proposal_title,
            "logo_path": io.BytesIO(logo_file.getvalue()) if logo_file else "placeholder_logo.png",
            "vertical_display": vertical_choice if vertical_key != "none" else "",
            "campaign_specs": {
                "GOALS_BULLETS": lines_to_bullets(goals_text) or ["--"],
                "AUDIENCE_BULLETS": lines_to_bullets(audience_text) or ["--"],
                "GEOGRAPHY_BULLETS": lines_to_bullets(geography_text) or ["--"],
                "BUDGET_BULLETS": lines_to_bullets(budget_text) or ["--"],
                "PLACEMENTS_BULLETS": lines_to_bullets(placements_text) or ["--"],
                "TIMING_BULLETS": lines_to_bullets(timing_text) or [flight_label],
            },
            "avails": {
                "rows": avails_rows_final,
                "total_avails": total_avails_str,
            },
            "media_plan": {
                "plan_title": proposal_title,
                "rows": media_plan_rows_final,
                "totals_label": "Monthly Totals",
                "total_impressions": f"{int(monthly_total_impressions):,}",
                "total_cost": f"${monthly_total_cost:,.0f}" + (" (Gross)" if agency_involved else ""),
                "full_flight_total": full_flight_total,
                "included_list": included_list,
            },
        }

        with st.spinner("Assembling deck..."):
            try:
                prs, original_count, kept_count = assembly.build_presentation(MASTER_DECK_PATH, selections)
                assembly.personalize(prs, fill_data)
                buffer = io.BytesIO()
                prs.save(buffer)
                buffer.seek(0)
            except Exception as exc:
                st.error(f"Assembly failed: {exc}")
                raise

        st.success(f"Assembled {kept_count} of {original_count} slides.")
        st.download_button(
            "Download .pptx",
            data=buffer,
            file_name=f"{(client_name or 'client').replace(' ', '_')}_proposal.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        )


main()

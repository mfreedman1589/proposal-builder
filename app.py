"""
app.py -- Phase 1 Streamlit form for the Premion Proposal Builder (build spec
section 5), wired directly to assembly.py's selection -> deletion -> fill
pipeline. No Supabase, no Claude API yet: the product list/CPMs are a
hardcoded dict below, and Campaign Specs bullets are typed in by hand.
"""

import io

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


def seed_media_plan_rows(selections, market_label):
    """Section C -> Section E: each selected product seeds a proposal line
    with its default CPM (spec section 5, Section C description)."""
    rows = []
    products = selections["products"]

    if selections.get("_premion_streaming_tv"):
        p = PRODUCTS["premion_streaming_tv"]
        rows.append({"Tactic": p["label"], "Flight": "", "Geo": market_label, "Targeting": "", "Impressions": 0, "CPM": p["default_cpm"]})

    if products.get("streaming_retargeting"):
        p = PRODUCTS["streaming_retargeting"]
        rows.append({"Tactic": p["label"], "Flight": "", "Geo": market_label, "Targeting": "", "Impressions": 0, "CPM": p["default_cpm"]})

    am = products.get("audience_marketplace", {})
    if am.get("enabled"):
        if am.get("audience_targeting"):
            p = PRODUCTS["audience_targeting"]
            rows.append({"Tactic": p["label"], "Flight": "", "Geo": market_label, "Targeting": "", "Impressions": 0, "CPM": p["default_cpm"]})
        if am.get("geofencing"):
            p = PRODUCTS["geofencing"]
            rows.append({"Tactic": p["label"], "Flight": "", "Geo": market_label, "Targeting": "", "Impressions": 0, "CPM": p["default_cpm"]})
        if am.get("site_retargeting_display"):
            p = PRODUCTS["site_retargeting_display"]
            rows.append({"Tactic": p["label"], "Flight": "", "Geo": market_label, "Targeting": "", "Impressions": 0, "CPM": p["default_cpm"]})
        if am.get("site_retargeting_preroll"):
            p = PRODUCTS["site_retargeting_preroll"]
            rows.append({"Tactic": p["label"], "Flight": "", "Geo": market_label, "Targeting": "", "Impressions": 0, "CPM": p["default_cpm"]})

    sports = products.get("live_sports", {})
    if sports.get("enabled"):
        for sport_key in sports.get("sports", []):
            label = next((k for k, v in SPORTS.items() if v == sport_key), sport_key)
            rows.append({"Tactic": f"Live Sports - {label}", "Flight": "", "Geo": market_label, "Targeting": "", "Impressions": 0, "CPM": PRODUCTS["live_sports"]["default_cpm"]})

    if products.get("total_tv"):
        p = PRODUCTS["broadcast_tv"]
        rows.append({"Tactic": p["label"], "Flight": "", "Geo": market_label, "Targeting": "", "Impressions": 0, "CPM": p["default_cpm"]})

    if not rows:
        rows.append({"Tactic": "", "Flight": "", "Geo": market_label, "Targeting": "", "Impressions": 0, "CPM": 0.0})

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


def lines_to_bullets(text):
    return [line.strip() for line in text.splitlines() if line.strip()]


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
        agency_involved = st.toggle("Ad agency involved? (gross markup x1.15)")
    with col2:
        logo_file = st.file_uploader("Client logo", type=["png", "jpg", "jpeg"])
        spanish_campaign = st.toggle("Spanish-language campaign?")
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
        preset = st.radio("Preset", ["Full Proposal", "Quick Pitch"], horizontal=True)
        preset_key = "full_proposal" if preset == "Full Proposal" else "quick_pitch"
        if preset_key == "quick_pitch":
            st.caption("Quick Pitch drops the full core-content deck; vertical-specific slides for the chosen vertical still come through automatically.")
    with col2:
        tegna_positioning = st.toggle("Include TEGNA media positioning slides", value=True)
        include_vertical_slides = True
        include_avails_template = False
        if vertical_key != "none":
            include_vertical_slides = st.toggle(f"Include {vertical_choice} vertical slides", value=True)
            include_avails_template = st.toggle("Include personalized targeting / avails table", value=True)

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
        total_tv = st.checkbox("Total TV")
        live_sports_enabled = st.checkbox("Live Sports")
        selected_sports = []
        if live_sports_enabled:
            selected_sports = st.multiselect("Sports packages", list(SPORTS.keys()))

    # ---------------- Section D: Targeting & attribution ----------------
    st.header("D. Targeting & attribution")
    st.caption("Standard (always included): Dashboard, Reporting, Web Attribution.")
    col1, col2 = st.columns(2)
    with col1:
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
    st.caption("Manually typed for phase 1 (no Claude API yet). One bullet per line.")
    spec_col1, spec_col2 = st.columns(2)
    with spec_col1:
        goals_text = st.text_area("Goals & Approach", height=90)
        audience_text = st.text_area("Audience", height=90)
        geography_text = st.text_area("Geography", height=90, value=market_label)
    with spec_col2:
        budget_text = st.text_area("Budget & Allocation", height=90)
        placements_text = st.text_area("Placements & Creative", height=90)
        timing_text = st.text_area("Timing", height=90)

    # ---------------- Section E: Proposal / media plan ----------------
    st.header("E. Proposal / media plan")
    st.caption("Phase 1 supports a single plan option (Option A). Cost = Impressions/1000 x CPM, gross x1.15 if agency toggle is on.")

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

    if "media_plan_rows" not in st.session_state or st.session_state.get("_seed_key") != str(seed_selections):
        st.session_state["media_plan_rows"] = seed_media_plan_rows(seed_selections, market_label)
        st.session_state["_seed_key"] = str(seed_selections)

    plan_df = pd.DataFrame(st.session_state["media_plan_rows"])
    edited_plan_df = st.data_editor(plan_df, num_rows="dynamic", key="media_plan_editor", use_container_width=True)

    markup = 1.15 if agency_involved else 1.0
    preview_rows = []
    total_impressions = 0
    total_cost = 0.0
    for _, row in edited_plan_df.iterrows():
        if not str(row.get("Tactic", "")).strip():
            continue
        impressions = float(row.get("Impressions") or 0)
        cpm = float(row.get("CPM") or 0)
        cost = (impressions / 1000.0) * cpm * markup
        total_impressions += impressions
        total_cost += cost
        preview_rows.append({
            "tactic": str(row["Tactic"]), "flight": str(row.get("Flight", "")), "geo": str(row.get("Geo", "")),
            "targeting": str(row.get("Targeting", "")), "impressions": f"{int(impressions):,}",
            "cost": f"${cost:,.0f}" + (" (Gross)" if agency_involved else ""),
        })

    preview_display = pd.DataFrame(preview_rows) if preview_rows else pd.DataFrame(
        columns=["tactic", "flight", "geo", "targeting", "impressions", "cost"])
    st.dataframe(preview_display, use_container_width=True)
    st.caption(f"Totals: {int(total_impressions):,} impressions / ${total_cost:,.0f}" + (" gross" if agency_involved else ""))

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
             "impressions": r["impressions"], "cost": r["cost"]}
            for r in preview_rows
        ] or [{"tactic": "", "flight": "", "geo": market_label, "targeting": "", "impressions": "0", "cost": "$0"}]

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
                "TIMING_BULLETS": lines_to_bullets(timing_text) or ["--"],
            },
            "avails": {
                "rows": avails_rows_final,
                "total_avails": total_avails_str,
            },
            "media_plan": {
                "plan_title": proposal_title,
                "rows": media_plan_rows_final,
                "totals_label": "Monthly Totals",
                "total_impressions": f"{int(total_impressions):,}",
                "total_cost": f"${total_cost:,.0f}" + (" (Gross)" if agency_involved else ""),
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

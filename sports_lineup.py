"""The sports lineup: which packages are sold, at what rate, in what season,
and the rules a live sports line has to follow.

Pure -- no DB, no Streamlit -- so the Proposal Builder, the drafting path,
setup_supabase.py and the Attribution Report Builder all read one table.

TEGNA sources live sports inventory directly through the Premion DSP
(2027 lineup, Matt, 2026-11). Rates are net CPM. The live rate card in
Supabase (`products`, rows keyed "sport:<key>") is what prices a line;
`cpm` here is its fallback and what `setup_supabase.py sports` writes.
Season dates, minimums and the NCAA restriction are product rules, not
rates, and live only here.

Retired packages stay known (RETIRED) so an old proposal still loads,
reprices and rebuilds as presented -- they're hidden from the picker, the
drafting prompt and anything new.
"""

import re
from dataclasses import dataclass
from datetime import date

# Every live package also has a minimum per geography (net, full flight).
GEO_MINIMUM = 2500.0
# The NCAAF Conferences package: the advertiser picks one; the choice is
# shown in the plan line's Targeting text.
NCAAF_CONFERENCES = ("SEC", "Big Ten", "Big 12", "ACC")
# What every live package's plan line says in its Targeting column.
LIVE_TARGETING = "100% Live, 100% In-Game, 100% CTV"
SPORTS_SELECT_TARGETING = "Live + on-demand sports content, 100% CTV"
# Creative a live sports line accepts -- added to Campaign Specs Placements
# whenever a live sports line is on the plan.
LIVE_CREATIVE_SPEC = "Live sports creative: :15 and :30 MP4 files only (no third-party VAST tags)."
# Reporting a live sports line gets -- an Included-with-Campaign entry.
LIVE_REPORTING = "Live Sports Reporting: Geo, Impressions, VCR & Channel-Level"


@dataclass(frozen=True)
class Package:
    key: str
    name: str                 # as the client reads it ("MLB Regular Season")
    cpm: float                # net
    order_minimum: float      # net, full flight, per order
    season: tuple = None      # (start, end) dates; None = full year
    tentative: bool = False   # season dates not yet confirmed
    live: bool = True         # Sports Select is a network package, not live
    ncaa: bool = False        # NCAA restriction applies
    picker: str = ""          # the form's multiselect label

    @property
    def tactic(self):
        """The media-plan Tactic a line for this package seeds with."""
        if not self.live or self.name.startswith("Live Sports"):
            return self.name
        return f"Live Sports - {self.name}"

    @property
    def label(self):
        return self.picker or self.name


PACKAGES = {p.key: p for p in [
    Package("all_live_sports", "Live Sports Select", 51.0, 10000.0,
            picker="Live Sports Select (RON, all sports)"),
    Package("mlb_reg", "MLB Regular Season", 52.0, 15000.0,
            (date(2027, 3, 24), date(2027, 9, 26))),
    Package("nba_reg", "NBA Regular Season", 57.0, 15000.0,
            (date(2026, 10, 20), date(2027, 4, 11))),
    Package("nhl_reg", "NHL Regular Season", 57.0, 15000.0,
            (date(2026, 10, 20), date(2027, 4, 10))),
    Package("ncaa_basketball", "NCAA Men's Basketball Regular Season", 55.0, 15000.0,
            (date(2026, 11, 1), date(2027, 3, 7)), ncaa=True),
    Package("ncaaf_reg", "NCAA Football Regular Season", 60.0, 15000.0,
            (date(2027, 8, 28), date(2027, 11, 27)), ncaa=True),
    Package("ncaaf_conference", "NCAA Football Conferences", 61.0, 15000.0,
            (date(2027, 8, 28), date(2027, 11, 27)), ncaa=True),
    Package("nfl_reg", "NFL Regular Season", 62.0, 15000.0,
            (date(2027, 9, 9), date(2028, 1, 9)), tentative=True),
    Package("golf", "Golf", 61.0, 15000.0,
            (date(2027, 1, 21), date(2027, 8, 29)), tentative=True,
            picker="Golf (PGA Tour & LIV)"),
    Package("motorsports", "Motorsports", 59.0, 15000.0,
            (date(2027, 2, 21), date(2027, 11, 14))),
    # A sports NETWORK package (live + on-demand), not live sports: none of
    # the live rules apply, and its delivery arrives as regular OTT. Order
    # minimum: the standard Premion CTV minimum until confirmed -- the app
    # has no number for that, so none is checked.
    Package("sports_select", "Sports Select", 38.0, 0.0, live=False,
            picker="Sports Select (sports network, live + on-demand)"),
]}

# Retired 2026-11: {key: (name as last sold, last CPM, closest current key)}.
# Kept so old proposals load and rebuild; never offered for anything new.
RETIRED = {
    "nfl_playoffs": ("NFL - Playoffs", 66.0, "nfl_reg"),
    "nfl_home_team": ("NFL - Home Team", 85.0, "nfl_reg"),
    "nba_playoffs": ("NBA - Playoffs", 63.0, "nba_reg"),
    "wnba_reg": ("WNBA - Regular Season", 59.0, "all_live_sports"),
    "wnba_playoffs": ("WNBA - Playoffs", 61.0, "all_live_sports"),
    "nhl_playoffs": ("NHL - Playoffs", 63.0, "nhl_reg"),
    "mlb_playoffs": ("MLB - Playoffs", 63.0, "mlb_reg"),
    "ncaaf_playoffs": ("NCAAF - Playoffs", 63.0, "ncaaf_reg"),
    "ncaaf_home_team": ("NCAAF - Home Team", 70.0, "ncaaf_conference"),
    "pga_majors": ("PGA Majors", 62.0, "golf"),
    "soccer_pro": ("Soccer (Professional)", 57.0, "all_live_sports"),
    "prestige_sports": ("Prestige Sports", 54.0, "all_live_sports"),
}
# Names the current packages used to carry, so an old proposal's tactic
# still reads as that package.
FORMER_TACTICS = {
    "all_live_sports": "Live Sports - All Live Sports",
    "nfl_reg": "Live Sports - NFL - Regular Season",
    "nba_reg": "Live Sports - NBA - Regular Season",
    "nhl_reg": "Live Sports - NHL - Regular Season",
    "mlb_reg": "Live Sports - MLB - Regular Season",
    "ncaaf_reg": "Live Sports - NCAAF - Regular Season",
    "ncaaf_conference": "Live Sports - NCAAF - Conference",
    "ncaa_basketball": "Live Sports - NCAA Basketball",
    "golf": "Live Sports - Live Golf",
}

CURRENT_KEYS = list(PACKAGES)
LIVE_KEYS = [k for k, p in PACKAGES.items() if p.live]


def is_retired(key):
    return key in RETIRED


def is_live(key):
    """A live sports package -- current or retired (every retired one was live)."""
    return (key in PACKAGES and PACKAGES[key].live) or key in RETIRED


def name(key):
    if key in PACKAGES:
        return PACKAGES[key].name
    if key in RETIRED:
        return RETIRED[key][0]
    return key


def picker_label(key):
    if key in PACKAGES:
        return PACKAGES[key].label
    if key in RETIRED:
        return f"{RETIRED[key][0]} (no longer offered)"
    return key


def tactic(key):
    """The Tactic a line for `key` seeds with."""
    if key in PACKAGES:
        return PACKAGES[key].tactic
    if key in RETIRED:
        return f"Live Sports - {RETIRED[key][0]}"
    return f"Live Sports - {key}"


def fallback_cpm(key):
    if key in PACKAGES:
        return PACKAGES[key].cpm
    if key in RETIRED:
        return RETIRED[key][1]
    return None


def targeting_copy(key, conference=None):
    """A sports line's Targeting text. NCAAF Conferences names the
    conference when one was chosen."""
    if key in PACKAGES and not PACKAGES[key].live:
        return SPORTS_SELECT_TARGETING
    if key == "ncaaf_conference" and conference:
        return f"{conference} games -- {LIVE_TARGETING}"
    return LIVE_TARGETING


def closest_current(key):
    return RETIRED[key][2] if key in RETIRED else (key if key in PACKAGES else None)


def tactic_labels():
    """{tactic: key} for every package's current and former tactic, longest
    first, for recognising a sports row by its Tactic (a drafted or merged
    row appends its own suffix, so it's a prefix match)."""
    out = {}
    for key in list(PACKAGES) + list(RETIRED):
        out[tactic(key)] = key
    for key, old in FORMER_TACTICS.items():
        out.setdefault(old, key)
    return dict(sorted(out.items(), key=lambda kv: -len(kv[0])))


def key_for_tactic(tactic_text):
    """The package a plan row's Tactic belongs to, or None."""
    text = str(tactic_text or "")
    for label, key in tactic_labels().items():
        if text.startswith(label):
            return key
    return None


def conference_from_text(text):
    """The NCAAF conference a Targeting cell or the notes name, or None."""
    for conf in sorted(NCAAF_CONFERENCES, key=len, reverse=True):
        if re.search(rf"\b{re.escape(conf)}\b", str(text or ""), re.IGNORECASE):
            return conf
    return None


# --- rules -------------------------------------------------------------------

def _fmt_day(d):
    return f"{d:%b} {d.day}, {d.year}"


def season_label(key):
    p = PACKAGES.get(key)
    if not p or not p.season:
        return None
    start, end = p.season
    label = f"{_fmt_day(start)} – {_fmt_day(end)}"
    return label + (" (tentative)" if p.tentative else "")


def season_fit(key, flight_start, flight_end):
    """"inside" / "partial" / "outside" for a live package with a season,
    None for a full-year or non-live package or no flight."""
    p = PACKAGES.get(key)
    if not p or not p.live or not p.season or not flight_start or not flight_end:
        return None
    start, end = p.season
    if flight_start >= start and flight_end <= end:
        return "inside"
    if flight_end < start or flight_start > end:
        return "outside"
    return "partial"


def season_message(key, flight_start, flight_end):
    """A rep-facing sentence when the flight runs outside the package's
    season, else None."""
    fit = season_fit(key, flight_start, flight_end)
    if fit in (None, "inside"):
        return None
    p = PACKAGES[key]
    flight = f"{_fmt_day(flight_start)} – {_fmt_day(flight_end)}"
    if fit == "outside":
        return (f"{p.name} runs {season_label(key)}, and this flight ({flight}) falls "
                f"entirely outside it -- the line can't deliver as planned. Move the flight "
                f"into the season or use a different package.")
    return (f"{p.name} runs {season_label(key)}; part of this flight ({flight}) falls "
            f"outside it, and those months can't deliver. Price the line for the months "
            f"inside the season only.")


def tentative_message(key):
    p = PACKAGES.get(key)
    if p and p.tentative and p.season:
        return f"{p.name} season dates are tentative ({season_label(key).replace(' (tentative)', '')})."
    return None


def minimum_messages(key, order_cost, geo_costs=()):
    """Rep-facing sentences for a live package under its order minimum or
    any geography under the per-geo minimum. `order_cost` is the package's
    full-flight net total on the plan; `geo_costs` [(geo label, net)]."""
    p = PACKAGES.get(key)
    if not p or not p.live:
        return []
    out = []
    if p.order_minimum and 0 < order_cost < p.order_minimum:
        out.append(f"{p.name} is ${order_cost:,.0f} net on this plan; its order minimum is "
                   f"${p.order_minimum:,.0f}.")
    for geo, cost in geo_costs:
        if 0 < cost < GEO_MINIMUM:
            out.append(f"{p.name} in {geo} is ${cost:,.0f} net; each geography needs at "
                       f"least ${GEO_MINIMUM:,.0f}.")
    return out


# Advertisers NCAA packages can't run for: sports betting, casino, gambling,
# alcohol. Whole-word matches on the advertiser's own name and description --
# a match is a reason to stop and check, not proof (a rep can confirm it's
# not one of these).
_RESTRICTED_TERMS = {
    "sports betting": ("sportsbook", "sports book", "sports betting", "betting", "bet365",
                       "draftkings", "fanduel", "betmgm", "caesars sportsbook", "wager",
                       "wagering", "parlay"),
    "casino": ("casino", "casinos", "racino", "slots", "resort casino"),
    "gambling": ("gambling", "gaming commission", "lottery", "poker", "bingo hall",
                 "off-track betting", "igaming"),
    "alcohol": ("alcohol", "beer", "brewery", "brewing", "brewpub", "winery", "wine",
                "vineyard", "distillery", "distilling", "spirits", "liquor", "vodka",
                "whiskey", "bourbon", "tequila", "hard seltzer", "taproom", "beverage alcohol"),
}


def restricted_category(*texts):
    """("alcohol", "brewery") -- the category and the word that matched --
    for the first restricted term in `texts`, else None."""
    joined = " ".join(str(t or "") for t in texts)
    for category, terms in _RESTRICTED_TERMS.items():
        for term in sorted(terms, key=len, reverse=True):
            if re.search(rf"\b{re.escape(term)}\b", joined, re.IGNORECASE):
                return category, term
    return None


def ncaa_restriction_message(keys, category):
    names = [PACKAGES[k].name for k in keys if k in PACKAGES and PACKAGES[k].ncaa]
    if not names or not category:
        return None
    pros = ", ".join(PACKAGES[k].name for k in ("nfl_reg", "nba_reg", "nhl_reg", "mlb_reg"))
    return (f"NCAA packages can't run for {category[0]} advertisers -- remove "
            f"{', '.join(names)}. Pro packages are available instead: {pros}.")

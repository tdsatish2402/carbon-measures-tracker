# BCA Tracker — Streamlit app
# Reads BCA_tracker_only_BCA.xlsx (multi-tab, normalised) and renders filterable views.
# Designed to fail gracefully on half-edited data: junk rows are dropped and reported, never shown as real records.

import os
from collections import Counter

import pandas as pd
import streamlit as st

DATA_FILE = "BCA_tracker_only_BCA.xlsx"
EXAMPLE_ID = "XX-EXAMPLE"

# Drop your Trade Law Observatory logo next to this file under one of these names.
# If none is present the app simply renders without it.
LOGO_CANDIDATES = ["logo.png", "TLO_Logo.png", "logo.jpg", "logo.jpeg", "logo.webp",
                   "trade_law_observatory.png", "tlo_logo.png"]
LOGO_WIDTH = 280  # the supplied wordmark is ~4.6:1, so this renders ~61px tall

st.set_page_config(page_title="BCA Tracker", layout="wide",
                   initial_sidebar_state="expanded")

# ---------- styling (restrained; Streamlit limits raw HTML) ----------
st.markdown("""
<style>
  /* --ink drives headings, timeline header bars and lane labels.
     Current value is the original teal. To match the Trade Law Observatory
     wordmark instead, swap it for the navy #1D2657 or the purple #3F106E. */
  :root { --ink:#0E3B43; --paper:#FBF8F2; --line:#E5DFD3; }
  .stApp { background:var(--paper); }
  h1,h2,h3 { color:var(--ink); font-family:Georgia,'Times New Roman',serif; letter-spacing:-.2px; }
  .cardwrap { border:1px solid var(--line); border-radius:6px; padding:1rem 1.15rem;
              background:#fff; margin-bottom:1.1rem; }
  .pill { display:inline-block; padding:.12rem .55rem; border-radius:999px;
          font-size:.72rem; font-weight:700; letter-spacing:.02em; }
  .lab { color:#7A6F5B; font-size:.72rem; text-transform:uppercase; letter-spacing:.08em;
         font-weight:700; margin-top:.5rem; }
  .val { color:#22303a; font-size:.92rem; margin-bottom:.2rem; }
  a { color:#1668a6; }
  .stDataFrame { border:1px solid var(--line); }
  /* breathing room between the collapsible sections on instrument cards */
  [data-testid="stExpander"] { margin-bottom:.6rem; border-radius:5px; }
  [data-testid="stExpander"] summary { font-size:.86rem; font-weight:600; color:var(--ink); }
  .tl-scroll { overflow-x:auto; border:1px solid var(--line); border-radius:6px; }
  .tl-table { border-collapse:collapse; width:100%; table-layout:fixed; }
  .tl-table th, .tl-table td { border:1px solid var(--line); vertical-align:top; }
  .tl-corner { background:var(--ink); color:#fff; width:130px; min-width:130px; padding:8px; font-size:12px; text-align:left; }
  .tl-yr { background:var(--ink); color:#fff; padding:8px; font-size:13px; font-family:Georgia,serif; min-width:150px; }
  .tl-lane { background:#F3EEE4; color:var(--ink); font-size:12.5px; font-weight:600; padding:8px; text-align:left; width:130px; min-width:130px; }
  .tl-table td { padding:4px; background:#fff; }
  .tl-empty { background:var(--paper) !important; }
  .tl-ev { font-size:11px; line-height:1.32; padding:3px 5px; margin:2px 0; background:#fafafa; border-radius:2px; color:#33403a; }
  .tl-mo { display:inline-block; background:#eee; border-radius:3px; padding:0 4px; margin-right:4px; font-weight:700; color:#555; font-size:10px; }
  .foot { color:#7A6F5B; font-size:.78rem; line-height:1.5; border-top:1px solid var(--line);
          padding-top:.9rem; margin-top:1.6rem; }
</style>
""", unsafe_allow_html=True)

CAT_COLOR = {
    "BCA": "#0E3B43",
    "Domestic carbon pricing": "#8A5A2B",
    "Enabling law": "#5B6B2E",
    "Proposal": "#6E5773",
}

# stage palette, shared by the map and the status pills
STAGE_COLOR = {"In Force": "#1F7A5A", "Draft": "#D98324", "Conceptual": "#3C7DA6"}
STAGE_RANK = {"In Force": 3, "Draft": 2, "Conceptual": 1}
MONTHS = {"01": "Jan", "02": "Feb", "03": "Mar", "04": "Apr", "05": "May", "06": "Jun",
          "07": "Jul", "08": "Aug", "09": "Sep", "10": "Oct", "11": "Nov", "12": "Dec"}

# jurisdiction -> ISO-3 codes for the choropleth. The EU expands to its 27 member states
# so the bloc colours in properly rather than vanishing from the map.
EU27 = ["AUT", "BEL", "BGR", "HRV", "CYP", "CZE", "DNK", "EST", "FIN", "FRA", "DEU",
        "GRC", "HUN", "IRL", "ITA", "LVA", "LTU", "LUX", "MLT", "NLD", "POL", "PRT",
        "ROU", "SVK", "SVN", "ESP", "SWE"]
ISO3 = {
    "European Union": EU27,
    "United Kingdom": ["GBR"],
    "Australia": ["AUS"],
    "Serbia": ["SRB"],
    "Norway": ["NOR"],
    "United States": ["USA"],
    "Thailand": ["THA"],
    "Turkiye": ["TUR"],
    "Türkiye": ["TUR"],
    "Chinese Taipei": ["TWN"],
    "Canada": ["CAN"],
}


def cat_pill(cat):
    c = CAT_COLOR.get(cat, "#555")
    return f'<span class="pill" style="background:{c}22;color:{c};border:1px solid {c}55">{cat}</span>'


def stage_pill(stage):
    c = STAGE_COLOR.get(stage, "#555")
    return f'<span class="pill" style="background:{c}22;color:{c};border:1px solid {c}66">{stage}</span>'


# ---------- column shim -------------------------------------------------------
# The workbook uses human-readable headers; the rest of this app uses short keys.
# Rename once, here, so nothing downstream has to change.
COLMAP = {
    "Entry into Force": "Implementation/Coming into Force Date",
    "Object and Purpose": "object_and_purpose",
    "Third Country Adjustment": "third_country_adjustment",
    "Default Values": "default_values",
    "De Minimis Threshold": "de_minimis",
    "Calculation": "calculation",
    "Revenue Use": "revenue_use",
    "One Line Summary of the Measure": "notes",
    "Scope 1 Coverage": "scope1",
    "Scope 2 Coverage": "scope2",
    "Scope 3 Coverage": "scope3",
    "Other Scope Coverage": "scope_other",
}


# ---------- load + clean ----------
@st.cache_data
def load():
    xl = pd.ExcelFile(DATA_FILE)
    def rd(tab): return pd.read_excel(DATA_FILE, tab, header=1) if tab in xl.sheet_names else pd.DataFrame()
    inst = rd("Instruments"); sec = rd("Sectors"); ev = rd("Events")
    if not inst.empty:
        inst.columns = [str(c).strip() for c in inst.columns]
        inst = inst.rename(columns=COLMAP)
    src = rd("Sources"); catleg = rd("Category_legend")

    valid_cats = set(catleg["category_value"].dropna()) if not catleg.empty else set()

    dropped = []
    if not inst.empty:
        inst = inst[inst["instrument_id"].notna()]
        inst = inst[~inst["instrument_id"].astype(str).str.strip().isin(["", "nan"])]
        inst = inst[inst["instrument_id"] != EXAMPLE_ID]
        # drop note/comment rows: id not code-like OR category not in legend
        def looks_real(r):
            iid = str(r["instrument_id"]).strip()
            if " " in iid or iid.startswith("NOTE"): return False
            if valid_cats and r.get("category") not in valid_cats: return False
            return True
        keep = inst.apply(looks_real, axis=1)
        dropped = inst[~keep]["instrument_id"].astype(str).tolist()
        inst = inst[keep]
        for c in inst.columns: inst[c] = inst[c].astype(str).replace("nan", "")
        # strip midnight timestamps Excel attaches to date cells (2026-01-01 00:00:00 -> 2026-01-01)
        for c in inst.columns:
            inst[c] = inst[c].str.replace(r" 00:00:00$", "", regex=True)
    for d in (sec, ev, src):
        if not d.empty:
            d.dropna(how="all", inplace=True)
            if "instrument_id" in d:
                d.drop(d[d["instrument_id"] == EXAMPLE_ID].index, inplace=True)
    return inst, sec, ev, src, dropped


inst, sec, ev, src, dropped = load()

# instrument_id -> jurisdiction, used wherever a raw code would otherwise be shown
JMAP = dict(zip(inst["instrument_id"], inst["jurisdiction"]))
CMAP = dict(zip(inst["instrument_id"], inst["category"]))
# sheet order = display order (EU, UK, Australia, ...)
JUR_ORDER = list(dict.fromkeys(inst["jurisdiction"]))


def label_for(ids):
    """Readable column/row labels: jurisdiction, disambiguated only when a
    jurisdiction has more than one instrument in the current selection."""
    jur = {i: JMAP.get(i, i) for i in ids}
    counts = Counter(jur.values())
    return {i: (j if counts[j] == 1 else f"{j} ({i})") for i, j in jur.items()}


# ---------- header ----------
logo = next((f for f in LOGO_CANDIDATES if os.path.exists(f)), None)
if logo:
    st.image(logo, width=LOGO_WIDTH)
st.title("BCA Tracker")
st.caption("Tracking border carbon measures worldwide.")

# ---------- sidebar filters ----------
st.sidebar.header("Filter")
cats = sorted([c for c in inst["category"].unique() if c])
stages = [s for s in ["In Force", "Draft", "Conceptual"] if s in set(inst["status_simple"])]
juris = sorted([j for j in inst["jurisdiction"].unique() if j])
f_cat = st.sidebar.multiselect("Category", cats, default=cats)
f_stage = st.sidebar.multiselect("Stage", stages, default=stages)
f_jur = st.sidebar.multiselect("Jurisdiction", juris, default=juris)

view = inst[inst["category"].isin(f_cat) & inst["status_simple"].isin(f_stage)
            & inst["jurisdiction"].isin(f_jur)]

if dropped:
    st.sidebar.markdown("---")
    st.sidebar.caption(f"⚠︎ {len(dropped)} non-data row(s) in the sheet were skipped: " + ", ".join(dropped[:8]))

# ---------- top metrics ----------
c1, c2, c3, c4 = st.columns(4)
c1.metric("Instruments shown", len(view))
c2.metric("Jurisdictions", view["jurisdiction"].nunique())
c3.metric("In force", (view["status_simple"] == "In Force").sum())
c4.metric("BCAs", (view["category"] == "BCA").sum())

tab0, tab1, tab2, tab3, tab4 = st.tabs(
    ["Overview", "Instruments", "Sector coverage", "Timeline", "Official sources"])

# ---- TAB 0: overview map ----
with tab0:
    st.subheader("Overview")
    st.caption("Where border carbon measures stand, by jurisdiction. "
               "Green = in force · Orange = draft or legislated but not yet in force · Blue = conceptual. "
               "Use the sidebar filters to narrow the map.")

    if view.empty:
        st.info("No instruments match the current filters. Widen the selection in the sidebar.")
    else:
        # one row per country code, taking the most advanced stage where a
        # jurisdiction has several instruments (e.g. the two US bills)
        recs = {}
        for _, r in view.iterrows():
            for iso in ISO3.get(r["jurisdiction"], []):
                rank = STAGE_RANK.get(r["status_simple"], 0)
                cur = recs.get(iso)
                if cur is None or rank > cur["rank"]:
                    recs[iso] = {"iso": iso, "Jurisdiction": r["jurisdiction"],
                                 "Stage": r["status_simple"], "rank": rank,
                                 "measures": list(cur["measures"]) if cur else []}
                recs[iso]["measures"].append(r["instrument_name"])
        mapdf = pd.DataFrame([{**v, "Measures": " · ".join(v["measures"])} for v in recs.values()])

        unmapped = sorted({r["jurisdiction"] for _, r in view.iterrows()
                           if not ISO3.get(r["jurisdiction"])})

        try:
            import plotly.express as px
            fig = px.choropleth(
                mapdf, locations="iso", locationmode="ISO-3", color="Stage",
                color_discrete_map=STAGE_COLOR,
                category_orders={"Stage": ["In Force", "Draft", "Conceptual"]},
                hover_name="Jurisdiction",
                hover_data={"iso": False, "Stage": True, "Measures": True},
            )
            fig.update_geos(showframe=False, showcoastlines=False, showcountries=True,
                            countrycolor="#E5DFD3", landcolor="#F1ECE1", lakecolor="#FBF8F2",
                            bgcolor="rgba(0,0,0,0)", projection_type="natural earth")
            fig.update_layout(
                margin=dict(l=0, r=0, t=0, b=0), height=520,
                paper_bgcolor="rgba(0,0,0,0)", geo_bgcolor="rgba(0,0,0,0)",
                legend=dict(orientation="h", yanchor="bottom", y=-0.04,
                            xanchor="left", x=0, title_text=""),
                font=dict(family="Arial", size=12, color="#22303a"),
            )
            st.plotly_chart(fig, use_container_width=True)
        except ModuleNotFoundError:
            st.warning("The map needs plotly. Install it with:  pip install plotly  "
                       "— open the table below in the meantime.")

        if unmapped:
            st.caption("Not shown on the map (no country code mapped): " + ", ".join(unmapped))
        st.caption(f"{len(view)} measure(s) across {view['jurisdiction'].nunique()} jurisdiction(s), current filters.")

        # sectors per instrument, collapsed to one cell
        def sectors_for(iid):
            if sec.empty: return ""
            s = sec[sec["instrument_id"] == iid]["sector"].dropna().astype(str).unique()
            return ", ".join(sorted(s))

        with st.expander("Show as table"):
            ov = view.copy()
            ov["Sectors covered"] = ov["instrument_id"].map(sectors_for)
            ov = ov.rename(columns={
                "jurisdiction": "Jurisdiction",
                "instrument_name": "Measure",
                "category": "Nature of measure",
                "status": "Status",
                "status_simple": "Stage",
                "Implementation/Coming into Force Date": "In force / from",
                "official_url": "Official link",
            })
            cols = [c for c in ["Jurisdiction", "Measure", "Nature of measure", "Status", "Stage",
                                "In force / from", "Sectors covered", "Official link"] if c in ov.columns]
            st.dataframe(
                ov[cols], hide_index=True, use_container_width=True,
                column_config={"Official link": st.column_config.LinkColumn("Official link", display_text="open ↗")},
            )

# ---- TAB 1: instrument cards ----
with tab1:
    if view.empty:
        st.info("No instruments match the current filters. Widen the selection in the sidebar.")
    else:
        if "expand_all" not in st.session_state:
            st.session_state.expand_all = False

        bcol, _sp = st.columns([1, 4])
        if bcol.button("Collapse all" if st.session_state.expand_all else "Expand all",
                       use_container_width=True):
            st.session_state.expand_all = not st.session_state.expand_all
            st.rerun()
        OPEN = st.session_state.expand_all

        main, side = st.columns([3, 1], gap="large")

        # right-side index: click a name to isolate that instrument
        with side:
            st.markdown('<div class="lab">Jump to instrument</div>', unsafe_allow_html=True)
            names = ["(show all)"] + [f'{r["jurisdiction"]} — {r["instrument_name"]}'
                                      for _, r in view.iterrows()]
            picked = st.radio("index", names, label_visibility="collapsed", key="idx")

        show = view
        if picked != "(show all)":
            show = view[view.apply(lambda r: f'{r["jurisdiction"]} — {r["instrument_name"]}' == picked, axis=1)]

        def val_block(v):
            st.markdown(f'<div class="val">{v}</div>', unsafe_allow_html=True)

        with main:
            for _, r in show.iterrows():
                st.markdown('<div class="cardwrap">', unsafe_allow_html=True)
                st.markdown(f'### {r["jurisdiction"]} — {r["instrument_name"]}')
                st.markdown(cat_pill(r["category"]) + "&nbsp;&nbsp;" +
                            stage_pill(r.get("status_simple", "")) +
                            f'&nbsp;&nbsp;<span class="pill" style="background:#eee;color:#333">{r["status"]}</span>',
                            unsafe_allow_html=True)
                # one-line summary sits directly under the title, always visible
                if r.get("notes"):
                    st.markdown('<div class="val" style="font-style:italic;color:#4a4030;'
                                f'margin:.55rem 0 .95rem">{r["notes"]}</div>', unsafe_allow_html=True)

                if r.get("Implementation/Coming into Force Date"):
                    with st.expander("Implementation / in force", expanded=OPEN):
                        val_block(r["Implementation/Coming into Force Date"])

                if r.get("object_and_purpose"):
                    with st.expander("Object & purpose", expanded=OPEN):
                        val_block(r["object_and_purpose"])

                with st.expander("Emissions scope", expanded=OPEN):
                    sc = st.columns(3)
                    for col, (lab, k) in zip(sc, [("Scope 1", "scope1"),
                                                  ("Scope 2", "scope2"),
                                                  ("Scope 3", "scope3")]):
                        col.markdown(f'<div class="lab">{lab}</div>'
                                     f'<div class="val">{r.get(k, "")}</div>', unsafe_allow_html=True)
                    if r.get("scope_other"):
                        st.markdown(f'<div class="lab">Other scope notes</div>'
                                    f'<div class="val">{r["scope_other"]}</div>', unsafe_allow_html=True)

                with st.expander("Adjustment, default values & thresholds", expanded=OPEN):
                    cc = st.columns(3)
                    for col, (lab, k) in zip(cc, [("3rd-country adjustment", "third_country_adjustment"),
                                                  ("Default values", "default_values"),
                                                  ("De minimis threshold", "de_minimis")]):
                        col.markdown(f'<div class="lab">{lab}</div>'
                                     f'<div class="val">{r.get(k, "")}</div>', unsafe_allow_html=True)

                if r.get("calculation"):
                    with st.expander("Calculation", expanded=OPEN):
                        val_block(r["calculation"])

                if r.get("revenue_use"):
                    with st.expander("Revenue use", expanded=OPEN):
                        val_block(r["revenue_use"])

                msec = sec[sec["instrument_id"] == r["instrument_id"]] if not sec.empty else pd.DataFrame()
                if not msec.empty:
                    with st.expander("Sectors covered", expanded=OPEN):
                        val_block(" · ".join(sorted(msec["sector"].dropna().astype(str).unique())))

                if r.get("official_url"):
                    with st.expander("Primary source", expanded=OPEN):
                        st.markdown(f'<div class="val">{r.get("primary_source", "")} — '
                                    f'<a href="{r["official_url"]}" target="_blank">official page ↗</a></div>',
                                    unsafe_allow_html=True)

                st.markdown('</div>', unsafe_allow_html=True)

# ---- TAB 2: sector coverage matrix ----
with tab2:
    st.subheader("Sector coverage")
    keep_ids = set(view["instrument_id"])
    msec = sec[sec["instrument_id"].isin(keep_ids)] if not sec.empty else pd.DataFrame()
    if msec.empty:
        st.info("No sector data for the current filter.")
    else:
        lab = label_for(sorted(set(msec["instrument_id"])))
        m = msec.copy()
        m["Jurisdiction"] = m["instrument_id"].map(lab)
        mat = pd.crosstab(m["sector"], m["Jurisdiction"])
        # order columns by the sheet's jurisdiction order, not alphabetically
        order = [c for j in JUR_ORDER for c in mat.columns if c == j or c.startswith(j + " (")]
        order += [c for c in mat.columns if c not in order]
        mat = mat[order]
        mat = (mat > 0).replace({True: "●", False: ""})
        st.dataframe(mat, use_container_width=True)

        with st.expander("Show HS codes behind each sector"):
            rank = {j: i for i, j in enumerate(order)}
            hs = m[["Jurisdiction", "sector", "hs_code", "scope_note"]].copy()
            hs["_o"] = hs["Jurisdiction"].map(rank).fillna(999)
            hs = (hs.sort_values(["_o", "sector"])
                    .drop(columns="_o")
                    .rename(columns={"sector": "Sector", "hs_code": "HS code",
                                     "scope_note": "Scope note"}))
            st.dataframe(
                hs, hide_index=True, use_container_width=True, height=560,
                column_config={
                    "Jurisdiction": st.column_config.TextColumn("Jurisdiction", width="small"),
                    "Sector": st.column_config.TextColumn("Sector", width="small"),
                    "HS code": st.column_config.TextColumn("HS code", width="small"),
                    "Scope note": st.column_config.TextColumn("Scope note", width="large"),
                },
            )

# ---- TAB 3: timeline ----
with tab3:
    st.subheader("Timeline of key events")
    keep_ids = set(view["instrument_id"])
    mev = ev[ev["instrument_id"].isin(keep_ids)].copy() if not ev.empty else pd.DataFrame()
    if mev.empty:
        st.info("No events match the current filter.")
    else:
        mev["date"] = mev["date"].astype(str)
        mev["year"] = mev["date"].str[:4]
        mev["mon"] = mev["date"].str[5:7].map(MONTHS).fillna("")
        mev["jurisdiction"] = mev["instrument_id"].map(JMAP)
        mev["category"] = mev["instrument_id"].map(CMAP)
        mev = mev[mev["jurisdiction"].notna() & (mev["year"].str.len() == 4)]

        all_years = sorted(mev["year"].unique())
        cL, cR = st.columns([3, 2])
        layout = cL.radio("Layout", ["Swimlane grid", "Chronological list"],
                          horizontal=True, label_visibility="collapsed")
        yr_sel = cR.select_slider("Year range", options=all_years,
                                  value=(all_years[0], all_years[-1])) if len(all_years) > 1 else (all_years[0], all_years[-1])
        y0, y1 = yr_sel
        tev = mev[(mev["year"] >= y0) & (mev["year"] <= y1)]
        years = sorted(tev["year"].unique())

        if tev.empty:
            st.info("No events in the selected year range.")
        elif layout == "Swimlane grid":
            st.caption("Years across the top · one lane per jurisdiction · each event in its year cell. "
                       "Read a lane left-to-right for one jurisdiction's arc; read a column top-to-bottom for one year. "
                       "Colour = category. Scrolls sideways if the range is wide.")
            # jurisdictions ordered by earliest event
            order_j = tev.groupby("jurisdiction")["date"].min().sort_values().index.tolist()

            def cellhtml(j, y):
                e = tev[(tev["jurisdiction"] == j) & (tev["year"] == y)]
                if e.empty:
                    return '<td class="tl-empty"></td>'
                items = []
                for _, r in e.sort_values("date").iterrows():
                    col = CAT_COLOR.get(r["category"], "#555")
                    ev_txt = str(r["event"])
                    link = f' <a href="{r["official_url"]}" target="_blank">↗</a>' if r.get("official_url") else ""
                    items.append(f'<div class="tl-ev" style="border-left:3px solid {col}">'
                                 f'<span class="tl-mo">{r["mon"]}</span>{ev_txt}{link}</div>')
                return f'<td>{"".join(items)}</td>'

            head = "<th class='tl-corner'>Jurisdiction</th>" + "".join(f"<th class='tl-yr'>{y}</th>" for y in years)
            rows_html = ""
            for j in order_j:
                rows_html += "<tr>" + f'<th class="tl-lane">{j}</th>' + "".join(cellhtml(j, y) for y in years) + "</tr>"
            st.markdown(f'<div class="tl-scroll"><table class="tl-table"><tr>{head}</tr>{rows_html}</table></div>',
                        unsafe_allow_html=True)
        else:
            st.caption("Every event in date order.")
            for _, r in tev.sort_values("date").iterrows():
                col = CAT_COLOR.get(r["category"], "#0E3B43")
                when = f'{r["mon"]} {r["year"]}'.strip()
                link = f' — <a href="{r["official_url"]}" target="_blank">source ↗</a>' if r.get("official_url") else ""
                st.markdown(f'<div style="border-left:3px solid {col};padding:.15rem 0 .55rem .8rem;margin-left:.3rem">'
                            f'<b>{when}</b> &nbsp;·&nbsp; <span style="color:#7A6F5B">{r["jurisdiction"]}</span><br>'
                            f'{r["event"]}{link}</div>', unsafe_allow_html=True)

# ---- TAB 4: sources ----
with tab4:
    st.subheader("Official pages & documents")
    keep_ids = set(view["instrument_id"])
    msrc = src[src["instrument_id"].isin(keep_ids)].copy() if not src.empty else pd.DataFrame()
    if msrc.empty:
        st.info("No sources match the current filter.")
    else:
        msrc["Jurisdiction"] = msrc["instrument_id"].map(JMAP)
        msrc["link"] = msrc["official_url"]
        msrc["_o"] = msrc["Jurisdiction"].map({j: i for i, j in enumerate(JUR_ORDER)}).fillna(999)
        show = (msrc.sort_values(["_o", "date"])
                    [["Jurisdiction", "doc_title", "doc_type", "date", "link"]]
                    .rename(columns={"doc_title": "Document", "doc_type": "Type", "date": "Year"}))
        st.dataframe(show, hide_index=True, use_container_width=True,
                     column_config={"link": st.column_config.LinkColumn("Official URL", display_text="open ↗")})

# ---------- footer ----------
st.markdown(
    '<div class="foot"><b>Disclaimer.</b> This is an independent tracker, compiled on a best-effort basis '
    'from official and other credible public sources. It is not affiliated with, or endorsed by, any government '
    'or organisation listed, and nothing here is legal advice. Entries may be incomplete or superseded — '
    'always check the linked official source before relying on it.</div>',
    unsafe_allow_html=True)

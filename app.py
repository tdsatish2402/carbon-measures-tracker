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
  /* Palette taken from the Trade Law Observatory wordmark: navy --ink,
     purple --accent, gold --gold. --paper is a neutral near-white (the old
     cream read as pink on screen). Change --ink here to restyle throughout. */
  :root { --ink:#1D2657; --accent:#3F106E; --gold:#C9A227;
          --paper:#F7F8FA; --card:#FFFFFF; --line:#E3E6EC; --muted:#5B6478; }
  html, body, [class*="css"] { font-size:17px; }
  .stApp { background:var(--paper); }
  .stApp p, .stApp li { font-size:1rem; line-height:1.6; }
  [data-testid="stCaptionContainer"] p { font-size:.92rem !important; color:var(--muted); }
  [data-testid="stHeader"] { background:transparent; }
  h1,h2,h3 { color:var(--ink); font-family:Georgia,'Times New Roman',serif; letter-spacing:-.2px; }
  .cardwrap { border:1px solid var(--line); border-radius:8px; padding:1.1rem 1.25rem;
              background:var(--card); margin-bottom:1.15rem;
              box-shadow:0 1px 2px rgba(29,38,87,.05); }
  .pill { display:inline-block; padding:.18rem .62rem; border-radius:999px;
          font-size:.78rem; font-weight:700; letter-spacing:.02em; }
  .lab { color:var(--muted); font-size:.76rem; text-transform:uppercase; letter-spacing:.08em;
         font-weight:700; margin-top:.5rem; }
  .val { color:#243049; font-size:.97rem; line-height:1.6; margin-bottom:.2rem; }
  a { color:var(--accent); }
  .stDataFrame { border:1px solid var(--line); }
  /* shared HTML tables — st.dataframe draws to canvas, so its text can't be
     resized with CSS; these render as real DOM and stay legible. */
  .dt-wrap { overflow:auto; border:1px solid var(--line); border-radius:6px;
             background:var(--card); margin:.3rem 0 .6rem; }
  .dt { border-collapse:collapse; width:100%; font-size:14.5px; }
  .dt th { background:var(--ink); color:#fff; font-weight:600; text-align:left;
           padding:10px 12px; position:sticky; top:0; z-index:2;
           font-size:13.5px; line-height:1.35; }
  .dt td { border-top:1px solid var(--line); padding:9px 12px; vertical-align:top;
           color:#243049; line-height:1.5; }
  .dt tbody tr:nth-child(even) td { background:#FAFBFD; }
  .dt td.sec { font-weight:600; color:var(--ink); white-space:nowrap;
               position:sticky; left:0; background:var(--card); z-index:1;
               border-right:1px solid var(--line); }
  .dt tbody tr:nth-child(even) td.sec { background:#FAFBFD; }
  .dt td.tick { text-align:center; color:#1B7A5A; font-size:17px; font-weight:700; }
  .dt th.jur { min-width:120px; white-space:normal; }
  [data-testid="stMetricValue"] { color:var(--ink); font-family:Georgia,serif; }
  [data-testid="stMetricLabel"] { color:var(--muted); }
  .stTabs [aria-selected="true"] { color:var(--accent) !important; }
  .logo-rule { border:0; border-top:2px solid var(--gold); width:230px;
               margin:.55rem 0 1.1rem; opacity:.85; }
  /* breathing room between the collapsible sections on instrument cards */
  [data-testid="stExpander"] { margin-bottom:.6rem; border-radius:6px;
                               border:1px solid var(--line); background:var(--card); }
  [data-testid="stExpander"] summary { font-size:.95rem; font-weight:600; color:var(--ink); }
  .tl-scroll { overflow-x:auto; border:1px solid var(--line); border-radius:6px; }
  .tl-table { border-collapse:collapse; width:100%; table-layout:fixed; }
  .tl-table th, .tl-table td { border:1px solid var(--line); vertical-align:top; }
  .tl-corner { background:var(--ink); color:#fff; width:150px; min-width:150px; padding:10px; font-size:13.5px; text-align:left; }
  .tl-yr { background:var(--ink); color:#fff; padding:10px; font-size:15.5px; font-family:Georgia,serif; min-width:190px; }
  .tl-lane { background:#EEF0F6; color:var(--ink); font-size:14px; font-weight:600; padding:10px; text-align:left; width:150px; min-width:150px; }
  .tl-table td { padding:4px; background:var(--card); }
  .tl-empty { background:var(--paper) !important; }
  .tl-ev { font-size:13.5px; line-height:1.45; padding:6px 8px; margin:4px 0; background:#F7F8FA; border-radius:3px; color:#243049; }
  .tl-mo { display:inline-block; background:#E3E6EC; border-radius:3px; padding:1px 6px; margin-right:5px; font-weight:700; color:var(--ink); font-size:12px; }
  .foot { color:var(--muted); font-size:.85rem; line-height:1.5; border-top:1px solid var(--line);
          padding-top:.9rem; margin-top:1.6rem; }
</style>
""", unsafe_allow_html=True)

# category colours, drawn from the wordmark's navy/purple/gold
CAT_COLOR = {
    "BCA": "#1D2657",
    "Domestic carbon pricing": "#8A6D1F",
    "Enabling law": "#3F106E",
    "Proposal": "#6B7280",
}

# stage palette, shared by the map and the status pills
STAGE_COLOR = {"In Force": "#1B7A5A", "Draft": "#C77A16", "Conceptual": "#3B5FA8"}
# self-explaining legend labels, so the map needs no separate colour key
STAGE_LABEL = {"In Force": "In force",
               "Draft": "Draft / legislated, not yet in force",
               "Conceptual": "Conceptual"}
LEGEND_COLOR = {STAGE_LABEL[k]: v for k, v in STAGE_COLOR.items()}
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


def html_table(headers, rows, max_height=None, first_col_sticky=False):
    """Render a real DOM table. st.dataframe paints to canvas, so its text can't
    be resized by CSS; these stay readable and let the scope-note column wrap."""
    import html as _h
    th = "".join(f'<th class="jur">{_h.escape(str(x))}</th>' for x in headers)
    body = []
    for r in rows:
        tds = []
        for i, cell in enumerate(r):
            if isinstance(cell, tuple):          # (css_class, raw_html)
                tds.append(f'<td class="{cell[0]}">{cell[1]}</td>')
            else:
                cls = ' class="sec"' if (first_col_sticky and i == 0) else ""
                tds.append(f"<td{cls}>{_h.escape(str(cell))}</td>")
        body.append("<tr>" + "".join(tds) + "</tr>")
    style = f' style="max-height:{max_height}px"' if max_height else ""
    st.markdown(f'<div class="dt-wrap"{style}><table class="dt"><thead><tr>{th}</tr></thead>'
                f'<tbody>{"".join(body)}</tbody></table></div>', unsafe_allow_html=True)


def cat_pill(cat):
    c = CAT_COLOR.get(cat, "#555")
    return f'<span class="pill" style="background:{c}22;color:{c};border:1px solid {c}55">{cat}</span>'


def stage_pill(text, stage=None):
    """Render `text` tinted by its stage. Pass stage separately to show the full
    status wording (e.g. "In force / operational") in the stage's colour."""
    c = STAGE_COLOR.get(stage if stage is not None else text, "#5B6478")
    return f'<span class="pill" style="background:{c}1F;color:{c};border:1px solid {c}66">{text}</span>'


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


NMAP = dict(zip(inst["instrument_id"], inst["instrument_name"]))
# sheet row order, so columns follow the register rather than the alphabet
ID_ORDER = {iid: n for n, iid in enumerate(inst["instrument_id"])}


def short_name(iid):
    """'Foreign Pollution Fee Act of 2025 (S. 1325, ...)' -> 'Foreign Pollution Fee Act'"""
    n = str(NMAP.get(iid, iid)).split(" of 20")[0].split(" (")[0].strip()
    return n if len(n) <= 34 else n[:31].rstrip() + "..."


def label_for(ids):
    """Readable column/row labels: jurisdiction alone, or jurisdiction plus a
    short instrument name when one jurisdiction has several instruments.
    Never shows a raw instrument_id."""
    jur = {i: JMAP.get(i, i) for i in ids}
    counts = Counter(jur.values())
    return {i: (j if counts[j] == 1 else f"{j} — {short_name(i)}") for i, j in jur.items()}


# ---------- header ----------
@st.cache_data
def prep_logo(path):
    """Knock the white background out of the wordmark and trim the margin, so it
    sits on the page rather than in a white box. Falls back to the raw file if
    Pillow/numpy are unavailable."""
    try:
        import numpy as np
        from PIL import Image
        a = np.array(Image.open(path).convert("RGBA"))
        lum = a[..., :3].astype(int).max(axis=2)
        alpha = np.clip((245 - lum) * (255 / 45.0), 0, 255)
        a[..., 3] = np.maximum(alpha, np.where(lum < 200, 255, 0)).astype(np.uint8)
        im = Image.fromarray(a, "RGBA")
        return im.crop(im.getbbox())
    except Exception:
        return path


logo = next((f for f in LOGO_CANDIDATES if os.path.exists(f)), None)
if logo:
    st.image(prep_logo(logo), width=LOGO_WIDTH)
    st.markdown('<hr class="logo-rule">', unsafe_allow_html=True)
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
c1, c2, c3 = st.columns(3)
c1.metric("Instruments", len(view))
c2.metric("Jurisdictions", view["jurisdiction"].nunique())
c3.metric("In force", (view["status_simple"] == "In Force").sum())

tab0, tab1, tab2, tab3, tab4 = st.tabs(
    ["Overview", "Instruments", "Sector coverage", "Timeline", "Official sources"])

# ---- TAB 0: overview map ----
with tab0:
    st.subheader("Overview")
    st.caption("Where border carbon measures stand, by jurisdiction. "
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
                                 "Stage": STAGE_LABEL.get(r["status_simple"], r["status_simple"]),
                                 "rank": rank,
                                 "measures": list(cur["measures"]) if cur else []}
                recs[iso]["measures"].append(r["instrument_name"])
        mapdf = pd.DataFrame([{**v, "Measures": " · ".join(v["measures"])} for v in recs.values()])

        unmapped = sorted({r["jurisdiction"] for _, r in view.iterrows()
                           if not ISO3.get(r["jurisdiction"])})

        try:
            import plotly.express as px
            fig = px.choropleth(
                mapdf, locations="iso", locationmode="ISO-3", color="Stage",
                color_discrete_map=LEGEND_COLOR,
                category_orders={"Stage": [STAGE_LABEL["In Force"], STAGE_LABEL["Draft"],
                                           STAGE_LABEL["Conceptual"]]},
                hover_name="Jurisdiction",
                hover_data={"iso": False, "Stage": True, "Measures": True},
            )
            fig.update_geos(showframe=False, showcoastlines=False, showcountries=True,
                            countrycolor="#D9DEE8", landcolor="#EBEEF4", lakecolor="#F7F8FA",
                            bgcolor="rgba(0,0,0,0)", projection_type="natural earth")
            fig.update_layout(
                margin=dict(l=0, r=0, t=0, b=0), height=520,
                paper_bgcolor="rgba(0,0,0,0)", geo_bgcolor="rgba(0,0,0,0)",
                legend=dict(orientation="h", yanchor="bottom", y=-0.06,
                            xanchor="left", x=0, title_text="",
                            font=dict(size=14)),
                font=dict(family="Arial", size=14, color="#243049"),
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
            cols = [c for c in ["Jurisdiction", "Measure", "Status",
                                "In force / from", "Sectors covered"] if c in ov.columns]
            rows = []
            for _, x in ov.iterrows():
                rows.append([x.get(c, "") for c in cols] +
                            [("", f'<a href="{x["Official link"]}" target="_blank">open ↗</a>'
                                  if x.get("Official link") else "")])
            html_table(cols + ["Official link"], rows, max_height=520)

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
                # one status pill only, tinted by its stage — the separate
                # "In Force" pill duplicated what "In force / operational" says
                st.markdown(cat_pill(r["category"]) + "&nbsp;&nbsp;" +
                            stage_pill(r["status"], r.get("status_simple", "")),
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
        # order columns by the sheet's instrument order, not alphabetically
        seen, order = set(), []
        for iid in sorted(lab, key=lambda i: ID_ORDER.get(i, 999)):
            c = lab[iid]
            if c in mat.columns and c not in seen:
                seen.add(c); order.append(c)
        order += [c for c in mat.columns if c not in seen]
        mat = mat[order]
        mat = mat > 0
        rows = [[sector] + [("tick", "✓" if mat.loc[sector, c] else "") for c in order]
                for sector in mat.index]
        html_table(["Sector"] + order, rows, max_height=560, first_col_sticky=True)

        with st.expander("Show HS codes behind each sector"):
            rank = {j: i for i, j in enumerate(order)}
            hs = m[["Jurisdiction", "sector", "hs_code", "scope_note"]].copy()
            hs["_o"] = hs["Jurisdiction"].map(rank).fillna(999)
            hs = (hs.sort_values(["_o", "sector"])
                    .drop(columns="_o")
                    .rename(columns={"sector": "Sector", "hs_code": "HS code",
                                     "scope_note": "Scope note"}))
            html_table(list(hs.columns), hs.values.tolist(), max_height=600)

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
                col = CAT_COLOR.get(r["category"], "#1D2657")
                when = f'{r["mon"]} {r["year"]}'.strip()
                link = f' — <a href="{r["official_url"]}" target="_blank">source ↗</a>' if r.get("official_url") else ""
                st.markdown(f'<div style="border-left:3px solid {col};padding:.25rem 0 .7rem .9rem;'
                            f'margin-left:.3rem;font-size:1rem;line-height:1.55">'
                            f'<b>{when}</b> &nbsp;·&nbsp; <span style="color:var(--muted)">{r["jurisdiction"]}</span><br>'
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
        show = msrc.sort_values(["_o", "date"])
        rows = [[x["Jurisdiction"], x["doc_title"], x["doc_type"], x["date"],
                 ("", f'<a href="{x["link"]}" target="_blank">open ↗</a>' if x.get("link") else "")]
                for _, x in show.iterrows()]
        html_table(["Jurisdiction", "Document", "Type", "Year", "Official URL"],
                   rows, max_height=620)

# ---------- footer ----------
st.markdown(
    '<div class="foot"><b>Disclaimer.</b> This is an independent tracker, compiled on a best-effort basis '
    'from official and other credible public sources. It is not affiliated with, or endorsed by, any government '
    'or organisation listed, and nothing here is legal advice. Entries may be incomplete or superseded — '
    'always check the linked official source before relying on it.</div>',
    unsafe_allow_html=True)

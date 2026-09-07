# Carbon Measures Tracker — Streamlit app
# Reads BCA_tracker.xlsx (multi-tab, normalised) and renders filterable views.
# Designed to fail gracefully on half-edited data: junk rows are dropped and reported, never shown as real records.

import pandas as pd
import streamlit as st

DATA_FILE = "BCA_tracker_only_BCA.xlsx"
EXAMPLE_ID = "XX-EXAMPLE"

st.set_page_config(page_title="Carbon Measures Tracker", layout="wide",
                   initial_sidebar_state="expanded")

# ---------- styling (restrained; Streamlit limits raw HTML) ----------
st.markdown("""
<style>
  :root { --ink:#0E3B43; --paper:#FBF8F2; --line:#E5DFD3; }
  .stApp { background:var(--paper); }
  h1,h2,h3 { color:var(--ink); font-family:Georgia,'Times New Roman',serif; letter-spacing:-.2px; }
  .eyebrow { text-transform:uppercase; letter-spacing:.18em; font-size:.72rem;
             color:#7A6F5B; font-weight:700; margin-bottom:.1rem; }
  .cardwrap { border:1px solid var(--line); border-radius:6px; padding:1rem 1.15rem;
              background:#fff; margin-bottom:.7rem; }
  .pill { display:inline-block; padding:.12rem .55rem; border-radius:999px;
          font-size:.72rem; font-weight:700; letter-spacing:.02em; }
  .lab { color:#7A6F5B; font-size:.72rem; text-transform:uppercase; letter-spacing:.08em;
         font-weight:700; margin-top:.5rem; }
  .val { color:#22303a; font-size:.92rem; margin-bottom:.2rem; }
  a { color:#1668a6; }
  .stDataFrame { border:1px solid var(--line); }
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
</style>
""", unsafe_allow_html=True)

CAT_COLOR = {
    "BCA": "#0E3B43",
    "Domestic carbon pricing": "#8A5A2B",
    "Enabling law": "#5B6B2E",
    "Proposal": "#6E5773",
}
def cat_pill(cat):
    c = CAT_COLOR.get(cat, "#555")
    return f'<span class="pill" style="background:{c}22;color:{c};border:1px solid {c}55">{cat}</span>'

# ---------- load + clean ----------
# --- column shim -------------------------------------------------------------
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

@st.cache_data
def load():
    xl = pd.ExcelFile(DATA_FILE)
    def rd(tab): return pd.read_excel(DATA_FILE, tab, header=1) if tab in xl.sheet_names else pd.DataFrame()
    inst = rd("Instruments"); sec = rd("Sectors"); ev = rd("Events")
    if not inst.empty:
        inst.columns = [str(c).strip() for c in inst.columns]
        inst = inst.rename(columns=COLMAP)
    src = rd("Sources"); catleg = rd("Category_legend")
    scopeleg = rd("Scope_legend"); watch = rd("Watchlist"); multi = rd("Multilateral")
    statleg = rd("Status_legend"); secleg = rd("Sector_legend"); meth = rd("Methodology")

    valid_cats = set(catleg["category_value"].dropna()) if not catleg.empty else set()
    valid_stats = set(statleg["status_value"].dropna()) if not statleg.empty else set()

    dropped = []
    if not inst.empty:
        n0 = len(inst)
        inst = inst[inst["instrument_id"].notna()]
        inst = inst[~inst["instrument_id"].astype(str).str.strip().isin(["", "nan"])]
        inst = inst[inst["instrument_id"] != EXAMPLE_ID]
        # drop note/comment rows: id not uppercase-code-like OR category not in legend
        def looks_real(r):
            iid = str(r["instrument_id"]).strip()
            if " " in iid or iid.startswith("NOTE"): return False
            if valid_cats and r.get("category") not in valid_cats: return False
            return True
        keep = inst.apply(looks_real, axis=1)
        dropped = inst[~keep]["instrument_id"].astype(str).tolist()
        inst = inst[keep]
        for c in inst.columns: inst[c] = inst[c].astype(str).replace("nan", "")
        # strip midnight timestamps Excel attaches to date cells (e.g. 2026-01-01 00:00:00 -> 2026-01-01)
        for c in inst.columns:
            inst[c] = inst[c].str.replace(r" 00:00:00$", "", regex=True)
    for d in (sec, ev, src):
        if not d.empty:
            d.dropna(how="all", inplace=True)
            if "instrument_id" in d:
                d.drop(d[d["instrument_id"] == EXAMPLE_ID].index, inplace=True)
    return inst, sec, ev, src, catleg, statleg, secleg, meth, dropped, scopeleg, watch, multi

inst, sec, ev, src, catleg, statleg, secleg, meth, dropped, scopeleg, watch, multi = load()

# ---------- header ----------
st.markdown('<div class="eyebrow">Trade &amp; Climate · Policy Registry</div>', unsafe_allow_html=True)
st.title("Carbon Measures Tracker")
st.caption("Border carbon adjustments and domestic carbon measures by jurisdiction — status, coverage, timeline and official sources. Independent; not affiliated with any government or organisation listed.")

# ---------- sidebar filters ----------
st.sidebar.header("Filter")
cats = sorted([c for c in inst["category"].unique() if c])
stats = sorted([s for s in inst["status"].unique() if s])
juris = sorted([j for j in inst["jurisdiction"].unique() if j])
f_cat = st.sidebar.multiselect("Category", cats, default=cats)
f_stat = st.sidebar.multiselect("Status", stats, default=stats)
f_jur = st.sidebar.multiselect("Jurisdiction", juris, default=juris)

view = inst[inst["category"].isin(f_cat) & inst["status"].isin(f_stat) & inst["jurisdiction"].isin(f_jur)]

if dropped:
    st.sidebar.markdown("---")
    st.sidebar.caption(f"⚠︎ {len(dropped)} non-data row(s) in the sheet were skipped: " + ", ".join(dropped[:8]))

# ---------- top metrics ----------
c1, c2, c3, c4 = st.columns(4)
c1.metric("Instruments shown", len(view))
c2.metric("Jurisdictions", view["jurisdiction"].nunique())
c3.metric("In force", (view["status"] == "In force / operational").sum())
c4.metric("BCAs", (view["category"] == "BCA").sum())

tab0, tab1, tab2, tab3, tab4, tab6, tab7, tab5 = st.tabs(
    ["Overview", "Instruments", "Sector coverage", "Timeline", "Official sources",
     "Watchlist", "Multilateral", "Methodology"])

# ---- TAB 0: overview register (clean flat table) ----
with tab0:
    st.subheader("Overview")
    st.caption("At-a-glance register of jurisdictions with carbon measures. Use the sidebar filters to narrow the list; open the Instruments tab for full detail on any measure.")
    if view.empty:
        st.info("No instruments match the current filters. Widen the selection in the sidebar.")
    else:
        # sectors per instrument, collapsed to one cell
        def sectors_for(iid):
            if sec.empty: return ""
            s = sec[sec["instrument_id"] == iid]["sector"].dropna().astype(str).unique()
            return ", ".join(sorted(s))
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
        cols = ["Jurisdiction", "Measure", "Nature of measure", "Status", "Stage",
                "In force / from", "Sectors covered", "Official link"]
        cols = [c for c in cols if c in ov.columns]
        st.dataframe(
            ov[cols].sort_values(["Nature of measure", "Jurisdiction"]),
            hide_index=True, use_container_width=True,
            column_config={"Official link": st.column_config.LinkColumn("Official link", display_text="open ↗")},
        )
        st.caption(f"{len(ov)} measure(s) across {ov['Jurisdiction'].nunique()} jurisdiction(s), current filters.")

# ---- TAB 1: instrument cards ----
with tab1:
    if view.empty:
        st.info("No instruments match the current filters. Widen the selection in the sidebar.")
    else:
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

        with main:
            for _, r in show.iterrows():
                st.markdown('<div class="cardwrap">', unsafe_allow_html=True)
                st.markdown(f'### {r["jurisdiction"]} — {r["instrument_name"]}')
                st.markdown(cat_pill(r["category"]) +
                            f'&nbsp;&nbsp;<span class="pill" style="background:#eee;color:#333">{r["status"]}</span>',
                            unsafe_allow_html=True)
                # (2) one-line summary now sits directly under the title, before implementation
                if r.get("notes"):
                    st.markdown(f'<div class="val" style="font-style:italic;color:#4a4030;margin-top:.4rem">{r["notes"]}</div>', unsafe_allow_html=True)
                impl = r.get("Implementation/Coming into Force Date", "")
                if impl: st.markdown(f'<div class="lab">Implementation / in force</div><div class="val">{impl}</div>', unsafe_allow_html=True)
                if r.get("object_and_purpose"):
                    st.markdown(f'<div class="lab">Object &amp; purpose</div><div class="val">{r["object_and_purpose"]}</div>', unsafe_allow_html=True)
                sc = st.columns(3)
                for col, (lab, key) in zip(sc, [("Scope 1","scope1"), ("Scope 2","scope2"), ("Scope 3","scope3")]):
                    col.markdown(f'<div class="lab">{lab}</div><div class="val">{r.get(key,"")}</div>', unsafe_allow_html=True)
                if r.get("scope_other"):
                    st.markdown(f'<div class="lab">Other scope notes</div><div class="val">{r["scope_other"]}</div>', unsafe_allow_html=True)
                cc = st.columns(3)
                for col, (lab, key) in zip(cc, [("3rd-country adjustment","third_country_adjustment"),
                                                ("Default values","default_values"),
                                                ("De minimis threshold","de_minimis")]):
                    col.markdown(f'<div class="lab">{lab}</div><div class="val">{r.get(key,"")}</div>', unsafe_allow_html=True)
                if r.get("calculation"):
                    st.markdown(f'<div class="lab">Calculation</div><div class="val">{r["calculation"]}</div>', unsafe_allow_html=True)
                if r.get("revenue_use"):
                    st.markdown(f'<div class="lab">Revenue use</div><div class="val">{r["revenue_use"]}</div>', unsafe_allow_html=True)
                msec = sec[sec["instrument_id"] == r["instrument_id"]] if not sec.empty else pd.DataFrame()
                if not msec.empty:
                    chips = " · ".join(sorted(msec["sector"].dropna().astype(str).unique()))
                    st.markdown(f'<div class="lab">Sectors</div><div class="val">{chips}</div>', unsafe_allow_html=True)
                srcline = r.get("primary_source",""); url = r.get("official_url","")
                if url: st.markdown(f'<div class="lab">Primary source</div><div class="val">{srcline} — <a href="{url}" target="_blank">official page ↗</a></div>', unsafe_allow_html=True)
                st.markdown('</div>', unsafe_allow_html=True)

# ---- TAB 2: sector coverage matrix ----
with tab2:
    st.subheader("Sector coverage")
    st.caption("Which instruments cover which sectors. Generated from the Sectors tab — rows are sectors, columns are instruments. ● = covered.")
    keep_ids = set(view["instrument_id"])
    msec = sec[sec["instrument_id"].isin(keep_ids)] if not sec.empty else pd.DataFrame()
    if msec.empty:
        st.info("No sector data for the current filter (domestic-pricing instruments have no product list).")
    else:
        mat = pd.crosstab(msec["sector"], msec["instrument_id"])
        mat = (mat > 0).replace({True: "●", False: ""})
        st.dataframe(mat, use_container_width=True)
        with st.expander("Show HS codes behind each sector"):
            st.dataframe(msec[["instrument_id","sector","hs_code","scope_note"]].sort_values(["sector","instrument_id"]),
                         hide_index=True, use_container_width=True)

# ---- TAB 3: timeline ----
with tab3:
    st.subheader("Timeline of key events")
    keep_ids = set(view["instrument_id"])
    mev = ev[ev["instrument_id"].isin(keep_ids)].copy() if not ev.empty else pd.DataFrame()
    if mev.empty:
        st.info("No events match the current filter.")
    else:
        jmap = dict(zip(inst["instrument_id"], inst["jurisdiction"]))
        cmap = dict(zip(inst["instrument_id"], inst["category"]))
        mev["date"] = mev["date"].astype(str)
        mev["year"] = mev["date"].str[:4]
        mev["jurisdiction"] = mev["instrument_id"].map(jmap)
        mev["category"] = mev["instrument_id"].map(cmap)
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
            order = tev.groupby("jurisdiction")["date"].min().sort_values().index.tolist()
            def cellhtml(j, y):
                e = tev[(tev["jurisdiction"] == j) & (tev["year"] == y)]
                if e.empty:
                    return '<td class="tl-empty"></td>'
                items = []
                for _, r in e.sort_values("date").iterrows():
                    col = CAT_COLOR.get(r["category"], "#555")
                    mo = r["date"][5:7] if len(r["date"]) >= 7 else ""
                    ev_txt = str(r["event"])
                    link = f' <a href="{r["official_url"]}" target="_blank">↗</a>' if r.get("official_url") else ""
                    items.append(f'<div class="tl-ev" style="border-left:3px solid {col}">'
                                 f'<span class="tl-mo">{mo}</span>{ev_txt}{link}</div>')
                return f'<td>{"".join(items)}</td>'
            head = "<th class='tl-corner'>Jurisdiction</th>" + "".join(f"<th class='tl-yr'>{y}</th>" for y in years)
            rows_html = ""
            for j in order:
                rows_html += "<tr>" + f'<th class="tl-lane">{j}</th>' + "".join(cellhtml(j, y) for y in years) + "</tr>"
            st.markdown(f'<div class="tl-scroll"><table class="tl-table"><tr>{head}</tr>{rows_html}</table></div>',
                        unsafe_allow_html=True)
        else:
            st.caption("Every event in date order.")
            for _, r in tev.sort_values("date").iterrows():
                col = CAT_COLOR.get(r["category"], "#0E3B43")
                link = f' — <a href="{r["official_url"]}" target="_blank">source ↗</a>' if r.get("official_url") else ""
                st.markdown(f'<div style="border-left:3px solid {col};padding:.15rem 0 .55rem .8rem;margin-left:.3rem">'
                            f'<b>{r["date"]}</b> &nbsp;·&nbsp; <span style="color:#7A6F5B">{r["jurisdiction"]}</span><br>'
                            f'{r["event"]}{link}</div>', unsafe_allow_html=True)

# ---- TAB 4: sources ----
with tab4:
    st.subheader("Official pages & documents")
    keep_ids = set(view["instrument_id"])
    msrc = src[src["instrument_id"].isin(keep_ids)].copy() if not src.empty else pd.DataFrame()
    if msrc.empty:
        st.info("No sources match the current filter.")
    else:
        jmap = dict(zip(inst["instrument_id"], inst["jurisdiction"]))
        msrc["jurisdiction"] = msrc["instrument_id"].map(jmap)
        msrc["link"] = msrc["official_url"]
        show = msrc[["jurisdiction","doc_title","doc_type","date","link"]].sort_values(["jurisdiction","date"])
        st.dataframe(show, hide_index=True, use_container_width=True,
                     column_config={"link": st.column_config.LinkColumn("Official URL")})

# ---- TAB 6: watchlist (jurisdictions checked, no instrument yet) ----
with tab6:
    st.subheader("Watchlist")
    st.caption("Jurisdictions checked that do not (yet) have a border carbon adjustment. "
               "Recorded so the negative finding is auditable. Not affected by the sidebar filters.")
    if watch.empty:
        st.info("Watchlist tab not found in the workbook.")
    else:
        st.dataframe(watch, hide_index=True, use_container_width=True,
                     column_config={"source_url": st.column_config.LinkColumn("Source", display_text="open ↗")})

# ---- TAB 7: multilateral landscape ----
with tab7:
    st.subheader("Multilateral landscape")
    st.caption("WTO discussions and disputes, standards work, and cooperation forums. Not instruments, "
               "so they sit outside the register. Not affected by the sidebar filters.")
    if multi.empty:
        st.info("Multilateral tab not found in the workbook.")
    else:
        st.dataframe(multi, hide_index=True, use_container_width=True,
                     column_config={"source_url": st.column_config.LinkColumn("Source", display_text="open ↗")})

# ---- TAB 5: methodology ----
with tab5:
    st.subheader("Methodology & attribution")
    if meth.empty:
        st.info("Methodology tab not found in the workbook.")
    else:
        for _, r in meth.iterrows():
            st.markdown(f'<div class="lab">{r["item"]}</div><div class="val">{r["detail"]}</div>', unsafe_allow_html=True)
    if not scopeleg.empty:
        st.markdown("---")
        st.markdown("#### Scope taxonomy")
        st.caption("The Scope 1/2/3 labels used in the register are a comparability overlay, not the legal test. "
                   "Each instrument defines embedded or embodied emissions in its own methodology annex.")
        st.dataframe(scopeleg, hide_index=True, use_container_width=True)

st.markdown("---")
st.caption("Independent tracker · built with Streamlit · data maintained in a version-controlled spreadsheet. Not affiliated with or endorsed by any organisation listed.")

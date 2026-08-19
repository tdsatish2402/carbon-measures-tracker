# Carbon Measures Tracker — Streamlit app
# Reads BCA_tracker.xlsx (multi-tab, normalised) and renders filterable views.
# Designed to fail gracefully on half-edited data: junk rows are dropped and reported, never shown as real records.

import pandas as pd
import streamlit as st

DATA_FILE = "BCA_tracker.xlsx"
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
@st.cache_data
def load():
    xl = pd.ExcelFile(DATA_FILE)
    def rd(tab): return pd.read_excel(DATA_FILE, tab, header=1) if tab in xl.sheet_names else pd.DataFrame()
    inst = rd("Instruments"); sec = rd("Sectors"); ev = rd("Events")
    src = rd("Sources"); catleg = rd("Category_legend")
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
    return inst, sec, ev, src, catleg, statleg, secleg, meth, dropped

inst, sec, ev, src, catleg, statleg, secleg, meth, dropped = load()

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

tab1, tab2, tab3, tab4, tab5 = st.tabs(
    ["Instruments", "Sector coverage", "Timeline", "Official sources", "Methodology"])

# ---- TAB 1: instrument cards ----
with tab1:
    if view.empty:
        st.info("No instruments match the current filters. Widen the selection in the sidebar.")
    for _, r in view.iterrows():
        st.markdown('<div class="cardwrap">', unsafe_allow_html=True)
        top = f'### {r["jurisdiction"]} — {r["instrument_name"]}'
        st.markdown(top)
        st.markdown(cat_pill(r["category"]) +
                    f'&nbsp;&nbsp;<span class="pill" style="background:#eee;color:#333">{r["status"]}</span>',
                    unsafe_allow_html=True)
        impl = r.get("Implementation/Coming into Force Date", "")
        if impl: st.markdown(f'<div class="lab">Implementation / in force</div><div class="val">{impl}</div>', unsafe_allow_html=True)
        if r.get("object_and_purpose"):
            st.markdown(f'<div class="lab">Object &amp; purpose</div><div class="val">{r["object_and_purpose"]}</div>', unsafe_allow_html=True)
        cc = st.columns(3)
        for col, (lab, key) in zip(cc, [("Emissions scope","emissions_scope"),
                                        ("3rd-country adjustment","third_country_adjustment"),
                                        ("Default values","default_values")]):
            col.markdown(f'<div class="lab">{lab}</div><div class="val">{r.get(key,"")}</div>', unsafe_allow_html=True)
        if r.get("calculation"):
            st.markdown(f'<div class="lab">Calculation</div><div class="val">{r["calculation"]}</div>', unsafe_allow_html=True)
        if r.get("revenue_use"):
            st.markdown(f'<div class="lab">Revenue use</div><div class="val">{r["revenue_use"]}</div>', unsafe_allow_html=True)
        # this instrument's sectors
        msec = sec[sec["instrument_id"] == r["instrument_id"]] if not sec.empty else pd.DataFrame()
        if not msec.empty:
            chips = " · ".join(sorted(msec["sector"].dropna().astype(str).unique()))
            st.markdown(f'<div class="lab">Sectors</div><div class="val">{chips}</div>', unsafe_allow_html=True)
        srcline = r.get("primary_source",""); url = r.get("official_url","")
        if url: st.markdown(f'<div class="lab">Primary source</div><div class="val">{srcline} — <a href="{url}" target="_blank">official page ↗</a></div>', unsafe_allow_html=True)
        if r.get("notes"): st.caption(r["notes"])
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
        mev["date"] = mev["date"].astype(str)
        mev = mev.sort_values("date")
        jmap = dict(zip(inst["instrument_id"], inst["jurisdiction"]))
        mev["jurisdiction"] = mev["instrument_id"].map(jmap)
        for _, r in mev.iterrows():
            link = f' — <a href="{r["official_url"]}" target="_blank">source ↗</a>' if r.get("official_url") else ""
            st.markdown(f'<div style="border-left:3px solid #0E3B43;padding:.15rem 0 .55rem .8rem;margin-left:.3rem">'
                        f'<b>{r["date"]}</b> &nbsp;·&nbsp; <span style="color:#7A6F5B">{r.get("jurisdiction","")}</span><br>'
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

# ---- TAB 5: methodology ----
with tab5:
    st.subheader("Methodology & attribution")
    if meth.empty:
        st.info("Methodology tab not found in the workbook.")
    else:
        for _, r in meth.iterrows():
            st.markdown(f'<div class="lab">{r["item"]}</div><div class="val">{r["detail"]}</div>', unsafe_allow_html=True)

st.markdown("---")
st.caption("Independent tracker · built with Streamlit · data maintained in a version-controlled spreadsheet. Not affiliated with or endorsed by any organisation listed.")

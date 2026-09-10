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
          --paper:#FFFFFF; --card:#FFFFFF; --line:#E1E5EC; --muted:#5B6478;
          --stripe:#F7F8FA; }
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
  .dt tbody tr:nth-child(even) td { background:var(--stripe); }
  .dt th { z-index:3; }
  .dt th.corner { position:sticky; left:0; z-index:5; border-right:1px solid #ffffff40; }
  .dt td.sec { font-weight:600; color:var(--ink); white-space:normal;
               position:sticky; left:0; background:var(--card); z-index:2;
               border-right:2px solid var(--line); min-width:190px; max-width:230px; }
  .dt tbody tr:nth-child(even) td.sec { background:var(--stripe); }
  .dt td.tick { text-align:center; font-size:21px; font-weight:700; color:#1B7A5A; line-height:1.1; }
  .dt td.pros { text-align:center; font-size:19px; color:#93A0B5; line-height:1.1; }
  .cvlegend { display:flex; flex-wrap:wrap; gap:1.1rem; align-items:center;
              font-size:.9rem; color:#243049; margin:.1rem 0 .9rem; }
  .cvlegend b { font-size:1.25rem; }
  .lead { color:var(--ink); font-size:1.03rem; line-height:1.62; font-weight:500;
           border-left:3px solid var(--gold); padding:.15rem 0 .15rem .85rem; margin:.1rem 0 .4rem; }
  .jchips { display:flex; flex-wrap:wrap; gap:.4rem; margin:.5rem 0 .2rem; }
  .jchip { display:inline-flex; align-items:center; gap:.45rem; text-decoration:none;
           border:1px solid var(--line); border-radius:999px; padding:.28rem .7rem;
           font-size:.87rem; color:var(--ink); background:var(--card); }
  .jchip:hover { border-color:var(--accent); color:var(--accent); }
  .jchip span { width:11px; height:11px; border-radius:50%; display:inline-block; }
  .bigprice { font-family:Georgia,serif; font-size:1.75rem; color:var(--ink);
              font-weight:600; line-height:1.15; }
  /* instrument card headers rendered as buttons: make them look like titles */
  .cardwrap div[data-testid="stButton"] button {
      background:transparent; border:none; padding:.1rem 0; box-shadow:none;
      font-family:Georgia,'Times New Roman',serif; font-size:1.2rem; font-weight:600;
      color:var(--ink); text-align:left; justify-content:flex-start; }
  .cardwrap div[data-testid="stButton"] button:hover { color:var(--accent); background:transparent; }
  .cardwrap div[data-testid="stButton"] button p { font-size:1.2rem; font-weight:600; }
  .dt th.jur { min-width:104px; white-space:normal; }
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
  .tl-ev { font-size:13.5px; line-height:1.45; padding:6px 8px; margin:4px 0; background:var(--stripe); border-radius:3px; color:#243049; }
  .tl-mo { display:inline-block; background:#E3E6EC; border-radius:3px; padding:1px 6px; margin-right:5px; font-weight:700; color:var(--ink); font-size:12px; }
  .foot { color:var(--muted); font-size:.85rem; line-height:1.5; border-top:1px solid var(--line);
          padding-top:.9rem; margin-top:1.6rem; }
</style>
""", unsafe_allow_html=True)

# Colours, labels and country codes all come from the workbook. The dicts below are
# only fallbacks for an older file that predates the control tabs.
FALLBACK_CAT = {"BCA": "#1D2657", "Domestic carbon pricing": "#8A6D1F",
                "Enabling law": "#3F106E", "Proposal": "#6B7280"}
FALLBACK_STAGE = {"In Force": ("In force", "#1B7A5A"),
                  "Draft": ("Draft / legislated, not yet in force", "#C77A16"),
                  "Conceptual": ("Conceptual", "#3B5FA8")}
MONTHS = {"01": "Jan", "02": "Feb", "03": "Mar", "04": "Apr", "05": "May", "06": "Jun",
          "07": "Jul", "08": "Aug", "09": "Sep", "10": "Oct", "11": "Nov", "12": "Dec"}

def html_table(headers, rows, max_height=None, first_col_sticky=False):
    """Render a real DOM table. st.dataframe paints to canvas, so its text can't
    be resized by CSS; these stay readable and let the scope-note column wrap."""
    import html as _h
    th = "".join(
        f'<th class="jur{" corner" if (first_col_sticky and i == 0) else ""}">{_h.escape(str(x))}</th>'
        for i, x in enumerate(headers))
    body = []
    for r in rows:
        tds = []
        for i, cell in enumerate(r):
            if isinstance(cell, tuple):          # (css_class, raw_html)
                tds.append(f'<td class="{cell[0]}">{cell[1]}</td>')
            else:
                cls = ' class="sec"' if (first_col_sticky and i == 0) else ""
                tds.append(f"<td{cls}>{_h.escape(str(cell))}</td>")
        # (tuples above carry pre-escaped HTML: (css_class, html))
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
    "Qualifying Carbon Prices": "qualifying_prices",
    "Verification": "verification",
    "Review and Appeal": "review_appeal",
    "Revenue Use": "revenue_use",
    "One Line Summary of the Measure": "notes",
    "Scope 1 Coverage": "scope1",
    "Scope 2 Coverage": "scope2",
    "Scope 3 Coverage": "scope3",
    "Other Scope Coverage": "scope_other",
}


# ---------- load + clean ----------
def file_stamp(path):
    """Fingerprint of the workbook. Passed into load() so that saving the Excel
    file changes the cache key and Streamlit re-reads it. Without this the
    @st.cache_data result is memoised forever and edits never appear."""
    try:
        st_ = os.stat(path)
        return (round(st_.st_mtime, 3), st_.st_size)
    except OSError:
        return (0, 0)


@st.cache_data(show_spinner="Reading workbook…")
def load(stamp):
    _ = stamp  # part of the cache key only
    xl = pd.ExcelFile(DATA_FILE)
    def rd(tab): return pd.read_excel(DATA_FILE, tab, header=1) if tab in xl.sheet_names else pd.DataFrame()
    inst = rd("Instruments"); sec = rd("Sectors"); ev = rd("Events")
    if not inst.empty:
        inst.columns = [str(c).strip() for c in inst.columns]
        inst = inst.rename(columns=COLMAP)
    src = rd("Sources"); catleg = rd("Category_legend"); link = rd("Carbon_price_recognition"); prices = rd("Prices")
    disp = rd("Display_text"); fields = rd("Field_display")
    stageleg = rd("Stage_legend"); covleg = rd("Coverage_legend"); catleg2 = rd("Category_legend")

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
    if not prices.empty and "price" in prices:
        prices["price"] = pd.to_numeric(prices["price"], errors="coerce")
        prices = prices[prices["price"].notna()]
    return (inst, sec, ev, src, link, prices, dropped,
            disp, fields, stageleg, covleg, catleg2)


(inst, sec, ev, src, link, prices, dropped,
 disp, fields, stageleg, covleg, catleg2) = load(file_stamp(DATA_FILE))

# ---------- everything below is driven by the workbook's control tabs ----------
def _col(df, name):
    return df[name] if (not df.empty and name in df.columns) else pd.Series(dtype=object)


TEXT = dict(zip(_col(disp, "key").astype(str), _col(disp, "value").astype(str)))


def cap(key):
    """Render a caption only if Display_text still holds a value for it, so a
    caption can be retired by clearing or deleting its row in the workbook."""
    v = T(key, "")
    if v:
        st.caption(v)


def T(key, default=""):
    """UI string from the Display_text tab, falling back to a built-in default."""
    v = TEXT.get(key, "")
    return v if v and v != "nan" else default


# stage: label + colour from Stage_legend
if not stageleg.empty and {"stage_value", "colour"} <= set(stageleg.columns):
    STAGE_LABEL = dict(zip(stageleg["stage_value"].astype(str),
                           stageleg["display_label"].astype(str)))
    STAGE_COLOR = dict(zip(stageleg["stage_value"].astype(str),
                           stageleg["colour"].astype(str)))
    STAGE_ORDER = list(stageleg["stage_value"].astype(str))
else:
    STAGE_LABEL = {k: v[0] for k, v in FALLBACK_STAGE.items()}
    STAGE_COLOR = {k: v[1] for k, v in FALLBACK_STAGE.items()}
    STAGE_ORDER = list(FALLBACK_STAGE)
LEGEND_COLOR = {STAGE_LABEL[k]: v for k, v in STAGE_COLOR.items() if k in STAGE_LABEL}
STAGE_RANK = {k: len(STAGE_ORDER) - i for i, k in enumerate(STAGE_ORDER)}

# category colours from Category_legend
if not catleg2.empty and {"category_value", "colour"} <= set(catleg2.columns):
    CAT_COLOR = dict(zip(catleg2["category_value"].astype(str), catleg2["colour"].astype(str)))
else:
    CAT_COLOR = dict(FALLBACK_CAT)

# coverage symbols + colours from Coverage_legend
if not covleg.empty and {"coverage_value", "symbol", "colour"} <= set(covleg.columns):
    COV = {r["coverage_value"]: (str(r["symbol"]), str(r["colour"]))
           for _, r in covleg.iterrows()}
else:
    COV = {"Current scope": ("\u2713", "#1B7A5A"), "Prospective": ("\u25cb", "#93A0B5")}
COV_CURRENT = next(iter(COV), "Current scope")

# instrument_id -> jurisdiction, used wherever a raw code would otherwise be shown
JMAP = dict(zip(inst["instrument_id"], inst["jurisdiction"]))
CMAP = dict(zip(inst["instrument_id"], inst["category"]))
SMAP = dict(zip(inst["instrument_id"], inst["status_simple"]))
# sheet order = display order (EU, UK, Australia, ...)
JUR_ORDER = list(dict.fromkeys(inst["jurisdiction"]))


NMAP = dict(zip(inst["instrument_id"], inst["instrument_name"]))
# sheet row order, so columns follow the register rather than the alphabet
ID_ORDER = {iid: n for n, iid in enumerate(inst["instrument_id"])}


def short_name(iid):
    """'Foreign Pollution Fee Act of 2025 (S. 1325, ...)' -> 'Foreign Pollution Fee Act'"""
    n = str(NMAP.get(iid, iid)).split(" of 20")[0].split(" (")[0].strip()
    return n if len(n) <= 34 else n[:31].rstrip() + "..."


def card_sections():
    """Ordered [(order, section_label, [field rows])] from the Field_display tab.
    Rows sharing an 'order' become one collapsible section."""
    if fields.empty or "column_name" not in fields.columns:
        return []
    f = fields.copy()
    f["show"] = f.get("show", "yes").astype(str).str.strip().str.lower()
    f = f[~f["show"].isin(["no", "false", "0", "n"])]
    f["order"] = pd.to_numeric(f["order"], errors="coerce").fillna(999)
    f["field_label"] = f.get("field_label", "").fillna("").astype(str).replace("nan", "")
    out = []
    for o in sorted(f["order"].unique()):
        g = f[f["order"] == o]
        out.append((o, str(g["section_label"].iloc[0]), g.to_dict("records")))
    return out


def iso3_for(row):
    """ISO-3 codes for the choropleth, read from the instrument's own iso3_codes cell.
    Semicolon-separated so a bloc (the EU) can list all its member states."""
    raw = str(row.get("iso3_codes", "") or "")
    return [c.strip().upper() for c in raw.replace(",", ";").split(";")
            if len(c.strip()) == 3]


def latest_price(iid):
    """Most recent published price row for an instrument, or None."""
    if prices.empty or "instrument_id" not in prices:
        return None
    d = prices[prices["instrument_id"] == iid]
    if d.empty:
        return None
    return d.sort_values("published_date").iloc[-1]


def fmt_price(r):
    return f'{r["currency"]} {r["price"]:,.2f} / {r["unit"]}'


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
st.title(T("app.title", "BCA Tracker"))
st.caption(T("app.tagline", ""))

def render_card_body(r, iid, OPEN):
    """Render one instrument's collapsible sections. Shared by the Instruments tab
    and the single-instrument deep-link page, so they can never drift apart."""
    # ---- sections are defined by the Field_display tab, not by this file ----
    for order, sec_label, items in card_sections():
        special = [i for i in items if str(i["layout"]).lower() == "special"]
        if special:
            kind = str(special[0]["column_name"]).strip().lower()
            if kind == "(prices)":
                lp = latest_price(iid)
                if lp is None:
                    continue
                hist = prices[prices["instrument_id"] == iid].sort_values("published_date")
                with st.expander(f'{sec_label} — {fmt_price(lp)} ({lp["period"]})', expanded=OPEN):
                    st.markdown(f'<div class="bigprice">{fmt_price(lp)}</div>'
                                f'<div class="lab" style="margin-top:0">{lp["period"]} · '
                                f'published {lp["published_date"]}</div>', unsafe_allow_html=True)
                    st.markdown(f'<div class="val" style="margin-top:.6rem">{lp["basis"]}</div>',
                                unsafe_allow_html=True)
                    if lp.get("notes"):
                        st.markdown(f'<div class="val" style="color:var(--muted)">{lp["notes"]}</div>',
                                    unsafe_allow_html=True)
                    if len(hist) > 1:
                        st.markdown(f'<div class="lab">{T("label.pricehistory", "Published history")}</div>',
                                    unsafe_allow_html=True)
                        rows = [[h["period"], h["published_date"], fmt_price(h),
                                 ("", f'<a href="{h["official_url"]}" target="_blank">open ↗</a>'
                                  if str(h.get("official_url", "")).startswith("http") else "")]
                                for _, h in hist[::-1].iterrows()]
                        html_table(["Period", "Published", "Price", "Source"], rows)
                    if len(hist) >= 4:
                        st.line_chart(hist.set_index("period")["price"], height=200)

            elif kind == "(sectors)":
                msec = sec[sec["instrument_id"] == iid] if not sec.empty else pd.DataFrame()
                if msec.empty:
                    continue
                with st.expander(sec_label, expanded=OPEN):
                    cur = sorted(msec[msec["coverage"] == COV_CURRENT]["sector"].astype(str).unique())
                    pro = sorted(msec[msec["coverage"] != COV_CURRENT]["sector"].astype(str).unique())
                    st.markdown(f'<div class="lab">{T("label.inscope", "In scope")}</div>'
                                f'<div class="val">{" · ".join(cur) if cur else "None defined yet"}</div>',
                                unsafe_allow_html=True)
                    if pro:
                        st.markdown(f'<div class="lab">{T("label.prospective", "Prospective")}</div>'
                                    f'<div class="val">{" · ".join(pro)}</div>', unsafe_allow_html=True)

            elif kind == "(legal)":
                lb = str(r.get("legal_basis", "") or "")
                url = str(r.get("official_url", "") or "")
                if not lb and not url:
                    continue
                with st.expander(sec_label, expanded=OPEN):
                    if lb:
                        st.markdown(f'<div class="val">{lb}</div>', unsafe_allow_html=True)
                    if url:
                        st.markdown('<div class="val" style="margin-top:.5rem">'
                                    f'<a href="{url}" target="_blank">official page ↗</a></div>',
                                    unsafe_allow_html=True)
            continue

        # ordinary text fields: 'full' spans the card, 'third' sits in a 3-up row
        vals = [(i, str(r.get(COLMAP.get(i["column_name"], i["column_name"]), "") or ""))
                for i in items]
        if not any(v for _, v in vals):
            continue
        with st.expander(sec_label, expanded=OPEN):
            thirds = [(i, v) for i, v in vals if str(i["layout"]).lower() == "third"]
            if thirds:
                cols = st.columns(max(len(thirds), 1))
                for col, (i, v) in zip(cols, thirds):
                    col.markdown(f'<div class="lab">{i["field_label"] or i["column_name"]}</div>'
                                 f'<div class="val">{v}</div>', unsafe_allow_html=True)
            for i, v in vals:
                lay = str(i["layout"]).lower()
                if lay == "third" or not v:
                    continue
                if i["field_label"]:
                    st.markdown(f'<div class="lab">{i["field_label"]}</div>', unsafe_allow_html=True)
                cls = "lead" if lay == "lead" else "val"
                st.markdown(f'<div class="{cls}">{v}</div>', unsafe_allow_html=True)



# ---------- deep link: ?instrument=EU-CBAM opens one instrument on its own page ----------
def instrument_url(iid):
    return f"?instrument={iid}"


_focus = str(st.query_params.get("instrument", "") or "").strip()
if _focus and _focus in set(inst["instrument_id"]):
    fr = inst[inst["instrument_id"] == _focus].iloc[0]
    st.markdown(f'<a href="./" style="font-size:.9rem">← back to the full tracker</a>',
                unsafe_allow_html=True)
    st.markdown(f'### {fr["jurisdiction"]} — {fr["instrument_name"]}')
    st.markdown(cat_pill(fr["category"]) + "&nbsp;&nbsp;" +
                stage_pill(fr["status"], fr.get("status_simple", "")), unsafe_allow_html=True)
    if fr.get("notes"):
        st.markdown('<div class="val" style="font-style:italic;color:#4a4030;'
                    f'margin:.55rem 0 .95rem">{fr["notes"]}</div>', unsafe_allow_html=True)
    OPEN = True
    view = inst[inst["instrument_id"] == _focus]
    render_card_body(fr, _focus, OPEN)
    st.markdown(f'<div class="foot">{T("app.footer", "")}</div>', unsafe_allow_html=True)
    st.stop()

# ---------- sidebar filters ----------
from datetime import datetime

if not os.path.exists(DATA_FILE):
    st.error(f"Cannot find **{DATA_FILE}**. It must sit in the same folder as app.py. "
             f"Currently looking in: `{os.getcwd()}`")
    st.stop()

_mt = datetime.fromtimestamp(os.path.getmtime(DATA_FILE))
st.sidebar.caption(f"Data file last saved  \n**{_mt:%d %b %Y, %H:%M:%S}**")
if st.sidebar.button(T("sidebar.reload", "Reload data"), use_container_width=True):
    st.cache_data.clear()
    st.rerun()
st.sidebar.markdown("---")
# Sidebar filters are hidden while the dataset is small - with 8 of 11 rows in one
# category they discriminated almost nothing. The Display_text keys and this block
# are kept so they can be switched back on once there is more to analyse.
SHOW_FILTERS = False
if SHOW_FILTERS:
    st.sidebar.header(T("sidebar.filter", "Filter"))
    cats = sorted([c for c in inst["category"].unique() if c])
    stages = [x for x in STAGE_ORDER if x in set(inst["status_simple"])]
    juris = sorted([j for j in inst["jurisdiction"].unique() if j])
    f_cat = st.sidebar.multiselect(T("sidebar.category", "Category"), cats, default=cats)
    f_stage = st.sidebar.multiselect(T("sidebar.stage", "Stage"), stages, default=stages)
    f_jur = st.sidebar.multiselect(T("sidebar.jurisdiction", "Jurisdiction"), juris, default=juris)
    view = inst[inst["category"].isin(f_cat) & inst["status_simple"].isin(f_stage)
                & inst["jurisdiction"].isin(f_jur)]
else:
    view = inst.copy()

@st.cache_data
def health(stamp):
    """Catch edits that would otherwise fail silently: a mistyped category drops the
    whole row, an unrecognised coverage value renders as an empty matrix cell, and an
    instrument_id typo orphans every child row."""
    _ = stamp
    xl = pd.ExcelFile(DATA_FILE)
    g = lambda t: pd.read_excel(DATA_FILE, t, header=1) if t in xl.sheet_names else pd.DataFrame()
    issues = []
    I, S = g("Instruments"), g("Sectors")
    legends = {"category": ("Category_legend", "category_value", I),
               "status": ("Status_legend", "status_value", I),
               "sector": ("Sector_legend", "sector_value", S),
               "coverage": ("Coverage_legend", "coverage_value", S)}
    for col, (sheet, key, df) in legends.items():
        L = g(sheet)
        if df.empty or L.empty or col not in df or key not in L:
            continue
        bad = sorted(set(df[col].dropna().astype(str)) - set(L[key].dropna().astype(str)))
        if bad:
            issues.append(f"**{col}** not in {sheet}: " + ", ".join(f"`{b}`" for b in bad[:5]))
    if not I.empty and "status_simple" in I:
        bad = sorted(set(I["status_simple"].dropna().astype(str)) - set(STAGE_COLOR))
        if bad:
            issues.append("**status_simple** must be In Force / Draft / Conceptual: "
                          + ", ".join(f"`{b}`" for b in bad[:5]))
    FD = g("Field_display")
    if not FD.empty and not I.empty and "column_name" in FD:
        have = set(I.columns.astype(str))
        for _, b in FD.iterrows():
            cn = str(b["column_name"]).strip()
            if cn.startswith("(") or str(b.get("show", "yes")).lower() in ("no", "false", "0"):
                continue
            if cn not in have:
                issues.append(f"**Field_display** refers to column `{cn}`, which is not a header in "
                              "the Instruments tab — that section will be skipped.")
    DT = g("Display_text")
    if not DT.empty and "key" in DT:
        dup = DT["key"].astype(str)[DT["key"].astype(str).duplicated()].unique()
        if len(dup):
            issues.append("**Display_text** has duplicate keys (the last one wins): "
                          + ", ".join(f"`{d}`" for d in dup[:5]))
    if not I.empty and "iso3_codes" in I:
        for _, b in I.iterrows():
            codes = [c.strip() for c in str(b.get("iso3_codes", "") or "").replace(",", ";").split(";") if c.strip()]
            bad = [c for c in codes if len(c) != 3]
            if bad:
                issues.append(f"**Instruments** `{b['instrument_id']}` has invalid iso3_codes "
                              + ", ".join(f"`{c}`" for c in bad[:3]) + " — use 3-letter codes separated by `;`.")
            elif not codes:
                issues.append(f"**Instruments** `{b['instrument_id']}` has no iso3_codes, so it will not "
                              "appear on the Overview map.")

    P = g("Prices")
    if not P.empty and "price" in P:
        bad = P[pd.to_numeric(P["price"], errors="coerce").isna()]
        for _, b in bad.iterrows():
            issues.append(f"**Prices** row `{b.get('instrument_id','?')} / {b.get('period','?')}` "
                          f"has a non-numeric price `{b['price']}` and will be ignored — "
                          "enter digits only, no currency symbol.")
        if "published_date" in P:
            badd = P[pd.to_datetime(P["published_date"], errors="coerce", format="mixed").isna()]
            for _, b in badd.iterrows():
                issues.append(f"**Prices** row `{b.get('period','?')}` has an unreadable "
                              f"published_date `{b['published_date']}` — use YYYY-MM-DD.")

    if not I.empty:
        ids = set(I["instrument_id"].dropna().astype(str))
        for t in ["Sectors", "Events", "Sources", "Carbon_price_recognition", "Prices"]:
            d = g(t)
            if d.empty or "instrument_id" not in d:
                continue
            orph = sorted(set(d["instrument_id"].dropna().astype(str)) - ids)
            if orph:
                issues.append(f"**{t}** references unknown instrument_id: "
                              + ", ".join(f"`{o}`" for o in orph[:5]))
    return issues


_issues = health(file_stamp(DATA_FILE))
if dropped or _issues:
    st.sidebar.markdown("---")
    st.sidebar.markdown(f'**{T("sidebar.datacheck", "Data check")}**')
    if dropped:
        st.sidebar.warning(f"{len(dropped)} row(s) skipped — the category is not in "
                           "Category_legend: " + ", ".join(dropped[:6]))
    for msg in _issues:
        st.sidebar.warning(msg)
    st.sidebar.caption("Fix these in the workbook, save, then press Reload data. "
                       "Values must match the legend tabs exactly.")

# ---------- top metrics ----------
c1, c2, c3 = st.columns(3)
c1.metric(T("metric.instruments", "Instruments"), len(view))
c2.metric(T("metric.jurisdictions", "Jurisdictions"), view["jurisdiction"].nunique())
c3.metric(T("metric.inforce", "In force"), (view["status_simple"] == STAGE_ORDER[0]).sum())

tab0, tab1, tab2, tab5, tab3, tab4 = st.tabs(
    [T("tab.overview", "Overview"), T("tab.instruments", "Instruments"),
     T("tab.sectors", "Sector coverage"), T("tab.recognition", "Carbon price recognition"),
     T("tab.timeline", "Timeline"), T("tab.sources", "Official sources")])

# ---- TAB 0: overview map ----
with tab0:
    st.subheader(T("tab.overview", "Overview"))
    cap("caption.overview")

    if view.empty:
        st.info(T("empty.filters", "No instruments match the current filters."))
    else:
        # one row per country code, taking the most advanced stage where a
        # jurisdiction has several instruments (e.g. the two US bills)
        recs = {}
        for _, r in view.iterrows():
            for iso in iso3_for(r):
                rank = STAGE_RANK.get(r["status_simple"], 0)
                cur = recs.get(iso)
                if cur is None or rank > cur["rank"]:
                    recs[iso] = {"iso": iso, "Jurisdiction": r["jurisdiction"],
                                 "Stage": STAGE_LABEL.get(r["status_simple"], r["status_simple"]),
                                 "rank": rank,
                                 "measures": list(cur["measures"]) if cur else []}
                recs[iso]["measures"].append(r["instrument_name"])
        mapdf = pd.DataFrame([{**v, "Measures": " · ".join(v["measures"])} for v in recs.values()])

        unmapped = sorted({r["jurisdiction"] for _, r in view.iterrows() if not iso3_for(r)})

        try:
            import plotly.express as px
            fig = px.choropleth(
                mapdf, locations="iso", locationmode="ISO-3", color="Stage",
                color_discrete_map=LEGEND_COLOR,
                category_orders={"Stage": [STAGE_LABEL[k] for k in STAGE_ORDER
                                           if k in STAGE_LABEL]},
                hover_name="Jurisdiction",
                hover_data={"iso": False, "Stage": True, "Measures": True},
            )
            # crop Antarctica and the empty high Arctic: on a laptop this removes a
            # third of the dead space, and on a phone it makes the map legible at all
            fig.update_geos(showframe=False, showcoastlines=False, showcountries=True,
                            countrycolor="#D9DEE8", landcolor="#EDF0F5", lakecolor="#FFFFFF",
                            bgcolor="rgba(0,0,0,0)", projection_type="natural earth",
                            lataxis_range=[-56, 84], lonaxis_range=[-168, 190])
            fig.update_layout(
                margin=dict(l=0, r=0, t=0, b=0), height=430, dragmode=False,
                paper_bgcolor="rgba(0,0,0,0)", geo_bgcolor="rgba(0,0,0,0)",
                legend=dict(orientation="h", yanchor="bottom", y=-0.06,
                            xanchor="left", x=0, title_text="",
                            font=dict(size=14)),
                font=dict(family="Arial", size=14, color="#243049"),
            )
            st.plotly_chart(fig, use_container_width=True,
                            config={"displayModeBar": False, "scrollZoom": False,
                                    "staticPlot": False, "responsive": True})
        except ModuleNotFoundError:
            st.warning("The map needs plotly. Install it with:  pip install plotly  "
                       "— open the table below in the meantime.")

        if unmapped:
            st.caption("Not shown on the map — add an ISO-3 code in the iso3_codes column of the "
                       "Instruments tab for: " + ", ".join(unmapped))

        # jurisdiction chips: each opens that instrument on its own page in a new browser tab.
        # Streamlit cannot switch the active tab programmatically, so a deep link is used instead.
        # grouped by stage, alphabetical within each group, so the dots read as runs
        chip_rank = {k: i for i, k in enumerate(STAGE_ORDER)}
        ordered = sorted(
            view.to_dict("records"),
            key=lambda x: (chip_rank.get(x.get("status_simple"), 99), str(x.get("jurisdiction", ""))))
        chips = "".join(
            f'<a class="jchip" href="{instrument_url(x["instrument_id"])}" target="_blank" '
            f'rel="noopener" title="{x["instrument_name"]}">'
            f'<span style="background:{STAGE_COLOR.get(x["status_simple"], "#5B6478")}"></span>'
            f'{x["jurisdiction"]} ↗</a>'
            for x in ordered)
        st.markdown(f'<div class="jchips">{chips}</div>', unsafe_allow_html=True)
        st.caption("Open any measure on its own page in a new browser tab.")

# ---- TAB 1: instrument cards ----
with tab1:
    if view.empty:
        st.info(T("empty.filters", "No instruments match the current filters."))
    else:
        # TWO independent levels. Streamlit forbids expanders inside expanders, so the
        # card level is a session-state toggle and only the sub-sections are expanders.
        st.session_state.setdefault("open_cards", set())
        st.session_state.setdefault("expand_all", False)

        pick_col, b1, b2 = st.columns([2, 1, 1])
        all_ids = list(view["instrument_id"])
        # type-to-search picker: the list itself tells the reader what is covered
        opts = [T("picker.all", "All instruments")] + [
            f'{x["jurisdiction"]} — {x["instrument_name"]}' for _, x in view.iterrows()]
        picked = pick_col.selectbox(T("picker.label", "Go to instrument"), opts,
                                    label_visibility="collapsed")
        cards_all_open = st.session_state.open_cards.issuperset(all_ids)
        if b1.button(T("button.collapse.instruments", "Collapse all instruments") if cards_all_open
                     else T("button.expand.instruments", "Expand all instruments"),
                     use_container_width=True):
            st.session_state.open_cards = set() if cards_all_open else set(all_ids)
            st.rerun()
        if b2.button(T("button.collapse.sections", "Collapse all sections") if st.session_state.expand_all
                     else T("button.expand.sections", "Expand all sections"),
                     use_container_width=True):
            st.session_state.expand_all = not st.session_state.expand_all
            st.rerun()
        OPEN = st.session_state.expand_all

        show = view
        if picked != T("picker.all", "All instruments"):
            show = view[view.apply(
                lambda r: f'{r["jurisdiction"]} — {r["instrument_name"]}' == picked, axis=1)]

        def val_block(v):
            st.markdown(f'<div class="val">{v}</div>', unsafe_allow_html=True)

        for _, r in show.iterrows():
            if True:
                iid = r["instrument_id"]
                is_open = iid in st.session_state.open_cards
                st.markdown('<div class="cardwrap">', unsafe_allow_html=True)
                if st.button(f'{"▾" if is_open else "▸"}\u2003{r["jurisdiction"]} — {r["instrument_name"]}',
                             key=f"card_{iid}", use_container_width=True):
                    st.session_state.open_cards ^= {iid}
                    st.rerun()
                # one status pill only, tinted by its stage — the separate
                # "In Force" pill duplicated what "In force / operational" says
                st.markdown(cat_pill(r["category"]) + "&nbsp;&nbsp;" +
                            stage_pill(r["status"], r.get("status_simple", "")),
                            unsafe_allow_html=True)
                if not is_open:
                    st.markdown('</div>', unsafe_allow_html=True)
                    continue
                # one-line summary sits directly under the title, always visible
                if r.get("notes"):
                    st.markdown('<div class="val" style="font-style:italic;color:#4a4030;'
                                f'margin:.55rem 0 .95rem">{r["notes"]}</div>', unsafe_allow_html=True)

                render_card_body(r, iid, OPEN)
                st.markdown('</div>', unsafe_allow_html=True)

# ---- TAB 2: sector coverage matrix ----
with tab2:
    st.subheader(T("tab.sectors", "Sector coverage"))
    keep_ids = set(view["instrument_id"])
    msec = sec[sec["instrument_id"].isin(keep_ids)] if not sec.empty else pd.DataFrame()
    if msec.empty:
        st.info("No sector data for the current filter.")
    else:
        lab = label_for(sorted(set(msec["instrument_id"])))
        m = msec.copy()
        m["Jurisdiction"] = m["instrument_id"].map(lab)
        mat = pd.crosstab(m["Jurisdiction"], m["sector"])
        # rows follow the register order, not the alphabet
        seen, order = set(), []
        for iid in sorted(lab, key=lambda i: ID_ORDER.get(i, 999)):
            c = lab[iid]
            if c in mat.index and c not in seen:
                seen.add(c); order.append(c)
        order += [c for c in mat.index if c not in seen]
        mat = mat.loc[order]
        sectors = list(mat.columns)
        cover = {(lab[r["instrument_id"]], r["sector"]): r["coverage"]
                 for _, r in m.iterrows()}
        stage_of = {lab[i]: SMAP.get(i, "") for i in lab}

        def cell(jur, sector):
            cv = cover.get((jur, sector))
            if cv in COV:
                sym, col = COV[cv]
                return ("tick", f'<span style="color:{col}">{sym}</span>')
            return ("tick", "")

        def row_label(jur):
            c = STAGE_COLOR.get(stage_of.get(jur, ""), "#5B6478")
            return ("sec", f'<span style="display:inline-block;width:14px;height:14px;'
                           f'border-radius:50%;background:{c};margin-right:9px;'
                           f'vertical-align:middle"></span>{jur}')

        rows = [[row_label(j)] + [cell(j, sc) for sc in sectors] for j in order]
        html_table(["Jurisdiction"] + sectors, rows, max_height=560, first_col_sticky=True)

        keys = "".join(
            f'<span><b style="color:{c}">{sym}</b> '
            f'{T("legend.inscope", "in scope") if k == COV_CURRENT else T("legend.prospective", "prospective")}'
            '</span>' for k, (sym, c) in COV.items())
        dots = "".join(
            f'<span><b style="color:{STAGE_COLOR[k]}">●</b> {v.lower()}</span>'
            for k, v in STAGE_LABEL.items() if k in set(view["status_simple"]))
        st.markdown(f'<div class="cvlegend">{keys}{dots}</div>', unsafe_allow_html=True)
        cap("caption.sectors")

        # HS codes here are chapter-level orientation, not the instruments' actual
        # commodity-code lists. Hidden until the dataset carries verified codes.

# ---- TAB 5: recognition of third-country carbon prices ----
# "recognition", not "linkage": linkage is a distinct legal concept (mutual recognition
# of allowances between two trading systems, as with the EU-Swiss ETS).
with tab5:
    st.subheader(T("tab.recognition", "Carbon price recognition"))
    keep_ids = set(view["instrument_id"])
    ml = link[link["instrument_id"].isin(keep_ids)].copy() if not link.empty else pd.DataFrame()
    if ml.empty:
        st.info("No recognition data for the current filter.")
    else:
        ml["Jurisdiction"] = ml["instrument_id"].map(JMAP)
        ml["_o"] = ml["instrument_id"].map(lambda i: ID_ORDER.get(i, 999))
        ml = ml.sort_values(["_o", "scheme_name"])
        rec = ml["recognition_status"].astype(str).str.startswith("Recognised")
        st.caption("Which foreign carbon prices each measure will credit against its own charge. "
                   f"{int(rec.sum())} scheme(s) are formally recognised, all by a single jurisdiction — "
                   "the UK is so far the only one to have published a list, while the EU's Article 9 "
                   "implementing act remains a draft. Emissions covered by free allowances never qualify, "
                   "because no effective price was paid on them.")

        st.session_state.setdefault("open_rec", set())
        groups = list(dict.fromkeys(ml["instrument_id"]))
        rb1, _rb = st.columns([1, 3])
        rec_all_open = st.session_state.open_rec.issuperset(groups)
        if rb1.button(T("button.collapse", "Collapse all") if rec_all_open
                      else T("button.expand", "Expand all"),
                      key="rec_toggle", use_container_width=True):
            st.session_state.open_rec = set() if rec_all_open else set(groups)
            st.rerun()

        for iid in groups:
            g = ml[ml["instrument_id"] == iid]
            n_rec = int(g["recognition_status"].astype(str).str.startswith("Recognised").sum())
            head = f"{JMAP.get(iid, iid)} — " + (f"{n_rec} schemes recognised" if n_rec
                                                 else str(g["recognition_status"].iloc[0]))
            with st.expander(head, expanded=(iid in st.session_state.open_rec)):
                rows = []
                for _, x in g.iterrows():
                    scheme = str(x["scheme_name"])
                    su = str(x.get("scheme_url", "") or "")
                    cell = ("", f'<a href="{su}" target="_blank">{scheme}</a>') \
                        if su.startswith("http") else scheme
                    link_cell = ("", f'<a href="{x["source_url"]}" target="_blank">open ↗</a>') \
                        if str(x.get("source_url", "")).startswith("http") else ""
                    rows.append([cell, x["scheme_jurisdiction"], x["recognition_status"],
                                 x["notes"], link_cell])
                html_table(["Scheme", "Scheme jurisdiction", "Status", "Notes", "Legal basis"],
                           rows, max_height=520)


# ---- TAB 3: timeline ----
with tab3:
    st.subheader(T("tab.timeline", "Timeline"))
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
            cap("caption.timeline.swimlane")
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
            cap("caption.timeline.list")
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
    st.subheader(T("tab.sources", "Sources"))
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
st.markdown(f'<div class="foot">{T("app.footer", "")}</div>', unsafe_allow_html=True)

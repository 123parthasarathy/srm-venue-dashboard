"""
SRM Ramapuram — public venue dashboard (Streamlit Cloud).

Anyone with the link can view. No login.

Update a semester: replace files in data/ (any filename), push to GitHub.
Optional override: data/<dept>/venues.csv  (copy data/_template_venues.csv).
New programme: one block in config.json + a folder under data/.
"""

from __future__ import annotations

import html
import json
import os
import re

import streamlit as st

from ingest import (
    DAYS,
    PERIOD_TIMES,
    data_stamp,
    load_all,
    occupancy_rows,
    search_catalog,
    year_sort,
)

st.set_page_config(
    page_title="SRM Ramapuram · Venue Dashboard",
    page_icon="🏫",
    layout="wide",
    initial_sidebar_state="collapsed",
)

ROOT = os.path.dirname(os.path.abspath(__file__))
DATA_DIR = os.path.join(ROOT, "data")
CONFIG_PATH = os.path.join(ROOT, "config.json")

DAY_NAMES = dict(zip(DAYS, ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday"]))
# Google brand palette for weekdays
DAY_COLORS = {"MON": "#4285F4", "TUE": "#EA4335", "WED": "#FBBC04", "THU": "#34A853", "FRI": "#4285F4"}
DAY_BG = {"MON": "#E8F0FE", "TUE": "#FCE8E6", "WED": "#FEF7E0", "THU": "#E6F4EA", "FRI": "#E8F0FE"}
GOOGLE = {"blue": "#4285F4", "red": "#EA4335", "yellow": "#FBBC04", "green": "#34A853"}


def hx(value) -> str:
    return html.escape(str(value or ""), quote=True)


def dept_slug(dept_id: str) -> str:
    return re.sub(r"[^A-Za-z0-9]+", "_", dept_id).strip("_")


def dept_theme(d: dict) -> tuple[str, str]:
    return d.get("color", GOOGLE["blue"]), d.get("bg", "#E8F0FE")


def inject_dept_button_css(meta: dict) -> None:
    rules = []
    for dk, d in meta.items():
        c, _ = dept_theme(d)
        slug = dept_slug(dk)
        rules.append(
            f'[data-testid="column"]:has(.g-folder[data-dept="{slug}"]) div.stButton > button {{'
            f"background:{c} !important;color:#fff !important;border:none !important;}}"
        )
    if rules:
        st.markdown(f"<style>{''.join(rules)}</style>", unsafe_allow_html=True)


def folder_card(d: dict, dk: str, n: int) -> str:
    c, bg = dept_theme(d)
    slug = dept_slug(dk)
    extra = f" · {hx(d['semester'])}" if d.get("semester") else ""
    return (
        f'<div class="g-folder" data-dept="{slug}" style="--c:{c};--bg:{bg}">'
        f'<div class="g-icon">{hx(d.get("icon", ""))}</div>'
        f'<div class="g-id">{hx(dk)}</div>'
        f'<div class="g-name">{hx(d["name"])}{extra}</div>'
        f'<span class="g-count">{n} sections</span></div>'
    )


def hero_block(title: str, subtitle: str) -> str:
    return (
        '<div class="g-stripe"><span class="b"></span><span class="r"></span>'
        '<span class="y"></span><span class="g"></span></div>'
        f'<div class="hero"><h1>{hx(title)}</h1>'
        f"<p>{hx(subtitle)} · Public link · no login</p></div>"
    )


def dept_banner_html(dept: dict, dept_key: str, records: list[dict]) -> str:
    c, bg = dept_theme(dept)
    sem = f" · {hx(dept['semester'])}" if dept.get("semester") else ""
    return (
        f'<div class="dept-banner" style="--banner-c:{c};--banner-bg:{bg}">'
        f'<h2>{hx(dept.get("icon",""))} {hx(dept_key)} — {hx(dept["name"])}</h2>'
        f"<p>{len(records)} sections · {len({r['year'] for r in records})} years{sem} · "
        f"share this view: the URL keeps the programme</p></div>"
    )


st.markdown(
    """
<style>
@import url('https://fonts.googleapis.com/css2?family=Google+Sans:wght@400;500;700&family=Roboto:wght@400;500;700&display=swap');
html, body, [class*="css"] { font-family: 'Roboto', 'Google Sans', sans-serif; }
.main .block-container { padding-top: 1rem; max-width: 1320px; background: #FAFAFA; }
#MainMenu, footer, .stDeployButton { visibility: hidden; }
header[data-testid="stHeader"] { background: transparent; }

/* Google logo stripe */
.g-stripe {
    display: flex; height: 5px; border-radius: 8px 8px 0 0; overflow: hidden; margin-bottom: 0;
}
.g-stripe span { flex: 1; }
.g-stripe .b { background: #4285F4; }
.g-stripe .r { background: #EA4335; }
.g-stripe .y { background: #FBBC04; }
.g-stripe .g { background: #34A853; }

.hero {
    background: #fff; color: #202124;
    padding: 20px 26px 18px; border-radius: 0 0 16px 16px; margin-bottom: 16px;
    box-shadow: 0 1px 3px rgba(60,64,67,.18), 0 4px 8px rgba(60,64,67,.08);
    border: 1px solid #E8EAED; border-top: none;
}
.hero h1 { font-size: 21px; font-weight: 700; margin: 0; color: #202124; }
.hero p { margin: 6px 0 0; font-size: 13px; color: #5F6368; }

.stats { display: flex; gap: 10px; flex-wrap: wrap; margin: 14px 0 8px; }
.stat {
    background: #fff; border: 1px solid #E8EAED; border-radius: 12px;
    padding: 10px 16px; font-size: 12px; color: #5F6368; font-weight: 500;
    box-shadow: 0 1px 2px rgba(60,64,67,.06);
}
.stat b { font-size: 18px; margin-right: 6px; font-weight: 700; }
.stat:nth-child(1) b { color: #4285F4; }
.stat:nth-child(2) b { color: #34A853; }
.stat:nth-child(3) b { color: #EA4335; }

/* Programme folder cards */
.prog-wrap { margin-bottom: 4px; }
.g-folder {
    background: var(--bg, #fff); border: 1px solid #E8EAED;
    border-radius: 14px; padding: 14px 12px 10px; margin-bottom: 6px;
    border-top: 4px solid var(--c, #4285F4);
    box-shadow: 0 1px 3px rgba(60,64,67,.12);
    transition: transform .15s, box-shadow .15s;
}
.g-folder:hover { transform: translateY(-2px); box-shadow: 0 4px 12px rgba(60,64,67,.18); }
.g-folder .g-icon { font-size: 32px; line-height: 1; margin-bottom: 6px; }
.g-folder .g-id {
    font-size: 17px; font-weight: 700; color: var(--c, #4285F4);
    letter-spacing: -0.02em;
}
.g-folder .g-name {
    font-size: 11px; color: #5F6368; margin-top: 4px; line-height: 1.35;
}
.g-folder .g-count {
    display: inline-block; margin-top: 8px; font-size: 10px; font-weight: 600;
    background: var(--c, #4285F4); color: #fff; padding: 2px 8px; border-radius: 999px;
}

div.stButton > button {
    font-family: 'Roboto', sans-serif; font-weight: 600; font-size: 13px;
    border-radius: 999px; border: none; min-height: 36px;
    background: var(--btn-bg, #4285F4); color: var(--btn-fg, #fff);
    box-shadow: 0 1px 2px rgba(60,64,67,.2);
}
div.stButton > button:hover {
    filter: brightness(1.06);
    box-shadow: 0 2px 6px rgba(60,64,67,.25);
}
div.stButton > button[kind="secondary"] {
    background: #fff; color: #3C4043; border: 1px solid #DADCE0;
}

.year-header {
    background: #4285F4; color: #fff; padding: 8px 16px; border-radius: 999px;
    font-weight: 600; font-size: 12px; margin: 16px 0 8px; display: inline-block;
}
.vtable {
    width: 100%; border-collapse: separate; border-spacing: 0;
    border-radius: 12px; overflow: hidden; margin-bottom: 16px;
    box-shadow: 0 1px 3px rgba(60,64,67,.12); border: 1px solid #E8EAED;
}
.vtable thead th {
    background: #202124; color: #fff; padding: 9px 10px; font-size: 11px;
    font-weight: 600; text-align: center; white-space: nowrap;
}
.vtable thead th:first-child { text-align: left; }
.vtable tbody td {
    padding: 7px 9px; font-size: 12px; text-align: center;
    border-bottom: 1px solid #F1F3F4; color: #3C4043; vertical-align: middle;
    background: #fff;
}
.vtable tbody td:first-child { text-align: left; font-weight: 600; }
.vtable tbody tr:nth-child(even) td { background: #FAFAFA; }
.vtable tbody tr:hover td { background: #E8F0FE; }
.venue-tag {
    display: inline-block; background: #E6F4EA; color: #137333;
    padding: 2px 8px; border-radius: 999px; font-weight: 600; font-size: 11px;
}
.rot-tag {
    display: inline-block; background: #FEF7E0; color: #B06000;
    padding: 2px 7px; border-radius: 999px; font-weight: 600; font-size: 10px; margin: 1px;
}
.subj-cell { font-size: 11px; color: #5F6368; }
.dept-banner {
    background: var(--banner-bg, #E8F0FE); padding: 14px 20px; border-radius: 14px;
    border-left: 5px solid var(--banner-c, #4285F4); margin: 8px 0 16px;
    box-shadow: 0 1px 3px rgba(60,64,67,.1);
}
.dept-banner h2 { color: var(--banner-c, #4285F4); margin: 0; font-size: 20px; font-weight: 700; }
.dept-banner p { color: #5F6368; margin: 4px 0 0; font-size: 12px; }
.muted { color: #9AA0A6; font-size: 12px; }

.section-title {
    font-size: 15px; font-weight: 700; color: #202124; margin: 8px 0 12px;
}
</style>
""",
    unsafe_allow_html=True,
)


@st.cache_data
def load_config() -> dict:
    with open(CONFIG_PATH, encoding="utf-8") as f:
        return json.load(f)


@st.cache_data
def load_catalog(_stamp: str):
    cfg = load_config()
    catalog, errors = load_all(DATA_DIR, cfg["departments"])
    occ = occupancy_rows(catalog)
    return catalog, errors, occ


def dept_map(cfg: dict) -> dict:
    return {d["id"]: d for d in cfg["departments"]}


def apply_query_params() -> None:
    qp = st.query_params
    if "dept" in qp and qp["dept"]:
        st.session_state.dept = qp["dept"]
    if "day" in qp and qp["day"] in DAY_NAMES:
        st.session_state.day = qp["day"]
    if "q" in qp and qp["q"] and "q" not in st.session_state:
        st.session_state.q = qp["q"]


def nav(dept=None, day=None, q=None) -> None:
    st.session_state.dept = dept
    st.session_state.day = day
    params = {}
    if dept:
        params["dept"] = dept
    if day:
        params["day"] = day
    query = st.session_state.get("q") if q is None else q
    if query:
        params["q"] = query
    st.query_params.from_dict(params)


def table(html_body: str) -> None:
    st.markdown(html_body, unsafe_allow_html=True)


def render_day_table(records: list[dict], day_key: str) -> None:
    years = sorted({r["year"] for r in records}, key=year_sort)
    clr, bg = DAY_COLORS[day_key], DAY_BG[day_key]
    st.markdown(
        f'<div style="background:{bg};border-left:5px solid {clr};padding:10px 16px;'
        f'border-radius:10px;margin-bottom:12px;font-weight:800;color:{clr};font-size:18px;">'
        f"{hx(DAY_NAMES[day_key])}"
        f'<span class="muted" style="margin-left:10px;font-weight:500;">Home classroom and periods</span></div>',
        unsafe_allow_html=True,
    )
    for yr in years:
        rows = [r for r in records if r["year"] == yr]
        st.markdown(
            f'<div class="year-header">{hx(yr)} — {len(rows)} sections</div>',
            unsafe_allow_html=True,
        )
        parts = [
            '<table class="vtable"><thead><tr>',
            "<th>Section</th><th>Venue</th>",
        ]
        for p in range(1, 9):
            parts.append(
                f"<th>P{p}<br><span style='font-size:10px;font-weight:400'>{PERIOD_TIMES[p]}</span></th>"
            )
        parts.append("<th>Rotation</th></tr></thead><tbody>")
        for rec in rows:
            tt = rec.get("timetable", {}).get(day_key, {})
            parts.append("<tr>")
            parts.append(f'<td>{hx(rec["section"])}</td>')
            parts.append(f'<td><span class="venue-tag">{hx(rec["venue"])}</span></td>')
            for p in range(1, 9):
                subj = tt.get(p, "")
                if not subj:
                    parts.append('<td class="muted">—</td>')
                else:
                    display = " ".join(str(subj).split())
                    if len(display) > 32:
                        display = display[:30] + "…"
                    parts.append(
                        f'<td class="subj-cell" title="{hx(subj)}">{hx(display)}</td>'
                    )
            rots = rec.get("rotation_venues") or []
            if rots:
                tags = " ".join(f'<span class="rot-tag">{hx(v)}</span>' for v in rots)
                parts.append(f"<td>{tags}</td>")
            else:
                parts.append('<td class="muted">—</td>')
            parts.append("</tr>")
        parts.append("</tbody></table>")
        table("".join(parts))


def render_summary(records: list[dict]) -> None:
    years = sorted({r["year"] for r in records}, key=year_sort)
    for yr in years:
        rows = [r for r in records if r["year"] == yr]
        st.markdown(f'<div class="year-header">{hx(yr)}</div>', unsafe_allow_html=True)
        parts = [
            '<table class="vtable"><thead><tr>',
            "<th>Section</th><th>Venue</th><th>Faculty Advisor</th><th>Rotation</th>",
            "</tr></thead><tbody>",
        ]
        for rec in rows:
            rots = rec.get("rotation_venues") or []
            rot = " ".join(f'<span class="rot-tag">{hx(v)}</span>' for v in rots) or "—"
            parts.append(
                "<tr>"
                f'<td>{hx(rec["section"])}</td>'
                f'<td><span class="venue-tag">{hx(rec["venue"])}</span></td>'
                f'<td>{hx(rec.get("fa"))}</td>'
                f"<td>{rot}</td></tr>"
            )
        parts.append("</tbody></table>")
        table("".join(parts))


def render_search(catalog: dict, occ: list[dict], query: str) -> None:
    hits = search_catalog(catalog, query)
    q = query.strip().lower()
    occ_hits = [r for r in occ if q in (r.get("venue") or "").lower() or q in (r.get("subject") or "").lower()]
    st.markdown(f"### Results for `{hx(query)}`")
    c1, c2, c3 = st.columns(3)
    c1.metric("Sections", len(hits["sections"]))
    c2.metric("Faculty", len(hits["faculty"]))
    c3.metric("Period matches", len(occ_hits))

    if hits["rooms"]:
        st.markdown("#### Rooms")
        parts = ['<table class="vtable"><thead><tr><th>Venue</th><th>Programme</th><th>Year</th><th>Section</th></tr></thead><tbody>']
        for r in hits["rooms"][:80]:
            parts.append(
                "<tr>"
                f'<td><span class="venue-tag">{hx(r["venue"])}</span></td>'
                f'<td>{hx(r["dept"])}</td><td>{hx(r["year"])}</td><td>{hx(r["section"])}</td></tr>'
            )
        parts.append("</tbody></table>")
        table("".join(parts))

    if hits["faculty"]:
        st.markdown("#### Faculty")
        parts = ['<table class="vtable"><thead><tr><th>Faculty</th><th>Programme</th><th>Year</th><th>Section</th><th>Venue</th></tr></thead><tbody>']
        for r in hits["faculty"][:80]:
            parts.append(
                "<tr>"
                f'<td>{hx(r["fa"])}</td><td>{hx(r["dept"])}</td>'
                f'<td>{hx(r["year"])}</td><td>{hx(r["section"])}</td>'
                f'<td><span class="venue-tag">{hx(r["venue"])}</span></td></tr>'
            )
        parts.append("</tbody></table>")
        table("".join(parts))

    if occ_hits:
        st.markdown("#### Who is in this room / subject (by period)")
        parts = [
            '<table class="vtable"><thead><tr>',
            "<th>Day</th><th>P</th><th>Venue</th><th>Programme</th><th>Section</th><th>Subject</th>",
            "</tr></thead><tbody>",
        ]
        order = {d: i for i, d in enumerate(DAYS)}
        occ_hits = sorted(occ_hits, key=lambda r: (order.get(r["day"], 9), r["period"], r["dept"]))
        for r in occ_hits[:200]:
            parts.append(
                "<tr>"
                f'<td>{hx(DAY_NAMES.get(r["day"], r["day"]))}</td>'
                f'<td>P{r["period"]}</td>'
                f'<td><span class="venue-tag">{hx(r["venue"])}</span></td>'
                f'<td>{hx(r["dept"])}</td><td>{hx(r["section"])}</td>'
                f'<td class="subj-cell">{hx(" ".join(str(r["subject"]).split())[:60])}</td></tr>'
            )
        parts.append("</tbody></table>")
        table("".join(parts))

    if not hits["sections"] and not hits["faculty"] and not occ_hits:
        st.info("No matches. Try a room code (e.g. ADMIN 101), a section letter, or a faculty name.")


def main() -> None:
    if "dept" not in st.session_state:
        st.session_state.dept = None
    if "day" not in st.session_state:
        st.session_state.day = None
    apply_query_params()

    cfg = load_config()
    catalog, errors, occ = load_catalog(data_stamp(ROOT, DATA_DIR))
    meta = dept_map(cfg)
    n_sec = sum(len(v) for v in catalog.values())
    n_ok = sum(1 for v in catalog.values() if v)

    st.markdown(hero_block(cfg.get("title", ""), cfg.get("subtitle", "")), unsafe_allow_html=True)
    st.markdown(
        f'<div class="stats">'
        f'<div class="stat"><b>{len(cfg["departments"])}</b> programmes</div>'
        f'<div class="stat"><b>{n_ok}</b> with data</div>'
        f'<div class="stat"><b>{n_sec}</b> sections</div>'
        f"</div>",
        unsafe_allow_html=True,
    )

    query = st.text_input(
        "Search",
        placeholder="Room (ADMIN 101), faculty, section, or programme…",
        label_visibility="collapsed",
        key="q",
    )

    if errors:
        with st.expander("Parser notes"):
            for k, msg in errors.items():
                st.caption(f"{k}: {msg}")

    if query and query.strip():
        render_search(catalog, occ, query.strip())
        return

    dept_key = st.session_state.dept
    if dept_key and dept_key not in meta:
        st.session_state.dept = None
        dept_key = None

    if dept_key is None:
        st.markdown('<p class="section-title">Select a programme</p>', unsafe_allow_html=True)
        inject_dept_button_css(meta)
        keys = [d["id"] for d in cfg["departments"]]
        for start in range(0, len(keys), 4):
            cols = st.columns(4)
            for j, dk in enumerate(keys[start : start + 4]):
                d = meta[dk]
                n = len(catalog.get(dk, []))
                with cols[j]:
                    st.markdown(folder_card(d, dk, n), unsafe_allow_html=True)
                    if st.button("Open →", key=f"d_{dk}", use_container_width=True, help=d["name"]):
                        nav(dept=dk)
                        st.rerun()
        with st.expander("How to update when a new timetable arrives"):
            st.markdown(
                """
This link is **public** — anyone can open it. No login.

1. Replace the Excel / Word / `.xls` file in the matching folder under `data/`. Any filename works; year is read from I / II / III / IV in the name.
2. Push to GitHub. Streamlit Cloud reloads this public link automatically.
3. Unusual layout? Put `venues.csv` in that department folder (copy `data/_template_venues.csv`). CSV wins over parsing.
4. New programme: add one entry in `config.json` and a folder under `data/`.
5. Offline copy of every venue: run `python export_details.py` — files go to the `details/` folder.
                """
            )
        return

    dept = meta[dept_key]
    records = catalog.get(dept_key, [])

    b1, b2 = st.columns([1, 4])
    with b1:
        if st.button("← All programmes", use_container_width=True):
            nav(dept=None, day=None)
            st.rerun()

    st.markdown(dept_banner_html(dept, dept_key, records), unsafe_allow_html=True)

    if not records:
        st.warning("No sections found. Drop new files in `data/` (or add `venues.csv`) and refresh.")
        return

    day_cols = st.columns(5)
    day_dot = {"MON": "🔵", "TUE": "🔴", "WED": "🟡", "THU": "🟢", "FRI": "🔵"}
    for i, (dk, dn) in enumerate(DAY_NAMES.items()):
        with day_cols[i]:
            active = st.session_state.day == dk
            label = f"{day_dot[dk]} {dn}"
            if st.button(label, key=f"day_{dk}", use_container_width=True, type="primary" if active else "secondary"):
                nav(dept=dept_key, day=dk)
                st.rerun()

    if st.session_state.day:
        render_day_table(records, st.session_state.day)
    else:
        st.markdown("#### Home venues (all days)")
        render_summary(records)


if __name__ == "__main__":
    main()

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
DAY_COLORS = {"MON": "#E67E22", "TUE": "#27AE60", "WED": "#2980B9", "THU": "#C0392B", "FRI": "#8E44AD"}
DAY_BG = {"MON": "#fef5ec", "TUE": "#eafaf1", "WED": "#eaf2fb", "THU": "#fdecea", "FRI": "#f5eef8"}


def hx(value) -> str:
    return html.escape(str(value or ""), quote=True)


st.markdown(
    """
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
html, body, [class*="css"] { font-family: Inter, sans-serif; }
.main .block-container { padding-top: 1.1rem; max-width: 1280px; }
#MainMenu, footer, .stDeployButton { visibility: hidden; }
header[data-testid="stHeader"] { background: transparent; }

.hero {
    background: linear-gradient(135deg, #0d2b6b 0%, #1565c0 55%, #0288d1 100%);
    color: #fff; padding: 22px 28px; border-radius: 16px; margin-bottom: 18px;
}
.hero h1 { font-size: 22px; font-weight: 800; margin: 0; letter-spacing: -0.02em; }
.hero p { margin: 6px 0 0; font-size: 13px; opacity: 0.88; }
.stats { display: flex; gap: 10px; flex-wrap: wrap; margin: 12px 0 6px; }
.stat {
    background: #fff; border: 1px solid #e3eaf3; border-radius: 10px;
    padding: 8px 14px; font-size: 12px; color: #37474f; font-weight: 600;
}
.stat b { color: #0d47a1; font-size: 15px; margin-right: 4px; }

div.stButton > button {
    font-family: Inter, sans-serif; font-weight: 700; font-size: 15px;
    border-radius: 12px; border: 1px solid #dbe4f0; min-height: 78px;
    background: #fff;
}
div.stButton > button:hover { border-color: #1565c0; background: #eef5ff; }

.year-header {
    background: #0d47a1; color: #fff; padding: 8px 16px; border-radius: 8px;
    font-weight: 700; font-size: 13px; margin: 16px 0 8px; display: inline-block;
}
.vtable {
    width: 100%; border-collapse: separate; border-spacing: 0;
    border-radius: 10px; overflow: hidden; margin-bottom: 16px;
    box-shadow: 0 1px 8px rgba(13,71,161,0.08);
}
.vtable thead th {
    background: #263238; color: #fff; padding: 9px 10px; font-size: 11px;
    font-weight: 600; text-align: center; white-space: nowrap;
}
.vtable thead th:first-child { text-align: left; }
.vtable tbody td {
    padding: 7px 9px; font-size: 12px; text-align: center;
    border-bottom: 1px solid #eceff1; color: #37474f; vertical-align: middle;
}
.vtable tbody td:first-child { text-align: left; font-weight: 650; }
.vtable tbody tr:nth-child(even) { background: #f8fbff; }
.vtable tbody tr:hover { background: #e3f2fd; }
.venue-tag {
    display: inline-block; background: #e8f5e9; color: #1b5e20;
    padding: 2px 8px; border-radius: 6px; font-weight: 700; font-size: 11px;
}
.rot-tag {
    display: inline-block; background: #fff3e0; color: #e65100;
    padding: 2px 7px; border-radius: 5px; font-weight: 600; font-size: 10px; margin: 1px;
}
.subj-cell { font-size: 11px; color: #455a64; }
.dept-banner {
    background: #eef2fb; padding: 14px 20px; border-radius: 12px;
    border-left: 5px solid #0d47a1; margin: 8px 0 16px;
}
.dept-banner h2 { color: #0d47a1; margin: 0; font-size: 20px; font-weight: 800; }
.dept-banner p { color: #546e7a; margin: 4px 0 0; font-size: 12px; }
.muted { color: #78909c; font-size: 12px; }
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

    st.markdown(
        f'<div class="hero"><h1>{hx(cfg.get("title"))}</h1>'
        f'<p>{hx(cfg.get("subtitle"))} · Public link · no login</p></div>',
        unsafe_allow_html=True,
    )
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
        keys = [d["id"] for d in cfg["departments"]]
        for start in range(0, len(keys), 4):
            cols = st.columns(4)
            for j, dk in enumerate(keys[start : start + 4]):
                d = meta[dk]
                n = len(catalog.get(dk, []))
                with cols[j]:
                    if st.button(f"{d.get('icon', '')}  {dk}", key=f"d_{dk}", use_container_width=True, help=d["name"]):
                        nav(dept=dk)
                        st.rerun()
                    extra = f" · {d['semester']}" if d.get("semester") else ""
                    st.caption(f"{d['name']} · {n} sections{extra}")
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

    st.markdown(
        f'<div class="dept-banner"><h2>{hx(dept.get("icon",""))} {hx(dept_key)} — {hx(dept["name"])}</h2>'
        f'<p>{len(records)} sections · {len({r["year"] for r in records})} years'
        f'{(" · " + hx(dept["semester"])) if dept.get("semester") else ""} · '
        f"share this view: the URL keeps the programme</p></div>",
        unsafe_allow_html=True,
    )

    if not records:
        st.warning("No sections found. Drop new files in `data/` (or add `venues.csv`) and refresh.")
        return

    day_cols = st.columns(5)
    for i, (dk, dn) in enumerate(DAY_NAMES.items()):
        with day_cols[i]:
            if st.button(dn, key=f"day_{dk}", use_container_width=True, type="primary" if st.session_state.day == dk else "secondary"):
                nav(dept=dept_key, day=dk)
                st.rerun()

    if st.session_state.day:
        render_day_table(records, st.session_state.day)
    else:
        st.markdown("#### Home venues (all days)")
        render_summary(records)


if __name__ == "__main__":
    main()

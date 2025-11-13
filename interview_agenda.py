# interview_agenda.py
import re
import json
from urllib.parse import quote
from zoneinfo import ZoneInfo
from itertools import permutations
from datetime import timedelta, datetime, timezone, time as dtime

import pandas as pd
import streamlit as st

# -------------------- App setup --------------------
st.set_page_config(page_title="Interview Scheduler", layout="wide")
st.title("📅 Interview Scheduler Tool")

st.markdown("""
### Instructions
1) Upload the CSV from Power Automate (columns: **Interviewer, StartTime, EndTime**; 15-minute increments).  
2) Set each interviewer’s duration (15, 30, 45, 60).  
3) Enter **Candidate Name** and **Job Title**.  
4) Click **Generate Agendas**, then use **Prepare invitations** beside an option.
---
""")

# -------------------- Helpers --------------------
EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")

def is_email(s: str) -> bool:
    return isinstance(s, str) and EMAIL_RE.match(s) is not None

def outlook_web_link(to_email, start_dt_local, end_dt_local, subject, body="", location=""):
    """Outlook web compose deeplink with fields pre-filled (LOCAL datetimes, no tz suffix)."""
    fmt = "%Y-%m-%dT%H:%M:%S"
    params = {
        "path": "/calendar/action/compose",
        "rru": "addevent",
        "startdt": start_dt_local.strftime(fmt),
        "enddt":   end_dt_local.strftime(fmt),
        "subject": subject,
        "body": body,
        "location": location,
        "to": to_email,
    }
    base = "https://outlook.office.com/calendar/deeplink/compose?"
    q = "&".join(f"{k}={quote(str(v))}" for k, v in params.items() if v is not None)
    return base + q

def parse_time_series(s):
    return pd.to_datetime(s, errors="coerce", infer_datetime_format=True)

# Sidebar inputs: timezone + candidate/job + lunch toggle
with st.sidebar:
    st.header("Settings")
    tz_label = st.selectbox(
        "Display / invite timezone",
        ["America/New_York", "America/Chicago", "America/Denver", "America/Los_Angeles", "UTC"],
        index=0
    )
    USER_TZ = ZoneInfo(tz_label)

    st.markdown("---")
    candidate_name = st.text_input("Candidate Name", value="Candidate Name")
    job_title      = st.text_input("Job Title", value="Job Title")

    st.markdown("---")
    avoid_lunch = st.checkbox("Avoid lunch 12–1 (allow end by 12:30 or start at 12:30)", value=True)

uploaded_file = st.file_uploader("Upload CSV from Power Automate", type=["csv"])

if uploaded_file:
    # ---------- Load & normalize ----------
    df = pd.read_csv(uploaded_file)
    df.columns = [c.strip() for c in df.columns]
    if len(df.columns) != 3:
        st.error("CSV must have exactly 3 columns: Interviewer, StartTime, EndTime")
        st.stop()

    df.columns = ["Interviewer", "StartTime", "EndTime"]
    df["Interviewer"] = df["Interviewer"].astype(str).str.strip()

    df["StartTime"] = parse_time_series(df["StartTime"])
    df["EndTime"]   = parse_time_series(df["EndTime"])

    df = df.dropna(subset=["Interviewer", "StartTime", "EndTime"]).copy()

    # floor to 15-min grid
    df["StartTime"] = df["StartTime"].dt.floor("15min")
    df["EndTime"]   = df["EndTime"].dt.floor("15min")

    fifteen = timedelta(minutes=15)
    df.loc[df["EndTime"] <= df["StartTime"], "EndTime"] = df["StartTime"] + fifteen

    # localize to chosen timezone if naive
    if df["StartTime"].dt.tz is None:
        df["StartTime"] = df["StartTime"].dt.tz_localize(USER_TZ)
        df["EndTime"]   = df["EndTime"].dt.tz_localize(USER_TZ)

    df["Date"] = df["StartTime"].dt.date

    # ---------- UI: durations ----------
    interviewers = sorted(df["Interviewer"].unique())
    st.subheader("Set Duration for Each Interviewer")
    cols = st.columns(min(4, len(interviewers)) or 1)
    durations = {}
    for i, person in enumerate(interviewers):
        durations[person] = cols[i % len(cols)].selectbox(
            f"{person}",
            [15, 30, 45, 60],
            index=1,  # default 30
            key=f"d_{person}"
        )

    max_per_day = st.slider("Maximum number of agenda options per day", 1, 10, 2)

    # ---------- Contiguity helpers ----------
    def build_blocks_map(day_frame):
        blocks = {}
        candidate_starts = set()
        for person, sub in day_frame.groupby("Interviewer"):
            sub = sub.sort_values("StartTime")
            s = set(zip(sub["StartTime"], sub["EndTime"]))
            blocks[person] = s
            candidate_starts |= set(sub["StartTime"].tolist())
        return blocks, sorted(candidate_starts)

    def has_contiguous(block_set, start_ts, minutes):
        steps = minutes // 15
        t = start_ts
        for _ in range(steps):
            if (t, t + fifteen) not in block_set:
                return False
            t += fifteen
        return True

    def find_agendas_contiguous(df_day, durations, max_per_day):
        day_agendas_total = []
        blocks_map, candidate_starts = build_blocks_map(df_day)

        # If any interviewer lacks blocks that day, skip
        for person in durations.keys():
            if person not in blocks_map or len(blocks_map[person]) == 0:
                return []

        seen_keys = set()

        for order in permutations(durations.keys()):
            if len(day_agendas_total) >= max_per_day:
                break
            for start in candidate_starts:
                if len(day_agendas_total) >= max_per_day:
                    break
                current = start
                agenda = []
                ok = True
                for person in order:
                    need = durations[person]
                    if has_contiguous(blocks_map[person], current, need):
                        agenda.append((person, current, current + timedelta(minutes=need)))
                        current = current + timedelta(minutes=need)
                    else:
                        ok = False
                        break
                if ok:
                    sig = tuple((p, s.isoformat(), e.isoformat()) for p, s, e in agenda)
                    if sig not in seen_keys:
                        seen_keys.add(sig)
                        day_agendas_total.append(agenda)
        return day_agendas_total

    def find_all_days(df, durations, max_per_day):
        agendas_all = []
        for _, day_frame in df.groupby("Date"):
            day_results = find_agendas_contiguous(day_frame, durations, max_per_day)
            agendas_all.extend(day_results)
        return agendas_all

    # ---------- Lunch filter ----------
    def agenda_respects_lunch_rule(agenda):
        """If avoid_lunch is on, keep only agendas that end <= 12:30 or start >= 12:30 local."""
        if not avoid_lunch:
            return True
        # Determine the agenda's day (use first slot start)
        first_local = agenda[0][1].astimezone(USER_TZ)
        last_local  = agenda[-1][2].astimezone(USER_TZ)
        lunch_1230  = datetime.combine(first_local.date(), dtime(12, 30), tzinfo=USER_TZ)
        # Allowed if everything ends by 12:30, or everything starts at/after 12:30
        return (last_local <= lunch_1230) or (first_local >= lunch_1230)

    # ---------- Generate ----------
    if st.button("Generate Agendas"):
        if not interviewers:
            st.error("No interviewers found in the CSV.")
            st.stop()

        agendas = find_all_days(df, durations, max_per_day)
        # apply lunch rule if needed
        agendas = [a for a in agendas if agenda_respects_lunch_rule(a)]

        if not agendas:
            msg = "No valid sequential agendas found."
            if avoid_lunch:
                msg += " Try turning off the lunch filter or adjust durations."
            st.error(msg)
        else:
            st.success(f"✅ Found {len(agendas)} possible agendas.")
            html_blocks = []

            # Defaults for invites
            location_default = "Microsoft Teams"
            subject_prefix   = f"Interview: {candidate_name} - {job_title}"

            # ------- render each option -------
            for idx, agenda in enumerate(agendas, start=1):
                # Header
                date_str = agenda[0][1].astimezone(USER_TZ).strftime("%A, %B %d, %Y")
                st.markdown(f"### Option {idx} — {date_str}")

                # Markdown table (visible)
                md = "| Interviewer | Start | End |\n|---|---:|---:|\n"
                for person, start_ts, end_ts in agenda:
                    md += f"| {person} | {start_ts.astimezone(USER_TZ).strftime('%I:%M %p')} | {end_ts.astimezone(USER_TZ).strftime('%I:%M %p')} |\n"
                st.markdown(md)

                # Agenda HTML for invite body
                rows_html = "".join(
                    f"<tr><td>{p}</td><td>{s.astimezone(USER_TZ).strftime('%I:%M %p')}</td><td>{e.astimezone(USER_TZ).strftime('%I:%M %p')}</td></tr>"
                    for p, s, e in agenda
                )
                agenda_table_html = (
                    "<p><b>Interview Agenda</b></p>"
                    "<table border='1' cellpadding='6' cellspacing='0' style='border-collapse:collapse;'>"
                    "<tr><th>Interviewer</th><th>Start</th><th>End</th></tr>"
                    f"{rows_html}"
                    "</table>"
                    "<p>(If using Outlook desktop, click the <b>Teams meeting</b> button to add the Teams link.)</p>"
                )

                # Build Outlook compose links for each interviewer in this option
                compose_links = []
                skipped = []
                for person, start_ts, end_ts in agenda:
                    if is_email(person):
                        link = outlook_web_link(
                            to_email=person,
                            start_dt_local=start_ts.astimezone(USER_TZ).replace(tzinfo=None),
                            end_dt_local=end_ts.astimezone(USER_TZ).replace(tzinfo=None),
                            subject=subject_prefix,
                            body=agenda_table_html,
                            location=location_default
                        )
                        compose_links.append(link)
                    else:
                        skipped.append(person)

                urls_json = json.dumps(compose_links)
                links_html = "".join(
                    [f'<li><a href="{u}" target="_blank" rel="noopener noreferrer">{u}</a></li>' for u in compose_links]
                ) or "<li>No links</li>"

                # Real HTML button that opens popups from a direct click (reduces blocking)
                st.components.v1.html(
                    f"""
                    <div style="margin:8px 0 4px 0">
                      <button id="prep_btn_{idx}" style="padding:8px 12px;border-radius:6px;border:1px solid #999;cursor:pointer;">
                        Prepare invitations for Option {idx}
                      </button>
                      <div id="prep_msg_{idx}" style="margin-top:6px;color:#444;"></div>
                      <details style="margin-top:6px;">
                        <summary>If nothing opens, click these links (pop-ups were blocked)</summary>
                        <ul style="margin-top:6px">{links_html}</ul>
                      </details>
                    </div>
                    <script>
                      (function() {{
                        const urls = {urls_json};
                        const btn = document.getElementById("prep_btn_{idx}");
                        const msg = document.getElementById("prep_msg_{idx}");
                        if (btn) {{
                          btn.onclick = function(e) {{
                            e.preventDefault();
                            let opened = 0;
                            urls.forEach((u, i) => {{
                              setTimeout(() => {{
                                const w = window.open(u, "_blank");
                                if (w) opened++;
                                if (i === urls.length - 1) {{
                                  msg.textContent = opened
                                    ? "Opened " + opened + " compose window(s). If some were blocked, allow pop-ups and click again."
                                    : "Pop-ups were blocked. Allow pop-ups for this site or use the links below.";
                                }}
                              }}, 60 * i);
                            }});
                          }}
                        }}
                      }})();
                    </script>
                    """,
                    height=160,
                )

                if skipped:
                    st.caption("⚠️ Skipped (not valid emails): " + ", ".join(skipped))

                # Optional visual summary below
                html_blocks.append(
                    f"""
                    <div style="margin:12px 0;">
                      <h3 style="margin:0 0 6px 0;">Option {idx} — {date_str}</h3>
                      <table border="1" cellspacing="0" cellpadding="6" style="border-collapse:collapse;">
                        <tr><th>Interviewer</th><th>Start</th><th>End</th></tr>
                        {rows_html}
                      </table>
                    </div>
                    """
                )

            # Email preview summary
            full_html = "".join(html_blocks)
            st.subheader("📧 Email Preview")
            st.markdown(full_html, unsafe_allow_html=True)

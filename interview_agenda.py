import streamlit as st
import pandas as pd
from datetime import timedelta, datetime, timezone
from itertools import permutations
from urllib.parse import quote
from zoneinfo import ZoneInfo
import re

st.set_page_config(page_title="Interview Scheduler", layout="wide")
st.title("📅 Interview Scheduler Tool")

st.markdown("""
### Instructions
Before using this tool:
1. Generate the interviewer availability CSV using your **Power Automate** flow (15-minute increments).
2. Upload that CSV below (columns: **Interviewer, StartTime, EndTime**).
3. For each interviewer, choose their interview duration (15, 30, 45, or 60 minutes).
4. Choose the **maximum agenda options per day** to generate, then click **Generate Agendas**.
---
""")

# -------------------- Helpers --------------------
EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")

def is_email(s: str) -> bool:
    return isinstance(s, str) and EMAIL_RE.match(s) is not None

def outlook_web_link(to_email, start_dt_local, end_dt_local, subject, body="", location=""):
    """
    Outlook on the web compose deeplink with fields pre-filled.
    Supply LOCAL (user timezone) datetimes formatted as YYYY-MM-DDTHH:MM:SS (no tz suffix).
    """
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

uploaded_file = st.file_uploader("Upload CSV from Power Automate", type=["csv"])

def parse_time_series(s):
    return pd.to_datetime(s, errors="coerce", infer_datetime_format=True)

# Sidebar timezone (used only for display and deeplink times)
with st.sidebar:
    st.header("Timezone")
    tz_label = st.selectbox(
        "Display / invite timezone",
        ["America/New_York", "America/Chicago", "America/Denver", "America/Los_Angeles", "UTC"],
        index=0
    )
    USER_TZ = ZoneInfo(tz_label)

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

    # ---------- UI: durations (only dropdowns) ----------
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

    # ---------- Generate ----------
    if st.button("Generate Agendas"):
        if not interviewers:
            st.error("No interviewers found in the CSV.")
            st.stop()

        agendas = find_all_days(df, durations, max_per_day)

        if not agendas:
            st.error("No valid sequential agendas found. Make sure your CSV has 15-minute rows for each free slice per interviewer.")
        else:
            st.success(f"✅ Found {len(agendas)} possible agendas.")
            html_blocks = []

            subject_prefix = "Interview"
            body_template = "Proposed interview slot generated by the Interview Scheduler."
            location_default = "Teams"

            for idx, agenda in enumerate(agendas, start=1):
                # Header
                date_str = agenda[0][1].astimezone(USER_TZ).strftime("%A, %B %d, %Y")
                st.markdown(f"### Option {idx} — {date_str}")

                # Markdown table
                md = "| Interviewer | Start | End |\n|---|---:|---:|\n"
                for person, start_ts, end_ts in agenda:
                    md += f"| {person} | {start_ts.astimezone(USER_TZ).strftime('%I:%M %p')} | {end_ts.astimezone(USER_TZ).strftime('%I:%M %p')} |\n"
                st.markdown(md)

                # Build Outlook compose links for each interviewer in this option
                compose_links = []
                skipped = []
                for person, start_ts, end_ts in agenda:
                    # Use Interviewer value as email if it looks like an email
                    if is_email(person):
                        link = outlook_web_link(
                            to_email=person,
                            start_dt_local=start_ts.astimezone(USER_TZ).replace(tzinfo=None),
                            end_dt_local=end_ts.astimezone(USER_TZ).replace(tzinfo=None),
                            subject=subject_prefix,
                            body=body_template,
                            location=location_default
                        )
                        compose_links.append(link)
                    else:
                        skipped.append(person)

                # One button to open all prefilled Outlook compose windows
                if st.button(f"Prepare invitations for Option {idx}", key=f"prep_{idx}"):
                    if compose_links:
                        urls_js_array = "[" + ",".join(f"'{quote_url}'" for quote_url in compose_links) + "]"
                        st.components.v1.html(
                            f"""
                            <script>
                              const urls = {urls_js_array};
                              // Try to open each compose window in a new tab
                              urls.forEach(u => window.open(u, "_blank"));
                              const el = document.getElementById("prep_status_{idx}");
                              if (el) el.innerText = "Opened {len(compose_links)} compose windows. If pop-ups were blocked, allow pop-ups and click again.";
                            </script>
                            <div id="prep_status_{idx}" style="font-family:system-ui,Arial;margin-top:6px;color:#444;"></div>
                            """,
                            height=40
                        )
                    else:
                        st.info("No valid interviewer emails were found in this option.")

                if skipped:
                    st.caption("⚠️ These rows didn’t look like emails and were skipped for invites: " + ", ".join(skipped))

                # For email preview block (unchanged visual summary)
                rows_html = "".join(
                    f"<tr><td>{p}</td><td>{s.astimezone(USER_TZ).strftime('%I:%M %p')}</td><td>{e.astimezone(USER_TZ).strftime('%I:%M %p')}</td></tr>"
                    for p, s, e in agenda
                )
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

            full_html = "".join(html_blocks)

            st.subheader("📧 Email Preview")
            st.markdown(full_html, unsafe_allow_html=True)

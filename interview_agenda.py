# interview_agenda.py
import re
import json
from urllib.parse import quote
from zoneinfo import ZoneInfo
from itertools import permutations
from datetime import timedelta, datetime, time as dtime

import pandas as pd
import streamlit as st

# ---------------- App setup ----------------
st.set_page_config(page_title="Interview Scheduler", layout="wide")
st.title("📅 Interview Scheduler Tool")

st.markdown("""
This app expects a single CSV with headers:

**Type, Email, Name, Title, StartTime, EndTime**

- People rows: `Type=Person`, `Email` = interviewer email, `Name`, `Title`
- Room rows:   `Type=Room`,   `Email` = room mailbox,   `Name` = room name, `Title` blank

**Flow:** Upload → set durations → Generate Agendas → click “Prepare invitations” beside an option.
---
""")

# ---------------- Helpers ----------------
EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")
def is_email(s: str) -> bool:
    return isinstance(s, str) and EMAIL_RE.match(s) is not None

def parse_dt(s):
    return pd.to_datetime(s, errors="coerce", infer_datetime_format=True)

def outlook_web_link(to_email, start_dt_local, end_dt_local, subject, body="", location=""):
    """Create Outlook web compose deeplink. Pass LOCAL (naive) datetimes."""
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

# ---------------- Sidebar ----------------
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
    avoid_lunch = st.checkbox(
        "Avoid lunch 12–1 (allow end by 12:30 or start at 12:30)",
        value=True
    )

# ---------------- CSV upload ----------------
uploaded_file = st.file_uploader(
    "Upload single CSV (Type, Email, Name, Title, StartTime, EndTime)", type=["csv"]
)

if uploaded_file:
    # --- strict schema ---
    df = pd.read_csv(uploaded_file)
    df.columns = [c.strip() for c in df.columns]
    expected = ["Type", "Email", "Name", "Title", "StartTime", "EndTime"]
    if list(df.columns) != expected:
        st.error(f"CSV headers must be exactly: {', '.join(expected)}")
        st.stop()

    # --- normalize & clean ---
    df["Type"] = df["Type"].astype(str).str.strip()
    df["Email"] = df["Email"].astype(str).str.strip()
    df["Name"]  = df["Name"].fillna("").astype(str).str.strip()
    df["Title"] = df["Title"].fillna("").astype(str).str.strip()

    df["StartTime"] = parse_dt(df["StartTime"])
    df["EndTime"]   = parse_dt(df["EndTime"])
    df = df.dropna(subset=["Email", "StartTime", "EndTime"]).copy()

    fifteen = timedelta(minutes=15)
    df["StartTime"] = df["StartTime"].dt.floor("15min")
    df["EndTime"]   = df["EndTime"].dt.floor("15min")
    df.loc[df["EndTime"] <= df["StartTime"], "EndTime"] = df["StartTime"] + fifteen

    # localize to selected timezone if naive
    if df["StartTime"].dt.tz is None:
        df["StartTime"] = df["StartTime"].dt.tz_localize(USER_TZ)
        df["EndTime"]   = df["EndTime"].dt.tz_localize(USER_TZ)

    df["Date"] = df["StartTime"].dt.date

    # split people vs rooms
    people_df = df[df["Type"].str.lower() == "person"].copy()
    rooms_df  = df[df["Type"].str.lower() == "room"].copy()

    # label for display: "Name — Title" (fallback to email if missing)
    name_map  = people_df.groupby("Email")["Name"].first().to_dict()
    title_map = people_df.groupby("Email")["Title"].first().to_dict()
    def label_for(email: str) -> str:
        nm = name_map.get(email, "")
        tt = title_map.get(email, "")
        return f"{nm} — {tt}" if nm and tt else (nm or email)

    # ------------- UI: durations -------------
    interviewers = sorted(people_df["Email"].unique())
    st.subheader("Set Duration for Each Interviewer")
    cols = st.columns(min(4, len(interviewers)) or 1)
    durations = {}
    for i, person in enumerate(interviewers):
        durations[person] = cols[i % len(cols)].selectbox(
            f"{label_for(person)}",
            [15, 30, 45, 60],
            index=1,  # default 30
            key=f"d_{person}"
        )

    max_per_day = st.slider("Maximum number of agenda options per day", 1, 10, 2)

    # ------------- People helpers -------------
    def build_blocks_map(day_frame):
        blocks = {}
        candidate_starts = set()
        for person, sub in day_frame.groupby("Email"):
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

        # ensure each interviewer has blocks that day
        for person in durations.keys():
            if person not in blocks_map or len(blocks_map[person]) == 0:
                return []

        seen = set()
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
                    if sig not in seen:
                        seen.add(sig)
                        day_agendas_total.append(agenda)
        return day_agendas_total

    def find_all_days(people_df, durations, max_per_day):
        agendas_all = []
        for _, day_frame in people_df.groupby("Date"):
            agendas_all.extend(find_agendas_contiguous(day_frame, durations, max_per_day))
        return agendas_all

    # ------------- Rooms helpers -------------
    def build_room_blocks(rooms_df_day):
        blocks = {}
        for room, sub in rooms_df_day.groupby("Email"):
            sub = sub.sort_values("StartTime")
            blocks[room] = set(zip(sub["StartTime"], sub["EndTime"]))
        return blocks

    def room_has_contiguous(blocks_set, start_ts, end_ts):
        t = start_ts
        while t < end_ts:
            if (t, t + fifteen) not in blocks_set:
                return False
            t += fifteen
        return True

    def pick_room_for_agenda(agenda):
        """Return (room_email, room_label) that covers whole agenda window, else (None, None)."""
        if rooms_df.empty:
            return (None, None)
        day = agenda[0][1].date()
        day_rooms = rooms_df[rooms_df["Date"] == day]
        if day_rooms.empty:
            return (None, None)
        start_ts = agenda[0][1]
        end_ts   = agenda[-1][2]
        room_blocks = build_room_blocks(day_rooms)
        name_map_rooms = day_rooms.groupby("Email")["Name"].first().to_dict()
        for room_email, blocks_set in room_blocks.items():
            if room_has_contiguous(blocks_set, start_ts, end_ts):
                return (room_email, name_map_rooms.get(room_email, room_email))
        return (None, None)

    # ------------- Lunch filter -------------
    def agenda_respects_lunch_rule(agenda):
        if not avoid_lunch:
            return True
        first_local = agenda[0][1].astimezone(USER_TZ)
        last_local  = agenda[-1][2].astimezone(USER_TZ)
        lunch_1230  = datetime.combine(first_local.date(), dtime(12, 30), tzinfo=USER_TZ)
        return (last_local <= lunch_1230) or (first_local >= lunch_1230)

    # ------------- Generate -------------
    if st.button("Generate Agendas"):
        if not interviewers:
            st.error("No interviewers found (Type=Person rows).")
            st.stop()

        agendas = [a for a in find_all_days(people_df, durations, max_per_day) if agenda_respects_lunch_rule(a)]

        if not agendas:
            msg = "No valid sequential agendas found."
            if avoid_lunch:
                msg += " Try turning off the lunch filter or adjust durations."
            st.error(msg)
        else:
            st.success(f"✅ Found {len(agendas)} possible agendas.")

            subject_prefix = f"Interview: {candidate_name} - {job_title}"

            for idx, agenda in enumerate(agendas, start=1):
                date_str = agenda[0][1].astimezone(USER_TZ).strftime("%A, %B %d, %Y")

                # Suggest a room that covers the whole window
                room_email, room_label = pick_room_for_agenda(agenda)
                header = f"### Option {idx} — {date_str}"
                if room_label:
                    header += f"  \n*Suggested room:* **{room_label}**"
                st.markdown(header)

                # Visible table with Name — Title
                md = "| Interviewer (Name — Title) | Start | End |\n|---|---:|---:|\n"
                for person, start_ts, end_ts in agenda:
                    display = label_for(person)
                    md += (
                        f"| {display} | "
                        f"{start_ts.astimezone(USER_TZ).strftime('%I:%M %p')} | "
                        f"{end_ts.astimezone(USER_TZ).strftime('%I:%M %p')} |\n"
                    )
                st.markdown(md)

                # Invite body HTML
                rows_html = "".join(
                    f"<tr><td>{label_for(p)}</td>"
                    f"<td>{s.astimezone(USER_TZ).strftime('%I:%M %p')}</td>"
                    f"<td>{e.astimezone(USER_TZ).strftime('%I:%M %p')}</td></tr>"
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

                # Location: Teams + room if available
                location_value = "Microsoft Teams" + (f"; Room: {room_label}" if room_label else "")

                # Build compose links (one per interviewer)
                compose_links = []
                for person, start_ts, end_ts in agenda:
                    if is_email(person):
                        link = outlook_web_link(
                            to_email=person,
                            start_dt_local=start_ts.astimezone(USER_TZ).replace(tzinfo=None),
                            end_dt_local=end_ts.astimezone(USER_TZ).replace(tzinfo=None),
                            subject=subject_prefix,
                            body=agenda_table_html,
                            location=location_value
                        )
                        compose_links.append(link)

                urls_json = json.dumps(compose_links)
                links_html = "".join(
                    [f'<li><a href="{u}" target="_blank" rel="noopener noreferrer">{u}</a></li>' for u in compose_links]
                ) or "<li>No links</li>"

                # Real HTML button to open drafts (handles popup blockers)
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

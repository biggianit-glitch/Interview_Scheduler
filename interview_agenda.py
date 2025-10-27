import streamlit as st
import pandas as pd
from datetime import timedelta
from itertools import permutations

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

uploaded_file = st.file_uploader("Upload CSV from Power Automate", type=["csv"])

def parse_time_series(s):
    # robust datetime parsing with coercion
    return pd.to_datetime(s, errors="coerce", infer_datetime_format=True)

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

    # drop invalid rows
    df = df.dropna(subset=["Interviewer", "StartTime", "EndTime"]).copy()

    # normalize to 15-minute grid to avoid second/millisecond mismatches
    df["StartTime"] = df["StartTime"].dt.floor("15min")
    # If EndTime isn't exactly 15min after StartTime, still floor to grid and keep given EndTime
    df["EndTime"] = df["EndTime"].dt.floor("15min")

    # ensure End >= Start + 15min (availability rows should be 15-minute slices)
    fifteen = timedelta(minutes=15)
    # Fix any rows where End==Start by bumping to +15
    df.loc[df["EndTime"] <= df["StartTime"], "EndTime"] = df["StartTime"] + fifteen

    df["Date"] = df["StartTime"].dt.date

    # ---------- UI: durations ----------
    interviewers = sorted(df["Interviewer"].unique())
    st.subheader("Set Duration for Each Interviewer")
    cols = st.columns(len(interviewers) if interviewers else 1)
    durations = {}
    for i, person in enumerate(interviewers):
        durations[person] = cols[i].selectbox(
            f"{person}",
            [15, 30, 45, 60],
            index=1  # default 30
        )

    max_per_day = st.slider("Maximum number of agenda options per day", 1, 10, 2)

    # ---------- Contiguity helpers ----------
    def build_blocks_map(day_frame):
        """
        For each interviewer on a given day, build a set of (start, end) 15-minute blocks.
        Also return a sorted list of candidate start times from the union of all starts.
        """
        blocks = {}
        candidate_starts = set()
        for person, sub in day_frame.groupby("Interviewer"):
            sub = sub.sort_values("StartTime")
            s = set(zip(sub["StartTime"], sub["EndTime"]))
            blocks[person] = s
            candidate_starts |= set(sub["StartTime"].tolist())
        return blocks, sorted(candidate_starts)

    def has_contiguous(block_set, start_ts, minutes):
        """Check if a person has contiguous 15-min blocks covering [start_ts, start_ts+minutes)."""
        steps = minutes // 15
        t = start_ts
        for _ in range(steps):
            if (t, t + fifteen) not in block_set:
                return False
            t += fifteen
        return True

    def find_agendas_contiguous(df_day, durations, max_per_day):
        """
        For each day:
          - Build 15-min block sets per interviewer
          - Try every interviewer order (permutations)
          - Try each candidate start time from that day
          - Require contiguous blocks per interviewer for their duration
        Stop after max_per_day agendas per day.
        """
        day_agendas_total = []
        blocks_map, candidate_starts = build_blocks_map(df_day)

        # If any interviewer lacks blocks that day, skip
        for person in durations.keys():
            if person not in blocks_map or len(blocks_map[person]) == 0:
                return []

        # Dedup guard (avoid same schedule discovered via different orders)
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
                    # signature for dedup: (date, [(person, start, end), ...])
                    sig = tuple((p, s.isoformat(), e.isoformat()) for p, s, e in agenda)
                    if sig not in seen_keys:
                        seen_keys.add(sig)
                        day_agendas_total.append(agenda)
        return day_agendas_total

    def find_all_days(df, durations, max_per_day):
        agendas_all = []
        for date_val, day_frame in df.groupby("Date"):
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

            for idx, agenda in enumerate(agendas, start=1):
                # Date header from first slot
                date_str = agenda[0][1].strftime("%A, %B %d, %Y")
                st.markdown(f"### Option {idx} — {date_str}")

                md = "| Interviewer | Start | End |\n|---|---:|---:|\n"
                for person, start_ts, end_ts in agenda:
                    md += f"| {person} | {start_ts.strftime('%I:%M %p')} | {end_ts.strftime('%I:%M %p')} |\n"
                st.markdown(md)

                rows_html = "".join(
                    f"<tr><td>{p}</td><td>{s.strftime('%I:%M %p')}</td><td>{e.strftime('%I:%M %p')}</td></tr>"
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

            # Copy button using text/html clipboard when supported; fallback to plain text
            safe_html = full_html.replace("\\", "\\\\").replace("`", "\\`")
            copy_component = f"""
            <script>
              async function copyHTML() {{
                const html = `{safe_html}`;
                try {{
                  if (window.ClipboardItem) {{
                    const data = new ClipboardItem({{'text/html': new Blob([html], {{type: 'text/html'}})}});
                    await navigator.clipboard.write([data]);
                  }} else {{
                    await navigator.clipboard.writeText(html);
                  }}
                  const el = document.getElementById("copy-status");
                  if (el) {{ el.innerText = "✅ Copied! Paste into Outlook."; }}
                }} catch (e) {{
                  const el = document.getElementById("copy-status");
                  if (el) {{ el.innerText = "❌ Copy failed. Select & copy from the preview."; }}
                }}
              }}
            </script>
            <button onclick="copyHTML()">📋 Copy HTML to Clipboard</button>
            <span id="copy-status" style="margin-left:8px;"></span>
            """
            st.components.v1.html(copy_component, height=60)

import io
from datetime import datetime, timedelta
from itertools import permutations
import pandas as pd
import streamlit as st
import pyperclip

# ---------------------- PAGE CONFIG ----------------------
st.set_page_config(page_title="Amneal Interview Agenda Builder", layout="wide")

# ---------------------- HEADER ----------------------
col1, col2 = st.columns([6, 1])
with col1:
    st.title("📅 Amneal Interview Agenda Builder")
    st.markdown("""
    ### 🧾 Instructions  
    This tool helps you generate complete interview agendas for candidates.  

    **Before using this tool:**
    1. Run the **Power Automate flow** that generates availability in 15-minute increments for each interviewer.  
    2. Export the results as a **CSV file** (formatted as: *Interviewer, StartTime, EndTime*).  
    3. Upload that CSV below.  

    Once uploaded, you can:
    - Assign interview durations for each interviewer (15, 30, 45, or 60 minutes).  
    - Generate up to a chosen number of sequential agendas **per day** (no gaps).  
    - Copy the fully formatted email HTML directly into Outlook with one click.  
    """)
with col2:
    st.image(
        "https://upload.wikimedia.org/wikipedia/en/thumb/3/3c/Amneal_Pharmaceuticals_logo.svg/512px-Amneal_Pharmaceuticals_logo.svg.png",
        width=120,
    )

# ---------------------- FUNCTIONS ----------------------
def try_parse_datetime(x):
    if pd.isna(x):
        return pd.NaT
    s = str(x).strip()
    dt = pd.to_datetime(s, errors="coerce", infer_datetime_format=True)
    return dt

def merge_slots(df):
    merged = []
    for (person, day), group in df.groupby(["Interviewer", df["StartTime"].dt.date]):
        group = group.sort_values("StartTime")
        start = group.iloc[0]["StartTime"]
        end = group.iloc[0]["EndTime"]
        for _, row in group.iloc[1:].iterrows():
            if abs((row["StartTime"] - end).total_seconds()) <= 60:
                end = row["EndTime"]
            else:
                merged.append((person, start, end))
                start, end = row["StartTime"], row["EndTime"]
        merged.append((person, start, end))
    return pd.DataFrame(merged, columns=["Interviewer", "StartTime", "EndTime"])

def find_agendas(df, durations, allow_gap=0, max_results_per_day=2):
    df = merge_slots(df)
    results = []
    debug = []

    for day, subset in df.groupby(df["StartTime"].dt.date):
        slots = {p.lower(): list(g.itertuples(index=False))
                 for p, g in subset.groupby("Interviewer")}
        names = list(durations.keys())
        day_results = []

        for order in permutations(names):
            first_name = order[0].lower()
            if first_name not in slots:
                continue

            for _, s0, e0 in slots[first_name]:
                agenda = [(order[0], s0, s0 + timedelta(minutes=durations[order[0]]))]
                if agenda[0][2] > e0:
                    continue
                valid = True
                cur_end = agenda[0][2]

                for name in order[1:]:
                    dur = durations[name]
                    if name not in slots:
                        valid = False
                        break

                    found = False
                    for _, s, e in slots[name]:
                        earliest_start = cur_end
                        latest_start = cur_end + timedelta(minutes=allow_gap)
                        if s <= latest_start <= e - timedelta(minutes=dur):
                            start = max(cur_end, s)
                            end = start + timedelta(minutes=dur)
                            if end <= e:
                                agenda.append((name, start, end))
                                cur_end = end
                                found = True
                                break
                    if not found:
                        valid = False
                        break

                if valid:
                    day_results.append((day, order, agenda))

        # Select up to N distinct agendas per day
        if day_results:
            sorted_by_start = sorted(day_results, key=lambda x: x[2][0][1])
            sorted_by_end = sorted(day_results, key=lambda x: x[2][-1][2], reverse=True)
            chosen = []

            if sorted_by_start:
                chosen.append(sorted_by_start[0])
            if max_results_per_day > 1:
                for seq in sorted_by_end:
                    if seq not in chosen:
                        chosen.append(seq)
                    if len(chosen) >= max_results_per_day:
                        break

            results.extend(chosen)

    if not results:
        debug.append("⚠️ No valid sequential agendas found.")
    return results, debug

# ---------------------- MAIN UI ----------------------
uploaded = st.file_uploader("📤 Upload CSV from Power Automate", type=["csv"])

if uploaded:
    df = pd.read_csv(uploaded)
    df.columns = [c.strip() for c in df.columns]
    interviewer_col = next((c for c in df.columns if "interviewer" in c.lower() or "email" in c.lower()), df.columns[0])
    start_col = next((c for c in df.columns if "start" in c.lower()), df.columns[1])
    end_col = next((c for c in df.columns if "end" in c.lower()), df.columns[2])

    df = df[[interviewer_col, start_col, end_col]].rename(columns={
        interviewer_col: "Interviewer",
        start_col: "StartTime",
        end_col: "EndTime"
    })

    df["Interviewer"] = df["Interviewer"].astype(str).str.strip().str.lower()
    df["StartTime"] = df["StartTime"].apply(try_parse_datetime)
    df["EndTime"] = df["EndTime"].apply(try_parse_datetime)
    df = df.dropna(subset=["StartTime", "EndTime"])

    interviewers = sorted(df["Interviewer"].unique())
    st.subheader("🧍 Set Duration for Each Interviewer")

    cols = st.columns(min(4, len(interviewers)))
    duration_options = [15, 30, 45, 60]
    durations = {}
    for i, name in enumerate(interviewers):
        with cols[i % len(cols)]:
            durations[name] = st.selectbox(f"{name}", duration_options, index=2, key=name)

    st.subheader("⚙️ Scheduling Options")
    allow_gap_toggle = st.checkbox("Allow small gaps between interviews (up to 15 minutes)", value=False)
    allow_gap = 15 if allow_gap_toggle else 0
    max_results = st.slider("Maximum number of options per day", 1, 10, 2)

    if st.button("Generate Interview Agendas", type="primary"):
        agendas, debug = find_agendas(df, durations, allow_gap=allow_gap, max_results_per_day=max_results)

        if not agendas:
            st.error("⚠️ No valid sequential agendas found.")
            if debug:
                for d in debug:
                    st.markdown(f"- {d}")
        else:
            st.success(f"Generated {len(agendas)} agenda option(s) across {len(set([a[0] for a in agendas]))} day(s).")

            # Combine all agendas into one HTML message
            full_html = """
            <div style='font-family:Arial, sans-serif;'>
                <p>Dear Candidate,</p>
                <p>We are pleased to share the available interview schedule options below. 
                Please confirm your preferred option at your earliest convenience.</p>
            """
            for idx, (day, order, agenda) in enumerate(agendas, start=1):
                full_html += f"<h3 style='color:#004aad;'>Option {idx} – {day}</h3>"
                full_html += "<table style='border-collapse:collapse;width:100%;' border='1'>"
                full_html += "<tr style='background-color:#f2f2f2;'><th>Interviewer</th><th>Start</th><th>End</th></tr>"
                for n, s, e in agenda:
                    full_html += f"<tr><td>{n}</td><td>{s.strftime('%I:%M %p')}</td><td>{e.strftime('%I:%M %p')}</td></tr>"
                full_html += "</table><br>"
            full_html += """
                <p>We look forward to connecting with you.</p>
                <p>Best regards,<br><strong>Talent Acquisition Team</strong><br>Amneal Pharmaceuticals</p>
            </div>
            """

            st.markdown("### 📧 Email Preview (All Options)")
            st.components.v1.html(full_html, height=500, scrolling=True)

            # ---- Copy-to-clipboard functionality ----
            st.write("Click below to copy the formatted email to your clipboard:")
            copy_btn = st.button("📋 Copy HTML to Clipboard")
            if copy_btn:
                try:
                    pyperclip.copy(full_html)
                    st.success("✅ HTML copied! Open a new Outlook email and press Ctrl+V to paste.")
                except Exception:
                    st.warning("⚠️ Clipboard copy not supported in this browser. Please copy manually from the preview above.")

else:
    st.info("Please upload a CSV file generated from your Power Automate flow to begin.")

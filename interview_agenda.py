import streamlit as st
import pandas as pd
from datetime import datetime, timedelta
from itertools import permutations

st.set_page_config(page_title="Interview Scheduler", layout="wide")

st.title("📅 Interview Scheduler Tool")

st.markdown("""
### Instructions
Before using this tool:
1. Generate the interviewer availability CSV file using the **Power Automate** scheduling flow.
2. Upload that CSV file below.
3. For each interviewer, select their interview duration (15, 30, 45, or 60 minutes).
4. Choose how many **maximum agenda options per day** to generate.
5. Click **Generate Agendas** to view options and copy them to Outlook.

---
""")

# --- File Upload ---
uploaded_file = st.file_uploader("Upload CSV from Power Automate", type=["csv"])

if uploaded_file:
    df = pd.read_csv(uploaded_file)
    df.columns = [c.strip() for c in df.columns]

    # Ensure consistent column naming
    expected_cols = ["Interviewer", "StartTime", "EndTime"]
    if len(df.columns) == 3:
        df.columns = expected_cols
    else:
        st.error("CSV must have exactly 3 columns: Interviewer, StartTime, EndTime")
        st.stop()

    # Parse datetime safely
    def parse_time(t):
        try:
            return pd.to_datetime(t, errors="coerce", infer_datetime_format=True)
        except Exception:
            return None

    df["StartTime"] = df["StartTime"].apply(parse_time)
    df["EndTime"] = df["EndTime"].apply(parse_time)

    df = df.dropna(subset=["StartTime", "EndTime"])
    df["Date"] = df["StartTime"].dt.date

    # --- Dynamic interviewer list ---
    interviewers = sorted(df["Interviewer"].unique())
    durations = {}

    st.subheader("Set Duration for Each Interviewer")
    cols = st.columns(len(interviewers))
    for i, interviewer in enumerate(interviewers):
        durations[interviewer] = cols[i].selectbox(
            f"{interviewer}",
            [15, 30, 45, 60],
            index=1
        )

    max_per_day = st.slider("Maximum number of agenda options per day", 1, 10, 2)

    # --- Function to find valid sequential agendas ---
    def find_agendas(df, durations, max_per_day):
        all_agendas = []
        grouped = df.groupby("Date")

        for date, group in grouped:
            day_agendas = []
            slots = {i: group[group["Interviewer"] == i].sort_values("StartTime") for i in durations.keys()}

            # Try all interviewer orders
            for order in permutations(durations.keys()):
                base = slots[order[0]]
                for _, first_row in base.iterrows():
                    start_time = first_row["StartTime"]
                    agenda = []
                    valid = True

                    current_time = start_time
                    for person in order:
                        duration = durations[person]
                        needed = timedelta(minutes=duration)

                        person_slots = slots[person]
                        next_slot = person_slots[
                            (person_slots["StartTime"] <= current_time)
                            & (person_slots["EndTime"] >= current_time + needed)
                        ]

                        if next_slot.empty:
                            valid = False
                            break
                        agenda.append((person, current_time, current_time + needed))
                        current_time += needed

                    if valid:
                        day_agendas.append(agenda)
                        if len(day_agendas) >= max_per_day:
                            break

            all_agendas.extend(day_agendas)
        return all_agendas

    # --- Generate agendas ---
    if st.button("Generate Agendas"):
        agendas = find_agendas(df, durations, max_per_day)
        if not agendas:
            st.error("No valid sequential agendas found. Check the uploaded data.")
        else:
            st.success(f"✅ Found {len(agendas)} possible agendas.")
            html_blocks = []

            for idx, agenda in enumerate(agendas, start=1):
                st.markdown(f"### Option {idx}")
                table_md = "| Interviewer | Start | End |\n|--------------|-------|-----|\n"
                for person, start, end in agenda:
                    table_md += f"| {person} | {start.strftime('%I:%M %p')} | {end.strftime('%I:%M %p')} |\n"
                st.markdown(table_md)

                html_table = f"""
                <h3>Option {idx}</h3>
                <table border="1" cellspacing="0" cellpadding="4" style="border-collapse: collapse;">
                    <tr><th>Interviewer</th><th>Start</th><th>End</th></tr>
                    {''.join([f'<tr><td>{p}</td><td>{s.strftime("%I:%M %p")}</td><td>{e.strftime("%I:%M %p")}</td></tr>' for p, s, e in agenda])}
                </table>
                <br>
                """
                html_blocks.append(html_table)

            full_html = "".join(html_blocks)

            st.subheader("📧 Email Preview")
            st.markdown(full_html, unsafe_allow_html=True)

            # --- Browser Copy Button ---
            copy_script = f"""
                <script>
                function copyToClipboard() {{
                    const html = `{full_html.replace("`", "\\`")}`;
                    navigator.clipboard.writeText(html).then(() => {{
                        alert("✅ HTML copied to clipboard. You can now paste it into Outlook!");
                    }});
                }}
                </script>
                <button onclick="copyToClipboard()">📋 Copy HTML to Clipboard</button>
            """
            st.components.v1.html(copy_script, height=100)

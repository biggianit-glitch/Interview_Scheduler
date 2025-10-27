import io
from datetime import datetime, timedelta
from itertools import permutations
import pandas as pd
import streamlit as st

st.set_page_config(page_title="Interview Agenda Builder", layout="centered")
st.title("📅 Interview Agenda Builder")

st.write("""
Upload your Power Automate CSV (15-minute available slots).  
The app detects all interviewers automatically, lets you set each duration (15, 30, 45, 60 min),  
and generates sequential interview agendas (earliest AM, latest PM) with no gaps by default.
""")

# ---------------------------------------------------------------------
def try_parse_datetime(x: str):
    if pd.isna(x):
        return pd.NaT
    s = str(x).strip()
    dt = pd.to_datetime(s, errors="coerce", infer_datetime_format=True)
    if not pd.isna(dt):
        return dt
    fmts = [
        "%m/%d/%Y %I:%M %p",
        "%B %d, %Y %I:%M %p",
        "%Y-%m-%d %H:%M:%S",
        "%m/%d/%Y %H:%M"
    ]
    for f in fmts:
        try:
            return datetime.strptime(s, f)
        except Exception:
            pass
    return pd.NaT


def merge_slots(df):
    """Combine consecutive 15-minute blocks per interviewer/day."""
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
    """
    Finds sequential (optionally small-gap) agendas.  
    Returns earliest-start and latest-end per day, up to the specified limit.
    """
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

        if day_results:
            sorted_by_start = sorted(day_results, key=lambda x: x[2][0][1])
            sorted_by_end = sorted(day_results, key=lambda x: x[2][-1][2], reverse=True)

            chosen = []
            # Add earliest AM option
            if sorted_by_start:
                chosen.append(sorted_by_start[0])
            # Add latest PM option if different
            if sorted_by_end and sorted_by_end[0] not in chosen:
                chosen.append(sorted_by_end[0])

            # Fill additional if user requested more
            if len(day_results) > 2 and max_results_per_day > 2:
                extra = sorted_by_start[1:max_results_per_day]
                for opt in extra:
                    if opt not in chosen:
                        chosen.append(opt)

            results.extend(chosen[:max_results_per_day])

    if not results:
        debug.append("⚠️ No valid sequential agendas found.")
    return results, debug


# ---------------------------------------------------------------------
uploaded = st.file_uploader("Upload CSV", type=["csv"])

if uploaded:
    df = pd.read_csv(uploaded)
    df.columns = [c.strip() for c in df.columns]

    # Auto-detect expected columns
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
    st.subheader("🧍 Set duration for each interviewer")

    cols = st.columns(min(4, len(interviewers)))
    duration_options = [15, 30, 45, 60]
    durations = {}
    for i, name in enumerate(interviewers):
        with cols[i % len(cols)]:
            durations[name] = st.selectbox(f"{name}", duration_options, index=2, key=name)

    st.subheader("⚙️ Scheduling Options")
    allow_gap_toggle = st.checkbox("Allow small gaps between interviews (up to 15 minutes)", value=False)
    allow_gap = 15 if allow_gap_toggle else 0
    max_results = st.slider("Number of agendas to generate per day", 1, 10, 2)

    if st.button("Generate Interview Agendas", type="primary"):
        agendas, debug = find_agendas(df, durations, allow_gap=allow_gap, max_results_per_day=max_results)

        if not agendas:
            st.error("⚠️ No valid sequential agendas found.")
            if debug:
                for d in debug:
                    st.markdown(f"- {d}")
        else:
            st.success(f"Found {len(agendas)} agenda option(s).")
            html_blocks = []
            for idx, (day, order, agenda) in enumerate(agendas, start=1):
                st.markdown(f"### 📅 {day} — Option {idx}")
                table = pd.DataFrame(
                    [(n, s.strftime('%I:%M %p'), e.strftime('%I:%M %p')) for n, s, e in agenda],
                    columns=["Interviewer", "Start", "End"]
                )
                st.table(table)
                html = f"<h4>{day} — Option {idx}</h4><table border='1'><tr><th>Interviewer</th><th>Start</th><th>End</th></tr>"
                for n, s, e in agenda:
                    html += f"<tr><td>{n}</td><td>{s.strftime('%I:%M %p')}</td><td>{e.strftime('%I:%M %p')}</td></tr>"
                html += "</table>"
                html_blocks.append(html)

            combined = "<br>".join(html_blocks)
            st.download_button("⬇️ Download HTML Output", combined, "agendas.html")
            st.code(combined, language="html")
else:
    st.info("Upload a CSV to begin.")

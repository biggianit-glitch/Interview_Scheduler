# interview_agenda.py

import re
import json
from urllib.parse import quote
from zoneinfo import ZoneInfo
from itertools import permutations
from datetime import timedelta, datetime, time as dtime

import pandas as pd
import streamlit as st


# ---------- App setup ----------
st.set_page_config(page_title="Interview Scheduler", layout="wide")

st.title("📅 Interview Scheduler Tool")

st.markdown("""
### Instructions
1) Upload the CSV with columns: **Interviewer, Name, Title, StartTime, EndTime**.
2) Set each interviewer’s duration: 15, 30, 45, or 60 minutes.
3) Optionally set a required interviewer order.
4) Enter **Candidate Name** and **Job Title**.
5) Click **Generate Agendas**.
6) Use **Prepare invitations** beside an option.
---
""")


# ---------- Constants ----------
EMAIL_RE = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")
FIFTEEN_MINUTES = timedelta(minutes=15)


# ---------- Helper functions ----------
def is_email(value: str) -> bool:
    return isinstance(value, str) and EMAIL_RE.match(value.strip()) is not None


def clean_text(value):
    if pd.isna(value):
        return ""
    return str(value).strip()


def parse_time_series(series: pd.Series) -> pd.Series:
    cleaned = series.astype(str).str.strip()

    parsed = pd.to_datetime(
        cleaned,
        errors="coerce"
    )

    return parsed


def outlook_web_link(to_email, start_dt_local, end_dt_local, subject, body="", location=""):
    fmt = "%Y-%m-%dT%H:%M:%S"

    params = {
        "path": "/calendar/action/compose",
        "rru": "addevent",
        "startdt": start_dt_local.strftime(fmt),
        "enddt": end_dt_local.strftime(fmt),
        "subject": subject,
        "body": body,
        "location": location,
        "to": to_email,
    }

    base = "https://outlook.office.com/calendar/deeplink/compose?"
    query_string = "&".join(
        f"{key}={quote(str(value))}"
        for key, value in params.items()
        if value is not None
    )

    return base + query_string


def safe_localize_or_convert(series: pd.Series, user_tz: ZoneInfo) -> pd.Series:
    if series.dt.tz is None:
        return series.dt.tz_localize(user_tz)

    return series.dt.tz_convert(user_tz)


def label_for_interviewer(email: str, name_map: dict, title_map: dict) -> str:
    name = clean_text(name_map.get(email, ""))
    title = clean_text(title_map.get(email, ""))

    if name and title:
        return f"{name} - {title}"

    if name:
        return name

    return email


def build_person_lookup(interviewers, name_map, title_map):
    """
    Builds a lookup so users can type either:
    - interviewer email
    - interviewer name
    - interviewer display label
    """
    lookup = {}

    for email in interviewers:
        email_clean = clean_text(email)
        name = clean_text(name_map.get(email, ""))
        title = clean_text(title_map.get(email, ""))
        label = label_for_interviewer(email, name_map, title_map)

        possible_keys = [
            email_clean,
            email_clean.lower(),
            name,
            name.lower(),
            label,
            label.lower()
        ]

        if name and title:
            possible_keys.append(f"{name} - {title}")
            possible_keys.append(f"{name} - {title}".lower())

        for key in possible_keys:
            if key:
                lookup[key] = email

    return lookup


def parse_custom_order(order_text, interviewers, person_lookup):
    """
    Parse a user-entered custom order.

    Accepts comma-separated values:
    jane.doe@company.com, john.smith@company.com

    Also accepts names if the CSV has a Name column:
    Jane Doe, John Smith
    """
    if not order_text or not order_text.strip():
        return None, None

    raw_items = [item.strip() for item in order_text.split(",") if item.strip()]

    if not raw_items:
        return None, None

    resolved_order = []
    unresolved_items = []

    for item in raw_items:
        matched_person = person_lookup.get(item) or person_lookup.get(item.lower())

        if matched_person:
            resolved_order.append(matched_person)
        else:
            unresolved_items.append(item)

    if unresolved_items:
        return None, unresolved_items

    unique_order = []
    seen = set()

    for person in resolved_order:
        if person not in seen:
            unique_order.append(person)
            seen.add(person)

    missing_people = [person for person in interviewers if person not in unique_order]

    final_order = unique_order + missing_people

    return final_order, None


def build_allowed_orders(interviewers, custom_order=None, first_person=None, last_person=None):
    """
    Determine which interviewer orders the app is allowed to test.

    Priority:
    1) If custom_order is provided, use that exact order.
    2) Otherwise, generate permutations.
    3) Apply first_person and/or last_person filters if selected.
    """
    if custom_order:
        return [tuple(custom_order)]

    possible_orders = list(permutations(interviewers))

    if first_person and first_person != "No preference":
        possible_orders = [
            order for order in possible_orders
            if order[0] == first_person
        ]

    if last_person and last_person != "No preference":
        possible_orders = [
            order for order in possible_orders
            if order[-1] == last_person
        ]

    return possible_orders


def build_blocks_map(day_frame: pd.DataFrame):
    blocks = {}
    candidate_starts = set()

    for person, sub in day_frame.groupby("Interviewer"):
        sub = sub.sort_values("StartTime")

        available_blocks = set(zip(sub["StartTime"], sub["EndTime"]))
        blocks[person] = available_blocks

        candidate_starts |= set(sub["StartTime"].tolist())

    return blocks, sorted(candidate_starts)


def has_contiguous_availability(block_set, start_ts, minutes):
    steps = minutes // 15
    current = start_ts

    for _ in range(steps):
        next_time = current + FIFTEEN_MINUTES

        if (current, next_time) not in block_set:
            return False

        current = next_time

    return True


def find_agendas_contiguous(
    df_day: pd.DataFrame,
    durations: dict,
    max_per_day: int,
    allowed_orders
):
    day_agendas_total = []
    blocks_map, candidate_starts = build_blocks_map(df_day)

    for person in durations.keys():
        if person not in blocks_map or len(blocks_map[person]) == 0:
            return []

    seen_keys = set()

    for order in allowed_orders:
        if len(day_agendas_total) >= max_per_day:
            break

        for start in candidate_starts:
            if len(day_agendas_total) >= max_per_day:
                break

            current = start
            agenda = []
            agenda_is_valid = True

            for person in order:
                required_minutes = durations[person]

                if has_contiguous_availability(
                    blocks_map[person],
                    current,
                    required_minutes
                ):
                    agenda.append(
                        (
                            person,
                            current,
                            current + timedelta(minutes=required_minutes)
                        )
                    )
                    current = current + timedelta(minutes=required_minutes)
                else:
                    agenda_is_valid = False
                    break

            if agenda_is_valid:
                signature = tuple(
                    (person, start_time.isoformat(), end_time.isoformat())
                    for person, start_time, end_time in agenda
                )

                if signature not in seen_keys:
                    seen_keys.add(signature)
                    day_agendas_total.append(agenda)

    return day_agendas_total


def find_all_days(df: pd.DataFrame, durations: dict, max_per_day: int, allowed_orders):
    agendas_all = []

    for _, day_frame in df.groupby("Date"):
        agendas_all.extend(
            find_agendas_contiguous(
                day_frame,
                durations,
                max_per_day,
                allowed_orders
            )
        )

    return agendas_all


def agenda_respects_lunch_rule(agenda, avoid_lunch: bool, user_tz: ZoneInfo):
    if not avoid_lunch:
        return True

    first_local = agenda[0][1].astimezone(user_tz)
    last_local = agenda[-1][2].astimezone(user_tz)

    lunch_cutoff = datetime.combine(
        first_local.date(),
        dtime(12, 30),
        tzinfo=user_tz
    )

    return last_local <= lunch_cutoff or first_local >= lunch_cutoff


# ---------- Sidebar ----------
with st.sidebar:
    st.header("Settings")

    tz_label = st.selectbox(
        "Display / invite timezone",
        [
            "America/New_York",
            "America/Chicago",
            "America/Denver",
            "America/Los_Angeles",
            "UTC"
        ],
        index=0
    )

    USER_TZ = ZoneInfo(tz_label)

    st.markdown("---")

    candidate_name = st.text_input("Candidate Name", value="Candidate Name")
    job_title = st.text_input("Job Title", value="Job Title")

    st.markdown("---")

    avoid_lunch = st.checkbox(
        "Avoid lunch 12-1",
        value=True,
        help="If checked, agendas must end by 12:30 PM or start at/after 12:30 PM."
    )


# ---------- File upload ----------
uploaded_file = st.file_uploader("Upload CSV", type=["csv"])


if uploaded_file:
    # ---------- Load CSV ----------
    try:
        df = pd.read_csv(uploaded_file)
    except Exception as e:
        st.error("The CSV could not be read. Please check the file format.")
        st.exception(e)
        st.stop()

    df.columns = [str(column).strip() for column in df.columns]

    # ---------- Required columns ----------
    required_columns = {"Interviewer", "StartTime", "EndTime"}
    missing_columns = required_columns - set(df.columns)

    if missing_columns:
        st.error(
            "CSV is missing required column(s): "
            + ", ".join(sorted(missing_columns))
        )
        st.stop()

    # ---------- Optional columns ----------
    has_name = "Name" in df.columns
    has_title = "Title" in df.columns

    if not has_name:
        df["Name"] = ""

    if not has_title:
        df["Title"] = ""

    # ---------- Clean core columns ----------
    df["Interviewer"] = df["Interviewer"].apply(clean_text)
    df["Name"] = df["Name"].apply(clean_text)
    df["Title"] = df["Title"].apply(clean_text)

    # ---------- Parse dates ----------
    df["StartTime"] = parse_time_series(df["StartTime"])
    df["EndTime"] = parse_time_series(df["EndTime"])

    bad_rows = df[
        df["StartTime"].isna()
        | df["EndTime"].isna()
        | (df["Interviewer"] == "")
    ].copy()

    if not bad_rows.empty:
        st.error(
            "Some rows could not be used because StartTime, EndTime, or Interviewer "
            "was blank or could not be parsed. Review the rows below."
        )
        st.dataframe(bad_rows)
        st.stop()

    # ---------- Normalize to 15-minute grid ----------
    df["StartTime"] = df["StartTime"].dt.floor("15min")
    df["EndTime"] = df["EndTime"].dt.floor("15min")

    df.loc[
        df["EndTime"] <= df["StartTime"],
        "EndTime"
    ] = df["StartTime"] + FIFTEEN_MINUTES

    # ---------- Timezone handling ----------
    try:
        df["StartTime"] = safe_localize_or_convert(df["StartTime"], USER_TZ)
        df["EndTime"] = safe_localize_or_convert(df["EndTime"], USER_TZ)
    except Exception as e:
        st.error(
            "There was a timezone conversion issue. This usually happens when "
            "the CSV mixes timezone-aware and timezone-naive date values."
        )
        st.exception(e)
        st.stop()

    # ---------- Add Date ----------
    df["Date"] = df["StartTime"].dt.date

    # ---------- Build display label maps ----------
    name_map = df.groupby("Interviewer")["Name"].first().to_dict()
    title_map = df.groupby("Interviewer")["Title"].first().to_dict()

    def label_for(email: str) -> str:
        return label_for_interviewer(email, name_map, title_map)

    # ---------- Show upload summary ----------
    st.success("CSV loaded successfully.")

    with st.expander("Preview uploaded availability"):
        st.dataframe(df)

    # ---------- Interviewers ----------
    interviewers = sorted(df["Interviewer"].unique())

    if not interviewers:
        st.error("No interviewers found in the CSV.")
        st.stop()

    person_lookup = build_person_lookup(interviewers, name_map, title_map)

    # ---------- UI: durations ----------
    st.subheader("Set Duration for Each Interviewer")

    cols = st.columns(min(4, len(interviewers)) or 1)

    durations = {}

    for i, person in enumerate(interviewers):
        durations[person] = cols[i % len(cols)].selectbox(
            label_for(person),
            [15, 30, 45, 60],
            index=1,
            key=f"d_{person}"
        )

    max_per_day = st.slider(
        "Maximum number of agenda options per day",
        min_value=1,
        max_value=10,
        value=2
    )

    # ---------- UI: interviewer order ----------
    st.subheader("Optional Interviewer Order Rules")

    interviewer_display_options = ["No preference"] + [
        label_for(person) for person in interviewers
    ]

    display_to_email = {
        label_for(person): person
        for person in interviewers
    }

    order_cols = st.columns(2)

    first_choice_display = order_cols[0].selectbox(
        "Force this interviewer to go first",
        interviewer_display_options,
        index=0
    )

    last_choice_display = order_cols[1].selectbox(
        "Force this interviewer to go last",
        interviewer_display_options,
        index=0
    )

    first_person = (
        display_to_email.get(first_choice_display)
        if first_choice_display != "No preference"
        else "No preference"
    )

    last_person = (
        display_to_email.get(last_choice_display)
        if last_choice_display != "No preference"
        else "No preference"
    )

    custom_order_text = st.text_input(
        "Optional exact interviewer order",
        value="",
        placeholder="Example: Jane Smith, John Doe, Alex Brown",
        help=(
            "Use comma-separated names or email addresses. "
            "If this is filled in, it overrides the first/last dropdowns."
        )
    )

    if custom_order_text.strip():
        st.caption(
            "Custom order is active. The app will use this exact order first, "
            "then append any missing interviewers at the end."
        )
    else:
        st.caption(
            "Leave custom order blank if you only want to force one person first or last."
        )

    # ---------- Generate agendas ----------
    if st.button("Generate Agendas"):
        custom_order, unresolved_items = parse_custom_order(
            custom_order_text,
            interviewers,
            person_lookup
        )

        if unresolved_items:
            st.error(
                "The following custom order entries could not be matched to an interviewer: "
                + ", ".join(unresolved_items)
            )
            st.stop()

        if (
            first_person != "No preference"
            and last_person != "No preference"
            and first_person == last_person
            and not custom_order
        ):
            st.error("The same interviewer cannot be forced both first and last.")
            st.stop()

        allowed_orders = build_allowed_orders(
            interviewers=interviewers,
            custom_order=custom_order,
            first_person=first_person,
            last_person=last_person
        )

        if not allowed_orders:
            st.error("No valid interviewer order could be generated from the selected rules.")
            st.stop()

        agendas = find_all_days(
            df,
            durations,
            max_per_day,
            allowed_orders
        )

        agendas = [
            agenda
            for agenda in agendas
            if agenda_respects_lunch_rule(agenda, avoid_lunch, USER_TZ)
        ]

        if not agendas:
            message = "No valid sequential agendas found."

            if custom_order:
                message += " The custom order may not fit the available time blocks."

            elif first_person != "No preference" or last_person != "No preference":
                message += " The first/last interviewer rule may be too restrictive."

            if avoid_lunch:
                message += " You can also try turning off the lunch filter or adjusting durations."

            st.error(message)
            st.stop()

        st.success(f"✅ Found {len(agendas)} possible agenda option(s).")

        location_default = "Microsoft Teams"
        subject_prefix = f"Interview: {candidate_name} - {job_title}"

        for idx, agenda in enumerate(agendas, start=1):
            date_str = agenda[0][1].astimezone(USER_TZ).strftime(
                "%A, %B %d, %Y"
            )

            st.markdown(f"### Option {idx}: {date_str}")

            visible_rows = []

            for person, start_ts, end_ts in agenda:
                visible_rows.append(
                    {
                        "Interviewer": label_for(person),
                        "Start": start_ts.astimezone(USER_TZ).strftime("%I:%M %p"),
                        "End": end_ts.astimezone(USER_TZ).strftime("%I:%M %p")
                    }
                )

            st.table(pd.DataFrame(visible_rows))

            rows_html = ""

            for person, start_ts, end_ts in agenda:
                rows_html += (
                    "<tr>"
                    f"<td>{label_for(person)}</td>"
                    f"<td>{start_ts.astimezone(USER_TZ).strftime('%I:%M %p')}</td>"
                    f"<td>{end_ts.astimezone(USER_TZ).strftime('%I:%M %p')}</td>"
                    "</tr>"
                )

            agenda_table_html = (
                "<p><b>Interview Agenda</b></p>"
                "<table border='1' cellpadding='6' cellspacing='0' "
                "style='border-collapse:collapse;'>"
                "<tr><th>Interviewer</th><th>Start</th><th>End</th></tr>"
                f"{rows_html}"
                "</table>"
                "<p>If using Outlook desktop, click the <b>Teams meeting</b> "
                "button to add the Teams link.</p>"
            )

            compose_links = []

            for person, start_ts, end_ts in agenda:
                if is_email(person):
                    start_local_naive = (
                        start_ts
                        .astimezone(USER_TZ)
                        .replace(tzinfo=None)
                    )

                    end_local_naive = (
                        end_ts
                        .astimezone(USER_TZ)
                        .replace(tzinfo=None)
                    )

                    link = outlook_web_link(
                        to_email=person,
                        start_dt_local=start_local_naive,
                        end_dt_local=end_local_naive,
                        subject=subject_prefix,
                        body=agenda_table_html,
                        location=location_default
                    )

                    compose_links.append(link)

            urls_json = json.dumps(compose_links)

            links_html = "".join(
                [
                    f'<li><a href="{url}" target="_blank" '
                    f'rel="noopener noreferrer">{url}</a></li>'
                    for url in compose_links
                ]
            )

            if not links_html:
                links_html = "<li>No invitation links were created because no interviewer values were valid email addresses.</li>"

            st.components.v1.html(
                f"""
                <div style="margin:8px 0 4px 0">
                  <button
                    id="prep_btn_{idx}"
                    style="
                      padding:8px 12px;
                      border-radius:6px;
                      border:1px solid #999;
                      cursor:pointer;
                      background:#f7f7f7;
                    "
                  >
                    Prepare invitations for Option {idx}
                  </button>

                  <div
                    id="prep_msg_{idx}"
                    style="margin-top:6px;color:#444;"
                  ></div>

                  <details style="margin-top:6px;">
                    <summary>
                      If nothing opens, click these links. Pop-ups may have been blocked.
                    </summary>
                    <ul style="margin-top:6px">
                      {links_html}
                    </ul>
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

                        if (!urls || urls.length === 0) {{
                          msg.textContent = "No invitation links were available for this option.";
                          return;
                        }}

                        urls.forEach((url, i) => {{
                          setTimeout(() => {{
                            const w = window.open(url, "_blank");

                            if (w) {{
                              opened++;
                            }}

                            if (i === urls.length - 1) {{
                              msg.textContent = opened
                                ? "Opened " + opened + " compose window(s). If some were blocked, allow pop-ups and click again."
                                : "Pop-ups were blocked. Allow pop-ups for this site or use the links below.";
                            }}
                          }}, 75 * i);
                        }});
                      }};
                    }}
                  }})();
                </script>
                """,
                height=170,
            )

else:
    st.info("Upload a CSV file to begin.")

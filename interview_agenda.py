import json  # add near the top of your file if not present

# ... inside your loop over each agenda/option ...
compose_links = []
skipped = []
for person, start_ts, end_ts in agenda:
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

# ---------- Render a real HTML button that opens popups from a direct click ----------
urls_json = json.dumps(compose_links)  # safe for JS
links_html = "".join([f'<li><a href="{u}" target="_blank" rel="noopener noreferrer">{u}</a></li>' for u in compose_links]) or "<li>No links</li>"

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
            // open sequentially with a tiny delay to reduce popup blocking
            urls.forEach((u, i) => {{
              setTimeout(() => {{
                const w = window.open(u, "_blank");
                if (w) opened++;
                if (i === urls.length - 1) {{
                  msg.textContent = opened
                    ? `Opened ${opened} compose window(s). If some were blocked, allow pop-ups and click again.`
                    : "Pop-ups were blocked. Allow pop-ups for this site or use the links below.";
                }}
              }}, 60 * i);
            }});
          }}
        }}
      }})();
    </script>
    """,
    height=120,
)
if skipped:
    st.caption("⚠️ Skipped (not valid emails): " + ", ".join(skipped))

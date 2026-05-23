"""
onboarding.py
-------------
Handles the Setup section for custom user configuration.
Users can upload their own profile.json, resume (PDF), Excel tracker,
and provide their Gmail token to override the default (Oscar's) session.
"""

import json
import io
import streamlit as st


def render_setup_page():
    st.markdown("## ⚙️ Setup")
    st.markdown(
        "By default, this app runs with the owner's profile and data. "
        "Use this page if you want to use the app with **your own** resume, "
        "job tracker, and Gmail."
    )

    # ── Current session status ──────────────────────────────
    st.markdown("---")
    _render_session_status()

    # ── Upload sections ─────────────────────────────────────
    st.markdown("---")
    _render_profile_upload()

    st.markdown("---")
    _render_excel_upload()

    st.markdown("---")
    _render_resume_upload()

    st.markdown("---")
    _render_gmail_token_upload()

    # ── Reset ───────────────────────────────────────────────
    st.markdown("---")
    if st.button("🔄 Reset to default (owner profile)", use_container_width=False):
        for key in ["custom_profile", "custom_excel_bytes", "custom_resume_bytes", "custom_gmail_token"]:
            st.session_state.pop(key, None)
        st.session_state["using_custom_profile"] = False
        st.success("Reset to default profile.")
        st.rerun()


# ── Session status banner ───────────────────────────────────

def _render_session_status():
    using_custom = st.session_state.get("using_custom_profile", False)

    if using_custom:
        name = st.session_state.get("custom_profile", {}).get("name", "Custom user")
        st.success(f"✅ Currently using **{name}'s** custom profile.")
        cols = st.columns(4)
        cols[0].markdown("Profile JSON " + ("✅" if "custom_profile" in st.session_state else "❌"))
        cols[1].markdown("Excel tracker " + ("✅" if "custom_excel_bytes" in st.session_state else "❌"))
        cols[2].markdown("Resume PDF " + ("✅" if "custom_resume_bytes" in st.session_state else "❌"))
        cols[3].markdown("Gmail token " + ("✅" if "custom_gmail_token" in st.session_state else "❌"))
    else:
        st.info("ℹ️ Using **default (owner)** profile. Upload your files below to switch.")


# ── Profile JSON ────────────────────────────────────────────

def _render_profile_upload():
    st.markdown("### 1 · Profile JSON")
    st.markdown(
        "Upload your `profile.json`. "
        "Don't have one? Use the prompt below to generate it with any LLM."
    )

    with st.expander("📋 Prompt to generate your profile.json"):
        st.code(_PROFILE_GENERATION_PROMPT, language="markdown")

    uploaded = st.file_uploader(
        "Upload profile.json",
        type=["json"],
        key="uploader_profile"
    )

    if uploaded:
        try:
            profile = json.load(uploaded)
            required = ["name", "summary", "skills", "experience", "education"]
            missing = [k for k in required if k not in profile]
            if missing:
                st.warning(f"⚠️ Missing fields in profile.json: {', '.join(missing)}. "
                           "The app may not work as expected.")
            st.session_state["custom_profile"] = profile
            st.session_state["using_custom_profile"] = True
            st.success(f"✅ Loaded profile for **{profile.get('name', 'Unknown')}**.")
        except json.JSONDecodeError:
            st.error("❌ Invalid JSON file. Please check the format.")


# ── Excel tracker ───────────────────────────────────────────

def _render_excel_upload():
    st.markdown("### 2 · Job Tracker Excel")
    st.markdown(
        "Upload your job tracker `.xlsx` file. "
        "It must have these columns: **Company, Position, Applied on, Status, JD, Application Link**."
    )

    uploaded = st.file_uploader(
        "Upload tracker Excel",
        type=["xlsx"],
        key="uploader_excel"
    )

    if uploaded:
        st.session_state["custom_excel_bytes"] = uploaded.read()
        st.success("✅ Excel tracker loaded.")

    st.caption(
        "Don't have one? Download the "
        "[template](https://github.com/oscar10408/Langgraph-Multi-Agent-Job-Tracker) "
        "from the repo and fill it in."
    )


# ── Resume PDF ──────────────────────────────────────────────

def _render_resume_upload():
    st.markdown("### 3 · Resume PDF *(optional)*")
    st.markdown("Used by the Cover Letter and Job Match agents to personalise output.")

    uploaded = st.file_uploader(
        "Upload resume PDF",
        type=["pdf"],
        key="uploader_resume"
    )

    if uploaded:
        st.session_state["custom_resume_bytes"] = uploaded.read()
        st.success("✅ Resume loaded.")


# ── Gmail token ─────────────────────────────────────────────

def _render_gmail_token_upload():
    st.markdown("### 4 · Gmail Token *(optional)*")
    st.markdown(
        "Required only for the **Email Scanner** feature. "
        "Paste the contents of your `token.json` below."
    )

    with st.expander("How to get your token.json"):
        st.markdown("""
1. Go to [Google Cloud Console](https://console.cloud.google.com/) and create a project.
2. Enable the **Gmail API**.
3. Create OAuth 2.0 credentials (Desktop app type) and download `credentials.json`.
4. Run this locally once to authorise and generate `token.json`:

```python
from google_auth_oauthlib.flow import InstalledAppFlow
import json

SCOPES = ["https://www.googleapis.com/auth/gmail.readonly"]
flow = InstalledAppFlow.from_client_secrets_file("credentials.json", SCOPES)
creds = flow.run_local_server(port=0)
print(creds.to_json())  # paste this output below
```

5. Copy the printed JSON and paste it into the box below.
        """)

    token_input = st.text_area(
        "Paste token.json contents here",
        height=150,
        placeholder='{"token": "ya29.xxx", "refresh_token": "1//xxx", ...}',
        key="input_gmail_token"
    )

    if st.button("💾 Save Gmail token", key="btn_save_gmail"):
        if token_input.strip():
            try:
                token_dict = json.loads(token_input.strip())
                required = ["token", "refresh_token", "client_id", "client_secret"]
                missing = [k for k in required if k not in token_dict]
                if missing:
                    st.warning(f"⚠️ Token may be incomplete. Missing: {', '.join(missing)}")
                st.session_state["custom_gmail_token"] = token_dict
                st.success("✅ Gmail token saved for this session.")
            except json.JSONDecodeError:
                st.error("❌ Invalid JSON. Please paste the raw token.json content.")
        else:
            st.warning("Please paste your token.json content first.")


# ── Profile generation prompt ───────────────────────────────

_PROFILE_GENERATION_PROMPT = """\
Please generate a profile.json file based on the information I provide below.
Output ONLY valid JSON, no explanation, no markdown code fences.

[Basic Info]
Name:
Current city:
Email:
LinkedIn URL (optional):
GitHub URL (optional):
Personal website (optional):

[Education] (one per line: School | Degree | Major | Graduation Year)
1.
2.

[Work / Internship Experience] (one per line: Company | Title | Start (YYYY-MM) | End (YYYY-MM or present) | One-line summary)
1.
2.

[Skills] (comma-separated, be as detailed as possible)
Programming languages:
ML / AI frameworks:
Databases:
Cloud / tools:
Other:

[Projects] (one per line: Project name | Technologies used | One-line description)
1.
2.

[Publications] (optional, one per line: Title | Venue | Year)
1.

[Job Search Goals]
Target role types (e.g. Data Scientist, ML Engineer):
Target industries (e.g. Tech, Finance):
Preferred locations:
Status (actively looking / open to opportunities):

[Summary] (2-3 sentences describing your background and strengths):

---

Output format (follow exactly):
{
  "name": "Jane Doe",
  "location": "San Francisco, CA",
  "email": "jane@example.com",
  "linkedin": "https://linkedin.com/in/janedoe",
  "github": "https://github.com/janedoe",
  "website": "",
  "summary": "...",
  "target_roles": ["Data Scientist"],
  "target_industries": ["Tech"],
  "target_locations": ["California", "Remote"],
  "job_search_status": "actively looking",
  "education": [
    {
      "school": "UC Berkeley",
      "degree": "Master of Science",
      "major": "Data Science",
      "graduation_year": 2025
    }
  ],
  "experience": [
    {
      "company": "Acme Corp",
      "title": "Data Analyst Intern",
      "start": "2024-06",
      "end": "2024-08",
      "summary": "Built ETL pipelines and dashboards for marketing analytics."
    }
  ],
  "skills": {
    "languages": ["Python", "SQL"],
    "ml_frameworks": ["PyTorch", "Scikit-learn"],
    "databases": ["MySQL", "PostgreSQL"],
    "cloud_tools": ["AWS", "Tableau"],
    "other": ["Git", "Linux"]
  },
  "projects": [
    {
      "name": "Customer Churn Prediction",
      "tech": ["XGBoost", "Streamlit"],
      "description": "End-to-end churn prediction pipeline with Streamlit dashboard."
    }
  ],
  "publications": []
}
"""

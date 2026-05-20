"""
app.py
------
Streamlit Web UI for Job Application Tracker Agent.
Phase 1: Dashboard + Applications Table
Phase 2: AI Tools (Cover Letter, Interview Prep, Job Match)
Phase 3: Chat bar (Agent integration)
"""
from dotenv import load_dotenv
load_dotenv()

import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from openpyxl import load_workbook
from datetime import date
import os
import sys

def install_playwright_browser():
    try:
        subprocess.run(
            [sys.executable, "-m", "playwright", "install", "chromium"],
            check=True
        )
    except Exception as e:
        print(f"[Playwright Install] Failed: {e}", flush=True)

install_playwright_browser()


if "GROQ_API_KEY" in st.secrets:
    os.environ["GROQ_API_KEY"] = st.secrets["GROQ_API_KEY"]
    
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

# ── Page config ────────────────────────────────────────────
st.set_page_config(
    page_title="Job Tracker",
    page_icon="💼",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ── Custom CSS ─────────────────────────────────────────────
st.markdown("""
<style>
    /* Hide Streamlit default elements */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}

    /* Main background */
    .stApp { background-color: #f8f7f4; }

    /* Sidebar */
    [data-testid="stSidebar"] {
        background-color: #ffffff;
        border-right: 1px solid #e8e6e0;
    }

    /* Metric cards */
    [data-testid="stMetric"] {
        background: #ffffff;
        border: 1px solid #e8e6e0;
        border-radius: 10px;
        padding: 1rem 1.25rem;
    }

    /* Buttons */
    .stButton > button {
        background: #1a1a1a;
        color: white;
        border: none;
        border-radius: 8px;
        font-weight: 500;
        padding: 0.5rem 1.25rem;
    }
    .stButton > button:hover {
        background: #333333;
        color: white;
    }

    /* Status badges */
    .badge {
        display: inline-block;
        padding: 3px 10px;
        border-radius: 20px;
        font-size: 12px;
        font-weight: 500;
    }
    .badge-interviewed { background: #dbeafe; color: #1d4ed8; }
    .badge-applied { background: #f3f4f6; color: #4b5563; }
    .badge-offered { background: #dcfce7; color: #166534; }
    .badge-rejected { background: #fee2e2; color: #991b1b; }
    .badge-unknown { background: #f3f4f6; color: #6b7280; }
    .badge-wishlist { background: #fef9c3; color: #854d0e; }
    .badge-followup { background: #f3e8ff; color: #7e22ce; }

    /* Section headers */
    .section-header {
        font-size: 16px;
        font-weight: 600;
        color: #1a1a1a;
        margin-bottom: 0.75rem;
        margin-top: 1.5rem;
    }

    /* AI tool cards */
    .ai-card {
        background: white;
        border: 1px solid #e8e6e0;
        border-radius: 12px;
        padding: 1.25rem;
        cursor: pointer;
        transition: border-color 0.15s;
    }
    .ai-card:hover { border-color: #1a1a1a; }
    .ai-card-icon { font-size: 24px; margin-bottom: 8px; }
    .ai-card-title { font-size: 14px; font-weight: 600; color: #1a1a1a; }
    .ai-card-desc { font-size: 12px; color: #6b7280; margin-top: 4px; }

    /* Chat messages */
    .chat-msg-user {
        background: #1a1a1a;
        color: white;
        padding: 10px 14px;
        border-radius: 12px 12px 2px 12px;
        margin: 8px 0;
        margin-left: 20%;
        font-size: 14px;
    }
    .chat-msg-agent {
        background: white;
        border: 1px solid #e8e6e0;
        color: #1a1a1a;
        padding: 10px 14px;
        border-radius: 12px 12px 12px 2px;
        margin: 8px 0;
        margin-right: 20%;
        font-size: 14px;
    }

    /* Dataframe styling */
    .stDataFrame { border-radius: 10px; overflow: hidden; }

    /* Tab styling */
    .stTabs [data-baseweb="tab-list"] {
        gap: 8px;
        background: transparent;
    }
    .stTabs [data-baseweb="tab"] {
        background: white;
        border: 1px solid #e8e6e0;
        border-radius: 8px;
        color: #6b7280;
        font-size: 13px;
    }
    .stTabs [aria-selected="true"] {
        background: #1a1a1a !important;
        color: white !important;
        border-color: #1a1a1a !important;
    }
</style>
""", unsafe_allow_html=True)


# ── Data loading ───────────────────────────────────────────
EXCEL_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data", "Record_of_job_AI.xlsx")

@st.cache_data(ttl=30)
def load_data():
    """Load Excel data into DataFrame."""
    def normalize_status(s):
        s = str(s).lower().strip()
        if "follow" in s or "questionaire" in s or "questionare" in s:
            return "follow-up q"
        if "interview" in s:
            return "interviewed"
        if "record" in s or "phone screen" in s or "screening" in s:
            return "interviewed"
        return s

    try:
        df = pd.read_excel(EXCEL_PATH)
        df.columns = df.columns.str.strip()
        df = df[df["Company"].notna() & (df["Company"] != "")]
        df["Status"] = df["Status"].fillna("applied").str.lower().str.strip()
        df["Status"] = df["Status"].apply(normalize_status)

        status_map = {
            "unknown": "applied",
            "unk": "applied",
        }
        df["Status"] = df["Status"].map(lambda x: status_map.get(x, x))
        df["Applied on"] = df["Applied on"].fillna("Unknown")
        return df
    except Exception as e:
        st.error(f"Error loading data: {e}")
        return pd.DataFrame()


def get_status_counts(df):
    counts = df["Status"].value_counts().to_dict()
    return {
        "total": len(df),
        "interviewed": counts.get("interviewed", 0) + counts.get("1st interview, 2nd interview", 0),
        "offered": counts.get("offered", 0),
        "rejected": counts.get("rejected", 0),
        "applied": counts.get("applied", 0),
        "follow_up": sum(v for k, v in counts.items() if "follow" in k or "questionaire" in k),
    }


def status_badge(status):
    s = str(status).lower().strip()
    if "interview" in s:
        return f'<span class="badge badge-interviewed">Interviewed</span>'
    elif s == "offered":
        return f'<span class="badge badge-offered">Offered</span>'
    elif s == "rejected":
        return f'<span class="badge badge-rejected">Rejected</span>'
    elif s == "applied":
        return f'<span class="badge badge-applied">Applied</span>'
    elif "follow" in s or "questionaire" in s or "questionare" in s:
        return f'<span class="badge badge-followup">Follow-up Q</span>'
    elif s == "wishlist":
        return f'<span class="badge badge-wishlist">Wishlist</span>'
    else:
        return f'<span class="badge badge-unknown">{status}</span>'


# ── Sidebar ────────────────────────────────────────────────
with st.sidebar:
    st.markdown("### 💼 Job Tracker")
    st.markdown("**Oscar Shih** · Ann Arbor, MI")
    st.divider()

    page = st.radio(
        "Navigation",
        ["📊 Dashboard", "📋 Applications", "📧 Email Scanner", "📝 Cover Letter", "🎯 Interview Prep", "🔍 Job Match", "💬 Chat"],
        label_visibility="collapsed"
    )

    st.divider()

    if st.button("🔄 Refresh data", use_container_width=True):
        st.cache_data.clear()
        st.rerun()


# ── Load data ──────────────────────────────────────────────
df = load_data()
stats = get_status_counts(df) if not df.empty else {}


# ══════════════════════════════════════════════════════════
# PAGE: DASHBOARD
# ══════════════════════════════════════════════════════════
if page == "📊 Dashboard":
    st.markdown("## Dashboard")
    st.markdown(f"*Last updated: {date.today().strftime('%B %d, %Y')}*")

    # Stats row
    col1, col2, col3, col4, col5 = st.columns(5)
    with col1:
        st.metric("Total Applied", stats.get("total", 0))
    with col2:
        st.metric("Interviews", stats.get("interviewed", 0))
    with col3:
        rate = round(stats.get("interviewed", 0) / max(stats.get("total", 1), 1) * 100, 1)
        st.metric("Interview Rate", f"{rate}%")
    with col4:
        st.metric("Offers", stats.get("offered", 0))
    with col5:
        st.metric("Rejected", stats.get("rejected", 0))

    st.markdown("---")

    col_left, col_right = st.columns([1.5, 1])

    with col_left:
        st.markdown('<div class="section-header">Application status breakdown</div>', unsafe_allow_html=True)

        status_map = {
            "interviewed": "Interviewed",
            "applied": "Applied",
            "rejected": "Rejected",
            "offered": "Offered",
            "unknown": "Unknown",
            "wishlist": "Wishlist",
            "follow-up q": "Follow-up Q",
        }

        status_counts = {}
        for s, label in status_map.items():
            count = len(df[df["Status"].str.contains(s, case=False, na=False)])
            if count > 0:
                status_counts[label] = count

        if status_counts:
            fig = px.bar(
                x=list(status_counts.keys()),
                y=list(status_counts.values()),
                color=list(status_counts.keys()),
                color_discrete_map={
                    "Interviewed": "#3b82f6",
                    "Applied": "#9ca3af",
                    "Rejected": "#ef4444",
                    "Offered": "#22c55e",
                    "Unknown": "#d1d5db",
                    "Wishlist": "#f59e0b",
                    "Follow-up Q": "#a855f7",
                },
                labels={"x": "", "y": "Count"},
            )
            fig.update_layout(
                showlegend=False,
                plot_bgcolor="white",
                paper_bgcolor="white",
                margin=dict(l=0, r=0, t=10, b=0),
                height=280,
                font=dict(size=13),
            )
            fig.update_traces(marker_line_width=0)
            st.plotly_chart(fig, use_container_width=True)

    with col_right:
        st.markdown('<div class="section-header">Status distribution</div>', unsafe_allow_html=True)

        if status_counts:
            fig2 = go.Figure(data=[go.Pie(
                labels=list(status_counts.keys()),
                values=list(status_counts.values()),
                hole=0.55,
                marker_colors=["#3b82f6", "#9ca3af", "#ef4444", "#22c55e", "#d1d5db", "#f59e0b", "#a855f7"],
            )])
            fig2.update_layout(
                showlegend=True,
                plot_bgcolor="white",
                paper_bgcolor="white",
                margin=dict(l=0, r=0, t=10, b=0),
                height=280,
                font=dict(size=12),
                legend=dict(font=dict(size=11)),
            )
            fig2.update_traces(textinfo="none")
            st.plotly_chart(fig2, use_container_width=True)

    # Recent applications
    st.markdown('<div class="section-header">Recent applications</div>', unsafe_allow_html=True)

    recent = df.tail(10)[["Company", "Position", "Applied on", "Status"]].copy()
    recent = recent.iloc[::-1]

    for _, row in recent.iterrows():
        col1, col2, col3, col4 = st.columns([2, 2.5, 1.5, 1.5])
        with col1:
            st.markdown(f"**{row['Company']}**")
        with col2:
            pos = str(row['Position'])
            st.markdown(f"<span style='color: #6b7280; font-size: 13px;'>{pos[:50]}{'...' if len(pos) > 50 else ''}</span>", unsafe_allow_html=True)
        with col3:
            st.markdown(f"<span style='color: #9ca3af; font-size: 13px;'>{row['Applied on']}</span>", unsafe_allow_html=True)
        with col4:
            st.markdown(status_badge(row['Status']), unsafe_allow_html=True)


# ══════════════════════════════════════════════════════════
# PAGE: APPLICATIONS
# ══════════════════════════════════════════════════════════
elif page == "📋 Applications":
    st.markdown("## Applications")

    col1, col2, col3 = st.columns([2, 1.5, 1])
    with col1:
        search = st.text_input("🔍 Search", placeholder="Company or position...")
    with col2:
        status_filter = st.selectbox("Status", ["All", "Applied", "Interviewed", "Offered", "Rejected", "Unknown"])
    with col3:
        limit = st.selectbox("Show", [50, 100, 200, "All"])

    filtered = df.copy()

    if search:
        mask = (
            filtered["Company"].str.contains(search, case=False, na=False) |
            filtered["Position"].str.contains(search, case=False, na=False)
        )
        filtered = filtered[mask]

    if status_filter != "All":
        filtered = filtered[filtered["Status"].str.contains(status_filter.lower(), case=False, na=False)]

    filtered = filtered.iloc[::-1]

    if limit != "All":
        filtered = filtered.head(int(limit))

    st.markdown(f"*Showing {len(filtered)} applications*")

    display_df = filtered[["Company", "Position", "Applied on", "Status"]].copy()
    display_df.columns = ["Company", "Position", "Applied", "Status"]

    st.dataframe(
        display_df,
        use_container_width=True,
        height=500,
        column_config={
            "Company": st.column_config.TextColumn("Company", width="medium"),
            "Position": st.column_config.TextColumn("Position", width="large"),
            "Applied": st.column_config.TextColumn("Applied", width="small"),
            "Status": st.column_config.TextColumn("Status", width="small"),
        }
    )


# ══════════════════════════════════════════════════════════
# PAGE: EMAIL SCANNER
# ══════════════════════════════════════════════════════════
elif page == "📧 Email Scanner":
    st.markdown("## Email Scanner")
    st.markdown("Scan your Gmail for job application updates and auto-update your Excel.")

    col1, col2 = st.columns([1, 2])
    with col1:
        num_emails = st.number_input("Emails to scan", min_value=5, max_value=200, value=50, step=5)

    if st.button("🔍 Scan emails now", use_container_width=False):
        with st.spinner("Connecting to Gmail and scanning emails..."):
            try:
                from tools import scan_emails_for_status
                result = scan_emails_for_status(max_results=num_emails)
                st.success(result)
                st.cache_data.clear()
            except Exception as e:
                st.error(f"Error: {e}")


# ══════════════════════════════════════════════════════════
# PAGE: COVER LETTER
# ══════════════════════════════════════════════════════════
elif page == "📝 Cover Letter":
    st.markdown("## Cover Letter Generator")

    url = st.text_input("Job URL", placeholder="https://jobs.ashbyhq.com/...")
    jd_manual = st.text_area("Or paste JD manually", height=200, placeholder="Paste the job description here if the URL can't be fetched automatically...")

    st.write("DEBUG url:", url)
    st.write("DEBUG jd_manual length:", len(jd_manual))
    
    if st.button("✨ Generate Cover Letter"):
        st.write("DEBUG: Cover Letter button clicked")
        if not url and not jd_manual:
            st.warning("Please provide a URL or paste the JD.")
        else:
            with st.spinner("Generating cover letter..."):
                try:
                    from nodes import _generate_cover_letter
                    from langchain_core.messages import HumanMessage

                    if url and not jd_manual:
                        from tools import scrape_job_url
                        st.write("DEBUG: about to call scrape_job_url")
                        jd = scrape_job_url(url)
                        st.write("DEBUG: scraped JD length:", len(jd))
                        if jd.startswith("SCRAPE_ERROR:"):
                            st.error(jd)
                            st.stop()
                        
                        if not jd or len(jd) < 200:
                            jd = jd_manual
                    else:
                        jd = jd_manual

                    if not jd:
                        st.error("Could not fetch JD. Please paste it manually.")
                    else:
                        result = _generate_cover_letter(jd, url, {"messages": [HumanMessage(content="generate cover letter")]})
                        cover_letter_text = result["result"]

                        st.markdown("### Generated Cover Letter")
                        st.text_area("", value=cover_letter_text, height=400)

                        if "saved to:" in cover_letter_text.lower():
                            st.success("Cover letter saved to your cover_letters/ folder!")

                except Exception as e:
                    st.error(f"Error: {e}")


# ══════════════════════════════════════════════════════════
# PAGE: INTERVIEW PREP
# ══════════════════════════════════════════════════════════
elif page == "🎯 Interview Prep":
    st.markdown("## Interview Prep")

    url = st.text_input("Job URL", placeholder="https://...")
    jd_manual = st.text_area("Or paste JD manually", height=200, placeholder="Paste the job description here...")

    if st.button("🎯 Generate Interview Prep"):
        if not url and not jd_manual:
            st.warning("Please provide a URL or paste the JD.")
        else:
            with st.spinner("Generating interview prep guide..."):
                try:
                    from nodes import _generate_interview_prep
                    from langchain_core.messages import HumanMessage

                    if url and not jd_manual:
                        from tools import scrape_job_url
                        jd = scrape_job_url(url)
                        if not jd or len(jd) < 200:
                            jd = jd_manual
                    else:
                        jd = jd_manual

                    if not jd:
                        st.error("Could not fetch JD. Please paste it manually.")
                    else:
                        result = _generate_interview_prep(jd, url, {"messages": [HumanMessage(content="interview prep")]})
                        st.markdown(result["result"])

                except Exception as e:
                    st.error(f"Error: {e}")


# ══════════════════════════════════════════════════════════
# PAGE: JOB MATCH
# ══════════════════════════════════════════════════════════
elif page == "🔍 Job Match":
    st.markdown("## Job Match Scorer")

    tab1, tab2 = st.tabs(["Single job", "Batch score"])

    with tab1:
        company_input = st.text_input("Company name", placeholder="Caesars Entertainment")
        url = st.text_input("Or job URL", placeholder="https://...")
        jd_manual = st.text_area("Or paste JD manually", height=150, placeholder="Paste the job description here...")

        if st.button("🔍 Score my match"):
            with st.spinner("Analyzing job match..."):
                try:
                    from nodes import _generate_job_match
                    from tools import get_jd_from_excel, scrape_job_url
                    from langchain_core.messages import HumanMessage

                    jd, final_company, final_position, row_index = "", "", "", None

                    if company_input:
                        excel_data = get_jd_from_excel(company=company_input)
                        if excel_data and excel_data["jd"]:
                            jd = excel_data["jd"]
                            final_company = excel_data["company"]
                            final_position = excel_data["position"]
                            row_index = excel_data["row_index"]
                            if not url and excel_data["url"]:
                                url = excel_data["url"]

                    if not jd and url:
                        jd = scrape_job_url(url)

                    if not jd and jd_manual:
                        jd = jd_manual

                    if not jd:
                        st.error("Could not find JD. Please paste it manually.")
                    else:
                        result = _generate_job_match(jd, url, final_company, final_position, row_index, {"messages": [HumanMessage(content="job match")]})
                        st.markdown(result["result"])
                        st.cache_data.clear()

                except Exception as e:
                    st.error(f"Error: {e}")

    with tab2:
        limit = st.number_input("Score latest N jobs", min_value=1, max_value=20, value=5)

        if st.button("🔍 Batch score"):
            with st.spinner(f"Scoring {limit} jobs..."):
                try:
                    from tools import get_latest_jobs_with_jd, scrape_job_url
                    from nodes import _generate_job_match
                    from langchain_core.messages import HumanMessage

                    jobs = get_latest_jobs_with_jd(limit=int(limit))
                    results = []

                    progress = st.progress(0)
                    for idx, job in enumerate(jobs):
                        jd = str(job.get("JD") or "").strip()
                        url = str(job.get("Application Link") or "").strip()
                        job_company = str(job.get("Company") or "")
                        job_position = str(job.get("Position") or "")
                        row_index = job["row_index"]

                        if not jd and url:
                            jd = scrape_job_url(url)

                        if not jd:
                            results.append({"Company": job_company, "Position": job_position, "Score": "N/A", "Note": "No JD"})
                            continue

                        match_result = _generate_job_match(jd, url, job_company, job_position, row_index, {"messages": [HumanMessage(content="job match")]})

                        score = "N/A"
                        import re
                        for line in match_result["result"].split("\n"):
                            if "SCORE:" in line.upper():
                                m = re.search(r'(\d+)/10', line)
                                if m:
                                    score = f"{m.group(1)}/10"
                                    break

                        results.append({"Company": job_company, "Position": job_position[:50], "Score": score})
                        progress.progress((idx + 1) / len(jobs))

                    st.cache_data.clear()
                    st.dataframe(pd.DataFrame(results), use_container_width=True)

                except Exception as e:
                    st.error(f"Error: {e}")


# ══════════════════════════════════════════════════════════
# PAGE: CHAT
# ══════════════════════════════════════════════════════════
elif page == "💬 Chat":
    st.markdown("## Chat with your Agent")

    if "chat_history" not in st.session_state:
        st.session_state.chat_history = []
    if "agent_state" not in st.session_state:
        st.session_state.agent_state = {
            "messages": [],
            "next_agent": "",
            "intent": "",
            "result": "",
            "pending_action": "",
            "pending_url": "",
            "pending_jd": ""
        }

    # Display chat history
    for msg in st.session_state.chat_history:
        if msg["role"] == "user":
            st.markdown(f'<div class="chat-msg-user">{msg["content"]}</div>', unsafe_allow_html=True)
        else:
            st.markdown(f'<div class="chat-msg-agent">{msg["content"]}</div>', unsafe_allow_html=True)

    # Input
    with st.form("chat_form", clear_on_submit=True):
        col1, col2 = st.columns([5, 1])
        with col1:
            user_input = st.text_input("", placeholder="Ask anything... e.g. 'Show my latest 5 applications'", label_visibility="collapsed")
        with col2:
            submitted = st.form_submit_button("Send ↗", use_container_width=True)

    if submitted and user_input:
        from langchain_core.messages import HumanMessage

        st.session_state.chat_history.append({"role": "user", "content": user_input})
        st.session_state.agent_state["messages"].append(HumanMessage(content=user_input))

        with st.spinner("Thinking..."):
            try:
                from supervisor import supervisor
                result = supervisor.invoke({**st.session_state.agent_state})
                st.session_state.agent_state = {**st.session_state.agent_state, **result}
                response = result["messages"][-1].content
                st.session_state.chat_history.append({"role": "agent", "content": response})
                st.cache_data.clear()
            except Exception as e:
                st.session_state.chat_history.append({"role": "agent", "content": f"Error: {e}"})

        st.rerun()

    # Quick action buttons
    st.markdown("---")
    st.markdown("**Quick actions**")
    col1, col2, col3, col4 = st.columns(4)
    with col1:
        if st.button("📋 Show latest 5 apps", use_container_width=True):
            st.session_state.chat_history.append({"role": "user", "content": "show me my latest 5 applications"})
            st.rerun()
    with col2:
        if st.button("📊 Job search analysis", use_container_width=True):
            st.session_state.chat_history.append({"role": "user", "content": "how is my job search going?"})
            st.rerun()
    with col3:
        if st.button("📧 Scan 20 emails", use_container_width=True):
            st.session_state.chat_history.append({"role": "user", "content": "scan my last 20 emails for job updates"})
            st.rerun()
    with col4:
        if st.button("🗑️ Clear chat", use_container_width=True):
            st.session_state.chat_history = []
            st.session_state.agent_state = {
                "messages": [], "next_agent": "", "intent": "",
                "result": "", "pending_action": "", "pending_url": "", "pending_jd": ""
            }
            st.rerun()

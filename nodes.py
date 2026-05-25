"""
nodes.py
--------
Logic for each Graph node.

Each node is a function that receives AgentState and returns an updated AgentState.
Nodes communicate through State and never call each other directly.

Node design:
  router_node   → Classifies user intent
  query_node    → Fetches application data
  add_node      → Adds a new application
  analyze_node  → Analyzes job search and gives advice
  respond_node  → Packages result into final response
"""

import os
from langchain_groq import ChatGroq
from langchain_core.messages import AIMessage, HumanMessage

from state import AgentState
from tools import (
    get_all_applications,
    get_pending_followups,
    add_application,
    get_status_summary,
    update_application,
    scrape_job_url,
    get_incomplete_rows,
    update_excel_row,
    load_profile,
    save_cover_letter,
    scan_emails_for_status,
    save_interview_prep,
    get_jd_from_excel,
    save_match_score_to_excel,
)
import re
from dotenv import load_dotenv
import json

# Initialize LLM (shared across all nodes)
load_dotenv()  # Load .env before initializing LLM
llm = ChatGroq(model="meta-llama/llama-4-scout-17b-16e-instruct")   


# ──────────────────────────────────────────
# Node 1: Router
# Classifies user intent and decides which path to take
# ──────────────────────────────────────────
def router_node(state: AgentState) -> AgentState:
    """
    Analyzes the latest user message and classifies intent.
    Returns intent: "query" | "add" | "analyze" | "unknown"
    """
    user_message = state["messages"][-1].content
    
    prompt = f"""You are a router for a job application tracker.
Classify the user's intent into ONE of these categories:
- "query": user wants to see their applications or follow-ups
- "add": user wants to add a new job application
- "analyze": user wants advice, summary, or analysis of their job search
- "edit": user wants to update or change the status of an existing application
- "scrape": user wants to automatically fill in missing job details from URLs in the Excel file
- "cover_letter": user wants to generate a cover letter for a specific job
- "email_scan": user wants to scan emails to update application statuses
- "interview_prep": user wants to prepare for a job interview, generate interview questions and tips
- "job_match": user wants to score or analyze how well they match a job, mentions "match", "score", "fit", "should I apply"
- "unknown": none of the above

User message: "{user_message}"

Reply with ONLY one word: query, add, analyze, edit, scrape, cover_letter, email_scan, interview_prep, job_match, or unknown."""
    
    response = llm.invoke([HumanMessage(content=prompt)])
    intent = response.content.strip().lower()

    # Fallback: ensure only valid intents are returned
    if intent not in ["query", "add", "analyze", "edit", "scrape", "cover_letter", "email_scan", "interview_prep", "job_match"]:
        intent = "unknown"

    print(f"[Router] Detected intent: {intent}")
    return {"intent": intent}


# ──────────────────────────────────────────
# Node 2: Query
# Fetches application data from Excel
# ──────────────────────────────────────────
def query_node(state: AgentState) -> AgentState:
    user_message = state["messages"][-1].content

    if "follow" in user_message.lower() or "remind" in user_message.lower():
        result = get_pending_followups()
    else:
        # Use LLM to extract the number of records requested
        extract_prompt = f"""Extract the number of records the user wants to see.
        Return ONLY a number, or "all" if they want all records.

        Examples:
        - "show me latest 5 applications" → 5
        - "show my last 10 jobs" → 10
        - "show all my applications" → all
        - "show me my applications" → all

        Message: "{user_message}" """

        response = llm.invoke([HumanMessage(content=extract_prompt)])
        val = response.content.strip().lower()

        limit = None
        if val.isdigit():
            limit = int(val)

        result = get_all_applications(limit=limit)
        print(f"[Query] Fetched data (limit={limit}):\n{result}...")

    return {"result": result}


# ──────────────────────────────────────────
# Node 3: Add
# Adds a new job application record
# ──────────────────────────────────────────
def add_node(state: AgentState) -> AgentState:
    """
    Extracts company name and position from the user message and adds a new record.
    Uses LLM for information extraction instead of complex regex.
    """
    user_message = state["messages"][-1].content

    # Use LLM to extract structured information
    prompt = f"""Extract job application details from this message.
Return ONLY in this exact format (no extra text):
company: <company name>
role: <job role>
notes: <any notes, or empty>

Message: "{user_message}" """

    response = llm.invoke([HumanMessage(content=prompt)])
    text = response.content.strip()

    # Parse the LLM response text
    company, role, notes = "Unknown", "Unknown", ""
    for line in text.split("\n"):
        if line.startswith("company:"):
            company = line.split(":", 1)[1].strip()
        elif line.startswith("role:"):
            role = line.split(":", 1)[1].strip()
        elif line.startswith("notes:"):
            notes = line.split(":", 1)[1].strip()

    result = add_application(company, role, notes)
    print(f"[Add] {result}")
    return {"result": result}


# ──────────────────────────────────────────
# Node 4: Analyze
# Analyzes job search status and provides actionable advice
# ──────────────────────────────────────────
def analyze_node(state: AgentState) -> AgentState:
    """
    Fetches status summary and follow-up list, then uses LLM to generate actionable advice.
    """
    summary = get_status_summary()
    followups = get_pending_followups()
    user_message = state["messages"][-1].content

    prompt = f"""You are a career coach analyzing a job seeker's progress.

Current application data:
{summary}

Pending follow-ups:
{followups}

User's question: "{user_message}"

Give a concise, actionable response (3-5 sentences). Be encouraging but honest."""

    response = llm.invoke([HumanMessage(content=prompt)])
    result = response.content.strip()
    print(f"[Analyze] Generated advice.")
    return {"result": result}


# ──────────────────────────────────────────
# Node 5: Respond
# Packages result into final AI response
# ──────────────────────────────────────────
def respond_node(state: AgentState) -> AgentState:
    """
    Wraps the result into an AIMessage and appends it to the conversation history.
    This is the terminal node for all paths.
    """
    result = state.get("result", "Sorry, I couldn't process your request.")

    if state.get("intent") == "unknown":
        result = "I can help you query your applications, add new ones, or analyze your job search progress. What would you like to do?"

    return {"messages": [AIMessage(content=result)]}

def edit_node(state: AgentState) -> AgentState:
    print("[Edit] Node entered!")
    user_message = state["messages"][-1].content

    prompt = f"""Extract job application edit details from this message.
Return ONLY in this exact format (no extra text):
id: <job id number, or empty>
company: <company name, or empty>
role: <job title, or empty>
status: <new status, or empty>
notes: <notes, or empty>

Valid statuses: applied, interviewed, rejected, offered, wishlist

Message: "{user_message}" """

    response = llm.invoke([HumanMessage(content=prompt)])
    text = response.content.strip()

    job_id, company, role, status, notes = None, None, None, None, None
    for line in text.split("\n"):
        if line.startswith("id:"):
            val = line.split(":", 1)[1].strip()
            job_id = int(val) if val.isdigit() else None
        elif line.startswith("company:"):
            val = line.split(":", 1)[1].strip()
            company = val if val else None
        elif line.startswith("role:"):
            val = line.split(":", 1)[1].strip()
            role = val if val else None
        elif line.startswith("status:"):
            val = line.split(":", 1)[1].strip()
            status = val if val else None
        elif line.startswith("notes:"):
            val = line.split(":", 1)[1].strip()
            notes = val if val else None

    result = update_application(company, role, notes, status, job_id)
    print(f"[Edit] {result}")
    return {"result": result}

def scrape_node(state: AgentState) -> AgentState:
    """
    找出 Excel 中缺少資料但有連結的列，
    爬取每個網址，用 LLM 抽取資訊，填回 Excel。
    """
    incomplete = get_incomplete_rows()

    if not incomplete:
        return {"result": "✅ All rows are already complete. Nothing to update."}

    updated, skipped = 0, 0
    failed_urls = []

    for row in incomplete:
        print(f"[Scrape] Fetching: {row['url']}...")
        content = scrape_job_url(row["url"])

        print(f"[Scrape] content: {content}")

        if not content or len(content) < 200:
            print(f"[Scrape] ⚠️ Too little content ({len(content)} chars), likely requires login: {row['url']}")
            skipped += 1
            failed_urls.append(f"{row['url']} (only {len(content)} chars)")
            continue


        # 用 LLM 從網頁內容抽取資訊
        prompt = f"""Extract job posting details from the text below.
Return ONLY in this exact format (no extra text):
company: <company name, or empty if not found>
position: <job title, or empty if not found>
jd: <full job description as plain text, do NOT summarize, keep every word exactly as written, or empty if not found>

Text:
{content}"""

        response = llm.invoke([HumanMessage(content=prompt)])
        text = response.content.strip()

        company, position, jd = None, None, None
        lines = text.split("\n")
        jd_lines = []
        in_jd = False

        for line in lines:
            if line.startswith("company:"):
                val = line.split(":", 1)[1].strip()
                company = val if val else None
                in_jd = False
            elif line.startswith("position:"):
                val = line.split(":", 1)[1].strip()
                position = val if val else None
                in_jd = False
            elif line.startswith("jd:"):
                val = line.split(":", 1)[1].strip()
                jd_lines = [val] if val else []
                in_jd = True
            elif in_jd:
                jd_lines.append(line)

        jd = "\n".join(jd_lines).strip() if jd_lines else None

        # 只更新原本缺少的欄位
        update_excel_row(
            row["row_index"],
            company=company if not row["company"] else None,
            position=position if not row["position"] else None,
            jd=jd if not row["jd"] else None,
        )
        updated += 1
        print(f"[Scrape] Updated row {row['row_index']}: {company} - {position}")

    result = f"✅ Updated {updated} rows, skipped {skipped} rows.\n"
    if failed_urls:
        result += "\n⚠️ Failed to fetch these URLs:\n"
        result += "\n".join(f"  - {url}" for url in failed_urls)

        
    return {"result": result}


def cover_letter_node(state: AgentState) -> AgentState:
    """
    嘗試從用戶訊息抽取 URL，爬取 JD。
    爬到就直接生成 cover letter。
    爬不到就暫停，請用戶手動貼 JD。
    """
    user_message = state["messages"][-1].content

    # 從訊息裡抽取 URL
    urls = re.findall(r'https?://\S+', user_message)
    url = urls[0] if urls else ""

    if url:
        print(f"[CoverLetter] Fetching: {url}...")
        content = scrape_job_url(url)
    else:
        content = ""

    # 爬到足夠內容，直接生成
    if content and len(content) >= 200:
        return _generate_cover_letter(content, url, state)

    # 爬不到，請用戶手動貼 JD
    print(f"[CoverLetter] Cannot fetch JD, asking user to paste manually.")
    return {
        "pending_action": "waiting_for_jd",
        "pending_url": url,
        "result": "",
        "messages": [AIMessage(content=
            "I wasn't able to automatically fetch the job description from that URL "
            "(it may require login or JavaScript). Could you please paste the job description text directly? "
            "I'll generate your cover letter right away.")]
    }


def cover_letter_jd_node(state: AgentState) -> AgentState:
    """
    用戶手動貼上 JD 後，用這個節點生成 cover letter。
    """
    jd = state["messages"][-1].content
    url = state.get("pending_url", "")
    return _generate_cover_letter(jd, url, state)


def _generate_cover_letter(jd: str, url: str, state: AgentState) -> AgentState:
    """
    根據 JD 和 profile.json 生成 cover letter。
    """
    profile = load_profile()
    profile_str = json.dumps(profile, indent=2, ensure_ascii=False)

    prompt = f"""You are helping Oscar Shih write a cover letter. Your goal is to make it sound authentically human — someone who has genuinely read the job description and is responding to it specifically.

        STEP 1 - Analyze the JD first:
        Before writing, identify:
        - The 3 most important requirements or responsibilities in this JD
        - The tone of the company (formal/casual/mission-driven/technical)
        - Any specific pain points or challenges this role is meant to solve
        - Keywords or phrases that appear multiple times (these are priorities)

        STEP 2 - Match Oscar's experience to those specific requirements:
        For each of the 3 requirements you identified, find the MOST relevant evidence from Oscar's background below. Be specific — use real numbers, real project names, real outcomes.

        STEP 3 - Write the cover letter:
        - Opening: Reference something SPECIFIC from this JD or company that genuinely resonates. Do NOT use "I am excited to apply" or "I am passionate about".
        - Body paragraph 1: Address requirement #1 with Oscar's most relevant experience. Tell a brief story, don't just list qualifications.
        - Body paragraph 2: Address requirement #2 and #3 together, weaving in a second relevant experience.
        - Closing: Express genuine interest in THIS specific role's challenges. Mention one thing you want to learn or contribute.
        - Sign off with name and email.

        Oscar's background:
        {profile_str}

        Job Description:
        {jd}

        Style rules:
        - Tone: match the company's tone from the JD
        - Length: 3-4 paragraphs, no more
        - Never use: "leverage", "synergy", "passionate", "excited to apply", "I am writing to apply"
        - Vary sentence length — mix short punchy sentences with longer ones
        - Each paragraph should advance a single clear point
        - Read it back: if it sounds like it could apply to any company, rewrite it

        End with:
        {profile['name']}
        {profile['email']}

        Output ONLY the cover letter, no commentary, no "Here is your cover letter"."""

    print(f"[CoverLetter] Generating cover letter...")
    response = llm.invoke([HumanMessage(content=prompt)])
    result = response.content.strip()

    # 從 JD 抽取公司名稱和職位
    extract_prompt = f"""Extract company name and job title from this job description.
    Return ONLY in this exact format:
    company: <company name>
    position: <job title>

    Text: {jd}"""

    extract_response = llm.invoke([HumanMessage(content=extract_prompt)])
    extract_text = extract_response.content.strip()

    company, position = "Unknown Company", "Unknown Position"
    for line in extract_text.split("\n"):
        if line.startswith("company:"):
            company = line.split(":", 1)[1].strip()
        elif line.startswith("position:"):
            position = line.split(":", 1)[1].strip()

    # 存成 Word 檔
    filepath = save_cover_letter(company, position, result)
    save_msg = f"\n\n✅ Cover letter saved to: {filepath}"
    print(save_msg)

    return {
    "result": result,
    "pending_action": "",
    "pending_url": "",
    "pending_jd": "",
    "messages": [AIMessage(content=result + save_msg)],
    }


def email_scan_node(state: AgentState) -> AgentState:
    user_message = state["messages"][-1].content

    # Use LLM to extract the number of records requested
    extract_prompt = f"""Extract the number of emails the user wants to scan.
        Return ONLY a number, or "50" if not specified.

        Examples:
        - "scan my last 100 emails" → 100
        - "check 20 emails" → 20
        - "scan my emails" → 50

        Message: "{user_message}" """

    response = llm.invoke([HumanMessage(content=extract_prompt)])
    val = response.content.strip().lower()

    max_results = 50
    if val.isdigit():
        max_results = int(val)

    print(f"[EmailScan] Connecting to Gmail (scanning {max_results} emails)...")
    result = scan_emails_for_status(max_results=max_results)
    print(f"[EmailScan] Done.")
    return {"result": result}


def interview_prep_node(state: AgentState) -> AgentState:
    """
    嘗試從用戶訊息抽取 URL，爬取 JD。
    爬到就直接生成 interview prep。
    爬不到就暫停，請用戶手動貼 JD。
    """
    user_message = state["messages"][-1].content

    urls = re.findall(r'https?://\S+', user_message)
    url = urls[0] if urls else ""

    if url:
        print(f"[InterviewPrep] Fetching: {url}...")
        content = scrape_job_url(url)
    else:
        content = ""

    if content and len(content) >= 200:
        return _generate_interview_prep(content, url, state)

    print(f"[InterviewPrep] Cannot fetch JD, asking user to paste manually.")
    return {
        "pending_action": "waiting_for_jd_interview",
        "pending_url": url,
        "result": "",
        "messages": [AIMessage(content=
            "I wasn't able to automatically fetch the job description from that URL. "
            "Could you please paste the job description text directly? "
            "I'll generate your interview prep right away.")]
    }


def interview_prep_jd_node(state: AgentState) -> AgentState:
    """
    用戶手動貼上 JD 後，用這個節點生成 interview prep。
    """
    jd = state["messages"][-1].content
    url = state.get("pending_url", "")
    return _generate_interview_prep(jd, url, state)


def _generate_interview_prep(jd: str, url: str, state: AgentState) -> AgentState:
    """
    根據 JD 和 profile.json 生成 interview prep。
    """
    profile = load_profile()
    profile_str = json.dumps(profile, indent=2, ensure_ascii=False)

    prompt = f"""You are a senior career coach preparing Oscar Shih for a specific job interview. Your goal is to give him answers he can actually use, not generic advice.

        STEP 1 - Research the company from the JD:
        Extract and summarize:
        - Company mission and what they actually do
        - Their values or culture signals (look for clues in how they write the JD)
        - The team or department this role belongs to
        - What success looks like in this role based on the JD

        STEP 2 - Analyze the role deeply:
        - What are the top 3 skills/experiences they're looking for?
        - What problems will this person be solving day-to-day?
        - What would make a candidate stand out vs just qualify?

        STEP 3 - Generate the prep guide:

        ## Company & Role Overview
        2-3 sentences on what the company does, their mission, and what this role is really about. This is what Oscar should internalize before the interview.

        ## Fit Analysis
        - ✅ Strong matches: List 4-5 specific JD requirements Oscar clearly meets, with evidence from his background
        - ⚠️ Partial matches: List 2-3 areas where he partially qualifies, with a suggested talking point to bridge the gap
        - ❌ Gaps: Any requirements he lacks, and whether to address proactively or not

        ## Behavioral Questions (STAR Format)
        Generate 5 behavioral questions based specifically on THIS JD's requirements. For each:
        **Q: [Question]**
        Suggested answer using Oscar's background:
        - Situation: [specific situation from Oscar's experience]
        - Task: [what he needed to do]
        - Action: [specific steps he took, with details]
        - Result: [quantified outcome]
        Write this as a FULL DRAFT ANSWER Oscar can practice, not just bullet points. 2-3 sentences per section.

        ## Technical Questions
        Generate 5 technical questions based on the specific tools, methods, or skills mentioned in THIS JD. For each:
        **Q: [Question]**
        Key points to cover: [3-4 specific points]
        Oscar's relevant experience: [direct reference to his background]
        Draft answer: [2-3 sentence answer he can adapt]

        ## Questions to Ask the Interviewer
        Generate 5 thoughtful questions specific to THIS company and role. Each question should:
        - Show you've read the JD carefully
        - Demonstrate genuine curiosity about the work
        - Not be answerable by a quick Google search

        Oscar's background:
        {profile_str}

        Job Description:
        {jd}

        Be specific throughout. Generic advice is useless. Every answer should reference something real from Oscar's background or something specific from the JD."""

    print(f"[InterviewPrep] Generating interview prep...")
    response = llm.invoke([HumanMessage(content=prompt)])
    result = response.content.strip()

    # 抽取公司名稱和職位
    extract_prompt = f"""Extract company name and job title from this job description.
Return ONLY in this exact format:
company: <company name>
position: <job title>

Text: {jd}"""

    extract_response = llm.invoke([HumanMessage(content=extract_prompt)])
    extract_text = extract_response.content.strip()

    company, position = "Unknown Company", "Unknown Position"
    for line in extract_text.split("\n"):
        if line.startswith("company:"):
            company = line.split(":", 1)[1].strip()
        elif line.startswith("position:"):
            position = line.split(":", 1)[1].strip()

    filepath = save_interview_prep(company, position, result)
    save_msg = f"\n\n✅ Interview prep saved to: {filepath}"

    return {
        "result": result + save_msg,
        "pending_action": "",
        "pending_url": "",
        "pending_jd": "",
        "messages": [AIMessage(content=result + save_msg)],
    }

def job_match_node(state: AgentState) -> AgentState:
    """
    Supports single and batch job matching:
    - Single: score my match for Caesars
    - Batch: score my match for the latest 5 jobs
    """
    user_message = state["messages"][-1].content

    # Use LLM to determine single vs batch mode and extract company/limit
    extract_prompt = f"""Analyze this job match request.
Return ONLY in this exact format:
mode: <single or batch>
limit: <number if batch, or empty if single>
company: <company name if single, or empty if batch>
position: <job title if single, or empty if batch>

Examples:
- "score my match for Caesars" → mode: single, company: Caesars
- "score latest 5 jobs" → mode: batch, limit: 5
- "match score for the last 3 applications" → mode: batch, limit: 3

Message: "{user_message}" """

    response = llm.invoke([HumanMessage(content=extract_prompt)])
    text = response.content.strip()

    mode, limit, company, position = "single", 5, None, None
    for line in text.split("\n"):
        if line.startswith("mode:"):
            mode = line.split(":", 1)[1].strip().lower()
        elif line.startswith("limit:"):
            val = line.split(":", 1)[1].strip()
            limit = int(val) if val.isdigit() else 5
        elif line.startswith("company:"):
            val = line.split(":", 1)[1].strip()
            company = val if val else None
        elif line.startswith("position:"):
            val = line.split(":", 1)[1].strip()
            position = val if val else None

    # Batch mode
    if mode == "batch":
        from tools import get_latest_jobs_with_jd
        jobs = get_latest_jobs_with_jd(limit=limit)

        if not jobs:
            return {
                "result": "No jobs with JD or URL found in Excel.",
                "messages": [AIMessage(content="No jobs with JD or URL found in Excel.")]
            }

        print(f"[JobMatch] Batch mode: processing {len(jobs)} jobs...")
        results = []

        for job in jobs:
            jd = str(job.get("JD") or "").strip()
            url = str(job.get("Application Link") or "").strip()
            job_company = str(job.get("Company") or "")
            job_position = str(job.get("Position") or "")
            row_index = job["row_index"]

            # No JD found — try fetching from URL
            if not jd and url:
                print(f"[JobMatch] Fetching URL for {job_company}...")
                content = scrape_job_url(url)
                if content and len(content) >= 200:
                    jd = content

            if not jd:
                print(f"[JobMatch] ⚠️ No JD for {job_company}, skipping.")
                results.append(f"⚠️ {job_company} - skipped (no JD available)")
                continue

            # Generate match score
            match_result = _generate_job_match(
                jd, url, job_company, job_position, row_index, state
            )
            # Extract score from result text
            for line in match_result["result"].split("\n"):
                if "SCORE:" in line.upper():
                    results.append(f"✅ {job_company} - {job_position}: {line.strip()}")
                    break

        summary = f"Batch job match complete ({len(jobs)} jobs processed):\n" + "\n".join(results)
        return {
            "result": summary,
            "pending_action": "",
            "pending_url": "",
            "pending_jd": "",
            "messages": [AIMessage(content=summary)],
        }

    # Single job mode
    urls = re.findall(r'https?://\S+', user_message)
    url = urls[0] if urls else ""

    jd, final_company, final_position, row_index = "", "", "", None
    excel_data = get_jd_from_excel(company=company, position=position)

    if excel_data and excel_data["jd"]:
        print(f"[JobMatch] ✅ Found JD in Excel for {excel_data['company']}")
        jd = excel_data["jd"]
        final_company = excel_data["company"]
        final_position = excel_data["position"]
        row_index = excel_data["row_index"]
        if not url and excel_data["url"]:
            url = excel_data["url"]
    elif excel_data and not excel_data["jd"]:
        final_company = excel_data["company"]
        final_position = excel_data["position"]
        row_index = excel_data["row_index"]
        if not url and excel_data["url"]:
            url = excel_data["url"]

    if not jd and url:
        print(f"[JobMatch] Fetching URL: {url}")
        content = scrape_job_url(url)
        if content and len(content) >= 200:
            jd = content

    if not jd:
        return {
            "pending_action": "waiting_for_jd_match",
            "pending_url": url,
            "result": "",
            "messages": [AIMessage(content=
                "I couldn't find the JD in your Excel or fetch it from the URL. "
                "Could you paste the job description directly?")]
        }

    return _generate_job_match(jd, url, final_company, final_position, row_index, state)


def job_match_jd_node(state: AgentState) -> AgentState:
    """Generates job match report after user manually pastes JD."""
    jd = state["messages"][-1].content
    url = state.get("pending_url", "")
    return _generate_job_match(jd, url, "", "", None, state)


def _generate_job_match(jd: str, url: str, company: str, position: str, row_index: int, state: AgentState) -> AgentState:
    """Generates a job match report based on JD and profile, and saves the score to Excel."""
    profile = load_profile()
    profile_str = json.dumps(profile, indent=2, ensure_ascii=False)

    prompt = f"""You are a career coach analyzing how well Oscar Shih matches a job posting.

Oscar's background:
{profile_str}

Job Description:
{jd}

Generate a Job Match Report with these sections:

## Overall Match Score
Give a score from 1-10 with a one-line justification.
Format the score line exactly like this: SCORE: X/10

## Strong Matches ✅
List 4-5 JD requirements Oscar clearly meets. For each:
- Requirement from JD
- Specific evidence from Oscar's background with numbers/metrics

## Partial Matches ⚠️
List 2-3 requirements Oscar partially meets. For each:
- Requirement from JD
- What Oscar has that's related
- Suggested talking point to bridge the gap

## Missing Requirements ❌
List any requirements Oscar lacks. For each:
- Requirement from JD
- Honest assessment
- Whether it's a dealbreaker or learnable

## Should You Apply?
Give a clear recommendation: Yes / Yes, but... / No, because...
Include 2-3 actionable suggestions to strengthen the application.

Be specific and reference real numbers from Oscar's background wherever possible."""

    print(f"[JobMatch] Generating job match report...")
    response = llm.invoke([HumanMessage(content=prompt)])
    result = response.content.strip()

    # Extract company and position if not already available
    if not company or not position:
        extract_prompt = f"""Extract company name and job title.
Return ONLY in this exact format:
company: <company name>
position: <job title>

Text: {jd}"""
        extract_response = llm.invoke([HumanMessage(content=extract_prompt)])
        for line in extract_response.content.strip().split("\n"):
            if line.startswith("company:"):
                company = line.split(":", 1)[1].strip() or "Unknown Company"
            elif line.startswith("position:"):
                position = line.split(":", 1)[1].strip() or "Unknown Position"

    # Extract score from result text
    score = None
    for line in result.split("\n"):
        if "SCORE:" in line.upper():
            match = re.search(r'(\d+)/10', line)
            if match:
                score = f"{match.group(1)}/10"
                break

    # Save score back to Excel
    save_msg = ""
    if score and row_index:
        save_match_score_to_excel(row_index, score)
        save_msg = f"\n\n✅ Match score {score} saved to Excel."
    elif score:
        save_msg = f"\n\n📊 Match score: {score} (company not found in Excel, score not saved)"

    return {
        "result": result + save_msg,
        "pending_action": "",
        "pending_url": "",
        "pending_jd": "",
        "messages": [AIMessage(content=result + save_msg)],
    }
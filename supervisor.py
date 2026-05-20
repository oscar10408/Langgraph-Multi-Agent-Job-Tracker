"""
supervisor.py
-------------
Supervisor Agent is responsible for:
1. Understanding user intent
2. Routing requests to the appropriate sub-agent
3. Returning the result back to the user

Architecture:
- ExcelAgent:   query, add, edit, analyze
- EmailAgent:   email scan
- CareerAgent:  cover letter, interview prep, job match
- ScraperAgent: scrape JD from URLs
"""

import os
from dotenv import load_dotenv
from langchain_groq import ChatGroq
from langchain_core.messages import HumanMessage
from langgraph.graph import StateGraph, START, END
from langgraph.graph.message import add_messages
from typing import Annotated, TypedDict

load_dotenv()


class SupervisorState(TypedDict):
    messages: Annotated[list, add_messages]
    next_agent: str
    result: str
    intent: str
    pending_action: str
    pending_url: str
    pending_jd: str


def build_supervisor():
    llm = ChatGroq(model="meta-llama/llama-4-scout-17b-16e-instruct")

    from agents.excel_agent import build_excel_agent
    from agents.email_agent import build_email_agent
    from agents.career_agent import build_career_agent
    from agents.scraper_agent import build_scraper_agent

    excel_agent = build_excel_agent()
    email_agent = build_email_agent()
    career_agent = build_career_agent()
    scraper_agent = build_scraper_agent()

    def supervisor_node(state: SupervisorState) -> SupervisorState:
        user_message = state["messages"][-1].content

        prompt = f"""You are a supervisor for a job application tracking system.
Decide which agent should handle this request:

- "excel": query applications, add new application, edit/update application status, analyze job search progress
- "email": scan Gmail emails to update application statuses
- "career": generate cover letter, prepare for interview, score job match/fit
- "scraper": fill in missing job details from URLs in Excel

User message: "{user_message}"

Reply with ONLY one word: excel, email, career, or scraper."""

        response = llm.invoke([HumanMessage(content=prompt)])
        next_agent = response.content.strip().lower()

        if next_agent not in ["excel", "email", "career", "scraper"]:
            next_agent = "excel"

        print(f"[Supervisor] Routing to: {next_agent} agent")
        return {"next_agent": next_agent}

    def route_to_agent(state: SupervisorState) -> str:
        # Handle pending actions — route directly to career agent
        if state.get("pending_action") in [
            "waiting_for_jd",
            "waiting_for_jd_interview",
            "waiting_for_jd_match"
        ]:
            return "career"
        return state.get("next_agent", "excel")

    def run_excel_agent(state: SupervisorState) -> SupervisorState:
        result = excel_agent.invoke({
            "messages": state["messages"],
            "intent": "",
            "result": "",
        })
        return {"messages": [result["messages"][-1]]}

    def run_email_agent(state: SupervisorState) -> SupervisorState:
        result = email_agent.invoke({
            "messages": state["messages"],
            "result": "",
        })
        return {"messages": [result["messages"][-1]]}

    def run_career_agent(state: SupervisorState) -> SupervisorState:
        result = career_agent.invoke({
            "messages": state["messages"],
            "intent": "",
            "result": "",
            "pending_action": state.get("pending_action", ""),
            "pending_url": state.get("pending_url", ""),
            "pending_jd": state.get("pending_jd", ""),
        })
        return {
            "messages": [result["messages"][-1]],
            "pending_action": result.get("pending_action", ""),
            "pending_url": result.get("pending_url", ""),
            "pending_jd": result.get("pending_jd", ""),
        }

    def run_scraper_agent(state: SupervisorState) -> SupervisorState:
        result = scraper_agent.invoke({
            "messages": state["messages"],
            "result": "",
        })
        return {"messages": [result["messages"][-1]]}

    # Build the Supervisor Graph
    builder = StateGraph(SupervisorState)

    builder.add_node("supervisor", supervisor_node)
    builder.add_node("excel_agent", run_excel_agent)
    builder.add_node("email_agent", run_email_agent)
    builder.add_node("career_agent", run_career_agent)
    builder.add_node("scraper_agent", run_scraper_agent)

    builder.add_edge(START, "supervisor")
    builder.add_conditional_edges(
        "supervisor",
        route_to_agent,
        {
            "excel": "excel_agent",
            "email": "email_agent",
            "career": "career_agent",
            "scraper": "scraper_agent",
        }
    )
    builder.add_edge("excel_agent", END)
    builder.add_edge("email_agent", END)
    builder.add_edge("career_agent", END)
    builder.add_edge("scraper_agent", END)

    return builder.compile()


supervisor = build_supervisor()
"""
agents/excel_agent.py
---------------------
ExcelAgent 負責所有跟申請記錄相關的操作：
- Query applications
- Add application
- Edit application
- Analyze job search
"""

import sys
import os
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from typing import Annotated, TypedDict
from langgraph.graph import StateGraph, START, END
from langgraph.graph.message import add_messages
from langchain_core.messages import HumanMessage


class ExcelState(TypedDict):
    messages: Annotated[list, add_messages]
    intent: str
    result: str


def build_excel_agent():
    from nodes import query_node, add_node, edit_node, analyze_node, respond_node
    from langchain_groq import ChatGroq
    from dotenv import load_dotenv
    load_dotenv()

    llm = ChatGroq(model="meta-llama/llama-4-scout-17b-16e-instruct")
    
    def excel_router(state: ExcelState) -> ExcelState:
        user_message = state["messages"][-1].content

        prompt = f"""Classify this request into ONE category:
- "query": user wants to see applications or follow-ups
- "add": user wants to add a new job application
- "edit": user wants to update or change an existing application
- "analyze": user wants advice, coaching, summary, or overall assessment of their job search progress (e.g. "how am I doing", "how is my job search going", "give me advice", "what should I focus on")

Message: "{user_message}"
Reply with ONLY one word: query, add, edit, or analyze."""

        response = llm.invoke([HumanMessage(content=prompt)])
        intent = response.content.strip().lower()
        if intent not in ["query", "add", "edit", "analyze"]:
            intent = "query"

        print(f"[ExcelAgent Router] Intent: {intent}")
        return {"intent": intent}

    def route_by_intent(state: ExcelState) -> str:
        return state.get("intent", "query")

    builder = StateGraph(ExcelState)

    builder.add_node("router", excel_router)
    builder.add_node("query", query_node)
    builder.add_node("add", add_node)
    builder.add_node("edit", edit_node)
    builder.add_node("analyze", analyze_node)
    builder.add_node("respond", respond_node)

    builder.add_edge(START, "router")
    builder.add_conditional_edges(
        "router",
        route_by_intent,
        {
            "query": "query",
            "add": "add",
            "edit": "edit",
            "analyze": "analyze",
        }
    )
    builder.add_edge("query", "respond")
    builder.add_edge("add", "respond")
    builder.add_edge("edit", "respond")
    builder.add_edge("analyze", "respond")
    builder.add_edge("respond", END)

    return builder.compile()
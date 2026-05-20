"""
agents/career_agent.py
----------------------
CareerAgent 負責處理所有職涯相關任務：
- Cover letter 生成
- Interview prep 生成
- Job match scoring
"""

import sys
import os
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from typing import Annotated, TypedDict
from langgraph.graph import StateGraph, START, END
from langgraph.graph.message import add_messages
from langchain_core.messages import AIMessage, HumanMessage


class CareerState(TypedDict):
    messages: Annotated[list, add_messages]
    intent: str
    result: str
    pending_action: str
    pending_url: str
    pending_jd: str


def build_career_agent():
    """
    建立並回傳 CareerAgent 的 compiled graph。
    """
    from nodes import (
        cover_letter_node,
        cover_letter_jd_node,
        interview_prep_node,
        interview_prep_jd_node,
        job_match_node,
        job_match_jd_node,
    )
    from langchain_groq import ChatGroq
    from dotenv import load_dotenv
    load_dotenv()

    llm = ChatGroq(model="meta-llama/llama-4-scout-17b-16e-instruct")

    def career_router(state: CareerState) -> CareerState:
        """判斷是 cover letter、interview prep 還是 job match。"""
        user_message = state["messages"][-1].content

        prompt = f"""Classify this request into ONE category:
            - "cover_letter": user wants to generate a cover letter
            - "interview_prep": user wants to prepare for an interview, generate interview questions
            - "job_match": user wants to score or analyze job fit, mentions match/score/fit/should I apply

            Message: "{user_message}"
            Reply with ONLY one word: cover_letter, interview_prep, or job_match."""

        response = llm.invoke([HumanMessage(content=prompt)])
        intent = response.content.strip().lower()
        if intent not in ["cover_letter", "interview_prep", "job_match"]:
            intent = "cover_letter"

        print(f"[CareerAgent Router] Intent: {intent}")
        return {"intent": intent}

    def route_by_intent(state: CareerState) -> str:
        return state.get("intent", "cover_letter")

    builder = StateGraph(CareerState)

    # 加入節點
    builder.add_node("router", career_router)
    builder.add_node("cover_letter", cover_letter_node)
    builder.add_node("interview_prep", interview_prep_node)
    builder.add_node("job_match", job_match_node)

    # 定義邊
    builder.add_edge(START, "router")
    builder.add_conditional_edges(
        "router",
        route_by_intent,
        {
            "cover_letter": "cover_letter",
            "interview_prep": "interview_prep",
            "job_match": "job_match",
        }
    )
    builder.add_edge("cover_letter", END)
    builder.add_edge("interview_prep", END)
    builder.add_edge("job_match", END)

    return builder.compile()
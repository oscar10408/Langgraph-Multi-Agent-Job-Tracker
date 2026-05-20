"""
agents/email_agent.py
---------------------
EmailAgent 負責掃描 Gmail 更新申請狀態。
"""

import sys
import os
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from typing import Annotated, TypedDict
from langgraph.graph import StateGraph, START, END
from langgraph.graph.message import add_messages


class EmailState(TypedDict):
    messages: Annotated[list, add_messages]
    result: str


def build_email_agent():
    from nodes import email_scan_node, respond_node

    builder = StateGraph(EmailState)

    builder.add_node("email_scan", email_scan_node)
    builder.add_node("respond", respond_node)

    builder.add_edge(START, "email_scan")
    builder.add_edge("email_scan", "respond")
    builder.add_edge("respond", END)

    return builder.compile()
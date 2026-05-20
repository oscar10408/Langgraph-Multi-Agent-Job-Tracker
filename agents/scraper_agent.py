"""
agents/scraper_agent.py
-----------------------
ScraperAgent 負責從 Excel 的 URL 自動抓取 JD 填入。
"""

import sys
import os
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from typing import Annotated, TypedDict
from langgraph.graph import StateGraph, START, END
from langgraph.graph.message import add_messages


class ScraperState(TypedDict):
    messages: Annotated[list, add_messages]
    result: str


def build_scraper_agent():
    from nodes import scrape_node, respond_node

    builder = StateGraph(ScraperState)

    builder.add_node("scrape", scrape_node)
    builder.add_node("respond", respond_node)

    builder.add_edge(START, "scrape")
    builder.add_edge("scrape", "respond")
    builder.add_edge("respond", END)

    return builder.compile()
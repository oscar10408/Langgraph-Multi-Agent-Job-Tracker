"""
state.py
--------
Defines the data structure passed between nodes in the Agent graph.

Every LangGraph node reads from State and returns an updated State.
Think of State as a shared whiteboard that all nodes can read and write.
"""

from typing import Annotated, TypedDict
from langgraph.graph.message import add_messages

class AgentState(TypedDict):
    """
    Agent state that persists throughout the graph lifecycle.

    Attributes:
        messages:       Conversation history
        intent:         User intent classified by the router node
        result:         Output text produced by each node
        pending_action: Tracks paused actions waiting for user input
        pending_url:    Stores the job URL while waiting for user input
        pending_jd:     Stores manually pasted JD content
    """
    messages: Annotated[list, add_messages]
    intent: str
    result: str
    pending_action: str
    pending_url: str
    pending_jd: str

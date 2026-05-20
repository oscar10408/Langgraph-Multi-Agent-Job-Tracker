"""
main.py
-------
Entry point: launches the Job Tracker Agent CLI interface.
Uses a Multi-agent architecture orchestrated by the Supervisor.
"""

import os
from dotenv import load_dotenv
from langchain_core.messages import HumanMessage

load_dotenv()

from supervisor import supervisor


def run():
    print("=" * 60)
    print("  💼 Job Application Tracker Agent")
    print("=" * 60)
    print("Commands you can try:")
    print("  📋 Query    - Show me all my applications")
    print("  ⏰ Follow   - Any follow-ups I should do?")
    print("  ➕ Add      - Add Netflix Data Scientist to my list")
    print("  ✏️  Edit     - Update Caesars status to interviewed")
    print("  📊 Analyze  - How is my job search going?")
    print("  🌐 Scrape   - Fill in missing job details from my Excel")
    print("  📧 Email    - Scan my emails for job application updates")
    print("  📝 Cover    - Generate cover letter for https://...")
    print("  🎯 Prep     - Help me prepare for interview at https://...")
    print("  🔍 Match    - Score my match for Caesars Entertainment")
    print("  ❌ Quit     - quit")
    print("=" * 60)

    state = {
        "messages": [],
        "next_agent": "",
        "intent": "",
        "result": "",
        "pending_action": "",
        "pending_url": "",
        "pending_jd": ""
    }

    while True:
        user_input = input("\nYou: ").strip()

        if not user_input:
            continue
        if user_input.lower() in ["quit", "exit", "q"]:
            print("Good luck with your job search! 🚀")
            break

        state["messages"].append(HumanMessage(content=user_input))

        result = supervisor.invoke({**state})

        # Update state with latest result
        state = {**state, **result}

        response = result["messages"][-1].content
        print(f"\nAgent: {response}")


if __name__ == "__main__":
    run()
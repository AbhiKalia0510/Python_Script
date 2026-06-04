# -*- coding: utf-8 -*-
"""
Created on Sat Mar 21 17:07:37 2026

@author: abhishek.a
"""

import os
import json
import gradio as gr
import anthropic

# ================= CONFIG =================
MODEL = "claude-3-5-sonnet-latest"

# ================= API =================
def get_client(api_key):
    if not api_key:
        raise gr.Error("Please enter API key")
    return anthropic.Anthropic(api_key=api_key)


# ================= PROMPT =================
def build_prompt(meeting_type, team_size, location, time_available, context):

    return f"""
You are a team engagement expert.

Generate 5 highly engaging team activity ideas.

Inputs:
- Meeting Type: {meeting_type}
- Team Size: {team_size}
- Location: {location}
- Time Available: {time_available}
- Context: {context}

For EACH activity return:

1. Activity Name
2. Description (2-3 lines)
3. Duration
4. Steps to execute
5. Ideal for (team type/use case)

Format STRICTLY as JSON:

[
  {{
    "name": "",
    "description": "",
    "duration": "",
    "steps": "",
    "ideal_for": ""
  }}
]
"""


# ================= GENERATE =================
def generate_activities(api_key, meeting_type, team_size, location, time_available, context):

    client = get_client(api_key)

    prompt = build_prompt(meeting_type, team_size, location, time_available, context)

    try:
        response = client.messages.create(
            model=MODEL,
            max_tokens=1200,
            messages=[{"role": "user", "content": prompt}]
        )

        raw = "".join([b.text for b in response.content if hasattr(b, "text")]).strip()

        raw = raw.replace("```json", "").replace("```", "")

        data = json.loads(raw)

        formatted = ""

        for i, act in enumerate(data, 1):
            formatted += f"""
### 🎯 Activity {i}: {act['name']}

**📝 Description:**  
{act['description']}

**⏱ Duration:** {act['duration']}

**📌 Steps:**  
{act['steps']}

**👥 Ideal For:** {act['ideal_for']}

---
"""

        return formatted

    except Exception as e:
        return f"❌ Error: {str(e)}"


# ================= UI =================
def build_ui():

    with gr.Blocks(title="Team Activity Generator", theme=gr.themes.Soft()) as app:

        gr.Markdown(
        """
        # 🧠 Team Activity Ideas
        Get custom activity ideas for your team — just describe your meeting.
        """
        )

        api_key = gr.Textbox(
            label="🔑 API Key",
            type="password",
            placeholder="Enter your Anthropic API key"
        )

        with gr.Row():
            meeting_type = gr.Dropdown(
                ["Standup", "Workshop", "Team Building", "Retrospective", "Icebreaker"],
                label="Meeting Type"
            )

            team_size = gr.Dropdown(
                ["2-5", "5-10", "10-20", "20+"],
                label="Team Size"
            )

        with gr.Row():
            location = gr.Dropdown(
                ["Remote", "Office", "Hybrid"],
                label="Location"
            )

            time_available = gr.Dropdown(
                ["5 min", "10 min", "15 min", "30 min", "1 hour"],
                label="Time Available"
            )

        context = gr.Textbox(
            label="Additional Context",
            placeholder="E.g., burnout, new joiners, engagement issues...",
            lines=3
        )

        generate_btn = gr.Button("🚀 Generate Activities")

        output = gr.Markdown()

        generate_btn.click(
            fn=generate_activities,
            inputs=[api_key, meeting_type, team_size, location, time_available, context],
            outputs=output
        )

    return app


# ================= RUN =================
if __name__ == "__main__":
    app = build_ui()
    app.launch(server_port=8888, inbrowser=True)
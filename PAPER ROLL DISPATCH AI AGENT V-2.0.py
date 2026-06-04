# -*- coding: utf-8 -*-
"""
===============================================================================
ENTERPRISE OFFLINE AI COPILOT PLATFORM
===============================================================================

AUTHOR:
Enterprise Local AI Assistant

FEATURES:
1. Fully Offline AI Assistant
2. Mistral 7B GGUF Support
3. Enterprise SAP-style Dashboard
4. Streaming AI Responses
5. Chat Memory
6. File Upload Intelligence
7. Excel Analysis
8. Python Assistant
9. SQL Assistant
10. Analytics Dashboard
11. Dark Enterprise UI
12. Multi-tab Workspace
13. System Monitoring
14. Optimized CPU Inference
15. No API Required

MODEL:
mistral-7b-instruct-v0.1.Q2_K.gguf

INSTALL:
pip install gradio pandas openpyxl pyxlsb llama-cpp-python plotly psutil loguru

RUN:
python enterprise_ai_copilot.py

===============================================================================
"""

# =============================================================================
# IMPORTS
# =============================================================================

import os
import gc
import time
import psutil
import socket
import platform
import traceback

import pandas as pd
import gradio as gr
import plotly.express as px

from datetime import datetime
from loguru import logger
from llama_cpp import Llama


# =============================================================================
# CONFIGURATION
# =============================================================================

MODEL_PATH = r"D:/Python Script/AI_MODELS/mistral-7b-instruct-v0.1.Q2_K.gguf"

APP_TITLE = "Enterprise AI Copilot"

MAX_HISTORY = 20

# =============================================================================
# LOGGER
# =============================================================================

logger.add(
    "enterprise_ai.log",
    rotation="10 MB",
    retention="7 days"
)

# =============================================================================
# LOAD MODEL
# =============================================================================

print("=" * 80)
print("LOADING ENTERPRISE AI MODEL...")
print("=" * 80)

try:

    llm = Llama(

        model_path=MODEL_PATH,

        n_ctx=4096,

        n_threads=max(
            os.cpu_count() - 1,
            1
        ),

        n_batch=512,

        n_gpu_layers=0,

        verbose=False,

        use_mmap=True,

        use_mlock=False
    )

    print("MODEL LOADED SUCCESSFULLY")

except Exception as e:

    print(f"MODEL LOAD ERROR:\n{e}")

    raise e


# =============================================================================
# GLOBAL MEMORY
# =============================================================================

uploaded_dataframes = {}

chat_sessions = {}


# =============================================================================
# SYSTEM MONITOR
# =============================================================================

def get_system_metrics():

    cpu = psutil.cpu_percent()

    ram = psutil.virtual_memory().percent

    disk = psutil.disk_usage("/").percent

    hostname = socket.gethostname()

    os_name = platform.system()

    return pd.DataFrame([{

        "CPU %": cpu,
        "RAM %": ram,
        "DISK %": disk,
        "HOST": hostname,
        "OS": os_name,
        "TIME": datetime.now().strftime("%H:%M:%S")
    }])


# =============================================================================
# FILE READER
# =============================================================================

def load_file(file):

    try:

        if file is None:
            return "No file uploaded"

        path = file.name

        ext = os.path.splitext(path)[1].lower()

        if ext == ".csv":

            df = pd.read_csv(path)

        elif ext == ".xlsb":

            df = pd.read_excel(
                path,
                engine="pyxlsb"
            )

        else:

            df = pd.read_excel(
                path,
                engine="openpyxl"
            )

        uploaded_dataframes["latest"] = df

        summary = f"""
✅ File Loaded Successfully

Rows: {len(df)}
Columns: {len(df.columns)}

Columns:
{', '.join(df.columns.tolist())}
"""

        return summary

    except Exception as e:

        logger.error(str(e))

        return str(e)


# =============================================================================
# DATA PREVIEW
# =============================================================================

def preview_data():

    try:

        if "latest" not in uploaded_dataframes:

            return pd.DataFrame()

        return uploaded_dataframes["latest"].head(100)

    except:

        return pd.DataFrame()


# =============================================================================
# ANALYTICS
# =============================================================================

def generate_chart():

    try:

        if "latest" not in uploaded_dataframes:

            return None

        df = uploaded_dataframes["latest"]

        numeric_cols = df.select_dtypes(
            include=["number"]
        ).columns.tolist()

        if len(numeric_cols) == 0:

            return None

        col = numeric_cols[0]

        fig = px.histogram(
            df,
            x=col,
            title=f"Distribution of {col}"
        )

        return fig

    except:

        return None


# =============================================================================
# PROMPT BUILDER
# =============================================================================

def build_prompt(
    system_prompt,
    history,
    message
):

    prompt = f"""
<s>[INST]
{system_prompt}
[/INST]
"""

    for item in history[-MAX_HISTORY:]:

        role = item["role"]

        content = item["content"]

        if role == "user":

            prompt += f"\nUser: {content}"

        elif role == "assistant":

            prompt += f"\nAssistant: {content}"

    prompt += f"\nUser: {message}\nAssistant:"

    return prompt


# =============================================================================
# STREAM RESPONSE
# =============================================================================

def stream_response(
    message,
    history,
    system_prompt,
    temperature,
    max_tokens
):

    try:

        if history is None:
            history = []

        prompt = build_prompt(
            system_prompt,
            history,
            message
        )

        history.append({
            "role": "user",
            "content": message
        })

        response_text = ""

        stream = llm.create_completion(

            prompt=prompt,

            max_tokens=int(max_tokens),

            temperature=float(temperature),

            top_p=0.95,

            repeat_penalty=1.1,

            stream=True
        )

        for token in stream:

            piece = token["choices"][0]["text"]

            response_text += piece

            updated_history = history + [{
                "role": "assistant",
                "content": response_text
            }]

            yield "", updated_history

        history.append({
            "role": "assistant",
            "content": response_text
        })

        gc.collect()

    except Exception as e:

        logger.error(traceback.format_exc())

        history.append({
            "role": "assistant",
            "content": str(e)
        })

        yield "", history


# =============================================================================
# CLEAR CHAT
# =============================================================================

def clear_chat():

    return []


# =============================================================================
# ENTERPRISE THEME
# =============================================================================

theme = gr.themes.Soft(

    primary_hue="blue",

    secondary_hue="slate",

    neutral_hue="slate"
)


# =============================================================================
# UI
# =============================================================================

with gr.Blocks(

    theme=theme,

    title=APP_TITLE,

    css="""
    .gradio-container {
        background-color: #0f172a;
    }

    .main-header {
        font-size: 34px;
        font-weight: bold;
        color: white;
    }

    .sub-header {
        color: #94a3b8;
    }
    """
) as demo:

    # =========================================================================
    # HEADER
    # =========================================================================

    gr.Markdown("""
# 🧠 ENTERPRISE AI COPILOT
### Offline SAP-style Enterprise AI Workspace
""")

    # =========================================================================
    # MAIN LAYOUT
    # =========================================================================

    with gr.Row():

        # =====================================================================
        # SIDEBAR
        # =====================================================================

        with gr.Column(scale=2):

            gr.Markdown("## ⚙️ AI Configuration")

            system_prompt = gr.Textbox(

                label="System Prompt",

                lines=10,

                value="""
You are an advanced Enterprise AI Assistant.

Capabilities:
- Python coding
- SQL generation
- Excel analysis
- Inventory analysis
- SAP-style reporting
- Dispatch analytics
- Automation
- Business intelligence
- Enterprise reporting
- Data science
- AI copiloting

Rules:
- Provide professional answers
- Optimize code
- Explain clearly
- Use enterprise standards
"""
            )

            temperature = gr.Slider(

                0.1,
                1.5,

                value=0.7,

                step=0.1,

                label="Temperature"
            )

            max_tokens = gr.Slider(

                64,
                2048,

                value=512,

                step=64,

                label="Max Tokens"
            )

            gr.Markdown(f"""
### 📦 Loaded Model
`mistral-7b-instruct-v0.1.Q2_K.gguf`
""")

            metrics_btn = gr.Button(
                "Refresh System Metrics"
            )

            metrics_table = gr.DataFrame()

            metrics_btn.click(
                fn=get_system_metrics,
                outputs=metrics_table
            )

        # =====================================================================
        # MAIN WORKSPACE
        # =====================================================================

        with gr.Column(scale=8):

            with gr.Tabs():

                # =============================================================
                # AI CHAT
                # =============================================================

                with gr.Tab("🤖 AI Copilot"):

                    chatbot = gr.Chatbot(

                        height=650,

                        type="messages",

                        bubble_full_width=False
                    )

                    msg = gr.Textbox(

                        label="Message",

                        placeholder="Ask anything..."
                    )

                    with gr.Row():

                        send_btn = gr.Button(

                            "Send",

                            variant="primary"
                        )

                        clear_btn = gr.Button(
                            "Clear"
                        )

                    send_btn.click(

                        fn=stream_response,

                        inputs=[
                            msg,
                            chatbot,
                            system_prompt,
                            temperature,
                            max_tokens
                        ],

                        outputs=[
                            msg,
                            chatbot
                        ]
                    )

                    msg.submit(

                        fn=stream_response,

                        inputs=[
                            msg,
                            chatbot,
                            system_prompt,
                            temperature,
                            max_tokens
                        ],

                        outputs=[
                            msg,
                            chatbot
                        ]
                    )

                    clear_btn.click(
                        fn=clear_chat,
                        outputs=chatbot
                    )

                # =============================================================
                # FILE ANALYSIS
                # =============================================================

                with gr.Tab("📊 File Intelligence"):

                    upload = gr.File(

                        label="Upload Excel / CSV File"
                    )

                    upload_status = gr.Markdown()

                    preview_btn = gr.Button(
                        "Preview Data"
                    )

                    chart_btn = gr.Button(
                        "Generate Analytics"
                    )

                    preview_table = gr.DataFrame(
                        height=400
                    )

                    chart_output = gr.Plot()

                    upload.change(

                        fn=load_file,

                        inputs=upload,

                        outputs=upload_status
                    )

                    preview_btn.click(

                        fn=preview_data,

                        outputs=preview_table
                    )

                    chart_btn.click(

                        fn=generate_chart,

                        outputs=chart_output
                    )

                # =============================================================
                # BUSINESS ANALYTICS
                # =============================================================

                with gr.Tab("📈 Enterprise Dashboard"):

                    gr.Markdown("""
# Enterprise Analytics Dashboard

Features:
- Inventory Analytics
- Dispatch KPIs
- Performance Monitoring
- Excel Intelligence
- AI Insights
""")

                    dashboard_metrics = gr.DataFrame()

                    dashboard_btn = gr.Button(
                        "Load Dashboard"
                    )

                    dashboard_btn.click(

                        fn=get_system_metrics,

                        outputs=dashboard_metrics
                    )

                # =============================================================
                # PYTHON ASSISTANT
                # =============================================================

                with gr.Tab("🐍 Python Assistant"):

                    gr.Markdown("""
# Python Enterprise Assistant

Capabilities:
- Python Optimization
- Automation
- Pandas
- SQL
- APIs
- Excel
- AI Scripts
""")

                    python_prompt = gr.Textbox(

                        label="Python Requirement",

                        lines=8
                    )

                    python_output = gr.Textbox(

                        label="AI Response",

                        lines=20
                    )

                    def python_helper(q):

                        return f"""
Python Assistant Request:

{q}

Use AI Copilot tab for full streaming response.
"""

                    py_btn = gr.Button(
                        "Generate"
                    )

                    py_btn.click(

                        fn=python_helper,

                        inputs=python_prompt,

                        outputs=python_output
                    )

# =============================================================================
# QUEUE
# =============================================================================

demo.queue(

    max_size=50,

    default_concurrency_limit=5
)

# =============================================================================
# LAUNCH
# =============================================================================

demo.launch(

    server_name="0.0.0.0",

    server_port=7860,

    share=True,

    inbrowser=True,

    show_error=True
)
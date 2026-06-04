"""
╔══════════════════════════════════════════════════════════════════════════════╗
║           🇮🇳  INDIAN ADDRESS VALIDATOR — Complete Python Script            ║
║        AI-powered · JustDial/Google Maps Format · Bulk CSV Support          ║
╚══════════════════════════════════════════════════════════════════════════════╝

Requirements:
    pip install anthropic pandas gradio tqdm colorama rich

Usage:
    python indian_address_validator.py

Author  : Indian Address Validator
Version : 2.0.0
"""

# ──────────────────────────────────────────────────────────────────────────────
# IMPORTS
# ──────────────────────────────────────────────────────────────────────────────
import os
import json
import time
import asyncio
import threading
import traceback
from datetime import datetime
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor, as_completed

import anthropic
import pandas as pd
import gradio as gr
from tqdm import tqdm
from colorama import Fore, Style, init as colorama_init
from rich.console import Console
from rich.table import Table
from rich.panel import Panel
from rich.progress import Progress, SpinnerColumn, BarColumn, TextColumn, TimeRemainingColumn

colorama_init(autoreset=True)
console = Console()


# ──────────────────────────────────────────────────────────────────────────────
# CONFIGURATION
# ──────────────────────────────────────────────────────────────────────────────
class Config:
    MODEL_NAME      = "claude-sonnet-4-20250514"   # Claude model
    MAX_TOKENS      = 900                           # Max tokens per API call
    MAX_ROWS        = 100_000                       # Max CSV rows (1 Lakh)
    BATCH_SIZE      = 10                            # Parallel API calls per batch
    RETRY_ATTEMPTS  = 3                             # Retries on API failure
    RETRY_DELAY     = 2.0                           # Seconds between retries
    OUTPUT_DIR      = "output"                      # Output folder
    LOG_FILE        = "validator.log"               # Log file name
    APP_TITLE       = "🇮🇳 Indian Address Validator"
    APP_PORT        = 7860


# ──────────────────────────────────────────────────────────────────────────────
# STATE ABBREVIATION MAP
# ──────────────────────────────────────────────────────────────────────────────
STATE_MAP = {
    "Andhra Pradesh": "AP", "Arunachal Pradesh": "AR", "Assam": "AS",
    "Bihar": "BR", "Chhattisgarh": "CG", "Goa": "GA", "Gujarat": "GJ",
    "Haryana": "HR", "Himachal Pradesh": "HP", "Jharkhand": "JH",
    "Karnataka": "KA", "Kerala": "KL", "Madhya Pradesh": "MP",
    "Maharashtra": "MH", "Manipur": "MN", "Meghalaya": "ML",
    "Mizoram": "MZ", "Nagaland": "NL", "Odisha": "OD", "Punjab": "PB",
    "Rajasthan": "RJ", "Sikkim": "SK", "Tamil Nadu": "TN", "Telangana": "TS",
    "Tripura": "TR", "Uttar Pradesh": "UP", "Uttarakhand": "UK",
    "West Bengal": "WB", "Delhi": "DL", "Jammu & Kashmir": "JK",
    "Ladakh": "LA", "Puducherry": "PY", "Chandigarh": "CH",
    "Andaman & Nicobar": "AN", "Lakshadweep": "LD",
    "Dadra & Nagar Haveli": "DN", "Daman & Diu": "DD",
}


# ──────────────────────────────────────────────────────────────────────────────
# LOGGER
# ──────────────────────────────────────────────────────────────────────────────
class Logger:
    def __init__(self, log_file: str = Config.LOG_FILE):
        self.log_file = log_file

    def _write(self, level: str, message: str):
        ts = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        line = f"[{ts}] [{level}] {message}\n"
        with open(self.log_file, "a", encoding="utf-8") as f:
            f.write(line)

    def info(self, msg):
        self._write("INFO", msg)
        console.print(f"[cyan]ℹ[/cyan] {msg}")

    def success(self, msg):
        self._write("SUCCESS", msg)
        console.print(f"[green]✓[/green] {msg}")

    def warning(self, msg):
        self._write("WARNING", msg)
        console.print(f"[yellow]⚠[/yellow] {msg}")

    def error(self, msg):
        self._write("ERROR", msg)
        console.print(f"[red]✗[/red] {msg}")


logger = Logger()


# ──────────────────────────────────────────────────────────────────────────────
# CLAUDE API CLIENT
# ──────────────────────────────────────────────────────────────────────────────
class AddressValidatorAPI:
    """Handles all Claude API calls for address validation."""

    def __init__(self, api_key: str):
        if not api_key or not api_key.strip():
            raise ValueError("Anthropic API key is required.")
        self.client = anthropic.Anthropic(api_key=api_key.strip())

    def _build_prompt(self, address: str) -> str:
        return f"""You are an expert Indian address validator with deep knowledge of JustDial,
Google Maps, Indian geography and postal codes.

Analyze and fully standardize this Indian address:
Input: "{address}"

Tasks:
1. CLEAN  — Remove duplicate city/area names and redundant words
             (Area, Block, Zone, Near, Opp, Beside, Main, Central, Locality, Prime, Region).
2. FORMAT — Standard JustDial pattern:
             [Shop/Plot details], [Landmark], [Street], [City], [State Abbrev] - [PIN]
3. AUTO-FILL (very important — never skip):
   • City    — infer from landmarks, area names, or context if missing
   • State   — infer from city; use 2-letter abbreviation
              (MH GJ UP TN KA DL RJ MP HR PB WB AP TS KL BR OD AS JH CG UK HP GA
               MN ML TR SK AR NL MZ DN DD JK LA PY AN CH LD)
   • PIN     — provide the most accurate known 6-digit Indian PIN for that area
4. Fix city names  (e.g. "Noida U.P." → "Noida", "New Bombay" → "Navi Mumbai").
5. Return status "failed" ONLY if address is completely unrecognizable.

Respond with ONLY a raw JSON object — no markdown, no backticks, no explanation:
{{
  "original":       "{address}",
  "corrected":      "<full standardized address>",
  "city":           "<city name>",
  "state":          "<full state name>",
  "state_abbrev":   "<2-letter abbrev>",
  "pin":            "<6-digit PIN>",
  "city_inferred":  true or false,
  "state_inferred": true or false,
  "pin_inferred":   true or false,
  "source":         "Indian Address Validator",
  "status":         "success" or "failed",
  "changes":        ["change 1", "change 2"]
}}"""

    def validate(self, address: str, retries: int = Config.RETRY_ATTEMPTS) -> dict:
        """Call Claude API to validate a single address. Retries on failure."""
        address = address.strip()
        if not address:
            return self._empty_result(address, "Empty address")

        for attempt in range(1, retries + 1):
            try:
                response = self.client.messages.create(
                    model=Config.MODEL_NAME,
                    max_tokens=Config.MAX_TOKENS,
                    messages=[{"role": "user", "content": self._build_prompt(address)}],
                )
                raw = "".join(
                    block.text for block in response.content if hasattr(block, "text")
                ).strip()
                raw = raw.replace("```json", "").replace("```", "").strip()
                result = json.loads(raw)
                result.setdefault("source", "Indian Address Validator")
                return result

            except json.JSONDecodeError as e:
                logger.warning(f"JSON parse error (attempt {attempt}): {e}")
            except anthropic.RateLimitError:
                wait = Config.RETRY_DELAY * attempt
                logger.warning(f"Rate limit hit. Waiting {wait}s before retry {attempt}...")
                time.sleep(wait)
            except anthropic.APIError as e:
                logger.error(f"API error (attempt {attempt}): {e}")
                if attempt == retries:
                    return self._empty_result(address, str(e))
            except Exception as e:
                logger.error(f"Unexpected error (attempt {attempt}): {e}")
                if attempt == retries:
                    return self._empty_result(address, str(e))
            time.sleep(Config.RETRY_DELAY)

        return self._empty_result(address, "Max retries exceeded")

    @staticmethod
    def _empty_result(address: str, reason: str) -> dict:
        return {
            "original": address, "corrected": "", "city": "", "state": "",
            "state_abbrev": "", "pin": "", "city_inferred": False,
            "state_inferred": False, "pin_inferred": False,
            "source": "Indian Address Validator", "status": "failed",
            "changes": [], "error": reason,
        }


# ──────────────────────────────────────────────────────────────────────────────
# BULK CSV PROCESSOR
# ──────────────────────────────────────────────────────────────────────────────
class BulkProcessor:
    """Processes large CSV files with multi-threaded API calls."""

    def __init__(self, api: AddressValidatorAPI):
        self.api = api
        Path(Config.OUTPUT_DIR).mkdir(exist_ok=True)

    def process(
        self,
        df: pd.DataFrame,
        address_col: str,
        id_col: str | None,
        progress_callback=None,
    ) -> tuple[pd.DataFrame, pd.DataFrame]:
        """
        Process a DataFrame of addresses.
        Returns: (output_df, error_df)
        """
        total = len(df)
        results = [None] * total
        errors = []
        done_count = [0]
        lock = threading.Lock()

        def process_row(idx: int, row):
            raw = str(row[address_col]).strip() if pd.notna(row[address_col]) else ""
            result = self.api.validate(raw)
            with lock:
                results[idx] = result
                done_count[0] += 1
                if progress_callback:
                    progress_callback(done_count[0], total, result)
            return idx, result

        with ThreadPoolExecutor(max_workers=Config.BATCH_SIZE) as executor:
            futures = {
                executor.submit(process_row, i, row): i
                for i, row in df.iterrows()
            }
            for future in as_completed(futures):
                try:
                    future.result()
                except Exception as e:
                    logger.error(f"Row processing error: {e}")

        # Build output DataFrame
        output_rows = []
        for i, (_, row) in enumerate(df.iterrows()):
            r = results[i] or AddressValidatorAPI._empty_result(
                str(row.get(address_col, "")), "Processing error"
            )
            out = row.to_dict()
            out.update({
                "corrected_address": r.get("corrected", ""),
                "city":              r.get("city", ""),
                "state":             r.get("state", ""),
                "state_abbrev":      r.get("state_abbrev", ""),
                "pin":               r.get("pin", ""),
                "city_inferred":     "Yes" if r.get("city_inferred") else "No",
                "state_inferred":    "Yes" if r.get("state_inferred") else "No",
                "pin_inferred":      "Yes" if r.get("pin_inferred") else "No",
                "validation_status": r.get("status", "failed"),
                "changes":           " | ".join(r.get("changes", [])),
            })
            output_rows.append(out)
            if r.get("status") == "failed":
                row_id = str(row[id_col]) if id_col and id_col in row else str(i + 1)
                errors.append({
                    "row_number": i + 1,
                    "row_id":     row_id,
                    "original_address": str(row.get(address_col, "")),
                    "reason":     r.get("error", "Validation failed"),
                })

        output_df = pd.DataFrame(output_rows)
        error_df  = pd.DataFrame(errors) if errors else pd.DataFrame(
            columns=["row_number", "row_id", "original_address", "reason"]
        )
        return output_df, error_df

    @staticmethod
    def save(df: pd.DataFrame, filename: str) -> str:
        path = Path(Config.OUTPUT_DIR) / filename
        df.to_csv(path, index=False, encoding="utf-8-sig")
        return str(path)


# ──────────────────────────────────────────────────────────────────────────────
# GRADIO UI
# ──────────────────────────────────────────────────────────────────────────────
def build_ui():
    """Build and return the Gradio app."""

    # ── Shared state ──────────────────────────────────────────────────────────
    api_holder   = {"instance": None}
    bulk_state   = {"output_df": None, "error_df": None}

    # ── Helpers ───────────────────────────────────────────────────────────────
    def get_api(key: str):
        if not key.strip():
            raise gr.Error("❌ Please enter your Anthropic API key first.")
        if (
            api_holder["instance"] is None
            or api_holder["instance"].client.api_key != key.strip()
        ):
            api_holder["instance"] = AddressValidatorAPI(key)
        return api_holder["instance"]

    def fmt_bool(val) -> str:
        return "✅ Yes" if val else "—"

    # ── Single validation ─────────────────────────────────────────────────────
    def validate_single(api_key: str, address: str):
        if not address.strip():
            return (
                "⚠️ Please enter an address.", "", "", "", "", "", "", "",
                gr.update(visible=False), gr.update(visible=False),
            )
        try:
            api = get_api(api_key)
            r   = api.validate(address)
            ok  = r["status"] == "success"
            status_html = (
                f'<div style="padding:14px 20px;border-radius:12px;font-weight:700;font-size:1.05em;'
                f'background:{"linear-gradient(135deg,#138808,#27ae60)" if ok else "linear-gradient(135deg,#e74c3c,#c0392b)"};'
                f'color:white">{"✅ Address standardized & auto-completed" if ok else "❌ Could not validate address"}</div>'
            )
            changes_md = (
                "\n".join(f"- {c}" for c in r.get("changes", [])) or "_No changes recorded._"
            )
            inferred_html = ""
            if ok:
                tags = []
                if r.get("city_inferred"):
                    tags.append(f'<span style="background:#fff3cd;color:#856404;padding:4px 12px;border-radius:20px;font-size:.85em;font-weight:700;border:1px solid #ffc107">🏙️ City Auto-filled</span>')
                if r.get("state_inferred"):
                    tags.append(f'<span style="background:#d1ecf1;color:#0c5460;padding:4px 12px;border-radius:20px;font-size:.85em;font-weight:700;border:1px solid #17a2b8">📍 State Auto-filled</span>')
                if r.get("pin_inferred"):
                    tags.append(f'<span style="background:#d4edda;color:#155724;padding:4px 12px;border-radius:20px;font-size:.85em;font-weight:700;border:1px solid #28a745">📮 PIN Auto-filled</span>')
                if tags:
                    inferred_html = '<div style="display:flex;gap:8px;flex-wrap:wrap;margin-top:8px">' + "".join(tags) + "</div>"

            json_out = json.dumps({
                "original":     r.get("original"),
                "corrected":    r.get("corrected"),
                "city":         r.get("city"),
                "state":        r.get("state"),
                "state_abbrev": r.get("state_abbrev"),
                "pin":          r.get("pin"),
                "status":       r.get("status"),
                "source":       r.get("source"),
            }, indent=2, ensure_ascii=False)

            return (
                status_html,
                r.get("corrected", ""),
                r.get("city", ""),
                f'{r.get("state","")} ({r.get("state_abbrev","")})',
                r.get("pin", ""),
                fmt_bool(r.get("city_inferred")),
                fmt_bool(r.get("state_inferred")),
                fmt_bool(r.get("pin_inferred")),
                gr.update(value=changes_md,  visible=True),
                gr.update(value=json_out,    visible=True),
            )
        except Exception as e:
            logger.error(traceback.format_exc())
            return (f'<div style="padding:14px;border-radius:12px;background:#e74c3c;color:white;font-weight:700">❌ Error: {e}</div>',
                    "", "", "", "", "", "", "",
                    gr.update(visible=False), gr.update(visible=False))

    # ── Bulk CSV ──────────────────────────────────────────────────────────────
    def load_csv(file):
        if file is None:
            return gr.update(choices=[], value=None), gr.update(choices=[], value=None), "No file uploaded."
        try:
            df = pd.read_csv(file.name, encoding="utf-8", on_bad_lines="skip")
            if len(df) > Config.MAX_ROWS:
                return (gr.update(choices=[], value=None), gr.update(choices=[], value=None),
                        f"❌ File has {len(df):,} rows. Maximum allowed is {Config.MAX_ROWS:,} (1 Lakh).")
            cols = list(df.columns)
            addr_default = next(
                (c for c in cols if any(k in c.lower() for k in ["address","addr","location","add"])),
                cols[0] if cols else None,
            )
            id_default = next((c for c in cols if c.lower() in ["id","row_id","sr"]), None)
            info = (
                f"✅ **File loaded:** `{Path(file.name).name}` · "
                f"**{len(df):,} rows** · **{len(cols)} columns** · "
                f"Columns: `{'`, `'.join(cols[:8])}{'...' if len(cols)>8 else ''}`"
            )
            return (
                gr.update(choices=cols, value=addr_default),
                gr.update(choices=["— None —"] + cols, value=id_default or "— None —"),
                info,
            )
        except Exception as e:
            return gr.update(choices=[], value=None), gr.update(choices=[], value=None), f"❌ Error reading file: {e}"

    def run_bulk(api_key: str, file, addr_col: str, id_col: str, progress=gr.Progress()):
        if file is None:
            raise gr.Error("Please upload a CSV file.")
        if not addr_col:
            raise gr.Error("Please select the address column.")

        try:
            api = get_api(api_key)
            df  = pd.read_csv(file.name, encoding="utf-8", on_bad_lines="skip")
            if len(df) > Config.MAX_ROWS:
                raise gr.Error(f"File exceeds {Config.MAX_ROWS:,} rows limit.")

            total   = len(df)
            id_col_real = id_col if id_col != "— None —" else None
            done    = [0]; success = [0]; failed = [0]

            def cb(d, t, r):
                done[0] = d
                if r.get("status") == "success":
                    success[0] += 1
                else:
                    failed[0] += 1
                progress(d / t, desc=f"Processing {d:,}/{t:,} · ✅ {success[0]:,} · ❌ {failed[0]:,}")

            processor = BulkProcessor(api)
            output_df, error_df = processor.process(df, addr_col, id_col_real, cb)

            ts      = datetime.now().strftime("%Y%m%d_%H%M%S")
            out_path = processor.save(output_df, f"validated_output_{ts}.csv")
            err_path = processor.save(error_df,  f"error_log_{ts}.csv") if not error_df.empty else None

            bulk_state["output_df"] = output_df
            bulk_state["error_df"]  = error_df

            summary = (
                f"### ✅ Bulk Validation Complete\n\n"
                f"| Metric | Value |\n|---|---|\n"
                f"| Total Rows | {total:,} |\n"
                f"| ✅ Success | {success[0]:,} |\n"
                f"| ❌ Failed  | {failed[0]:,} |\n"
                f"| Success Rate | {success[0]/total*100:.1f}% |\n"
                f"| Output File | `{out_path}` |\n"
                f"| Error Log   | `{err_path or 'None'}` |"
            )

            preview = output_df[[
                c for c in [addr_col,"corrected_address","city","state","pin","validation_status"]
                if c in output_df.columns
            ]].head(10)

            return (
                summary,
                preview,
                gr.update(value=out_path, visible=True),
                gr.update(value=err_path, visible=err_path is not None),
            )

        except gr.Error:
            raise
        except Exception as e:
            logger.error(traceback.format_exc())
            raise gr.Error(f"Processing failed: {e}")

    def make_sample():
        path = Path(Config.OUTPUT_DIR) / "sample_input.csv"
        path.parent.mkdir(exist_ok=True)
        sample = pd.DataFrame({
            "id": range(1, 7),
            "raw_address": [
                "SHREEJI INFOTECH JAMNAGAR JAMNAGAR JAMNAGAR, Jamnagar Area Block",
                "Plot No 12, Sector 5, Noida, Uttar Pradesh",
                "Shop No 45, MG Road, Mumbai, Near Churchgate Station",
                "12, Linking Road, Bandra",
                "Survey No 32, Koramangala, Bengaluru",
                "House No 7, Civil Lines, Allahabad",
            ],
        })
        sample.to_csv(path, index=False)
        return str(path)

    # ── GRADIO LAYOUT ─────────────────────────────────────────────────────────
    with gr.Blocks(
        title=Config.APP_TITLE,
        theme=gr.themes.Soft(
            primary_hue="orange",
            secondary_hue="green",
            neutral_hue="slate",
        ),
        css="""
        .header-html { text-align:center; padding:10px 0 4px; }
        .header-html h1 { font-size:2em; color:#FF9933; margin-bottom:4px; }
        .header-html p { color:#555; font-size:1em; }
        .tricolor { height:6px; background:linear-gradient(90deg,#FF9933 33%,white 33%,white 66%,#138808 66%);
                    border-radius:4px; margin:10px 0; }
        .result-box { border:2px solid #138808 !important; border-radius:12px !important; }
        .inferred-box { font-size:.85em; color:#856404; }
        footer { display:none !important; }
        """
    ) as app:

        # Header
        gr.HTML("""
        <div class="header-html">
          <h1>🇮🇳 Indian Address Validator</h1>
          <p>AI-powered standardization · Auto-fills City, State & PIN · Bulk CSV up to 1 Lakh rows</p>
        </div>
        <div class="tricolor"></div>
        """)

        # API Key
        with gr.Row():
            api_key_box = gr.Textbox(
                label="🔑 Anthropic API Key",
                placeholder="sk-ant-api03-...",
                type="password",
                info="Get your key from console.anthropic.com. Never shared or stored.",
                scale=3,
            )
            gr.HTML("""<div style="padding:30px 0 0 10px;font-size:.88em;color:#888">
                Your key is used only for API calls in this session.</div>""")

        gr.HTML('<div class="tricolor"></div>')

        with gr.Tabs():

            # ── Tab 1: Single Address ─────────────────────────────────────────
            with gr.TabItem("🔍 Single Address"):
                with gr.Row():
                    with gr.Column(scale=1):
                        addr_input = gr.Textbox(
                            label="📋 Raw Indian Address",
                            placeholder="Paste any Indian address — even incomplete ones!\ne.g., Shop No 3, Madhav Complex, Jamnagar Jamnagar\nMissing fields will be auto-filled by AI...",
                            lines=5,
                        )
                        gr.Examples(
                            examples=[
                                ["SHREEJI INFOTECH JAMNAGAR JAMNAGAR JAMNAGAR, Jamnagar Area Block"],
                                ["Plot No 12, Sector 5, Noida, Uttar Pradesh"],
                                ["Shop No 45, MG Road, Mumbai, Near Churchgate Station"],
                                ["12, Linking Road, Bandra"],
                                ["Survey No 32, Koramangala, Bengaluru Karnataka"],
                            ],
                            inputs=[addr_input],
                            label="📌 Example Addresses",
                        )
                        validate_btn = gr.Button("✨ Validate & Standardize", variant="primary", size="lg")

                    with gr.Column(scale=1):
                        status_html  = gr.HTML(label="Status")
                        corrected_out = gr.Textbox(label="✅ Standardized Address", lines=3, elem_classes=["result-box"])
                        with gr.Row():
                            city_out  = gr.Textbox(label="🏙️ City",  scale=2)
                            state_out = gr.Textbox(label="📍 State", scale=2)
                            pin_out   = gr.Textbox(label="📮 PIN",   scale=1)
                        with gr.Row():
                            city_inf  = gr.Textbox(label="City Auto-filled?",  interactive=False, scale=1)
                            state_inf = gr.Textbox(label="State Auto-filled?", interactive=False, scale=1)
                            pin_inf   = gr.Textbox(label="PIN Auto-filled?",   interactive=False, scale=1)
                        changes_md = gr.Markdown(label="🔧 Changes Made", visible=False)
                        json_out   = gr.Code(language="json", label="{ } JSON Output", visible=False)

                validate_btn.click(
                    fn=validate_single,
                    inputs=[api_key_box, addr_input],
                    outputs=[status_html, corrected_out, city_out, state_out, pin_out,
                             city_inf, state_inf, pin_inf, changes_md, json_out],
                )

            # ── Tab 2: Bulk CSV ───────────────────────────────────────────────
            with gr.TabItem("📂 Bulk CSV (up to 1 Lakh rows)"):
                with gr.Row():
                    sample_btn  = gr.Button("⬇️ Download Sample Input CSV", size="sm", variant="secondary")
                    sample_file = gr.File(label="Sample CSV", visible=False, interactive=False)

                sample_btn.click(fn=make_sample, outputs=[sample_file]).then(
                    fn=lambda p: gr.update(visible=True, value=p), inputs=[sample_file], outputs=[sample_file]
                )

                csv_upload = gr.File(label="📁 Upload CSV File", file_types=[".csv"])
                csv_info   = gr.Markdown("_Upload a CSV file to begin._")

                with gr.Row():
                    col_addr = gr.Dropdown(label="Address Column *",        choices=[], interactive=True, scale=2)
                    col_id   = gr.Dropdown(label="Row ID Column (optional)", choices=[], interactive=True, scale=1)

                csv_upload.change(fn=load_csv, inputs=[csv_upload], outputs=[col_addr, col_id, csv_info])

                bulk_btn = gr.Button("🚀 Start Bulk Validation", variant="primary", size="lg")

                bulk_summary = gr.Markdown()
                bulk_preview = gr.Dataframe(
                    label="📋 Preview (first 10 rows)",
                    interactive=False,
                    wrap=True,
                )
                with gr.Row():
                    dl_output = gr.File(label="⬇️ Download Output CSV", visible=False, interactive=False)
                    dl_errors = gr.File(label="⬇️ Download Error Log",  visible=False, interactive=False)

                bulk_btn.click(
                    fn=run_bulk,
                    inputs=[api_key_box, csv_upload, col_addr, col_id],
                    outputs=[bulk_summary, bulk_preview, dl_output, dl_errors],
                )

            # ── Tab 3: State Reference ────────────────────────────────────────
            with gr.TabItem("🗺️ State Code Reference"):
                gr.Markdown("### 📋 Indian State Abbreviation Map")
                state_data = [[state, abbrev] for state, abbrev in sorted(STATE_MAP.items())]
                gr.Dataframe(
                    value=state_data,
                    headers=["State / UT", "Abbreviation"],
                    interactive=False,
                    col_count=(2, "fixed"),
                )

        gr.HTML("""
        <div class="tricolor"></div>
        <div style="text-align:center;padding:16px;color:#aaa;font-size:.85em">
          🇮🇳 Indian Address Validator · Powered by Claude AI (Anthropic) ·
          <a href="https://console.anthropic.com" target="_blank" style="color:#FF9933">Get API Key</a>
        </div>""")

    return app


# ──────────────────────────────────────────────────────────────────────────────
# CLI MODE (no UI — direct Python usage)
# ──────────────────────────────────────────────────────────────────────────────
def cli_single(api_key: str, address: str):
    """Validate a single address from command line."""
    api = AddressValidatorAPI(api_key)
    console.print(Panel(f"[bold]Input:[/bold] {address}", title="🇮🇳 Indian Address Validator", style="orange3"))
    r = api.validate(address)
    table = Table(show_header=True, header_style="bold green")
    table.add_column("Field", style="cyan", min_width=18)
    table.add_column("Value")
    table.add_row("Corrected",      r.get("corrected","—"))
    table.add_row("City",           r.get("city","—") + (" ★" if r.get("city_inferred") else ""))
    table.add_row("State",          f'{r.get("state","—")} ({r.get("state_abbrev","—")})' + (" ★" if r.get("state_inferred") else ""))
    table.add_row("PIN",            r.get("pin","—") + (" ★" if r.get("pin_inferred") else ""))
    table.add_row("Status",         "✅ success" if r["status"]=="success" else "❌ failed")
    table.add_row("Changes",        "\n".join(r.get("changes",[])) or "—")
    console.print(table)
    console.print("[dim](★ = auto-filled by AI)[/dim]")
    return r


def cli_bulk(api_key: str, input_csv: str, address_col: str = "raw_address", id_col: str | None = None):
    """Process a CSV file from command line."""
    console.print(Panel(f"[bold]File:[/bold] {input_csv}\n[bold]Address column:[/bold] {address_col}", title="📂 Bulk CSV Processor"))
    df  = pd.read_csv(input_csv, encoding="utf-8", on_bad_lines="skip")
    api = AddressValidatorAPI(api_key)
    processor = BulkProcessor(api)

    with Progress(SpinnerColumn(), BarColumn(), TextColumn("{task.description}"), TimeRemainingColumn()) as prog:
        task = prog.add_task("Processing...", total=len(df))

        def cb(done, total, r):
            prog.update(task, completed=done, description=f"Processing {done}/{total}")

        output_df, error_df = processor.process(df, address_col, id_col, cb)

    ts   = datetime.now().strftime("%Y%m%d_%H%M%S")
    out  = processor.save(output_df, f"validated_output_{ts}.csv")
    err  = processor.save(error_df,  f"error_log_{ts}.csv") if not error_df.empty else None
    console.print(f"\n[green]✅ Done![/green] Output → {out}")
    if err:
        console.print(f"[yellow]⚠ Error log → {err}[/yellow]")
    return output_df, error_df


# ──────────────────────────────────────────────────────────────────────────────
# ENTRY POINT
# ──────────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="🇮🇳 Indian Address Validator")
    parser.add_argument("--mode",        choices=["ui","single","bulk"], default="ui",
                        help="Run mode: ui (Gradio), single, or bulk")
    parser.add_argument("--api-key",     default=os.environ.get("ANTHROPIC_API_KEY",""),
                        help="Anthropic API key (or set ANTHROPIC_API_KEY env var)")
    parser.add_argument("--address",     help="[single mode] Address to validate")
    parser.add_argument("--input-csv",   help="[bulk mode] Path to input CSV")
    parser.add_argument("--address-col", default="raw_address", help="[bulk mode] Address column name")
    parser.add_argument("--id-col",      default=None,          help="[bulk mode] ID column name")
    parser.add_argument("--port",        type=int, default=Config.APP_PORT, help="[ui mode] Port")
    parser.add_argument("--share",       action="store_true", help="[ui mode] Create public Gradio link")
    args = parser.parse_args()

    if args.mode == "ui":
        console.print(Panel(
            f"[bold orange3]Starting Indian Address Validator UI[/bold orange3]\n"
            f"[dim]Port: {args.port} · Share: {args.share}[/dim]\n"
            f"[dim]Set ANTHROPIC_API_KEY env var or enter in the UI[/dim]",
            title="🇮🇳 Indian Address Validator"
        ))
        app = build_ui()
        app.launch(
            server_port=7861,
            share=args.share,
          inbrowser=True,
        )

    elif args.mode == "single":
        if not args.address:
            console.print("[red]--address is required in single mode[/red]")
        else:
            cli_single(args.api_key, args.address)

    elif args.mode == "bulk":
        if not args.input_csv:
            console.print("[red]--input-csv is required in bulk mode[/red]")
        else:
            cli_bulk(args.api_key, args.input_csv, args.address_col, args.id_col)

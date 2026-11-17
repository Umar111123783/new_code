import os
import json
import re
from pathlib import Path
from typing import Dict, Any, List, Optional, Tuple
import requests
import streamlit as st

from parser import (
    extract_zip,
    build_xml_context,
    get_dll_files,
    get_xaml_files,
    get_important_files,
    get_project_tree,
    sanitize_robin_script,
    apply_robin_node_metadata,
    extract_robin_metadata_context,
)
from llm import SYSTEM_PROMPT

# ==================================================
# Streamlit Page Config
# ==================================================
st.set_page_config(
    page_title="UiPath → PAD Robin (Google Gemini)",
    page_icon="✨",
    layout="wide",
)

st.title("✨ UiPath XAML → PAD Robin Converter (Google Gemini Engine)")
st.caption(
    "Direct Google Gemini conversion of UiPath projects into clean, directly pasteable Power Automate Desktop (Robin) scripts."
)

DEFAULT_GEMINI_ENDPOINT = "https://generativelanguage.googleapis.com/v1beta"
DEFAULT_MODELS = [
    "gemini-1.5-flash",
    "gemini-1.5-flash-8b",
    "gemini-3.8-flash",
    "gemini-1.5-pro",
    "gemini-2.5-flash",
    "Custom...",
]

# State initialization
if "gemini_api_key" not in st.session_state:
    st.session_state["gemini_api_key"] = (
        os.environ.get("GEMINI_API_KEY", "")
        or os.environ.get("GOOGLE_API_KEY", "")
    )
if "gemini_model_choice" not in st.session_state:
    st.session_state["gemini_model_choice"] = "gemini-1.5-flash"
if "gemini_endpoint" not in st.session_state:
    st.session_state["gemini_endpoint"] = os.environ.get(
        "GEMINI_ENDPOINT", DEFAULT_GEMINI_ENDPOINT
    )


def compress_xaml_for_speed(xaml_content: str) -> str:
    """Strip visual layout designer bloat (HintSize, ViewState, IdRef) to speed up generation by 3x-5x."""
    # Remove ViewState XML blocks
    cleaned = re.sub(r'<sap:WorkflowViewStateService\.ViewState>.*?</sap:WorkflowViewStateService\.ViewState>', '', xaml_content, flags=re.DOTALL)
    cleaned = re.sub(r'<sap2010:WorkflowViewState\.IdRef>.*?</sap2010:WorkflowViewState\.IdRef>', '', cleaned, flags=re.DOTALL)
    # Remove inline layout attributes
    cleaned = re.sub(r'\s*sap2010:WorkflowViewState\.IdRef="[^"]*"', '', cleaned)
    cleaned = re.sub(r'\s*sap:VirtualizedContainerService\.HintSize="[^"]*"', '', cleaned)
    cleaned = re.sub(r'\s*mc:Ignorable="[^"]*"', '', cleaned)
    cleaned = re.sub(r'\n\s*\n', '\n', cleaned)
    return cleaned


def _build_gemini_url(base_endpoint: str, model_name: str, api_key: str) -> str:
    """Build full Gemini generateContent URL."""
    base = (base_endpoint or DEFAULT_GEMINI_ENDPOINT).strip().rstrip("/")
    if ":generatecontent" in base.lower():
        url = base
    else:
        if not base.lower().endswith("/models"):
            base = f"{base}/models"
        url = f"{base}/{model_name}:generateContent"

    sep = "&" if "?" in url else "?"
    if api_key and "key=" not in url.lower():
        url = f"{url}{sep}key={api_key.strip()}"
    return url


def call_gemini_generate(
    api_key: str,
    model_name: str,
    prompt_content: str,
    temperature: float = 0.1,
    base_endpoint: str = DEFAULT_GEMINI_ENDPOINT,
) -> Tuple[str, Optional[str], str]:
    """Call Google Gemini generateContent REST endpoint with auto-fallback for deprecated models."""
    candidates_to_try = [model_name]
    for fallback in ["gemini-3.8-flash", "gemini-1.5-flash", "gemini-1.5-pro", "gemini-2.5-flash"]:
        if fallback not in candidates_to_try:
            candidates_to_try.append(fallback)

    last_error = ""
    used_model = model_name

    payload = {
        "contents": [{"role": "user", "parts": [{"text": prompt_content}]}],
        "systemInstruction": {"parts": [{"text": SYSTEM_PROMPT}]},
        "generationConfig": {
            "temperature": float(temperature),
            "maxOutputTokens": 8192,
        },
    }

    for current_model in candidates_to_try:
        url = _build_gemini_url(base_endpoint, current_model, api_key)
        try:
            resp = requests.post(
                url,
                json=payload,
                headers={"Content-Type": "application/json"},
                timeout=180,
            )
        except requests.RequestException as exc:
            return "", f"Connection error: {exc}", current_model

        if resp.status_code == 200:
            try:
                data = resp.json()
            except ValueError:
                return "", "Invalid JSON received from Gemini.", current_model

            candidates = data.get("candidates", [])
            if not candidates:
                return "", f"No candidates returned: {resp.text}", current_model

            parts = candidates[0].get("content", {}).get("parts", [])
            response_text = "\n".join(p.get("text", "") for p in parts if isinstance(p, dict))
            if not response_text.strip():
                return "", "Gemini returned empty candidate text.", current_model

            return response_text, None, current_model

        # If model not found (404), try next supported model
        if resp.status_code == 404:
            last_error = f"Gemini API Error (404): {resp.text}"
            continue

        return "", f"Gemini API Error ({resp.status_code}): {resp.text}", current_model

    return "", last_error, used_model


def parse_gemini_response(
    response_text: str,
    include_node_metadata: bool = False,
    workflow_name: str = "Workflow",
    var_types: Optional[Dict[str, str]] = None,
) -> Tuple[str, str]:
    """Extract and sanitize summary and robin script from raw Gemini output."""
    clean_resp = response_text.strip()
    if "```json" in clean_resp:
        clean_resp = clean_resp.split("```json", 1)[1].split("```", 1)[0].strip()
    elif "```" in clean_resp:
        clean_resp = clean_resp.split("```", 1)[1].split("```", 1)[0].strip()

    summary = ""
    robin_script: Any = ""

    try:
        parsed = json.loads(clean_resp)
        if isinstance(parsed, dict):
            summary = str(parsed.get("summary", ""))
            robin_script = parsed.get("robin_script", "")
    except Exception:
        if "```" in response_text:
            blocks = re.findall(r'```(?:robin|text)?\s*(.*?)\s*```', response_text, re.DOTALL)
            if blocks:
                robin_script = blocks[0].strip()
                summary = re.sub(r'```(?:robin|text)?\s*.*?\s*```', '', response_text, flags=re.DOTALL).strip()
        if not robin_script and (
            "SET " in response_text
            or "Display." in response_text
            or "IF " in response_text
        ):
            robin_script = response_text.strip()
            summary = "Robin script generated directly by Gemini."

    if isinstance(robin_script, (list, dict)):
        robin_script = json.dumps(robin_script, indent=2, ensure_ascii=False)

    cleaned = sanitize_robin_script(
        str(robin_script or ""),
        keep_node_metadata=False,
        preserve_blank_lines=False,
    )
    if cleaned and include_node_metadata:
        cleaned = apply_robin_node_metadata(
            cleaned,
            workflow_name=workflow_name,
            var_types=var_types,
            reset_per_subflow=True,
        )
        cleaned = sanitize_robin_script(
            cleaned,
            keep_node_metadata=True,
            preserve_blank_lines=True,
        )
    return summary, cleaned


# ==================================================
# Configuration Form
# ==================================================
st.subheader("🔑 Gemini Configuration")
col_model, col_key, col_ep = st.columns([1.5, 2.0, 2.0])

with col_model:
    chosen_model = st.selectbox(
        "Model",
        DEFAULT_MODELS,
        index=0,
        help="Select Google Gemini model.",
    )
    if chosen_model == "Custom...":
        active_model = st.text_input(
            "Custom Model Name",
            value="gemini-3.8-flash",
            help="Type custom Gemini model ID (e.g. gemini-3.8-flash).",
        )
    else:
        active_model = chosen_model

with col_key:
    api_key = st.text_input(
        "Gemini API Key",
        type="password",
        value=st.session_state["gemini_api_key"],
        placeholder="AIzaSy...",
        help="Google AI Studio Gemini API Key.",
    )

with col_ep:
    base_endpoint = st.text_input(
        "Base Endpoint",
        value=st.session_state["gemini_endpoint"],
        help="Default is Google v1beta endpoint. Modify if using a corporate proxy.",
    )

# ==================================================
# Sidebar
# ==================================================
with st.sidebar:
    st.header("⚙️ Generation Settings")
    temperature = st.slider(
        "Temperature",
        min_value=0.0,
        max_value=1.0,
        value=0.1,
        step=0.05,
        help="Lower values (0.0 - 0.2) give deterministic code output.",
    )
    st.markdown("---")
    st.subheader("🧾 Output Format")
    include_node_metadata = st.checkbox(
        "Include # node_id metadata lines",
        value=False,
        help="Default is False for clean canvas paste without comments.",
    )
    st.markdown("---")
    if api_key.strip():
        st.success(f"✅ Ready with `{active_model}`")
    else:
        st.warning("⚠️ Enter your Gemini API Key to convert.")

# ==================================================
# Upload & Conversion
# ==================================================
uploaded = st.file_uploader(
    "Upload UiPath Project (.zip or .nupkg)",
    type=["zip", "nupkg"],
)

if uploaded is not None:
    st.write(f"📄 **File:** `{uploaded.name}` | 📦 **Type:** `{uploaded.type}`")

    try:
        with st.spinner("Extracting project files..."):
            project_folder = extract_zip(uploaded)
            xaml_files = get_xaml_files(project_folder)
            dlls = get_dll_files(project_folder)
            important_files = get_important_files(project_folder)
            project_tree = get_project_tree(project_folder)

        col1, col2, col3 = st.columns(3)
        col1.metric("XAML Workflows", len(xaml_files))
        col2.metric("DLL Dependencies", len(dlls))
        col3.metric("Key Files", len(important_files))

        with st.expander("📁 Project Structure & Workflows"):
            st.code(project_tree, language="text")
            st.write("**Detected Workflows:**")
            for x in xaml_files:
                st.write(f"- `{Path(x).name}`")

        # ==================================================
        # Generate Button
        # ==================================================
        if st.button("🚀 Generate with Gemini", type="primary", use_container_width=True):
            if not api_key.strip():
                st.error("🔑 Please enter your Gemini API Key before generating.")
                st.stop()
            if not active_model.strip():
                st.error("🤖 Please select or enter a Gemini model name.")
                st.stop()

            with st.spinner("Building optimized project context for Gemini..."):
                raw_xml = build_xml_context(project_folder)
                xml_context = compress_xaml_for_speed(raw_xml)
                st.caption(f"Optimized context: {len(xml_context):,} chars (reduced by {len(raw_xml) - len(xml_context):,} chars for fast inference).")

            prompt_content = (
                f"UiPath Project Extracted Context:\n{xml_context}\n\n"
                f"Detected DLL Dependencies:\n{', '.join(dlls) if dlls else 'None'}\n\n"
                "Convert this into clean, directly pasteable Power Automate Desktop (Robin) script and business summary in JSON format."
            )

            with st.spinner(f"Generating Robin Script with {active_model}..."):
                raw_response, error_msg, used_model = call_gemini_generate(
                    api_key=api_key.strip(),
                    model_name=active_model.strip(),
                    prompt_content=prompt_content,
                    temperature=temperature,
                    base_endpoint=base_endpoint.strip(),
                )

            Path("model_response.txt").write_text(
                raw_response or error_msg or "No response.",
                encoding="utf-8",
            )

            if error_msg:
                st.error(f"❌ **Gemini Generation Failed:**\n\n{error_msg}")
                st.stop()

            workflow_name, var_types = extract_robin_metadata_context(xml_context)
            summary, robin_script = parse_gemini_response(
                raw_response,
                include_node_metadata=include_node_metadata,
                workflow_name=workflow_name,
                var_types=var_types,
            )

            if not robin_script.strip():
                st.error("Gemini responded, but did not return a valid Robin script.")
                with st.expander("🔍 View Raw Response"):
                    st.text(raw_response)
                st.stop()

            st.success(f"✨ **Generated by Google Gemini (`{used_model}`)**")

            if summary.strip():
                st.subheader("📋 Executive Summary")
                st.write(summary)

            st.subheader("⚡ Generated Power Automate Desktop (Robin) Script")
            st.info(
                "💡 **How to Copy Without HTML Corruption (`<br>`, `&gt;`, `<a href>`):**\n"
                "1. Click inside the **Plain Text** box below and press **Ctrl + A → Ctrl + C**, or click **Download Robin Script (.txt)**.\n"
                "2. Open your desktop flow in Power Automate Desktop.\n"
                "3. Click on the central canvas and press **Ctrl + V**."
            )

            st.text_area(
                "Plain Text Robin Script (Copy-Safe: Ctrl+A → Ctrl+C — never adds <br> or &gt;)",
                value=robin_script,
                height=340,
            )

            st.download_button(
                label="📥 Download Robin Script (.txt)",
                data=robin_script.encode("utf-8"),
                file_name="robin_script_gemini.txt",
                mime="text/plain",
                use_container_width=True,
            )

            with st.expander("Formatted Code View (Use Plain Text Box Above to Copy)"):
                st.code(robin_script, language="text")

            with st.expander("🔍 View Raw Gemini Model Response"):
                st.text(raw_response)

    except Exception as ex:
        st.error(f"Error processing project: {str(ex)}")


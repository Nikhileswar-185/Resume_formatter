"""
This file is used to build the interface for uploading documents and generate a downloadable standardized resume.

It uses Streamlit for the UI and integrates with the resume_builder module to create the standardized resume.
"""

import os
import json
from datetime import datetime
import streamlit as st

from utils.data_parser import extract_text_from_file
from utils.llm import parse_json
from config.prompts import PromptHolder
from src.resume_builder import resume_builder


st.set_page_config(page_title="Resume Standardization", page_icon="📄", layout="wide")

# Sidebar: configuration and flags
with st.sidebar:
    st.header("Settings")
    
 
    tone = st.selectbox(
        "Tone",
        options=["professional", "formal", "concise", "impactful", "casual"],
        index=0,
    )
    refine_wording = st.toggle(
        "Refine wording via LLM (bonus)",
        value=False,
        help="Optional extra pass to tune phrasing/conciseness",
    )

st.title("Resume Standardizer")

st.markdown(
    "Upload resume files in .pdf or .docx. The app extracts information, standardizes the structure, and produces a new .docx."
)

# Uploads allow multiple resumes in a single run
uploaded_files = st.file_uploader(
    "Upload one or more resumes",
    type=["pdf", "docx"],
    accept_multiple_files=True,
)

# Optional style/voice hints are forwarded to the LLM extraction prompt
extra_instructions = st.text_area(
    "Optional style/voice instructions",
    value=f"Use a {tone} tone. Keep it succinct. Emphasize measurable impact.",
    height=100,
)

col_run, col_opts = st.columns([1, 3])
with col_run:
    run = st.button("Convert to Standardized DOCX", type="primary")

if "latest_resume" not in st.session_state:
    st.session_state.latest_resume = None

if run:
    if not uploaded_files:
        st.warning("Please upload at least one file.")
        st.stop()

    for file in uploaded_files:
        with st.spinner(f"Processing {file.name}..."):
            try:
                file_bytes = file.read()
                ext = os.path.splitext(file.name)[1].lower().lstrip('.')
                
                # Extract and parse
                raw_text = extract_text_from_file(file_bytes, ext)
                data = parse_json(PromptHolder.STRUCTURE_SCHEMA_PROMPT, raw_text)
                
                resume_bytes = resume_builder(data)
                dt = datetime.now().strftime("%Y%m%d_%H%M%S")
                out_name = f"standard_resume_{dt}.docx"

                # Keep only the latest resume
                st.session_state.latest_resume = {
                    "file_name": out_name,
                    "bytes": resume_bytes
                }

            except Exception as e:
                st.exception(e)

# Show only the latest download button
if st.session_state.latest_resume:
    res = st.session_state.latest_resume
    st.download_button(
        label=f"Download {res['file_name']}",
        data=res['bytes'],
        file_name=res['file_name'],
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

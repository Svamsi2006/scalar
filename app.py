"""
app.py - Enterprise PII Redactor Streamlit Web Application

A production-ready web wrapper for the modular PII detection, redaction,
and context-preserving faking engine for Microsoft Word (.docx) documents.
Processes files 100% in-memory without server disk persistence.
"""

import io
import os
import re
import time
import json
import base64
import logging
from typing import Set, Dict, List, Any, Optional

import streamlit as st
import pandas as pd
import docx

# Clean modular imports from existing pipeline
from pii_redactor.document_parser import DocumentParser, extract_full_text
from pii_redactor.detection_engine import DetectionEngine
from pii_redactor.replacement_manager import ReplacementManager

# Configure root logger for Streamlit runtime
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s"
)
logger = logging.getLogger("pii_redactor.app")

@st.cache_resource
def load_spacy_model():
    import spacy
    import subprocess
    import sys
    try:
        return spacy.load("en_core_web_sm")
    except OSError:
        logger.info("spaCy model 'en_core_web_sm' not found. Attempting dynamic download...")
        subprocess.run([sys.executable, "-m", "spacy", "download", "en_core_web_sm"])
        try:
            return spacy.load("en_core_web_sm")
        except Exception as e:
            logger.error(f"Failed to load spaCy model after download: {e}")
            raise e

# ==============================================================================
# 1. PAGE CONFIGURATION & THEME STYLING
# ==============================================================================
st.set_page_config(
    page_title="Enterprise PII Redactor",
    page_icon="🔒",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Custom enterprise CSS styling
st.markdown("""
<style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700;800&display=swap');
    
    html, body, [class*="css"] {
        font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif;
    }
    
    /* Header hero gradient badge */
    .hero-container {
        background: linear-gradient(135deg, #0f172a 0%, #1e293b 50%, #0f172a 100%);
        border: 1px solid #334155;
        border-radius: 16px;
        padding: 2rem 2.5rem;
        margin-bottom: 2rem;
        color: #f8fafc;
        box-shadow: 0 10px 25px -5px rgba(0, 0, 0, 0.3), 0 8px 10px -6px rgba(0, 0, 0, 0.2);
    }
    
    .hero-title {
        font-size: 2.25rem;
        font-weight: 800;
        letter-spacing: -0.025em;
        margin-bottom: 0.5rem;
        display: flex;
        align-items: center;
        gap: 0.75rem;
        background: linear-gradient(135deg, #ffffff 0%, #cbd5e1 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
    }
    
    .hero-subtitle {
        font-size: 1.05rem;
        color: #94a3b8;
        line-height: 1.6;
        margin-bottom: 1.25rem;
        max-width: 900px;
    }
    
    .badge-pill-container {
        display: flex;
        flex-wrap: wrap;
        gap: 0.6rem;
    }
    
    .badge-pill {
        display: inline-flex;
        align-items: center;
        gap: 0.4rem;
        background: rgba(255, 255, 255, 0.08);
        border: 1px solid rgba(255, 255, 255, 0.15);
        padding: 0.35rem 0.85rem;
        border-radius: 9999px;
        font-size: 0.8rem;
        font-weight: 600;
        color: #e2e8f0;
        backdrop-filter: blur(8px);
    }
    
    /* Feature cards */
    .feature-card {
        background: #ffffff;
        border: 1px solid #e2e8f0;
        border-radius: 12px;
        padding: 1.25rem;
        height: 100%;
        box-shadow: 0 1px 3px rgba(0, 0, 0, 0.05);
        transition: transform 0.2s ease, box-shadow 0.2s ease;
    }
    
    .feature-card:hover {
        transform: translateY(-2px);
        box-shadow: 0 8px 16px -4px rgba(0, 0, 0, 0.08);
    }
    
    .card-title {
        font-size: 1rem;
        font-weight: 700;
        color: #0f172a;
        margin-bottom: 0.35rem;
        display: flex;
        align-items: center;
        gap: 0.5rem;
    }
    
    .card-body {
        font-size: 0.875rem;
        color: #64748b;
        line-height: 1.5;
    }
    
    /* Primary action button styling */
    div.stButton > button[kind="primary"],
    div.stButton > button:first-child:not([kind="secondary"]) {
        background: linear-gradient(135deg, #2563eb 0%, #1d4ed8 100%) !important;
        color: #ffffff !important;
        border: none !important;
        border-radius: 10px !important;
        padding: 0.75rem 2rem !important;
        font-weight: 700 !important;
        font-size: 1.05rem !important;
        letter-spacing: 0.01em !important;
        box-shadow: 0 4px 12px rgba(37, 99, 235, 0.35) !important;
        transition: all 0.2s ease !important;
        width: 100%;
    }
    
    div.stButton > button:hover {
        background: linear-gradient(135deg, #1d4ed8 0%, #1e40af 100%) !important;
        transform: translateY(-1px) !important;
        box-shadow: 0 6px 18px rgba(37, 99, 235, 0.45) !important;
    }
    
    /* Export card highlight border styling */
    .export-card {
        background: #ffffff;
        border: 2px solid #3b82f6;
        border-radius: 16px;
        padding: 2rem 2.5rem;
        max-width: 650px;
        margin: 1.5rem auto;
        box-shadow: 0 10px 25px -5px rgba(59, 130, 246, 0.08), 0 8px 10px -6px rgba(59, 130, 246, 0.04);
        text-align: center;
    }
    
    /* Direct download HTML button */
    .direct-download-btn {
        display: block;
        width: 100%;
        text-align: center;
        background: linear-gradient(135deg, #059669 0%, #047857 100%);
        color: #ffffff !important;
        padding: 0.85rem 1.75rem;
        font-weight: 700;
        font-size: 1.05rem;
        border-radius: 10px;
        text-decoration: none !important;
        box-shadow: 0 4px 14px rgba(5, 150, 105, 0.35);
        transition: all 0.2s ease;
        margin-top: 0.5rem;
        margin-bottom: 0.5rem;
    }
    
    .direct-download-btn:hover {
        background: linear-gradient(135deg, #047857 0%, #065f46 100%);
        transform: translateY(-1px);
        box-shadow: 0 6px 20px rgba(5, 150, 105, 0.45);
        color: #ffffff !important;
    }
    
    /* Metric container styling */
    div[data-testid="stMetric"] {
        background: #f8fafc;
        border: 1px solid #e2e8f0;
        border-radius: 12px;
        padding: 1rem 1.25rem;
        box-shadow: 0 1px 2px rgba(0, 0, 0, 0.04);
    }
    
    div[data-testid="stMetricLabel"] {
        font-weight: 600;
        color: #475569;
        font-size: 0.875rem;
    }
    
    div[data-testid="stMetricValue"] {
        font-weight: 800;
        color: #0f172a;
    }
    
    /* Success banner */
    .success-banner {
        background: linear-gradient(135deg, #ecfdf5 0%, #d1fae5 100%);
        border: 1px solid #6ee7b7;
        border-radius: 12px;
        padding: 1.25rem 1.5rem;
        color: #065f46;
        margin-bottom: 1.5rem;
        display: flex;
        align-items: center;
        gap: 0.85rem;
    }
</style>
""", unsafe_allow_html=True)

# ==============================================================================
# 2. PII METADATA & STATIC OPTIONS DEFINITIONS
# ==============================================================================
CATEGORY_DEFINITIONS = {
    "FULL_NAME": {
        "label": "Full Names",
        "icon": "👤",
        "description": "Names of individuals, executives, stakeholders, and personnel"
    },
    "EMAIL": {
        "label": "Email Addresses",
        "icon": "📧",
        "description": "Electronic mail addresses (@domain.com)"
    },
    "PHONE": {
        "label": "Phone Numbers",
        "icon": "📞",
        "description": "International and domestic telephone numbers with country codes"
    },
    "COMPANY": {
        "label": "Companies & Organizations",
        "icon": "🏢",
        "description": "Corporate entities, institutions, and legal bodies"
    },
    "ADDRESS": {
        "label": "Physical Addresses",
        "icon": "📍",
        "description": "Street addresses, suites, postal codes, and geographic locations"
    },
    "SSN": {
        "label": "SSNs / National IDs",
        "icon": "🪪",
        "description": "Social Security Numbers, Tax IDs, and national identifiers"
    },
    "CREDIT_CARD": {
        "label": "Credit Card Numbers",
        "icon": "💳",
        "description": "13-19 digit payment card numbers (Visa, Mastercard, Amex)"
    },
    "DATE": {
        "label": "Dates of Birth / Sensitive Dates",
        "icon": "📅",
        "description": "Birthdates, transaction dates, and calendar timestamps"
    },
    "IP_ADDRESS": {
        "label": "IP Addresses",
        "icon": "🌐",
        "description": "IPv4 network addresses and host indicators"
    },
    "CIN": {
        "label": "Corporate Identity Numbers (CIN)",
        "icon": "🏢",
        "description": "Indian Corporate Identity Numbers"
    },
    "PAN": {
        "label": "Permanent Account Numbers (PAN)",
        "icon": "🪪",
        "description": "Indian Income Tax Permanent Account Numbers"
    },
    "GSTIN": {
        "label": "GSTIN",
        "icon": "📜",
        "description": "Indian Goods and Services Tax Identification Numbers"
    },
    "AADHAAR": {
        "label": "Aadhaar Numbers",
        "icon": "👤",
        "description": "12-digit Indian national identity numbers"
    }
}

REDACTION_MODES = {
    "Contextual Synthetic Replacements (Faker)": "synthetic",
    "Entity Reference Tokens ([NAME_1])": "token",
    "Standard Redaction Mask ([REDACTED: FULL_NAME])": "mask",
    "Visual Blackout (█████)": "blackout",
}

LOCALE_OPTIONS = {
    "English (United States) - en_US": "en_US",
    "English (India) - en_IN": "en_IN",
    "English (United Kingdom) - en_GB": "en_GB",
    "English (Canada) - en_CA": "en_CA",
    "English (Australia) - en_AU": "en_AU",
}

STYLE_PREVIEW_MAP = {
    "synthetic": {
        "desc": "Replaces sensitive data with seeded, contextually consistent fake names, company names, and addresses.",
        "original": "Kushal Hegde from KSH International emailed cs.connect@kshinternational.com.",
        "anon": "John Doe from Acme Corporation emailed john.doe@example.com."
    },
    "token": {
        "desc": "Replaces PII with incrementing serial tags. Preserves relationships and structural occurrences without using random names.",
        "original": "Kushal Hegde from KSH International emailed cs.connect@kshinternational.com.",
        "anon": "[NAME_1] from [COMPANY_1] emailed [EMAIL_1]."
    },
    "mask": {
        "desc": "Replaces PII with generic compliance placeholders representing the data category.",
        "original": "Kushal Hegde from KSH International emailed cs.connect@kshinternational.com.",
        "anon": "[REDACTED: FULL_NAME] from [REDACTED: COMPANY] emailed [REDACTED: EMAIL]."
    },
    "blackout": {
        "desc": "Mimics a physical blackout marker by replacing sensitive characters with standard solid blackout characters (█).",
        "original": "Kushal Hegde from KSH International emailed cs.connect@kshinternational.com.",
        "anon": "████████████ from ████████████████████ emailed ██████████████████████████████."
    }
}

# ==============================================================================
# 3. MAIN HEADER & BRANDING
# ==============================================================================
st.markdown("""
<div class="hero-container">
    <div class="hero-title">
        <span>🔒</span> Enterprise PII Redactor
    </div>
    <div class="hero-subtitle">
        High-performance, format-preserving PII anonymization and contextual faking for Microsoft Word (.docx) documents.
    </div>
    <div class="badge-pill-container">
        <span class="badge-pill">🛡️ Format-Preserved Runs</span>
        <span class="badge-pill">🧠 Hybrid spacy-NER</span>
        <span class="badge-pill">🔁 Stateful Cache</span>
        <span class="badge-pill">⚡ 100% In-Memory Stream</span>
    </div>
</div>
""", unsafe_allow_html=True)

# 3-Step Workflow Banner (Clean dashboard feel)
col_w1, col_w2, col_w3 = st.columns(3)
with col_w1:
    st.markdown("""
    <div class="feature-card">
        <div class="card-title">1️⃣ Upload Any DOCX</div>
        <div class="card-body">Upload any Word document (.docx). The file is processed completely in-memory without server disk storage.</div>
    </div>
    """, unsafe_allow_html=True)
with col_w2:
    st.markdown("""
    <div class="feature-card">
        <div class="card-title">2️⃣ Configure Redaction</div>
        <div class="card-body">Select your preferred redaction style (Standard Mask, Entity Tokens, Blackout) and exclude categories if needed.</div>
    </div>
    """, unsafe_allow_html=True)
with col_w3:
    st.markdown("""
    <div class="feature-card">
        <div class="card-title">3️⃣ Redact & Download</div>
        <div class="card-body">Execute the run-splitting engine, review verified entity mappings, and instantly download your clean document.</div>
    </div>
    """, unsafe_allow_html=True)

st.markdown("<div style='height: 1.5rem;'></div>", unsafe_allow_html=True)

# ==============================================================================
# 4. INITIAL STATE HANDLERS
# ==============================================================================
if "redaction_results" not in st.session_state:
    st.session_state["redaction_results"] = None
if "active_file_bytes" not in st.session_state:
    st.session_state["active_file_bytes"] = None
if "active_filename" not in st.session_state:
    st.session_state["active_filename"] = ""
if "active_filesize_kb" not in st.session_state:
    st.session_state["active_filesize_kb"] = 0.0

# ==============================================================================
# 5. SIDEBAR CONFIGURATOR (COOL UI)
# ==============================================================================
with st.sidebar:
    st.markdown("## ⚙️ Advanced Settings")
    st.caption("Engine parameter tuning & health diagnostics.")
    st.divider()

    # Visual Redaction Style
    st.markdown("#### 🎨 Visual Redaction Style")
    selected_mode_label = st.selectbox(
        "Choose Redaction Mode:",
        options=list(REDACTION_MODES.keys()),
        index=0,
        help="Determines the visual representation of redacted PII."
    )
    selected_mode = REDACTION_MODES[selected_mode_label]

    # Optional seed & locale if synthetic mode is selected
    if selected_mode == "synthetic":
        seed_value = st.slider(
            "Faker Random Seed",
            min_value=0,
            max_value=10000,
            value=42,
            step=1,
            help="Controls Faker's seed for reproducible synthetic values."
        )
        selected_locale_label = st.selectbox(
            "Synthetic Locale Style",
            options=list(LOCALE_OPTIONS.keys()),
            index=1,
            help="Determines regional styling for synthetic names and addresses."
        )
        selected_locale = LOCALE_OPTIONS[selected_locale_label]
    else:
        seed_value = 42
        selected_locale = "en_US"
        st.info("💡 Real Redaction Active: No random or fictitious data will be generated.")

    # Style Preview Container
    preview = STYLE_PREVIEW_MAP[selected_mode]
    st.markdown(f"""
    <div style="background-color: #f8fafc; border: 1px solid #e2e8f0; border-radius: 10px; padding: 0.9rem; margin-top: 0.5rem; margin-bottom: 0.5rem;">
        <span style="font-weight: 700; color: #1e293b; font-size: 0.85rem; display: flex; align-items: center; gap: 0.35rem;">✨ Redaction Preview</span>
        <p style="font-size: 0.775rem; color: #64748b; margin-top: 0.2rem; margin-bottom: 0.6rem; line-height: 1.4;">
            {preview['desc']}
        </p>
        <div style="font-size: 0.775rem; color: #475569; margin-bottom: 0.4rem; line-height: 1.4;">
            <strong>Original:</strong><br/>
            <span style="font-family: monospace; color: #ef4444; background-color: #fef2f2; padding: 0.1rem 0.25rem; border-radius: 4px; display: block; margin-top: 0.15rem; word-break: break-all;">{preview['original']}</span>
        </div>
        <div style="font-size: 0.775rem; color: #475569; line-height: 1.4;">
            <strong>Anonymized:</strong><br/>
            <span style="font-family: monospace; color: #22c55e; background-color: #f0fdf4; padding: 0.1rem 0.25rem; border-radius: 4px; display: block; margin-top: 0.15rem; word-break: break-all;">{preview['anon']}</span>
        </div>
    </div>
    """, unsafe_allow_html=True)

    st.divider()

    # Engine Health & Diagnostics
    st.markdown("#### 🤖 Engine Diagnostics")
    try:
        nlp_test = load_spacy_model()
        st.markdown("""
        <div style="background-color: #f0fdf4; border: 1px solid #bbf7d0; border-radius: 8px; padding: 0.6rem 0.8rem; margin-bottom: 0.5rem; display: flex; align-items: center; gap: 0.5rem; color: #166534; font-size: 0.8rem;">
            <span style="height: 10px; width: 10px; background-color: #22c55e; border-radius: 50%; display: inline-block;"></span>
            <strong>spaCy NER Model Active:</strong> en_core_web_sm loaded
        </div>
        """, unsafe_allow_html=True)
    except Exception:
        st.markdown("""
        <div style="background-color: #fffbeb; border: 1px solid #fef3c7; border-radius: 8px; padding: 0.6rem 0.8rem; margin-bottom: 0.5rem; display: flex; align-items: center; gap: 0.5rem; color: #92400e; font-size: 0.8rem;">
            <span style="height: 10px; width: 10px; background-color: #f59e0b; border-radius: 50%; display: inline-block;"></span>
            <strong>spaCy Model Unavailable:</strong> Regex-only mode
        </div>
        """, unsafe_allow_html=True)

    with st.expander("ℹ️ Architecture Highlights", expanded=False):
        st.markdown("""
        - **Real Document Parsing:** Ingests any Word document, scanning body paragraphs, tables, headers, and footers.
        - **XML Run Splitting:** Splits XML `<w:r>` runs at boundary offsets without corrupting styles.
        - **In-Memory Streams:** Zero disk footprint processing via `io.BytesIO`.
        """)
        
    st.caption("Enterprise PII Redaction Suite • v1.2.0")

# ==============================================================================
# 6. SINGLE-PAGE CONTINUOUS FLOW WORKSPACE
# ==============================================================================

# ------------------------------------------------------------------------------
# SECTION 1: 📂 Ingest & Process
# ------------------------------------------------------------------------------
st.markdown("## 📂 Ingest & Process")

col_upload, col_sample = st.columns([3, 1])

with col_upload:
    uploaded_file = st.file_uploader(
        "Choose any Microsoft Word document (.docx)",
        type=["docx"],
        help="Upload standard DOCX documents containing body paragraphs, tables, headers, or footers."
    )
    if uploaded_file is not None:
        file_data = uploaded_file.getvalue()
        st.session_state["active_file_bytes"] = io.BytesIO(file_data)
        st.session_state["active_filename"] = uploaded_file.name
        st.session_state["active_filesize_kb"] = len(file_data) / 1024.0
        if "use_sample" in st.session_state:
            del st.session_state["use_sample"]
        
sample_file_path = "Red Herring Prospectus.docx"
with col_sample:
    st.markdown("**Quick Demo Document**")
    if os.path.exists(sample_file_path):
        if st.button("📄 Load Demo Prospectus", use_container_width=True):
            st.session_state["use_sample"] = True
            with open(sample_file_path, "rb") as f:
                content = f.read()
                st.session_state["active_file_bytes"] = io.BytesIO(content)
                st.session_state["active_filename"] = sample_file_path
                st.session_state["active_filesize_kb"] = len(content) / 1024.0
    else:
        st.caption("Demo sample document not found in workspace.")
        
# Active Document Card display
if st.session_state["active_file_bytes"] is not None:
    st.markdown(f"""
    <div style="background-color: #f1f5f9; border: 1px solid #cbd5e1; border-radius: 12px; padding: 1.25rem 1.5rem; margin-top: 1rem; margin-bottom: 1.5rem;">
        <h4 style="margin-top: 0; color: #0f172a; font-weight: 700; display: flex; align-items: center; gap: 0.5rem;">
            📄 Active Document Summary
        </h4>
        <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(200px, 1fr)); gap: 1rem; margin-top: 0.75rem;">
            <div>
                <span style="color: #64748b; font-size: 0.8rem; font-weight: 600;">FILENAME</span><br/>
                <code style="font-size: 0.9rem; font-weight: 700; color: #2563eb;">{st.session_state["active_filename"]}</code>
            </div>
            <div>
                <span style="color: #64748b; font-size: 0.8rem; font-weight: 600;">FILE SIZE</span><br/>
                <span style="font-size: 0.9rem; font-weight: 700; color: #334155;">{st.session_state["active_filesize_kb"]:.2f} KB</span>
            </div>
            <div>
                <span style="color: #64748b; font-size: 0.8rem; font-weight: 600;">TARGET REDACTION STYLE</span><br/>
                <span style="font-size: 0.9rem; font-weight: 700; color: #0f172a;">{selected_mode.upper()}</span>
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)
    
    start_redaction = st.button("🚀 Run Redaction Engine", type="primary", use_container_width=True)
    
    if start_redaction:
        start_time = time.time()
        
        with st.spinner("Analyzing document structure, detecting real PII, and performing format-preserving redaction..."):
            try:
                # 1. Reset BytesIO stream position
                st.session_state["active_file_bytes"].seek(0)
                
                # 2. Instantiate core components cleanly
                parser = DocumentParser()
                detector = DetectionEngine()
                replacement_mgr = ReplacementManager(mode=selected_mode, seed=seed_value, locale=selected_locale)
                
                # 3. Load DOCX directly from in-memory stream
                parser.load(st.session_state["active_file_bytes"])
                
                # 4. Extract all paragraphs across Body, Tables, Headers, Footers
                paragraphs = parser.iter_all_paragraphs()
                total_paras = len(paragraphs)
                
                if total_paras == 0:
                    st.warning("⚠️ The uploaded document contains no readable text paragraphs.")
                    st.stop()
                
                # 5. Execute Run-Splitting Redaction with real-time progress
                progress_bar = st.progress(0, text="Parsing document loops...")
                
                for idx, paragraph in enumerate(paragraphs, start=1):
                    full_para_text = extract_full_text(paragraph)
                    
                    if full_para_text.strip():
                        if idx % 3 == 0:
                            progress_text = f"Evaluating Named Entities... Elements: {idx}/{total_paras}"
                        elif idx % 3 == 1:
                            progress_text = f"Replacing formatting runs... Elements: {idx}/{total_paras}"
                        else:
                            progress_text = f"Scanning paragraph blocks... Elements: {idx}/{total_paras}"
                        
                        detections = detector.detect_pii(full_para_text)
                        if detections:
                            parser.process_paragraph(paragraph, detections, replacement_mgr)
                    else:
                        progress_text = f"Scanning paragraph blocks... Elements: {idx}/{total_paras}"
                    
                    if idx % 25 == 0 or idx == total_paras:
                        pct = idx / total_paras
                        progress_bar.progress(pct, text=progress_text)
                
                # 6. Save modified document directly to an in-memory BytesIO buffer
                output_docx_bytes = io.BytesIO()
                parser.save(output_docx_bytes)
                output_docx_bytes.seek(0)
                
                processed_bytes_data = output_docx_bytes.getvalue()
                elapsed_time = time.time() - start_time
                progress_bar.empty()
                
                # Sanitize output filename to prevent browser download GUID issues
                base_name = os.path.splitext(st.session_state["active_filename"])[0]
                safe_name = re.sub(r'[^a-zA-Z0-9_\-]', '_', base_name)
                safe_name = re.sub(r'_+', '_', safe_name).strip('_')
                if not safe_name:
                    safe_name = "document"
                clean_output_filename = f"{safe_name}_redacted.docx"
                
                # Store results in session state for stateful rendering
                st.session_state["redaction_results"] = {
                    "processed_bytes": processed_bytes_data,
                    "filename": st.session_state["active_filename"],
                    "clean_output_filename": clean_output_filename,
                    "elapsed_time": elapsed_time,
                    "parser_stats": parser.get_stats(),
                    "manager_stats": replacement_mgr.get_stats(),
                    "audit_report": replacement_mgr.get_report(),
                    "total_paras": total_paras,
                    "mode": selected_mode
                }
                
                st.success("🎉 Redaction engine execution complete! Scroll down to view the performance metrics, entity audit mapping, and secure export panel.")
                st.rerun()
                
            except Exception as e:
                logger.exception(f"Unexpected error during redaction: {e}")
                st.error(f"❌ An error occurred during document redaction: {str(e)}")
                st.info("💡 Tip: Verify that the document is a valid Microsoft Word .docx file.")
else:
    st.info("💡 Upload any Word document (.docx) or click 'Load Demo Prospectus' above to begin.")

# ------------------------------------------------------------------------------
# SECTION 2: 📊 Performance Metrics
# ------------------------------------------------------------------------------
if st.session_state["redaction_results"] is not None:
    st.divider()
    st.markdown("## 📊 Performance Metrics")
    
    results = st.session_state["redaction_results"]
    manager_stats: Dict[str, int] = results["manager_stats"]
    parser_stats: Dict[str, int] = results["parser_stats"]
    audit_report: Dict[str, Dict[str, str]] = results["audit_report"]
    total_replacements = parser_stats.get("replacements_made", 0)
    total_paras = results["total_paras"]
    elapsed_time = results["elapsed_time"]
    
    # Success Banner
    st.markdown(f"""
    <div class="success-banner">
        <span style="font-size: 1.5rem;">🎉</span>
        <div>
            <div style="font-weight: 700; font-size: 1.1rem;">Anonymization Metrics Ready!</div>
            <div style="font-size: 0.9rem; color: #047857;">
                Processed <strong>{total_paras:,}</strong> document elements and executed <strong>{total_replacements:,}</strong> redactions in <strong>{elapsed_time:.2f}s</strong>.
            </div>
        </div>
    </div>
    """, unsafe_allow_html=True)
    
    # Summary KPI Metric Cards
    st.markdown("### Executive Summary")
    kpi1, kpi2, kpi3, kpi4 = st.columns(4)
    
    with kpi1:
        st.metric(
            label="Elements Processed",
            value=f"{total_paras:,}",
            help="Total body paragraphs, table cells, headers, and footers scanned."
        )
    with kpi2:
        st.metric(
            label="Total Redactions Made",
            value=f"{total_replacements:,}",
            help="Total PII instances replaced with compliance redaction tags."
        )
    with kpi3:
        unique_entities = sum(len(v) for v in audit_report.values())
        st.metric(
            label="Unique PII Entities Mapped",
            value=f"{unique_entities:,}",
            help="Number of distinct real PII values identified and mapped."
        )
    with kpi4:
        st.metric(
            label="Processing Speed",
            value=f"{total_paras / max(elapsed_time, 0.01):.0f} elem/sec",
            delta=f"{elapsed_time:.2f}s elapsed"
        )
        
    st.divider()
    
    # Category Breakdown and Bar Chart
    st.markdown("### Redactions by PII Category")
    if manager_stats:
        cat_cols = st.columns(min(len(manager_stats), 4))
        for i, (cat_key, count) in enumerate(manager_stats.items()):
            col = cat_cols[i % len(cat_cols)]
            meta = CATEGORY_DEFINITIONS.get(cat_key, {"label": cat_key, "icon": "🏷️"})
            col.metric(label=f"{meta['icon']} {meta['label']}", value=count)
        
        st.markdown("<div style='height: 1.5rem;'></div>", unsafe_allow_html=True)
        
        chart_data = []
        for cat_key, count in manager_stats.items():
            meta = CATEGORY_DEFINITIONS.get(cat_key, {"label": cat_key, "icon": "🏷️"})
            chart_data.append({
                "PII Category": f"{meta['icon']} {meta['label']}",
                "Unique Entities Redacted": count
            })
        df_chart = pd.DataFrame(chart_data)
        st.bar_chart(df_chart.set_index("PII Category"), color="#2563eb", use_container_width=True)
    else:
        st.info("ℹ️ No sensitive PII entities were detected in this document based on your active filter criteria.")

# ------------------------------------------------------------------------------
# SECTION 3: 🔍 Interactive Audit Log
# ------------------------------------------------------------------------------
if st.session_state["redaction_results"] is not None:
    st.divider()
    st.markdown("## 🔍 Interactive Audit Log")
    
    results = st.session_state["redaction_results"]
    audit_report: Dict[str, Dict[str, str]] = results["audit_report"]
    clean_out_filename = results["clean_output_filename"]
    
    st.caption("Review the exact mapping between real sensitive data found in the document and the applied redaction.")
    
    audit_rows: List[Dict[str, str]] = []
    for cat_key, mapping in audit_report.items():
        meta = CATEGORY_DEFINITIONS.get(cat_key, {"label": cat_key, "icon": "🏷️"})
        for real_val, fake_val in mapping.items():
            audit_rows.append({
                "Category": f"{meta['icon']} {meta['label']}",
                "Original Sensitive Value": real_val,
                "Redaction / Replacement Applied": fake_val
            })
    
    if audit_rows:
        df_audit = pd.DataFrame(audit_rows)
        
        search_query = st.text_input("🔍 Search translation table (type any name, company, email...)", "")
        if search_query:
            mask = df_audit.apply(lambda row: row.astype(str).str.contains(search_query, case=False).any(), axis=1)
            filtered_df = df_audit[mask]
        else:
            filtered_df = df_audit
            
        st.dataframe(filtered_df, use_container_width=True, height=365)
    else:
        st.info("No PII mappings were generated.")

# ------------------------------------------------------------------------------
# SECTION 4: ⬇️ Secure Export
# ------------------------------------------------------------------------------
if st.session_state["redaction_results"] is not None:
    st.divider()
    st.markdown("## ⬇️ Secure Export")
    
    results = st.session_state["redaction_results"]
    clean_out_filename = results["clean_output_filename"]
    processed_bytes = results["processed_bytes"]
    audit_report = results["audit_report"]
    
    # Highlight-Bordered Card
    st.markdown(f"""
    <div class="export-card">
        <span style="font-size: 3rem;">🔒</span>
        <h3 style="margin-top: 1rem; color: #0f172a; font-weight: 800;">Export Anonymized Output</h3>
        <p style="color: #64748b; font-size: 0.9rem; margin-bottom: 2rem; line-height: 1.5;">
            Your redacted file is ready. The document has been processed in-memory with zero local disk logging. Choose a secure download method below:
        </p>
    </div>
    """, unsafe_allow_html=True)
    
    # Center-positioned columns for download buttons
    col_down1, col_down2 = st.columns(2)
    
    with col_down1:
        st.markdown("<div style='text-align: center; font-weight: 700; margin-bottom: 0.5rem;'>📄 Anonymized Document</div>", unsafe_allow_html=True)
        
        # 1. Direct Base64 Instant Download Link (100% immune to browser GUID issues)
        b64_docx = base64.b64encode(processed_bytes).decode()
        direct_download_html = f'''
        <a href="data:application/vnd.openxmlformats-officedocument.wordprocessingml.document;base64,{b64_docx}" 
           download="{clean_out_filename}" 
           class="direct-download-btn">
           📥 Click to Download {clean_out_filename}
        </a>
        '''
        st.markdown(direct_download_html, unsafe_allow_html=True)
        
        # 2. Standard Streamlit Download Button (with sanitized filename)
        st.download_button(
            label=f"💾 Alternative Download ({clean_out_filename})",
            data=processed_bytes,
            file_name=clean_out_filename,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True
        )
        
    with col_down2:
        st.markdown("<div style='text-align: center; font-weight: 700; margin-bottom: 0.5rem;'>📥 Mapping Logs</div>", unsafe_allow_html=True)
        
        # Form audit table records
        audit_rows_logs = []
        for cat_key, mapping in audit_report.items():
            meta = CATEGORY_DEFINITIONS.get(cat_key, {"label": cat_key, "icon": "🏷️"})
            for real_val, fake_val in mapping.items():
                audit_rows_logs.append({
                    "Category": f"{meta['icon']} {meta['label']}",
                    "Original Sensitive Value": real_val,
                    "Redaction / Replacement Applied": fake_val
                })
        
        df_audit_logs = pd.DataFrame(audit_rows_logs)
        
        # CSV Download Button
        csv_data = df_audit_logs.to_csv(index=False).encode('utf-8')
        st.download_button(
            label="📥 Download Audit Log (CSV)",
            data=csv_data,
            file_name=f"{os.path.splitext(clean_out_filename)[0]}_audit_log.csv",
            mime="text/csv",
            use_container_width=True
        )
        
        # JSON Download Button
        json_data = json.dumps(audit_report, indent=2).encode('utf-8')
        st.download_button(
            label="📥 Download Entity Map (JSON)",
            data=json_data,
            file_name=f"{os.path.splitext(clean_out_filename)[0]}_entity_map.json",
            mime="application/json",
            use_container_width=True
        )

# Footer
st.markdown("<div style='height: 2.5rem;'></div>", unsafe_allow_html=True)
st.markdown("""
<div style="text-align: center; color: #94a3b8; font-size: 0.825rem; border-top: 1px solid #e2e8f0; padding-top: 1.5rem;">
    Enterprise PII Redaction Engine • Built for secure document anonymization • Zero-storage in-memory architecture
</div>
""", unsafe_allow_html=True)

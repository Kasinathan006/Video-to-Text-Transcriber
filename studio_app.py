#!/usr/bin/env python3
"""
studio_app.py - AI Transcription Studio (Interactive Web App)

A Streamlit web application that turns video and audio recordings into
professionally formatted Microsoft Word (.docx) documents using
Sarvam AI (Cloud SOTA) or faster-whisper (Local GPU/CPU, fully offline).

To launch:
    pip install streamlit
    streamlit run studio_app.py
"""

import glob
import json
import os
import time
import uuid
from pathlib import Path

import streamlit as st

import sarvam_transcribe_to_docx as sarvam_stt

APP_DIR = Path(__file__).resolve().parent
WORKSPACE_ROOT = APP_DIR / "studio_workspace"

LANGUAGE_OPTIONS = [
    "en-IN (English/Tanglish/Multilingual)",
    "hi-IN (Hindi)",
    "ta-IN (Tamil)",
    "te-IN (Telugu)",
    "ml-IN (Malayalam)",
    "kn-IN (Kannada)",
]

FONT_MAP = {
    "Calibri": "Calibri",
    "Arial": "Arial",
    "Nirmala UI (Best for Tamil/Hindi)": "Nirmala UI",
}

# Page configuration
st.set_page_config(
    page_title="AI Video-to-Text Studio",
    page_icon="🎙️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Custom Styling (Rich Modern Aesthetics & Glassmorphism feel)
st.markdown("""
<style>
    .main-header {
        font-family: 'Inter', sans-serif;
        font-weight: 800;
        font-size: 2.8rem;
        background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
        -webkit-background-clip: text;
        -webkit-text-fill-color: transparent;
        margin-bottom: 0.2rem;
    }
    .sub-header {
        font-size: 1.1rem;
        color: #4A5568;
        margin-bottom: 2rem;
    }
    .stProgress > div > div > div > div {
        background-image: linear-gradient(to right, #667eea, #764ba2);
    }
</style>
""", unsafe_allow_html=True)


def transcribe_with_sarvam(audio_files, api_key, language_code, json_work_dir, progress_bar):
    """Send audio chunks to Sarvam AI batch STT and download JSON transcripts."""
    import sarvamai

    client = sarvamai.SarvamAI(api_subscription_key=api_key)
    stt_job = client.speech_to_text_job.create_job(
        model="saaras:v3",
        language_code=language_code,
        with_diarization=False,
    )
    job_id = stt_job.job_id
    stt_job.upload_files(audio_files)
    stt_job.start()
    progress_bar.progress(50)

    st.write("⏳ **Step 3:** Processing AI speech recognition on Sarvam cloud...")
    while True:
        job_status = client.speech_to_text_job.get_status(job_id)
        if job_status.job_state in ("Completed", "Failed"):
            break
        time.sleep(5)

    if job_status.job_state == "Failed":
        raise RuntimeError("Sarvam AI job failed on cloud servers. Check your API key quota and audio files.")

    os.makedirs(json_work_dir, exist_ok=True)
    stt_job.download_outputs(json_work_dir)
    json_files = sorted(glob.glob(os.path.join(json_work_dir, "*.json")))
    if not json_files:
        raise RuntimeError("Sarvam AI returned no transcript files.")
    return json_files


def transcribe_with_whisper(audio_files, model_size, json_work_dir, progress_bar):
    """Transcribe locally with faster-whisper (GPU if available, else CPU)."""
    from faster_whisper import WhisperModel

    st.write(f"💻 Loading faster-whisper `{model_size}` model (GPU if available)...")
    try:
        model = WhisperModel(model_size, device="cuda", compute_type="int8_float16")
        st.caption("Running on CUDA GPU (int8_float16).")
    except Exception:
        model = WhisperModel(model_size, device="cpu", compute_type="int8")
        st.caption("CUDA unavailable — running on CPU (int8).")

    os.makedirs(json_work_dir, exist_ok=True)
    json_files = []
    for idx, a_file in enumerate(audio_files):
        st.write(f"Transcribing chunk {idx + 1}/{len(audio_files)}...")
        segments, _info = model.transcribe(a_file, beam_size=5)
        text = " ".join(seg.text.strip() for seg in segments)
        j_path = os.path.join(json_work_dir, f"chunk_{idx:02d}.json")
        with open(j_path, "w", encoding="utf-8") as f:
            json.dump({"transcript": text}, f, ensure_ascii=False)
        json_files.append(j_path)
        progress_bar.progress(min(50 + int(35 * (idx + 1) / len(audio_files)), 85))
    return json_files


def run_pipeline(uploaded_file, settings):
    """Full pipeline: save upload -> extract audio -> transcribe -> compile docx.

    Returns dict with docx bytes + stats, stored in session_state so the
    download button survives Streamlit reruns.
    """
    job_dir = WORKSPACE_ROOT / f"job_{uuid.uuid4().hex[:8]}"
    audio_work_dir = job_dir / "extracted_audio"
    json_work_dir = job_dir / "json_transcripts"
    job_dir.mkdir(parents=True, exist_ok=True)

    input_path = job_dir / uploaded_file.name
    with open(input_path, "wb") as f:
        f.write(uploaded_file.getbuffer())

    output_docx_name = f"{Path(uploaded_file.name).stem}_AI_Transcript.docx"
    output_docx_path = job_dir / output_docx_name

    with st.status("⚡ Running Studio Pipeline...", expanded=True) as status:
        # Step A: Video to Audio Extraction
        st.write("🎵 **Step 1:** Extracting & standardizing audio (16kHz Mono PCM WAV)...")
        progress_bar = st.progress(15)
        audio_files, duration = sarvam_stt.extract_and_segment_audio(
            str(input_path),
            str(audio_work_dir),
            chunk_minutes=settings["chunk_minutes"],
        )
        st.success(f"✅ Audio extracted into {len(audio_files)} chunk(s). "
                   f"Total Duration: {duration / 60:.1f} minutes.")
        progress_bar.progress(35)

        # Step B: Transcription
        if settings["engine"] == "sarvam":
            st.write("☁️ **Step 2:** Uploading to Sarvam AI (`saaras:v3`)...")
            json_files = transcribe_with_sarvam(
                audio_files, settings["api_key"], settings["language_code"],
                str(json_work_dir), progress_bar,
            )
            model_label = "Sarvam AI saaras:v3 Batch Speech-to-Text"
        else:
            st.write("🔒 **Step 2:** Running local offline transcription (faster-whisper)...")
            json_files = transcribe_with_whisper(
                audio_files, settings["whisper_model"], str(json_work_dir), progress_bar,
            )
            model_label = f"OpenAI Whisper {settings['whisper_model']} (Local faster-whisper)"
        progress_bar.progress(85)

        # Step C: Word Document Compilation
        st.write("📄 **Step 4:** Compiling formatted Microsoft Word document...")
        sarvam_stt.create_transcription_docx(
            json_files,
            str(output_docx_path),
            uploaded_file.name,
            duration,
            font_name=settings["font_name"],
            sentences_per_paragraph=settings["para_sentences"],
            model_label=model_label,
            chunk_minutes=settings["chunk_minutes"],
        )
        progress_bar.progress(100)
        status.update(label="🎉 Transcription Complete!", state="complete", expanded=False)

    docx_bytes = output_docx_path.read_bytes()
    word_count = 0
    for jf in json_files:
        with open(jf, "r", encoding="utf-8") as f:
            word_count += len(json.load(f).get("transcript", "").split())

    return {
        "docx_bytes": docx_bytes,
        "docx_name": output_docx_name,
        "duration_min": duration / 60,
        "chunks": len(audio_files),
        "words": word_count,
        "model_label": model_label,
    }


def main():
    st.markdown('<div class="main-header">🎙️ AI Video-to-Text Studio</div>', unsafe_allow_html=True)
    st.markdown(
        '<div class="sub-header">Transform massive video & audio recordings into professionally '
        'formatted Word reports with Sarvam AI & Whisper.</div>',
        unsafe_allow_html=True,
    )

    # --- SIDEBAR CONFIGURATION ---
    with st.sidebar:
        st.header("⚙️ Studio Settings")

        model_choice = st.radio(
            "Select Transcription Engine:",
            options=[
                "Sarvam AI (Cloud SOTA - Best for Indian Languages)",
                "OpenAI Whisper (Local Offline)",
            ],
            index=0,
            help="Sarvam AI is lightning fast and tuned for Tamil, Hindi, Tanglish, Malayalam, etc. "
                 "Whisper runs 100% locally on your PC — audio never leaves your machine.",
        )
        engine = "sarvam" if "Sarvam" in model_choice else "whisper"

        st.divider()

        api_key = os.getenv("SARVAM_API_KEY", "")
        language_code = "en-IN"
        whisper_model = "large-v3"
        chunk_minutes = 45

        if engine == "sarvam":
            api_key = st.text_input(
                "Sarvam AI API Key:",
                value=api_key,
                type="password",
                help="Enter your Sarvam AI API subscription key (or set the SARVAM_API_KEY environment variable).",
            )
            language_choice = st.selectbox("Target Language:", options=LANGUAGE_OPTIONS, index=0)
            language_code = language_choice.split(" ")[0]
            chunk_minutes = st.slider(
                "Audio Segmentation Chunk Size (mins):",
                min_value=15, max_value=45, value=45,
                help="Recordings longer than this are automatically split to prevent API limits.",
            )
        else:
            whisper_model = st.selectbox(
                "Whisper Model Size:",
                ["large-v3", "medium", "small", "base", "tiny"],
                index=0,
            )
            st.info("💡 Local model uses your GPU if CUDA is available, otherwise CPU.")

        st.divider()
        st.markdown("### 📊 Document Styling")
        font_choice = st.selectbox("Word Font Family:", list(FONT_MAP.keys()), index=0)
        para_sentences = st.slider("Sentences per Paragraph:", min_value=3, max_value=10, value=6)

    settings = {
        "engine": engine,
        "api_key": api_key,
        "language_code": language_code,
        "whisper_model": whisper_model,
        "chunk_minutes": chunk_minutes,
        "font_name": FONT_MAP[font_choice],
        "para_sentences": para_sentences,
    }

    # --- MAIN CONTENT AREA ---
    col1, col2 = st.columns([3, 2])

    with col1:
        st.markdown("### 1️⃣ Upload Video or Audio Recording")
        uploaded_file = st.file_uploader(
            "Drop your media file here (supports MP4, MKV, MOV, WAV, MP3, M4A)...",
            type=["mp4", "mkv", "mov", "avi", "flv", "wmv", "webm", "wav", "mp3", "m4a", "flac", "ogg"],
        )

        if uploaded_file is not None:
            st.json({
                "Filename": uploaded_file.name,
                "File Size": f"{uploaded_file.size / (1024 * 1024):.2f} MB",
                "File Type": uploaded_file.type or Path(uploaded_file.name).suffix,
            })

            st.markdown("### 2️⃣ Start Transcription Pipeline")
            if st.button("🚀 Generate Formatted Word Document", type="primary", use_container_width=True):
                if engine == "sarvam" and not settings["api_key"].strip():
                    st.error("❌ Please enter your Sarvam AI API key in the sidebar "
                             "(or switch to the local Whisper engine).")
                else:
                    try:
                        st.session_state["result"] = run_pipeline(uploaded_file, settings)
                    except Exception as e:
                        st.session_state.pop("result", None)
                        st.error(f"❌ Pipeline failed: {e}")

        # --- SUCCESS & DOWNLOAD AREA (survives reruns via session_state) ---
        result = st.session_state.get("result")
        if result:
            st.success("✨ Your Microsoft Word transcript is ready!")
            m1, m2, m3 = st.columns(3)
            m1.metric("Duration", f"{result['duration_min']:.1f} min")
            m2.metric("Audio Chunks", result["chunks"])
            m3.metric("Words Transcribed", f"{result['words']:,}")
            st.caption(f"Engine: {result['model_label']}")

            st.download_button(
                label="📥 Download Formatted Word Document (.docx)",
                data=result["docx_bytes"],
                file_name=result["docx_name"],
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                type="primary",
                use_container_width=True,
            )

    with col2:
        st.markdown("### 💡 Why AI Studio?")
        st.info("""
        **⚡ Memory-Efficient Extraction**
        Our engine pipes video directly from disk to `ffmpeg`, allowing you to transcribe massive 50GB+ recordings without running out of RAM.

        **🇮🇳 Indian Language Mastery**
        Powered by Sarvam AI (`saaras:v3`), specifically tuned for Tamil, Hindi, Tanglish, Malayalam, Telugu, Kannada, and Indian English accents.

        **🔒 Privacy Mode**
        Switch to local Whisper and your audio never leaves this machine — ideal for legal, medical, and confidential recordings.

        **📑 Verbatim Formatting**
        Transcripts are cleanly broken into paragraphs with an executive summary header and Unicode font protection for Indian scripts.
        """)

        st.markdown("### 🌐 REST API & Dashboard")
        st.code("python api_server.py", language="bash")
        st.caption("Launch the FastAPI backend + web dashboard at http://localhost:8000 "
                   "for background job processing (Phase 2).")

        st.markdown("### 🛠️ Quick CLI Tip")
        st.code('python sarvam_transcribe_to_docx.py "your_video.mp4"', language="bash")
        st.caption("You can also run the 1-click CLI pipeline directly in your terminal!")


if __name__ == "__main__":
    main()

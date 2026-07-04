#!/usr/bin/env python3
"""
api_server.py - VoxDoc AI — Production SaaS Server

Complete, sellable transcription product:

  * Accounts & sessions ....... signup / login, PBKDF2-hashed passwords, bearer tokens
  * Subscription tiers ........ Starter (free) / Creator Pro / Agency with monthly
                                minute quotas, upload-size caps & concurrency limits
  * License keys .............. mint with `python manage.py gen-keys pro`, sell them,
                                buyers redeem in-app to upgrade — no Stripe required
  * Persistent job queue ...... SQLite-backed jobs survive restarts; background worker
                                pool never blocks or times out HTTP requests
  * Transcription engines ..... Sarvam AI Batch API (cloud) / faster-whisper (local GPU)
  * Web product ............... marketing landing page at /  ·  app dashboard at /app
                                OpenAPI docs at /docs

Quick start:
    python manage.py create-admin you@example.com YourPassword123
    python api_server.py                      # http://localhost:8000

Configuration (.env file or environment variables):
    SARVAM_API_KEY   server-wide Sarvam AI key used for cloud transcription
    VOXDOC_PORT      listen port (default 8000)
"""

import glob
import json
import logging
import os
import shutil
import threading
import time
import uuid
from concurrent.futures import ThreadPoolExecutor
from logging.handlers import RotatingFileHandler
from pathlib import Path
from typing import Optional

import uvicorn
from fastapi import Depends, FastAPI, File, Form, Header, HTTPException, Request, UploadFile
from fastapi.responses import FileResponse, JSONResponse, Response
from pydantic import BaseModel

import database as db
import sarvam_transcribe_to_docx as pipeline

APP_DIR = Path(__file__).resolve().parent
STORAGE_ROOT = APP_DIR / "storage" / "jobs"
LOGS_DIR = APP_DIR / "logs"
DASHBOARD_HTML = APP_DIR / "dashboard.html"
LANDING_HTML = APP_DIR / "landing.html"

ALLOWED_EXTENSIONS = {
    ".mp4", ".mkv", ".mov", ".avi", ".flv", ".wmv", ".webm", ".m4v",
    ".ts", ".mts", ".m2ts", ".wav", ".mp3", ".m4a", ".flac", ".ogg", ".aac",
}

VERSION = "2.0.0"

# ----------------------------------------------------------------------------
# Logging
# ----------------------------------------------------------------------------
LOGS_DIR.mkdir(parents=True, exist_ok=True)
logger = logging.getLogger("voxdoc")
logger.setLevel(logging.INFO)
_fh = RotatingFileHandler(LOGS_DIR / "server.log", maxBytes=5_000_000,
                          backupCount=3, encoding="utf-8")
_fh.setFormatter(logging.Formatter("%(asctime)s %(levelname)s %(message)s"))
logger.addHandler(_fh)
logger.addHandler(logging.StreamHandler())

app = FastAPI(
    title="VoxDoc AI — Media-to-Document API",
    description="Transcribe massive video/audio recordings into formatted Word documents "
                "using Sarvam AI (cloud) or faster-whisper (local, offline).",
    version=VERSION,
)

EXECUTOR = ThreadPoolExecutor(max_workers=2, thread_name_prefix="voxdoc-worker")

# In-memory only — user-supplied API keys are never written to the database
_JOB_SECRETS: dict = {}

_WHISPER_CACHE: dict = {}
_WHISPER_LOCK = threading.Lock()


# ----------------------------------------------------------------------------
# Auth dependencies
# ----------------------------------------------------------------------------

def current_user(authorization: Optional[str] = Header(None)) -> dict:
    if not authorization or not authorization.lower().startswith("bearer "):
        raise HTTPException(401, "Not authenticated. Provide 'Authorization: Bearer <token>'.",
                            headers={"WWW-Authenticate": "Bearer"})
    user = db.get_user_by_token(authorization[7:].strip())
    if not user:
        raise HTTPException(401, "Session expired or invalid — please log in again.",
                            headers={"WWW-Authenticate": "Bearer"})
    return user


def admin_user(user: dict = Depends(current_user)) -> dict:
    if not user.get("is_admin"):
        raise HTTPException(403, "Administrator access required.")
    return user


def _public_user(user: dict) -> dict:
    return {
        "id": user["id"],
        "email": user["email"],
        "full_name": user.get("full_name", ""),
        "is_admin": bool(user.get("is_admin")),
        "quota": db.quota_summary(user),
    }


# ----------------------------------------------------------------------------
# Global error handling
# ----------------------------------------------------------------------------

@app.exception_handler(Exception)
async def unhandled_exception_handler(request: Request, exc: Exception):
    logger.exception("Unhandled error on %s %s: %s", request.method, request.url.path, exc)
    return JSONResponse(status_code=500, content={"detail": "Internal server error."})


# ----------------------------------------------------------------------------
# Transcription workers
# ----------------------------------------------------------------------------

def _get_whisper_model(model_size: str):
    from faster_whisper import WhisperModel

    with _WHISPER_LOCK:
        if model_size not in _WHISPER_CACHE:
            try:
                _WHISPER_CACHE[model_size] = WhisperModel(
                    model_size, device="cuda", compute_type="int8_float16")
                logger.info("Loaded faster-whisper %s on CUDA", model_size)
            except Exception:
                _WHISPER_CACHE[model_size] = WhisperModel(
                    model_size, device="cpu", compute_type="int8")
                logger.info("Loaded faster-whisper %s on CPU", model_size)
        return _WHISPER_CACHE[model_size]


def _transcribe_sarvam(job_id: str, audio_files, params: dict, json_dir: Path):
    import sarvamai

    api_key = _JOB_SECRETS.pop(job_id, None) or os.getenv("SARVAM_API_KEY", "")
    client = sarvamai.SarvamAI(api_subscription_key=api_key)
    stt_job = client.speech_to_text_job.create_job(
        model=params.get("sarvam_model", "saaras:v3"),
        language_code=params.get("language_code", "en-IN"),
        with_diarization=False,
    )
    stt_job.upload_files(audio_files)
    stt_job.start()
    db.update_job(job_id, progress=55, detail="Sarvam AI cloud processing...")

    while True:
        status = client.speech_to_text_job.get_status(stt_job.job_id)
        if status.job_state in ("Completed", "Failed"):
            break
        time.sleep(5)

    if status.job_state == "Failed":
        raise RuntimeError("Sarvam AI job failed on cloud servers. Check API key quota and audio.")

    json_dir.mkdir(parents=True, exist_ok=True)
    stt_job.download_outputs(str(json_dir))
    json_files = sorted(glob.glob(str(json_dir / "*.json")))
    if not json_files:
        raise RuntimeError("Sarvam AI returned no transcript files.")
    return json_files


def _transcribe_whisper(job_id: str, audio_files, params: dict, json_dir: Path):
    model = _get_whisper_model(params.get("whisper_model", "large-v3"))
    json_dir.mkdir(parents=True, exist_ok=True)
    json_files = []
    with _WHISPER_LOCK:  # one GPU transcription at a time
        for idx, a_file in enumerate(audio_files):
            db.update_job(job_id, progress=50 + int(35 * idx / len(audio_files)),
                          detail=f"Transcribing chunk {idx + 1}/{len(audio_files)} locally...")
            segments, _info = model.transcribe(a_file, beam_size=5)
            text = " ".join(seg.text.strip() for seg in segments)
            j_path = json_dir / f"chunk_{idx:02d}.json"
            with open(j_path, "w", encoding="utf-8") as f:
                json.dump({"transcript": text}, f, ensure_ascii=False)
            json_files.append(str(j_path))
    return json_files


def _run_job(job_id: str):
    """Background worker: extract -> quota check -> transcribe -> compile -> bill."""
    job = db.get_job(job_id)
    if not job:
        return
    params = json.loads(job["params"])
    user = db.get_user(job["user_id"])

    job_dir = STORAGE_ROOT / job_id
    input_path = job_dir / job["filename"]
    audio_dir = job_dir / "audio"
    json_dir = job_dir / "json"
    docx_name = f"{Path(job['filename']).stem}_VoxDoc_Transcript.docx"
    docx_path = job_dir / docx_name

    try:
        db.update_job(job_id, status="processing", stage="extracting", progress=10,
                      detail="Extracting & standardizing audio (16kHz mono WAV)...",
                      started_at=time.time())

        audio_files, duration = pipeline.extract_and_segment_audio(
            str(input_path), str(audio_dir), chunk_minutes=params.get("chunk_minutes", 45))
        duration_min = duration / 60.0

        # Hard quota enforcement now that the true duration is known
        quota = db.quota_summary(user)
        if duration_min > quota["minutes_remaining"] + 0.05:
            raise RuntimeError(
                f"Monthly quota exceeded: this recording is {duration_min:.1f} min but your "
                f"'{quota['tier_name']}' plan has {quota['minutes_remaining']:.1f} min left. "
                f"Upgrade your plan or redeem a license key on the Account page.")

        db.update_job(job_id, stage="transcribing", progress=40,
                      duration_sec=round(duration, 1),
                      detail=f"Audio ready ({len(audio_files)} chunk(s)). Starting transcription...")

        if job["engine"] == "sarvam":
            json_files = _transcribe_sarvam(job_id, audio_files, params, json_dir)
            model_label = f"Sarvam AI {params.get('sarvam_model', 'saaras:v3')} Batch Speech-to-Text"
        else:
            json_files = _transcribe_whisper(job_id, audio_files, params, json_dir)
            model_label = f"OpenAI Whisper {params.get('whisper_model', 'large-v3')} (Local faster-whisper)"

        db.update_job(job_id, stage="compiling", progress=90,
                      detail="Compiling formatted Word document...")

        pipeline.create_transcription_docx(
            json_files, str(docx_path), job["filename"], duration,
            font_name=params.get("font_name", "Calibri"),
            sentences_per_paragraph=params.get("sentences_per_paragraph", 6),
            model_label=model_label,
            chunk_minutes=params.get("chunk_minutes", 45),
        )

        words = 0
        for jf in json_files:
            with open(jf, "r", encoding="utf-8") as f:
                words += len(json.load(f).get("transcript", "").split())

        db.record_usage(job["user_id"], job_id, duration_min)
        db.update_job(job_id, status="completed", stage="done", progress=100,
                      detail="Transcript ready for download.",
                      docx_name=docx_name, words=words, model_label=model_label,
                      finished_at=time.time())
        logger.info("Job %s completed for %s (%.1f min, %d words)",
                    job_id, user["email"], duration_min, words)
    except Exception as e:
        _JOB_SECRETS.pop(job_id, None)
        db.update_job(job_id, status="failed", stage="error", progress=100,
                      detail=str(e), finished_at=time.time())
        logger.warning("Job %s failed for %s: %s", job_id, user["email"] if user else "?", e)


def _job_response(job: dict) -> dict:
    j = {k: v for k, v in job.items() if k not in ("params", "user_id")}
    j["download_url"] = f"/api/v1/jobs/{job['id']}/download" if job["status"] == "completed" else None
    return j


# ----------------------------------------------------------------------------
# Public routes
# ----------------------------------------------------------------------------

@app.get("/api/v1/health")
def health():
    ffmpeg_ok = shutil.which("ffmpeg") is not None
    return {"status": "ok" if ffmpeg_ok else "degraded", "ffmpeg": ffmpeg_ok,
            "version": VERSION}


@app.get("/api/v1/pricing")
def pricing():
    return {"tiers": db.TIERS}


FAVICON_SVG = ('<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100">'
               '<text y=".9em" font-size="90">\U0001F399</text></svg>')


@app.get("/favicon.ico", include_in_schema=False)
def favicon():
    return Response(content=FAVICON_SVG, media_type="image/svg+xml",
                    headers={"Cache-Control": "public, max-age=86400"})


@app.get("/")
def landing():
    if LANDING_HTML.exists():
        return FileResponse(LANDING_HTML, media_type="text/html")
    return JSONResponse({"message": "VoxDoc AI API. App: /app · Docs: /docs"})


@app.get("/app")
def dashboard():
    if DASHBOARD_HTML.exists():
        return FileResponse(DASHBOARD_HTML, media_type="text/html")
    return JSONResponse({"message": "Dashboard file missing. See /docs for the API."})


# ----------------------------------------------------------------------------
# Auth & account
# ----------------------------------------------------------------------------

class SignupBody(BaseModel):
    email: str
    password: str
    full_name: str = ""


class LoginBody(BaseModel):
    email: str
    password: str


class RedeemBody(BaseModel):
    key: str


@app.post("/api/v1/auth/signup", status_code=201)
def signup(body: SignupBody):
    try:
        user = db.create_user(body.email, body.password, body.full_name)
    except ValueError as e:
        raise HTTPException(400, str(e))
    token = db.create_session(user["id"])
    logger.info("New signup: %s", user["email"])
    return {"token": token, "user": _public_user(user)}


@app.post("/api/v1/auth/login")
def login(body: LoginBody):
    user = db.verify_login(body.email, body.password)
    if not user:
        raise HTTPException(401, "Incorrect email or password.")
    token = db.create_session(user["id"])
    return {"token": token, "user": _public_user(user)}


@app.post("/api/v1/auth/logout")
def logout(authorization: Optional[str] = Header(None),
           user: dict = Depends(current_user)):
    db.delete_session(authorization[7:].strip())
    return {"ok": True}


@app.get("/api/v1/account")
def account(user: dict = Depends(current_user)):
    return _public_user(user)


@app.post("/api/v1/license/redeem")
def redeem(body: RedeemBody, user: dict = Depends(current_user)):
    try:
        result = db.redeem_license_key(user["id"], body.key)
    except ValueError as e:
        raise HTTPException(400, str(e))
    logger.info("License redeemed by %s -> %s", user["email"], result["tier"])
    return {"ok": True, **result, "user": _public_user(db.get_user(user["id"]))}


# ----------------------------------------------------------------------------
# Transcription jobs
# ----------------------------------------------------------------------------

@app.post("/api/v1/transcribe")
async def create_transcription_job(
    user: dict = Depends(current_user),
    file: UploadFile = File(...),
    engine: str = Form("sarvam"),
    language_code: str = Form("en-IN"),
    sarvam_model: str = Form("saaras:v3"),
    whisper_model: str = Form("large-v3"),
    chunk_minutes: int = Form(45),
    font_name: str = Form("Calibri"),
    sentences_per_paragraph: int = Form(6),
    api_key: Optional[str] = Form(None),
):
    ext = Path(file.filename or "").suffix.lower()
    if ext not in ALLOWED_EXTENSIONS:
        raise HTTPException(400, f"Unsupported file type '{ext}'. "
                                 f"Allowed: {', '.join(sorted(ALLOWED_EXTENSIONS))}")
    if engine not in ("sarvam", "whisper"):
        raise HTTPException(400, "engine must be 'sarvam' or 'whisper'")
    if sarvam_model not in ("saaras:v3", "saarika:v2"):
        raise HTTPException(400, "sarvam_model must be 'saaras:v3' or 'saarika:v2'")

    resolved_key = (api_key or os.getenv("SARVAM_API_KEY", "")).strip()
    if engine == "sarvam" and not resolved_key:
        raise HTTPException(400, "The Sarvam cloud engine needs an API key: pass the "
                                 "api_key field or set SARVAM_API_KEY on the server.")

    quota = db.quota_summary(user)
    if quota["minutes_remaining"] <= 0:
        raise HTTPException(402, f"Monthly quota exhausted ({quota['monthly_minutes']} min on "
                                 f"the {quota['tier_name']} plan). Upgrade or redeem a license key.")
    if db.count_active_jobs(user["id"]) >= quota["concurrent_jobs"]:
        raise HTTPException(429, f"Your plan allows {quota['concurrent_jobs']} concurrent job(s). "
                                 "Please wait for the current job to finish.")

    max_bytes = quota["max_upload_mb"] * 1024 * 1024
    job_id = uuid.uuid4().hex[:12]
    job_dir = STORAGE_ROOT / job_id
    job_dir.mkdir(parents=True, exist_ok=True)
    safe_name = Path(file.filename).name
    dest = job_dir / safe_name

    size = 0
    with open(dest, "wb") as out:
        while chunk := await file.read(4 * 1024 * 1024):
            size += len(chunk)
            if size > max_bytes:
                out.close()
                shutil.rmtree(job_dir, ignore_errors=True)
                raise HTTPException(413, f"File exceeds your plan's {quota['max_upload_mb']} MB "
                                         "upload limit. Upgrade for larger files.")
            out.write(chunk)

    params = {
        "language_code": language_code,
        "sarvam_model": sarvam_model,
        "whisper_model": whisper_model,
        "chunk_minutes": max(5, min(45, chunk_minutes)),
        "font_name": font_name,
        "sentences_per_paragraph": max(1, min(12, sentences_per_paragraph)),
    }
    job = db.create_job(user["id"], safe_name, size / (1024 * 1024), engine, params,
                        job_id=job_id)
    if api_key:
        _JOB_SECRETS[job["id"]] = resolved_key

    EXECUTOR.submit(_run_job, job["id"])
    logger.info("Job %s queued by %s (%s, %.1f MB)", job["id"], user["email"],
                engine, size / (1024 * 1024))
    return JSONResponse(status_code=202, content={
        "job_id": job["id"],
        "status": "queued",
        "status_url": f"/api/v1/jobs/{job['id']}",
        "download_url": f"/api/v1/jobs/{job['id']}/download",
    })


def _get_owned_job(job_id: str, user: dict) -> dict:
    job = db.get_job(job_id)
    if not job or (job["user_id"] != user["id"] and not user.get("is_admin")):
        raise HTTPException(404, f"Job '{job_id}' not found.")
    return job


@app.get("/api/v1/jobs")
def list_jobs(user: dict = Depends(current_user)):
    return {"jobs": [_job_response(j) for j in db.list_jobs_for_user(user["id"])]}


@app.get("/api/v1/jobs/{job_id}")
def get_job(job_id: str, user: dict = Depends(current_user)):
    return _job_response(_get_owned_job(job_id, user))


@app.get("/api/v1/jobs/{job_id}/download")
def download_job(job_id: str, user: dict = Depends(current_user)):
    job = _get_owned_job(job_id, user)
    if job["status"] != "completed":
        raise HTTPException(409, f"Job is '{job['status']}' — no document available yet.")
    docx_path = STORAGE_ROOT / job_id / job["docx_name"]
    if not docx_path.exists():
        raise HTTPException(410, "The document file has been removed from storage.")
    return FileResponse(
        docx_path, filename=job["docx_name"],
        media_type="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    )


@app.delete("/api/v1/jobs/{job_id}")
def delete_job(job_id: str, user: dict = Depends(current_user)):
    job = _get_owned_job(job_id, user)
    if job["status"] == "processing":
        raise HTTPException(409, "Cannot delete a job while it is processing.")
    db.delete_job(job_id)
    shutil.rmtree(STORAGE_ROOT / job_id, ignore_errors=True)
    return {"deleted": job_id}


# ----------------------------------------------------------------------------
# Admin
# ----------------------------------------------------------------------------

class TierBody(BaseModel):
    tier: str
    days: Optional[int] = 30      # None or 0 = never expires


class KeysBody(BaseModel):
    tier: str
    days: int = 30
    count: int = 1


@app.get("/api/v1/admin/users")
def admin_list_users(admin: dict = Depends(admin_user)):
    return {"users": [_public_user(u) for u in db.list_users()]}


@app.post("/api/v1/admin/users/{user_id}/tier")
def admin_set_tier(user_id: str, body: TierBody, admin: dict = Depends(admin_user)):
    try:
        days = None if not body.days else body.days
        db.set_user_tier(user_id, body.tier, duration_days=days)
    except ValueError as e:
        raise HTTPException(400, str(e))
    logger.info("Admin %s set %s -> %s", admin["email"], user_id, body.tier)
    return {"ok": True, "user": _public_user(db.get_user(user_id))}


@app.post("/api/v1/admin/license-keys")
def admin_gen_keys(body: KeysBody, admin: dict = Depends(admin_user)):
    try:
        keys = db.generate_license_keys(body.tier, duration_days=body.days, count=body.count)
    except ValueError as e:
        raise HTTPException(400, str(e))
    logger.info("Admin %s minted %d %s key(s)", admin["email"], len(keys), body.tier)
    return {"keys": keys}


@app.get("/api/v1/admin/license-keys")
def admin_list_keys(admin: dict = Depends(admin_user)):
    return {"keys": db.list_license_keys()}


# ----------------------------------------------------------------------------
# Startup
# ----------------------------------------------------------------------------

def startup():
    db.load_env()
    db.init_db()
    STORAGE_ROOT.mkdir(parents=True, exist_ok=True)
    interrupted = db.fail_interrupted_jobs()
    if interrupted:
        logger.info("Marked %d interrupted job(s) as failed after restart", interrupted)
    if not shutil.which("ffmpeg"):
        logger.warning("FFmpeg not found on PATH — extraction will fail until installed!")


startup()


if __name__ == "__main__":
    port = int(os.getenv("VOXDOC_PORT", "8000"))
    print("=" * 62)
    print(" VOXDOC AI — PRODUCTION SERVER v" + VERSION)
    print(f" Landing page : http://localhost:{port}")
    print(f" App          : http://localhost:{port}/app")
    print(f" API Docs     : http://localhost:{port}/docs")
    print(" First run?     python manage.py create-admin <email> <password>")
    print("=" * 62)
    uvicorn.run(app, host="0.0.0.0", port=port)

# 🎙️ VoxDoc AI — AI Media-to-Document Studio (Production)

A complete, sellable SaaS product that turns video & audio recordings (even 10GB+ files)
into beautifully formatted Microsoft Word documents using **Sarvam AI** (cloud SOTA for
Indian languages) or **Faster-Whisper** (local GPU/CPU — 100% offline & private).

> 📘 Product blueprint & architecture: [PRODUCT_BUILD_GUIDE.md](PRODUCT_BUILD_GUIDE.md)

---

## 🏪 What makes this sellable

| Capability | How it works |
| :--- | :--- |
| **User accounts** | Signup/login with PBKDF2-hashed passwords & 30-day bearer sessions |
| **Subscription tiers** | Starter (free, 60 min/mo) · Creator Pro ($19, 900 min/mo) · Agency ($49, 3000 min/mo) |
| **Quota enforcement** | Monthly minute ledger, per-tier upload caps (500MB / 10GB / 50GB) & concurrency limits |
| **License keys** | Mint keys offline, sell via Gumroad / Lemon Squeezy / invoice — buyers redeem in-app, plan activates instantly. No Stripe required. |
| **Persistent jobs** | SQLite-backed queue survives restarts; interrupted jobs fail honestly, completed docs stay downloadable |
| **Marketing site** | Landing page with pricing at `/`, full app at `/app`, OpenAPI docs at `/docs` |
| **Admin panel** | Manage users, change plans, mint license keys — in the web UI or via `manage.py` |

---

## 🚀 Launch in 3 commands

```bash
pip install -r requirements.txt
python manage.py create-admin you@yourdomain.com YourStrongPassword
python api_server.py                # or double-click start_voxdoc.bat
```

Then open **http://localhost:8000** — landing page, app, and API are all live.

**Configuration** — copy `.env.example` to `.env`:
```
SARVAM_API_KEY=sk_your_key_here     # server-wide key for the cloud engine
VOXDOC_PORT=8000
```
(FFmpeg must be on PATH: [ffmpeg.org](https://ffmpeg.org/download.html) · `brew install ffmpeg` · `apt install ffmpeg`)

---

## 💰 Selling workflow

1. **Mint keys**: `python manage.py gen-keys pro --count 10 --days 30`
   (or use the Admin Panel in the web app)
2. **Sell them** on Gumroad / Lemon Squeezy / by invoice.
3. **Buyer redeems** the key on their Account page → plan upgrades instantly for the key's duration.
4. Track everything: `python manage.py list-users` · `python manage.py list-keys --unredeemed`

Manual overrides: `python manage.py set-tier customer@email.com agency --days 365`

---

## 🌐 REST API (for Agency-tier customers & integrations)

All endpoints need `Authorization: Bearer <token>` (from `/api/v1/auth/login`).

| Endpoint | Description |
| :--- | :--- |
| `POST /api/v1/auth/signup` · `/login` · `/logout` | Account & session management |
| `GET /api/v1/account` | Profile, plan & live quota |
| `POST /api/v1/license/redeem` | Redeem a license key |
| `POST /api/v1/transcribe` | Upload media (multipart, streamed) → `202` + `job_id` |
| `GET /api/v1/jobs` / `/jobs/{id}` | List / poll jobs (owner-scoped) |
| `GET /api/v1/jobs/{id}/download` | Download the finished `.docx` |
| `DELETE /api/v1/jobs/{id}` | Delete a job & its files |
| `GET /api/v1/admin/*` | Users, tiers & license keys (admin only) |
| `GET /api/v1/health` · `/pricing` | Public status & plans |

Example:
```bash
curl -X POST http://localhost:8000/api/v1/transcribe \
  -H "Authorization: Bearer $TOKEN" \
  -F "file=@lecture.mp4" -F "engine=whisper" -F "whisper_model=large-v3"
```

---

## ✨ Pipeline features

- **Resilient extraction**: hardened FFmpeg (`-err_detect ignore_err -fflags +discardcorrupt`)
  streams from disk — corrupted MP4s survive, RAM stays flat on huge files, audio is
  standardized to 16kHz mono WAV and auto-chunked at ≤45 minutes.
- **Hybrid AI**: Sarvam AI batch API (`saaras:v3` English output / `saarika:v2` native script,
  8 Indian languages) or faster-whisper (`large-v3`→`tiny`) on CUDA GPU with CPU fallback.
- **Word-perfect documents**: executive metadata table, clean paragraphing, per-part word
  counts, and complex-script font fallback (`Nirmala UI`) so Tamil/Hindi never render as boxes.

## 🖥️ Also included

- **`app.py`** — standalone Streamlit studio (no accounts; includes Google Drive import &
  speaker diarization): `streamlit run app.py`
- **`studio_app.py`** — original Phase 1 prototype
- **CLI pipeline** — `python sarvam_transcribe_to_docx.py "video.mp4"`

## 🔒 Production deployment notes

- Put the server behind HTTPS (Caddy/nginx reverse proxy or Cloudflare Tunnel).
- Back up `storage/voxdoc.db` (accounts, keys, jobs) and `storage/jobs/` (documents).
- Logs rotate automatically in `logs/server.log`.
- Scale-out path (per the build guide): swap the thread pool for Celery+Redis and local
  storage for S3/R2 — the job interface is already shaped for it.

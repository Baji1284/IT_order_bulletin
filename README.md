# FM Orders Bulletin Pipeline

Cloud Run Job that daily:

1. Fetches the latest Gmail message with subject **FM Orders Bulletin**
2. Downloads the PDF attachment in memory
3. Extracts tabular data via the **Gemini API** (exact structure + numeric precision)
4. Clears and writes the result into a **Google Sheet**

Scheduled for **9:05 AM Asia/Kolkata** via Cloud Scheduler.

## Project files

| File | Role |
|------|------|
| `main.py` | Job entrypoint (Gmail → Gemini → Sheets) |
| `requirements.txt` | Pinned Python dependencies |
| `Dockerfile` | Container image for Cloud Run Job |
| `DEPLOY.md` | Step-by-step GCP CLI / Console setup |
| `.env.example` | Env var names only (no secrets) |

## Quick start

See **[DEPLOY.md](DEPLOY.md)** for enabling APIs, Secret Manager, Artifact Registry, Cloud Run Job, and Scheduler.

## Security

- Never commit API keys, passwords, or service account JSON.
- Inject secrets via Secret Manager → Cloud Run Job environment variables.
- Rotate any credentials that were pasted into chat or tickets before go-live.

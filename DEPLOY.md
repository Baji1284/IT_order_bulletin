# FM Orders Bulletin — GCP Deployment Guide

Daily Cloud Run Job: fetch the **FM Orders Bulletin** PDF from Gmail → extract tables with Gemini → write to Google Sheets. Triggered at **9:05 AM Asia/Kolkata** by Cloud Scheduler.

> **Security:** Do not put passwords or API keys in source code, Docker images, or git.  
> Credentials shared in chat should be **rotated** (Gemini API key and mailbox password) before production use.

---

## Prerequisites

- Google Cloud project with billing enabled
- Google Workspace admin access (for Gmail domain-wide delegation)
- `gcloud` CLI installed and authenticated
- Docker installed (or use Cloud Build)
- Target Google Sheet ID
- Gemini API key ([Google AI Studio](https://aistudio.google.com/apikey))
- Service account JSON key

**Note on “Gmail App Password”:** The Gmail **API** does not use app passwords. This job uses a **service account + domain-wide delegation** to impersonate `itdocumentation@forbesmarshall.com`. App passwords apply only to IMAP/SMTP, which this pipeline does not use.

---

## 0. Set project variables

```bash
export PROJECT_ID="YOUR_GCP_PROJECT_ID"
export REGION="asia-south1"   # Mumbai; change if needed
export REPO="fm-orders"
export IMAGE="fm-orders-job"
export JOB_NAME="fm-orders-bulletin"
export SCHEDULER_SA="fm-orders-scheduler"
export JOB_SA="fm-orders-runtime"

gcloud config set project "$PROJECT_ID"
```

---

## 1. Enable required APIs

```bash
gcloud services enable \
  run.googleapis.com \
  cloudscheduler.googleapis.com \
  secretmanager.googleapis.com \
  artifactregistry.googleapis.com \
  cloudbuild.googleapis.com \
  gmail.googleapis.com \
  sheets.googleapis.com \
  iam.googleapis.com \
  iamcredentials.googleapis.com
```

**Console:** APIs & Services → Enable APIs and Services → enable the same APIs.

---

## 2. Create a service account (runtime)

```bash
gcloud iam service-accounts create "$JOB_SA" \
  --display-name="FM Orders Bulletin Cloud Run Job"

export JOB_SA_EMAIL="${JOB_SA}@${PROJECT_ID}.iam.gserviceaccount.com"
```

### 2a. Grant project roles used by the job

```bash
# Read secrets at runtime
gcloud projects add-iam-policy-binding "$PROJECT_ID" \
  --member="serviceAccount:${JOB_SA_EMAIL}" \
  --role="roles/secretmanager.secretAccessor"

# Optional: write Cloud Logging (usually already available)
gcloud projects add-iam-policy-binding "$PROJECT_ID" \
  --member="serviceAccount:${JOB_SA_EMAIL}" \
  --role="roles/logging.logWriter"
```

### 2b. Create a JSON key for Gmail/Sheets API auth (store in Secret Manager)

```bash
gcloud iam service-accounts keys create ./sa-key.json \
  --iam-account="$JOB_SA_EMAIL"
```

Keep `sa-key.json` only long enough to create the secret, then delete the local file.

### 2c. Domain-wide delegation (Gmail) — Workspace Admin

1. Google Workspace Admin → **Security** → **Access and data control** → **API controls** → **Manage Domain Wide Delegation**.
2. **Add new** with the service account **Client ID** (from the JSON `client_id` field, or IAM → service account → Advanced → Client ID).
3. OAuth scopes (comma-separated):

   ```
   https://www.googleapis.com/auth/gmail.readonly,https://www.googleapis.com/auth/gmail.send
   ```

4. Save.

Without `gmail.readonly` the job cannot read `itdocumentation@forbesmarshall.com`. `gmail.send`
is what lets it mail the link to the finished sheet (see `NOTIFY_EMAILS`); without it the
notification step fails with `unauthorized_client` while the sheet is still written correctly.

### 2d. Share the Google Sheet

1. Open the target spreadsheet.
2. Share as **Editor** with: `fm-orders-runtime@YOUR_GCP_PROJECT_ID.iam.gserviceaccount.com`
3. Copy the Sheet ID from the URL:  
   `https://docs.google.com/spreadsheets/d/<GOOGLE_SHEET_ID>/edit`

---

## 3. Create and populate Secret Manager secrets

```bash
# Gemini API key (paste when prompted; do not echo into shell history if avoidable)
printf '%s' 'YOUR_GEMINI_API_KEY' | gcloud secrets create gemini-api-key --data-file=-

# Full service account JSON
gcloud secrets create google-sa-json --data-file=./sa-key.json
rm -f ./sa-key.json

# Google Sheet ID
printf '%s' 'YOUR_GOOGLE_SHEET_ID' | gcloud secrets create google-sheet-id --data-file=-

# Allow the job SA to access these secrets
for SECRET in gemini-api-key google-sa-json google-sheet-id; do
  gcloud secrets add-iam-policy-binding "$SECRET" \
    --member="serviceAccount:${JOB_SA_EMAIL}" \
    --role="roles/secretmanager.secretAccessor"
done
```

**Console:** Security → Secret Manager → Create secret → paste value / upload JSON.

To update a secret later:

```bash
printf '%s' 'NEW_VALUE' | gcloud secrets versions add gemini-api-key --data-file=-
```

---

## 4. Artifact Registry — build and push the image

```bash
gcloud artifacts repositories create "$REPO" \
  --repository-format=docker \
  --location="$REGION" \
  --description="FM Orders Bulletin images"

gcloud auth configure-docker "${REGION}-docker.pkg.dev"
```

### Option A — Cloud Build (recommended)

From the project directory containing `Dockerfile` and `main.py`:

```bash
gcloud builds submit --tag \
  "${REGION}-docker.pkg.dev/${PROJECT_ID}/${REPO}/${IMAGE}:latest"
```

### Option B — Local Docker

```bash
docker build -t "${REGION}-docker.pkg.dev/${PROJECT_ID}/${REPO}/${IMAGE}:latest" .
docker push "${REGION}-docker.pkg.dev/${PROJECT_ID}/${REPO}/${IMAGE}:latest"
```

---

## 5. Create the Cloud Run Job

```bash
gcloud run jobs create "$JOB_NAME" \
  --image="${REGION}-docker.pkg.dev/${PROJECT_ID}/${REPO}/${IMAGE}:latest" \
  --region="$REGION" \
  --service-account="$JOB_SA_EMAIL" \
  --task-timeout=15m \
  --max-retries=1 \
  --memory=1Gi \
  --cpu=1 \
  --set-env-vars="GMAIL_USER_EMAIL=itdocumentation@forbesmarshall.com,EMAIL_SUBJECT=FM Orders Bulletin,GEMINI_MODEL=gemini-3.1-pro-preview,SHEET_RANGE=Sheet1,TEMPLATE_SHEET_ID=17YLBsDFyJ3ejKc3y12hNd6mb1htcwovRR0lXiUrux4U,TEMPLATE_WORKSHEET_GID=1145057086,DRIVE_PARENT_FOLDER_ID=0AFfpBsQN7VH4Uk9PVA" \
  --set-secrets="GEMINI_API_KEY=gemini-api-key:latest,GOOGLE_SERVICE_ACCOUNT_JSON=google-sa-json:latest,GOOGLE_SHEET_ID=google-sheet-id:latest"
```

**Console:** Cloud Run → Jobs → Create → select image → Variables & Secrets → map secrets to env vars as above.

### Field mapping

`ORDERS BULLETIN_data_mapping.xlsx` decides which PDF line feeds which template field, using
its `tags_from_the_order_bulletin_pdf` and `MAPPING` columns. It is copied into the image, so
**editing the mapping means rebuilding and redeploying** — there is no runtime lookup. Its last
block also defines the company totals: `FMPL TOTAL` and `JV TOTAL` are summed from the section
totals listed there. Extra rules are computed in Python rather than by the model:

- INTOPS FMPL `MECH STDS` = `SSD+PAPER` + `SSD+PAPER-SG`
- Company totals = sums of the mapped section totals
- Process Domestic / CORE DOMESTIC `DIGITAL SUSTENANCE BUSINESS` copies the matching FMPL
  TOTAL's digital-sustenance cells (`D` as-is, `E` = `D`/12, current-FY `S`, and
  current-month projections `Y` + achmnt `AA`)
- CORE `MECH STDS + BOILERS + ENERGY SERVICES` = sum of those three division rows
- `GROUP TOTAL WITHOUT INTERCO` = PDF Grand Total; `GROUP TOTAL (including JV Interco)` =
  Grand Total + `INTERCO-JV` Total

Percentage columns are left blank in those computed rows, because percentages cannot be
meaningfully added.

### Drive location for the monthly files

`DRIVE_PARENT_FOLDER_ID` must point at a **Shared Drive** (or a folder inside one), and the
runtime service account must be added to that Shared Drive as **Content Manager**. A regular
My Drive folder does not work — the service account has no Drive storage quota of its own, so
copying the template there fails with `storageQuotaExceeded`. The only alternative is to set
`DRIVE_IMPERSONATE_USER` and grant the service account domain-wide delegation for
`https://www.googleapis.com/auth/drive`, so files are created under that user's quota instead.

### Deploy from GitHub (Cloud Build trigger)

Trigger `order-bulletin` watches `Baji1284/IT_order_bulletin` branch `main` and runs as
`fm-orders-cloudbuild@hr-docker.iam.gserviceaccount.com`.

1. Keep `cloudbuild.yaml` in the repo root (builds the image, pushes to Artifact Registry,
   updates Cloud Run Job `fm-orders-bulletin`).
2. Push to `main` — the trigger runs automatically.
3. Check: Cloud Build → History, then confirm the job image updated.

Manual run of the same config:

```bash
gcloud builds triggers run order-bulletin --region=asia-south1 --branch=main --project=hr-docker
```

### Manual test run

```bash
gcloud run jobs execute "$JOB_NAME" --region="$REGION" --wait
gcloud run jobs executions list --job="$JOB_NAME" --region="$REGION"
```

Check logs:

```bash
gcloud logging read \
  "resource.type=cloud_run_job AND resource.labels.job_name=${JOB_NAME}" \
  --limit=50 --format="value(textPayload)"
```

---

## 6. Cloud Scheduler — daily 9:05 AM IST

Create a dedicated SA that can invoke the job:

```bash
gcloud iam service-accounts create "$SCHEDULER_SA" \
  --display-name="FM Orders Scheduler Invoker"

export SCHEDULER_SA_EMAIL="${SCHEDULER_SA}@${PROJECT_ID}.iam.gserviceaccount.com"

# Allow Scheduler SA to run this Cloud Run Job
gcloud run jobs add-iam-policy-binding "$JOB_NAME" \
  --region="$REGION" \
  --member="serviceAccount:${SCHEDULER_SA_EMAIL}" \
  --role="roles/run.invoker"

# Scheduler needs permission to mint OIDC tokens as that SA
gcloud projects add-iam-policy-binding "$PROJECT_ID" \
  --member="serviceAccount:${SCHEDULER_SA_EMAIL}" \
  --role="roles/iam.serviceAccountUser"
```

Create the schedule (cron `5 9 * * *` = 09:05 daily, timezone `Asia/Kolkata`):

```bash
gcloud scheduler jobs create http fm-orders-daily-905 \
  --location="$REGION" \
  --schedule="5 9 * * *" \
  --time-zone="Asia/Kolkata" \
  --uri="https://${REGION}-run.googleapis.com/apis/run.googleapis.com/v1/namespaces/${PROJECT_ID}/jobs/${JOB_NAME}:run" \
  --http-method=POST \
  --oauth-service-account-email="$SCHEDULER_SA_EMAIL" \
  --oauth-token-scope="https://www.googleapis.com/auth/cloud-platform"
```

**Console:** Cloud Scheduler → Create Job → Frequency `5 9 * * *` → Timezone Asia/Kolkata → Target: HTTP → URL as above → Auth: OAuth / service account.

### Force a scheduler test

```bash
gcloud scheduler jobs run fm-orders-daily-905 --location="$REGION"
```

---

## 7. Environment variable reference

| Variable | Source | Purpose |
|----------|--------|---------|
| `GEMINI_API_KEY` | Secret `gemini-api-key` | Gemini API auth |
| `GOOGLE_SERVICE_ACCOUNT_JSON` | Secret `google-sa-json` | SA key for Gmail + Sheets |
| `GOOGLE_SHEET_ID` | Secret `google-sheet-id` | Target spreadsheet |
| `GMAIL_USER_EMAIL` | Env | Mailbox to impersonate |
| `EMAIL_SUBJECT` | Env | Exact subject: `FM Orders Bulletin` |
| `GEMINI_MODEL` | Env | Default `gemini-3.1-pro-preview` |
| `SHEET_RANGE` | Env | Worksheet name (default `Sheet1`) |

---

## 8. Update image after code changes

```bash
gcloud builds submit --tag \
  "${REGION}-docker.pkg.dev/${PROJECT_ID}/${REPO}/${IMAGE}:latest"

gcloud run jobs update "$JOB_NAME" \
  --region="$REGION" \
  --image="${REGION}-docker.pkg.dev/${PROJECT_ID}/${REPO}/${IMAGE}:latest"
```

---

## Troubleshooting

| Symptom | Likely cause |
|---------|----------------|
| `unauthorized_client` / Gmail 403 | Domain-wide delegation missing or wrong Client ID / scopes |
| `Not found` / Sheets 404 | Sheet not shared with SA email, or wrong `GOOGLE_SHEET_ID` |
| No email found | Subject mismatch, email not yet arrived, or mailbox empty |
| Gemini empty / invalid JSON | PDF too large / model issue — try `gemini-1.5-pro` via `GEMINI_MODEL` |
| Scheduler 403 | Scheduler SA missing `roles/run.invoker` on the job |

---

## Local dry-run (optional)

```bash
python -m venv .venv
# Windows PowerShell:
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt

$env:GEMINI_API_KEY="..."
$env:GOOGLE_APPLICATION_CREDENTIALS="C:\path\to\sa-key.json"
$env:GOOGLE_SHEET_ID="..."
$env:GMAIL_USER_EMAIL="itdocumentation@forbesmarshall.com"
python main.py
```

# ATS candidate seeder

Python script that reads candidate rows from an Excel workbook, posts each Google Drive resume link to the ATS API, and writes the result back into the same row.

The API parses the resume (this can take about 4–7 seconds per candidate). The script is sequential and **resume-safe**: you can stop it and run it again later. Rows that already succeeded, failed, or were skipped are not called again.

## Requirements

- Python 3.10+
- An Excel file with a `Resume Link` column (this repo includes `ProfilesData.xlsx`)
- A valid ATS Bearer token

## Setup

```text
python -m pip install -r requirements.txt
copy .env.example .env
```

Edit `.env` and set your token:

```text
ATS_API_URL=https://ats-api-agbpc2drfqccg4gg.southindia-01.azurewebsites.net/api/google-drive
ATS_BEARER_TOKEN=your_jwt_here
```

`ATS_API_URL` is optional; if omitted, the script uses the default URL above. `ATS_BEARER_TOKEN` is required.

Do not commit `.env`. It is listed in `.gitignore`.

## How to run

Keep the Excel file **closed** in Excel (and avoid OneDrive locking it) while the script runs.

Smoke test a few pending rows:

```text
python seed_candidates.py --file ProfilesData.xlsx --limit 3
```

Process the full file:

```text
python seed_candidates.py --file ProfilesData.xlsx
```

`--file` defaults to `ProfilesData.xlsx`, so this is equivalent:

```text
python seed_candidates.py
```

Stop with Ctrl+C at any time. The last completed row is saved before exit. Run the same command again to continue.

Re-attempt rows that previously failed (for example after a token or API fix). `SUCCESS` and `SKIPPED` rows are still left alone:

```text
python seed_candidates.py --file ProfilesData.xlsx --retry-failed
```

## Command-line options

| Option | Default | Description |
| --- | --- | --- |
| `--file` | `ProfilesData.xlsx` | Path to the workbook |
| `--limit N` | none | Process at most N pending rows this run |
| `--retry-failed` | off | Call the API again for rows marked `FAILED` |

## What the script sends

Each pending row with a Google Drive URL in **Resume Link** is posted as:

```json
{
  "driveLink": "https://drive.google.com/file/d/.../view"
}
```

Other Excel columns (name, skills, CTC, and so on) are not sent. The API is expected to parse the resume and store the candidate.

A row is treated as **success** when the API returns HTTP **200 or 201** and JSON with non-empty:

- `documentId`
- `candidateId`
- `importId`

Request timeout is 90 seconds.

## Excel columns written by the script

Existing columns are left unchanged. These columns are added on first run (or reused if they already exist):

| Column | Meaning |
| --- | --- |
| `documentId` | ID from a successful API response |
| `candidateId` | ID from a successful API response |
| `importId` | ID from a successful API response |
| `seed_status` | `SUCCESS`, `FAILED`, or `SKIPPED` |
| `seed_http_status` | HTTP status from the API (empty on network/timeout errors) |
| `seed_error` | Error text for failed or skipped rows |
| `seed_started_at` | Local time the API POST was sent (ISO 8601 with timezone offset) |
| `seed_finished_at` | Local time the response or error returned |
| `seed_duration_seconds` | Elapsed seconds for the API call (empty on skipped rows) |

### Status rules

- **SUCCESS** — 200/201 and all three IDs present. Will not be called again.
- **FAILED** — any other HTTP status, timeout, network error, or 200/201 without the three IDs. Will not be called again unless you pass `--retry-failed`.
- **SKIPPED** — empty `Resume Link`, or a URL that is not a Google Drive link. Will not be called again.

A later run only posts rows with an empty `seed_status` (plus `FAILED` when `--retry-failed` is set).

## Large files (10,000+ rows)

At ~5 seconds per API call, 10,000 new rows is on the order of 14 hours. Typical workflow:

1. Run with `--limit 3` to confirm the token and API.
2. Run without `--limit` for as long as you want (1 hour, overnight, etc.).
3. Stop and start again; already-written statuses are skipped.
4. Inspect `seed_status` / `seed_error` in Excel for failures.

The workbook is saved after every processed row so a crash does not lose completed work.

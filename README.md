# BRUHsailer interactive guide — repo setup

This repo hosts the interactive BRUHsailer guide and rebuilds it automatically
**every hour** by polling the live Google Docs for changes. No manual export
or git commits required for normal updates.

## How it works

1. Every hour, a GitHub Actions workflow runs in the background.
2. It downloads the latest .docx export of each chapter from Google Docs
   (using the public `export?format=docx` endpoint — no API keys needed).
3. It runs `build.py`, which parses the docx files and produces a fresh
   `index.html`.
4. If anything changed (the docs were edited since last run), it commits the
   refreshed `source/Chapter*.docx` and `index.html` back to the repo.
5. GitHub Pages picks up the new `index.html` within ~1 minute.

End-to-end lag between editing a Google Doc and the site catching up: at most
~70 minutes, usually less.

## Repo layout

```
.
├── .github/workflows/rebuild.yml   # GitHub Actions workflow
├── source/
│   ├── Chapter1.docx               # cached copies, refreshed hourly
│   ├── Chapter2.docx               #   from the configured Google Doc IDs
│   └── Chapter3.docx
├── base.html                       # site template (CSS / JS / EOC notes / sidebar)
├── build.py                        # fetcher + parser + splicer
├── index.html                      # generated output, served by GitHub Pages
├── requirements.txt
└── README.md                       # this file
```

## Where the Google Doc IDs live

In `build.py`, near the top:

```python
GOOGLE_DOC_IDS = {
    1: "1gCez5XG5FA1kmmBYydur3RaI_cr-dYNJlnigRrByEX8",
    2: "1YQiZ6curEYPpgm3DtjZcWHPoEEkGpYdXZ-I0gCM5p10",
    3: "1O1VeAkwS6VAzGVy0GT205GqiNaOAbw17H5uyuMwz39o",
}
```

If the source documents ever move to new Google Docs, edit those IDs and push.
The doc ID is the long string between `/d/` and `/edit` in the doc's URL.

## Requirements for the source docs

Each Google Doc must be shared so that **"Anyone with the link can view"** (or
more open). The build will fail with HTTP 403 if a doc is private or shared
only with specific people. The current docs are already publicly viewable, so
this should just work.

To verify a doc is publicly accessible: paste the export URL into an incognito
window:

```
https://docs.google.com/document/d/<DOC_ID>/export?format=docx
```

If a .docx download starts, you're good. If you see a sign-in page, the doc
needs its sharing settings adjusted.

## What gets auto-updated vs. what doesn't

| Part of the site | Updates automatically? | How to change |
|---|---|---|
| The 226 numbered steps and their contents | **Yes**, every hour | Edit the Google Docs |
| The CSS, sidebar, search box, filters | No — lives in `base.html` | Edit `base.html`, push |
| The End-of-Chapter notes (stats, references, links) | No — lives in `base.html` | Edit `base.html`, push |
| The "Mark Chapter X as complete" button | No — lives in `base.html` | Edit `base.html`, push |

The docx files are the source of truth for what's in each step. Everything
else is a static template you maintain by hand.

## Manual triggers

You don't normally need to do anything. But you can:

- **Force a rebuild right now**: go to the Actions tab → "Rebuild guide" →
  "Run workflow" button.
- **Trigger a rebuild by editing `base.html`**: any push that changes
  `build.py`, `base.html`, `requirements.txt`, or the workflow file itself
  triggers an immediate rebuild.

## Running locally to test before pushing

```bash
pip install -r requirements.txt
python build.py                  # fetches docx + rebuilds
python build.py --no-fetch       # uses existing local docx (offline)
```

Useful flags:
```bash
python build.py --source my_docx_folder --base custom_base.html --output out.html
```

## Polling frequency

The current schedule is `cron: '0 * * * *'` (every hour at minute 0). To change:

- Every 15 minutes: `'*/15 * * * *'`
- Every 6 hours: `'0 */6 * * *'`
- Daily at 06:00 UTC: `'0 6 * * *'`

Edit `.github/workflows/rebuild.yml` and commit. GitHub Actions is free for
public repos with no realistic cap, so even every 5 minutes would be fine
budget-wise.

## Troubleshooting

- **Workflow says "No changes — Google Docs unchanged"**: working as intended.
  The docs haven't been edited since the last run.
- **HTTP 403 on fetch**: a doc has been made private, or its ID has changed.
  Verify the export URL works in incognito. If the doc moved, update
  `GOOGLE_DOC_IDS` in `build.py`.
- **"Downloaded data is not a .docx file"**: Google returned an HTML page
  instead of the file. Usually means the doc requires login. Check sharing
  settings.
- **Site shows old content after a successful workflow**: browser caching.
  Hard refresh (Ctrl+Shift+R / Cmd+Shift+R), or wait ~60 seconds for the
  Pages CDN.
- **`python build.py` locally fails on `python-docx` import**:
  `pip install -r requirements.txt`.

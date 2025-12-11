# google-docs-resume-assistant

# Resume Assistant for Google Docs

A quick, simple tool to personalize resumes from inside Google Docs, via one UI, instead of copy/pasting into ChatGPT and manually reapplying changes.

## What it does
- Sidebar: paste a job description and get tailored suggestions inserted into your Doc.
- Improve Selection: highlight text and get a stronger, quantified rewrite.
- Formatting preserved: replacements keep your bold/italic/colors.

## Setup (bound script)
1) Open your resume in Google Docs.
2) Extensions → Apps Script (this creates a bound project).
3) Replace the default files with `Code.gs` and `sidebar.html` from this repo.
4) In `Code.gs`, set your key:  
   `const OPENAI_API_KEY = "YOUR_API_KEY_HERE";`  
   (Do not commit real keys.)
5) Save.

## Authorize
1) In Apps Script editor, select `authorizeScript` and click Run (▶️).
2) Approve the permissions (your own script).

## Use
- In the Doc: Resume Assistant → Open Assistant (sidebar).
- Set to 'Suggesting Mode'.
- Job Description: paste JD → Generate Suggestions.
- Improve Selection: select text in the Doc → Improve Selection.

## Notes / troubleshooting
- Browser blockers can prevent sidebar calls. If you see `PERMISSION_DENIED` or `google.script.run` failures, try incognito or a clean profile (no extensions, allow third-party cookies).
- The menu-driven functions run with full permissions; if the sidebar is blocked, use a clean profile or add a menu flow for JD processing.
- Model: `gpt-4o-mini`, temperature 0.2.

## Safety
- Keep your API key out of source control. Use the placeholder in `Code.gs` or set Script Properties in your own copy.

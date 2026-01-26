# Resume & Cover Letter Assistant for Google Docs

Tools to personalize resumes and cover letters from inside Google Docs, via one UI, instead of copy/pasting into ChatGPT and manually reapplying changes.

## Project Structure

```
resume_editor_googledocs/
├── resume/           # Resume Assistant tool
│   ├── Code.gs
│   └── sidebar.html
├── coverletter/      # Cover Letter Assistant tool
│   ├── Code.gs
│   ├── sidebar.html
│   └── README.md
└── README.md         # This file
```

## Resume Assistant

### What it does
- Sidebar: paste a job description and get tailored suggestions inserted into your Doc.
- Improve Selection: highlight text and get a stronger, quantified rewrite.
- Match Analysis: see how well your resume matches the job description.
- Formatting preserved: replacements keep your bold/italic/colors.
- Keyword optimization: automatically incorporates JD keywords into existing bullets.

### Setup (bound script)
1) Open your resume in Google Docs.
2) Extensions → Apps Script (this creates a bound project).
3) Replace the default files with `Code.gs` and `sidebar.html` from the `resume/` directory.
4) In `Code.gs`, set your key:  
   `const OPENAI_API_KEY = "YOUR_API_KEY_HERE";`  
   (Do not commit real keys.)
5) Save.

### Use
- In the Doc: Resume Assistant → Open Assistant (sidebar).
- Job Description: paste JD → Generate Suggestions.
- Improve Selection: select text in the Doc → Improve Selection.

## Cover Letter Assistant

### What it does
- Sidebar: paste job description, company name (optional), and resume (optional) to get tailored suggestions.
- Company Research: automatically fetches company overview from Perplexity API.
- Resume Alignment: optionally paste your resume for better alignment between cover letter and resume.
- Paragraph-level suggestions: optimizes cover letter paragraphs while maintaining narrative flow.
- Formatting preserved: replacements keep your formatting.

### Setup (bound script)
1) Open your cover letter in Google Docs.
2) Extensions → Apps Script (this creates a bound project).
3) Replace the default files with `Code.gs` and `sidebar.html` from the `coverletter/` directory.
4) In `Code.gs`, set your API keys:
   - `const OPENAI_API_KEY = "YOUR_API_KEY_HERE";`
   - `const PERPLEXITY_API_KEY = "YOUR_PERPLEXITY_API_KEY_HERE";` (optional, for company research)
   (Do not commit real keys.)
5) Save.

### Use
- In the Doc: Cover Letter Assistant → Open Assistant (sidebar).
- Resume (Optional): paste your resume for better alignment.
- Company Name (Optional): enter company name to fetch company overview.
- Job Description: paste JD → Generate Suggestions.

## Authorize (Both Tools)
1) In Apps Script editor, select `authorizeScript` and click Run (▶️).
2) Approve the permissions (your own script).

## Notes / troubleshooting
- Browser blockers can prevent sidebar calls. If you see `PERMISSION_DENIED` or `google.script.run` failures, try incognito or a clean profile (no extensions, allow third-party cookies).
- The menu-driven functions run with full permissions; if the sidebar is blocked, use a clean profile or add a menu flow for JD processing.
- Model: `gpt-4o-mini`, temperature 0.2.
- Perplexity API is optional for cover letters - if not configured, the tool will work without company research.

## Safety
- Keep your API keys out of source control. Use the placeholders in `Code.gs` or set Script Properties in your own copy.

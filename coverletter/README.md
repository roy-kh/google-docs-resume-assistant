# Cover Letter Assistant for Google Docs

A tool to personalize cover letters from inside Google Docs, similar to the Resume Assistant. Takes a general cover letter template and refines it for each position using the job description and company research.

## What it does
- Sidebar: paste job description, company name (optional), and resume (optional) to get tailored suggestions
- Company Research: automatically fetches company overview from Perplexity API
- Resume Alignment: optionally paste your resume for better alignment between cover letter and resume
- Formatting preserved: replacements keep your formatting
- Selective application: review suggestions before applying

## Setup (bound script)
1) Open your cover letter in Google Docs.
2) Extensions → Apps Script (this creates a bound project).
3) Replace the default files with `Code.gs` and `sidebar.html` from this directory.
4) In `Code.gs`, set your API keys:
   - `const OPENAI_API_KEY = "YOUR_API_KEY_HERE";`
   - `const PERPLEXITY_API_KEY = "YOUR_PERPLEXITY_API_KEY_HERE";` (optional, for company research)
   (Do not commit real keys.)
5) Save.

## Authorize
1) In Apps Script editor, select `authorizeScript` and click Run (▶️).
2) Approve the permissions (your own script).

## Use
- In the Doc: Cover Letter Assistant → Open Assistant (sidebar).
- Resume (Optional): paste your resume for better alignment
- Company Name (Optional): enter company name to fetch company overview
- Job Description: paste JD → Generate Suggestions.
- Review suggestions and apply selected ones.

## Features

### Resume Highlights Extraction
If you paste your resume, the tool extracts the most relevant 3-5 experiences that align with the job description. This helps ensure your cover letter references the right experiences.

### Company Research
Enter a company name to automatically fetch company overview including:
- Company culture and values
- Recent developments
- What makes them unique

This information is used to personalize your cover letter.

### Paragraph-Level Suggestions
Unlike resume bullets, cover letter suggestions work at the paragraph level, maintaining the narrative flow while optimizing for:
- Keyword alignment with job description
- Company-specific customization
- Clarity and professionalism

## Notes / troubleshooting
- Browser blockers can prevent sidebar calls. If you see `PERMISSION_DENIED` or `google.script.run` failures, try incognito or a clean profile.
- Perplexity API is optional - if not configured, the tool will work without company research.
- Model: `gpt-4o-mini`, temperature 0.2.

## Safety
- Keep your API keys out of source control. Use the placeholders in `Code.gs` or set Script Properties in your own copy.

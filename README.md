# Study Guide

Upload your lecture files (PDF, PPTX, DOCX) for any course and get back an interactive study guide: a full structured summary section-by-section, plain-language explanations of every hard concept, a searchable glossary of jargon, and an auto-generated quiz covering everything — instead of re-reading slides one by one. Remembers what you've studied and every quiz score, per course, across sessions.

## Setup (one time)

1. **Install Python dependencies** — open a terminal in this folder:
   ```
   pip install -r requirements.txt
   ```
   (If you already have the `.venv` from the La Liga Value Radar project, you can reuse it — activate it first, then run the same command.)

2. **Get an Anthropic API key** — https://console.anthropic.com/ → Settings → API Keys. This app calls the Claude API directly to build each study guide, so it needs your own key. You pay Anthropic directly for usage (a full course's worth of slides typically costs a small fraction of a dollar to process).

3. **Give the app your key**, either way:
   - Paste it into the sidebar each time you run the app (not saved to disk), or
   - Set it once as an environment variable so you never have to paste it:
     - Windows (PowerShell): `setx ANTHROPIC_API_KEY "sk-ant-..."` (restart your terminal after)

## Running it

```
streamlit run app.py
```

Opens in your browser automatically. Leave the terminal window open while you use it.

## How to use it

1. In the sidebar: name the course, upload every file you have for it (you can select multiple at once), click **Build study guide**. Takes a minute or two depending on how much material there is.
2. Pick the course from the sidebar list. Four tabs:
   - **Study guide** — every section as a collapsible card: full explanation, key points, worked examples, and a "Hard words in this section" popover. Search box filters sections live. Checkbox marks a section studied — saved permanently.
   - **Glossary** — every jargon term from the whole course, alphabetical, searchable.
   - **Quiz** — multiple-choice, covers the entire course. Submit to see your score, what you got wrong, and why — saved to your history.
   - **Progress** — quiz score history over time, and which sections are marked studied.
3. Add more courses any time the same way — the sidebar list grows, and each course's data (guide + glossary + quiz + your progress) is kept separately, permanently, in a local database file (`study_guide.db`, created automatically next to `app.py`).

## Notes

- **Model choice** (sidebar): Sonnet 5 is the default and the right balance for this. Haiku 4.5 is faster/cheaper if you're processing a lot of courses and don't need maximum depth; Opus 5 if a particularly dense course needs the most capable pass.
- **Scanned PDFs** (photographed slides, no real text layer) won't extract — text-based PDFs, exported PPTX, and DOCX all work.
- **Very large courses**: if one course's combined material is unusually huge, the app truncates to a safe size and tells you — split it into two "courses" (e.g. by semester half) if that happens often.
- Everything (the generated guides, your progress, your scores) lives in `study_guide.db` in this folder — back it up if you want to keep your history safe, and don't delete the file unless you want a clean slate.

# ESPER — Excel-based Spaced-repetition Program for Efficient Retention

Anki is extremely effective, but many people give up on it because it is hard to learn. ESPER was made for them: **ESPER is basically Anki made easy.** You can import an Anki `.apkg` file directly and ESPER converts it automatically.

ESPER is a static web app (GitHub Pages) for studying any set of information organized as key, value, and hint. Use the flashcard-based **Study page** for active recall or open the printable **A4 view** to review a whole set at a glance. FSRS manages card state and review scheduling locally in the browser.

On first launch, choose **한국어 사용자** or **English user**. Change the language at any time from the settings menu. Language changes the interface, AI prompt, and support QR only; it does not change workbook contents or study progress.

## Introduction video

[![Watch the ESPER introduction video](https://img.youtube.com/vi/J7beSHbJy94/hqdefault.jpg)](https://www.youtube.com/watch?v=J7beSHbJy94&t=1s)

## Usage

Open https://sunghokim128.github.io/ESPER/ in a Chromium-based browser (Chrome or Edge), which is required for the folder-access API. Projects are saved in a folder on your own device; the site stores nothing.

## Getting started

1. Choose a language when prompted. The format guide is available in Korean and English.
2. Add an Excel, CSV, TXT, TSV, Anki `.apkg`, or ESPER `.espr` file. Opening a cartridge takes you directly to its Study page.
3. On the Study page, try to recall the value for each key. Flip the card to reveal the value, then open its hint if needed. Rate your recall **Again**, **Hard**, **Good**, or **Easy**.
4. Use the header to switch between **Study**, **Learning status**, and **A4 view** once a project is open (the home page shows only the project list). The header always shows the save-location status: a green dot with the folder name when access is granted, or a blinking red dot with "Not set" when it is not. Folder errors appear in red under the button.
   - **Learning status** has three sections. **Settings** holds the exam D-day, target recall on the exam day, and daily new-card count; each row is edited on its own with its own save/cancel, and the first-pass completion date is derived from them. **My Log** has a **Simple / Detailed** toggle: Simple shows this month's calendar, streak, and today's study time and card count; Detailed shows a one-year contribution graph plus total study time and total cards studied. **Card understanding** also has Simple / Detailed: Simple is a tile grid in workbook order whose opacity reflects estimated recall (hover for key and probability), Detailed is a searchable, filterable card list where selecting a card shows its value, hint, recall probability, and FSRS details.
   - Browser Back is disabled; use the in-app navigation.
5. In A4 view, click a key to reveal or hide its hint. Choose **Learning priority** to put reviewed cards first from lowest estimated recall to highest, then start unreviewed cards on a separate section/page. Adjust image size and other display settings in the grouped settings panel. Use the header to return to Study.

Study rounds include up to 100 cards due for review, including new cards. During a round, **Undo rating** (button or ←) restores the previous card's FSRS state, removes that review, and shows the card again. The settings panel's **Edit current card** button ends the round, saves its record, and opens the edit page scrolled to and highlighting that card's row.

Settings let you change the item count, choose whether hints appear on the back of a card by default, and customize shortcuts. Default shortcuts are Space (flip), ↓ (hint), 1-4 (Again through Easy), ← (undo the last rating), S (pronounce key), and Esc (exit). FSRS supplies each card's next due time and the rating previews.

Each project card on the home page shows the number of items to review now: cards whose review is due plus today's remaining new-card allowance (the same count a Study round would contain).

The project list shows each configured exam's estimated average recall probability on that date. It is averaged over cards with review history; new cards without reviews are excluded.

## Universal workbook format

The first row is a header and is skipped. Data starts in row 2:

| Column | Contents |
|---|---|
| A | key |
| B | value |
| C | hint |

The key, value, and hint can represent any topic. For example, a vocabulary workbook can use an English word as the key, a Korean meaning or English definition as the value, and an English sentence as the hint. A medical-concept workbook can use a term as the key, its definition as the value, and a memory cue or clinical scenario as the hint. Enter line breaks in a hint cell when you want multiple lines displayed.

Exam dates use an in-app calendar localized to the selected language. The AI generation prompts are available in the format guide; they create three-column workbooks and do not generate example translations.

## Saving and records

ESPER processes files in the browser. Project settings, interface language, and study settings are kept in IndexedDB. Connect a save folder to keep each project and its files together:

```text
selected-folder/
└── vocabulary/
    ├── vocabulary.xlsx
    ├── vocabulary.json
    ├── vocabulary study record.json
    └── images/
        └── <image ID>.jpg
```

The project folder name is the uploaded workbook filename without its extension. The workbook filename is kept as uploaded; do not rename the project folder or its files in the save folder. Project names must be unique, and ESPER does not add suffixes or rename conflicting uploads. FSRS data, study records, and images are saved inside that project's folder. The bundled `ts-fsrs` 6.0.0 library runs FSRS-7 on-device; FSRS does not use a CDN, remote API, or runtime network connection.

- Uninstalling the extension can remove its IndexedDB project list and folder permission. After reinstalling, choose the same save folder again; ESPER scans project folders and restores supported study sets. Matching FSRS project IDs are reused, and study records and images stay with their project. The browser requires you to grant folder access again.
- The project-folder layout is the new pre-release format; ESPER does not migrate files from the earlier shared-folder layout.
- Each project folder has one matching `<workbook name>.json` file. It stores FSRS card state, IDs, and review logs, not ESPER's session duration or aggregate statistics.
- `<workbook name> study record.json` stores only My Log data: `{ "version": 1, "sessions": [{ "date", "itemCount", "durationSeconds" }] }`, one entry per finished study round. Per-key counts and other details are not saved. Earlier `.xlsx` study records are not read.
- Deleting a project removes its entire folder, including its workbook, FSRS data, study record, and images.
- Cards are matched across workbook edits by a SHA-256 fingerprint of key, value, and hint, not by row number. Deleted or ambiguous cards are archived rather than removed or silently assigned another item's history.
- Legacy workbooks with a recognized status/ignored column in A are read as the old B=key, C=value, D=hint layout. Opening them in the editor and saving converts them to the new A=key, B=value, C=hint layout.
- A review is written to the FSRS JSON before the Study session advances. Invalid or damaged FSRS JSON is not reset automatically; failed saves remain available for retry while the page stays open.
- Switching between Study, Learning status, and A4 view preserves the open project and current session.
- **Export to PDF** prints the A4 view; values that are hidden on screen remain hidden in the PDF.

## Project files

| File | Purpose |
|---|---|
| `index.html` | App document |
| `tutorial.html`, `landing.css` | Tutorial (demo video) page |
| `app.css` | Interface, Study cards, A4 view, and print styles |
| `app.js` | Language, project, navigation, Study, and A4 view behavior |
| `assets/guide/downloads/` | Bilingual AI prompts |
| `lib/projectStore.js` | IndexedDB settings and project files (workbook, FSRS JSON, study record JSON, images) in the connected save folder |
| `lib/studyRecords.js` | Study record JSON normalization and session append |
| `lib/studyPace.js` | Daily new-card pace and first-pass completion calculation |
| `lib/parseSheet.js` | Workbook parsing, editing, and export |
| `tests/` | Node test suites (`node --test tests/<file>`) |
| `lib/fsrs.js` | FSRS-7 scheduling, card synchronization, and JSON validation |
| `vendor/ts-fsrs/` | Locally bundled official `ts-fsrs` 6.0.0 browser build and license |
| `lib/sheetPages.js` | A4 pagination and item ordering |

## Acknowledgements

ESPER's review scheduling is built on [`ts-fsrs`](https://github.com/open-spaced-repetition/ts-fsrs) (v6.0.0), the TypeScript implementation of the FSRS (Free Spaced Repetition Scheduler) algorithm by [Open Spaced Repetition](https://github.com/open-spaced-repetition). The official browser (UMD) build is bundled unmodified in `vendor/ts-fsrs/` and runs entirely on-device, scheduling cards with FSRS-7 (`DefaultScheduler`). `ts-fsrs` is distributed under the MIT License; its copyright notice and license text are kept in `vendor/ts-fsrs/LICENSE`.

# Welcome to ESPER!

## Introduction
ESPER stands for **Excel-based Spaced-repetition Program for Efficient Retention**.
It is a spaced-repetition learning tool powered by FSRS-7, designed for people who want a simpler alternative to Anki's interface.
ESPER works for any kind of memorization that can be organized as key–value pairs, from vocabulary and formulas to scientific concepts and historical facts.
Import an existing workbook or Anki deck to get started without making flashcards one by one.
Supported imports: Excel workbooks (.xlsx, .xlsm), Anki decks (.apkg), and ESPER share files (.espr).

## ESPER learning process
1. **Open ESPER**
Open ESPER in a current Chromium-based browser such as Chrome or Edge.
Choose your interface language and a folder for saving projects.
2. **Import study material.**
Choose an Excel workbook, Anki deck, or ESPER `.espr` file.
Anki decks are converted into ESPER workbooks.
3. **Set your study plan.**
In Learning status, set an exam date and desired recall level for exam day.
ESPER suggests a first-pass target at about 80% of the time remaining, leaving the final period for reviews. Use the recommended daily new-card amount or choose your own.
4. **Study with active recall.**
Try to remember the value before revealing it. Use a hint if needed, then rate your recall **Again**, **Hard**, **Good**, or **Easy**.
FSRS-7 uses your ratings to schedule future reviews.
5. **Continue from your progress.**
Return to the project for scheduled reviews and track your study activity in Learning status.

    Below are some additional features:
- **A4 view:** Review multiple items in a printable layout. Use **Export to PDF** to print or save the view as a PDF.
- **Learning status:** Review your study log and card understanding, and adjust your study plan.


## Workbook format
ESPER uses a simple three-column format. The first row is a header; study items begin on the next row.
The same format works for words, equations, scientific concepts, dates, and more.
Use the in-app **Format Guide** for help preparing a compatible workbook.

- **Column A — Key:** The question, term, or concept to remember.
- **Column B — Value:** The answer or information to recall.
- **Column C — Hint:** Optional supporting information.

| Key | Value | Hint |
|---|---|---|
| Action potential | Na⁺ influx depolarizes; K⁺ efflux repolarizes. | ![Action potential diagram](./assets/guide/action-potential.svg) |

In the study card, the **Key** appears on the front. Try to recall the **Value**, then flip the card to reveal it. The **Hint** is available when you need extra help; it can include an image, as shown here.

![Action potential study card front and back: recall the key, then flip to see the value and optional image hint](./assets/guide/action-potential-flashcard.svg)


## Project file structure
When you choose a save folder, ESPER organizes project files like this:

```text
chosen-folder/
├── esper-library.json
└── project-name/
    ├── study-workbook.xlsx
    ├── study-workbook.json
    ├── project-name study record.json
    └── images/  (when the workbook uses images)
```

- `study-workbook.xlsx` — your study material saved as a workbook. Anki decks are converted to `.xlsx`
- `study-workbook.json` — FSRS card states and review history, plus the project's exam-planning settings.
- `project-name study record.json` — My Log study-session summaries.
- `images/` — images referenced by the workbook.
- `esper-library.json` — project order, colors, and user preferences. The browser also keeps project-list details, some settings, and save-folder permission.

The screenshots below show how projects appear in ESPER and how they are organized in the selected save folder:

**ESPER project list:** Each card represents a study project. The save-location indicator shows the connected folder.

![ESPER project list with multiple study projects and a connected save location](./assets/readme/project-home.png)

**Selected save folder:** The root contains `esper-library.json` and a separate folder for each project.

![Selected ESPER save folder containing esper-library.json and project folders](./assets/readme/save-folder.png)

**A project folder:** The workbook and its FSRS data are stored together. Current My Log records are saved as `project-name study record.json`; an extra `.xlsx` record visible in this example is not required by the current storage format.

![Example project folder containing its workbook and related data files](./assets/readme/project-folder.png)

Keeping study files in a folder you choose gives you control over where they are stored and lets you back up or move the folder yourself.
ESPER does not upload study materials or progress to an ESPER account or sync them between devices.
Reconnect the same folder to continue your projects with their saved progress.
You can also choose different folders for separate sets of projects; each folder works like its own workspace, similar to having multiple accounts, but ESPER does not create or manage user accounts.


## Data storage and privacy
Study materials and progress stay in the save folder you choose; ESPER does not upload them to an ESPER server.
The app fetches its tutorial and guide assets from the ESPER site when needed.
External sites are contacted only when you choose to open their links, and their privacy policies then apply.


## Acknowledgements
ESPER uses [`ts-fsrs`](https://github.com/open-spaced-repetition/ts-fsrs) version 6.0.0, by [Open Spaced Repetition](https://github.com/open-spaced-repetition), to implement FSRS-7 scheduling.
`ts-fsrs` is distributed under the MIT License; see [`vendor/ts-fsrs/LICENSE`](./vendor/ts-fsrs/LICENSE).

# ESPER
**Focus on learning, not on managing it.**

ESPER stands for **Excel-based Spaced-repetition Program for Efficient Retention**. It combines Flashcards and Skimming in an adaptive memorization program for vocabulary, formulas, scientific concepts, historical facts, and other key–value material.

## Motivation

Memorizing a large amount of material takes more than mastering each item one at a time. Repeatedly skimming the material helps build familiarity, layer by layer.

Conventional word lists can waste time repeating items you already know, while manually tracking the state of every item takes attention away from learning. Flashcard apps such as Anki address this with spaced repetition, but are less suited to quickly reviewing large collections of material.

ESPER combines both approaches so each can complement the other.

## How It Works

1. **Excel Import** — Upload an Excel workbook (`.xlsx` or `.xlsm`) containing the material you want to study. You can also import Anki decks (`.apkg`) and ESPER share files (`.espr`).
2. **Study Planning** — Set an exam date and target recall level. ESPER recommends a daily number of new items based on the time remaining.
3. **Flashcards** — Practice active recall, using hints when needed. FSRS-7 estimates how well each item is remembered and uses your **Again**, **Hard**, **Good**, or **Easy** ratings to schedule reviews. Study rounds end after 30 minutes by default; change the time limit in study settings.
4. **Adaptive Skimming** — Generate an A4 study sheet and optionally sort it by learning priority, with items with the lowest estimated recall first. Reveal or hide values on screen, study one page at a time in fullscreen, or export the sheets to PDF for printing. Skimming time is recorded in My Log and included in daily, monthly, and yearly study-time totals.

## Philosophy

You choose how to study; the algorithm decides what to prioritize.

ESPER aims to reduce the effort spent deciding what to memorize, when to review it, and which items may be slipping—so you can focus your time and attention on learning itself.

<a href="https://youtu.be/OC4Cubn3H78?si=3OSJvky_JRfOxakA"><img src="https://img.youtube.com/vi/OC4Cubn3H78/hqdefault.jpg" alt="Watch the ESPER tutorial video"></a>

## Getting started

1. Open [ESPER](https://sunghokim128.github.io/ESPER/) in a current Chromium-based browser such as Chrome or Edge.
2. Choose your interface language and a folder for your projects.
3. Import an Excel workbook, Anki deck, or ESPER `.espr` file. Anki decks are converted into ESPER workbooks.
4. In Learning status, set your exam date and target recall level. ESPER recommends finishing the first pass at about 80% of the time remaining, leaving the final period for reviews. Follow the recommended number of new items per day or set your own.
5. Choose **Flashcards** for scheduled active-recall practice or **Skimming** for an A4 overview. Learning status shows your plan and study history.

<br><br>

## Workbook format

ESPER uses a simple three-column format: The first row is a header, and study items start on the second row.<br>
Use this format for vocabulary, equations, scientific concepts, dates, and more.<br>
The in-app **Format Guide** can help you prepare a compatible workbook.

- **Column A — Key:** The question, term, or concept to remember.
- **Column B — Value:** The answer or information to recall.
- **Column C — Hint:** Optional supporting information.
<br>

| Key | Value | Hint |
|---|---|---|
| Action potential | Na⁺ influx depolarizes; K⁺ efflux repolarizes. | ![Action potential diagram](./assets/guide/action-potential.svg) |
<br>

During study, the **Key** appears on the front of the card. Try to recall the **Value**, then flip the card to reveal it. You can show the **Hint** when you need extra help; hints can include images.

![Action potential study card front and back: recall the key, then flip to see the value and optional image hint](./assets/guide/action-potential-flashcard.svg)

<br><br>

## Project file structure

Choose a save folder and ESPER will organize your projects like this:

```text
chosen-folder/
├── esper-library.json
└── project-name/
    ├── study-workbook.xlsx
    ├── study-workbook.json
    ├── project-name study record.json
    └── images/  (when the workbook uses images)
```

- `study-workbook.xlsx` — The study material. Anki decks are converted to `.xlsx`; imported Excel workbooks keep their supported file extension.
- `study-workbook.json` — FSRS card states and review history, along with exam-planning settings.
- `project-name study record.json` — My Log session dates, study mode, and duration. Flashcard sessions also record cards studied; Skimming sessions record study time without a card count. Monthly view shows a selectable calendar and deletable session records; Yearly view shows the annual activity graph and totals.
- `images/` — Images used in the workbook.
- `esper-library.json` — Project order, colors, and user preferences. Project-list details, some settings, and save-folder permission are also kept in the browser.

These screenshots show the project list and the corresponding folders on your device:

**ESPER project list** — Each card represents a study project. The save-location indicator shows the connected folder.

![ESPER project list with multiple study projects and a connected save location](./assets/readme/project-home.png)

**Selected save folder** — The root contains `esper-library.json` and a separate folder for each project.

![Selected ESPER save folder containing esper-library.json and project folders](./assets/readme/save-folder.png)

**A project folder** — The workbook and FSRS data are stored together. The current My Log file is `project-name study record.json`. The `.xlsx` record shown in the screenshot is not used by the current format.

![Example project folder containing its workbook and related data files](./assets/readme/project-folder.png)

Choose where your study files are stored, and copy or move the folder yourself when you need to back it up or transfer it.<br>
Reconnect the same folder to continue your projects with their saved progress.<br>
You can use different folders as separate workspaces for different sets of projects. These workspaces may feel like separate accounts, but ESPER does not create or manage user accounts.

<br><br>

## Data storage and privacy

ESPER fetches tutorial and guide assets from its website when needed.<br>
The main app uses Google Analytics 4 to measure visits and usage. Loading the app sends analytics data to Google, which may include browser, device, and interaction information. See [Google's Privacy Policy](https://policies.google.com/privacy).<br>
Third-party sites are contacted only when you choose to open their links; their privacy policies apply when you visit them.

<br><br>

## Acknowledgements

ESPER uses a vendored copy of [`ts-fsrs`](https://github.com/open-spaced-repetition/ts-fsrs) version 6.0.0, created by [Open Spaced Repetition](https://github.com/open-spaced-repetition), to implement FSRS-7 scheduling. The library's bundled source is [`vendor/ts-fsrs/index.umd.js`](./vendor/ts-fsrs/index.umd.js).<br>
`ts-fsrs` is distributed under the MIT License. Its copyright and license notice are preserved in [`vendor/ts-fsrs/LICENSE`](./vendor/ts-fsrs/LICENSE).

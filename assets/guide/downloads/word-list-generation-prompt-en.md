# English Vocabulary Study Set Excel Guide

Create an English vocabulary study set from the words or topic the user provides (for example TOEFL, IELTS, business English), and provide a real `.xlsx` file with a download link.

## Column layout

Use the first row for headers and start the words on row 2. Create exactly these three columns; do not add column D or beyond.

| Column | Header | Contents |
|---|---|---|
| A | key | The English word (e.g. `impair`) |
| B | value | A short, learner-friendly English definition (e.g. `to make something worse or weaker`) |
| C | hint | One natural English example sentence (e.g. `Noise can impair your concentration.`) |

## Writing rules

- Write value as a clear, concise English definition in simple words, without using the key word itself. If a word has several senses, give only the one or two most common.
- Write hint as an example sentence only. Do not add translations or notes in any other language.
- Make each example a natural, everyday sentence that is neither too long nor too difficult.
- If the user supplies specific words, use exactly those words; do not add or drop any. If they supply only a topic, choose suitable words in the requested number.
- Merge duplicate words into one row. Do not guess at corrections to uncertain spellings.
- Enable text wrapping and use readable column widths.

## Final checks

- Is A = word, B = English definition, C = English example sentence?
- Is column D and beyond empty?
- Are there no translations anywhere?
- Did you provide a download link to a real `.xlsx` file?

Now create the Excel file from the user's words or topic using these rules.

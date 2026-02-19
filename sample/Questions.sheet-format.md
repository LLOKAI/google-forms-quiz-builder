# Questions Sheet Format

This sheet drives the **Google Forms Quiz Builder**.
It supports a small **metadata block** followed by a **questions table**.

---

## 1) Metadata rows (top of sheet)

Place these in the first column (A), one per row:

| Key                | Value (Column B)                | Notes                                               |
| ------------------ | ------------------------------- | --------------------------------------------------- |
| `FormTitle`        | e.g., `Math Assignment1`        | Title for the generated Form copy.                  |
| `FormDescription`  | e.g., `Maths quiz for beginner` | Optional; appears under the title.                  |
| `LimitOneResponse` | `TRUE` / `FALSE`                | If `TRUE`, requires sign-in & limits to 1 per user. |

Add a **blank row** after metadata, then the questions header.

> The script looks for `FormTitle`, `FormDescription`, `LimitOneResponse` case-insensitively in column A.
> If any are missing, defaults are used (title: _Untitled Quiz_, description empty, limit false).

## 2) Questions header (required)

Create this header row exactly (order matters for A..D):

```
Section | Question | Type | Points | AnswerA | AnswerB | AnswerC | AnswerD | ImageURL | AnswerAImageURL | AnswerBImageURL | AnswerCImageURL | AnswerDImageURL
```

- Header matching is **case-insensitive** but spelling must match.
- Only `AnswerA..AnswerD` are parsed (max 4 options).
- `ImageURL` and `Answer*ImageURL` columns are **optional** — omit them if not needed.

## 3) Question rows (one per row)

| Column         | Required | Description                                                                                                               |
| -------------- | :------: | ------------------------------------------------------------------------------------------------------------------------- |
| **Section**    |    ✅    | Groups questions and inserts a page break titled: `"<section> Section — <sum pts> pts total"`.                            |
| **Question**   |    ✅    | The prompt shown to students.                                                                                             |
| **Type**       |    ✅    | Must be `MCQ` (single choice). Matching is case-insensitive. |
| **Points**     |    ✅    | Integer points for the item. Non-numeric or blank → treated as **0**.                                                     |
| **AnswerA..D** |    ✅    | Used for `MCQ`. Provide at least 2 options.                                                |
| **ImageURL**   |    ❌    | Optional. A public URL to an image. If provided, an image is inserted before the question in the form.                    |
| **AnswerAImageURL..AnswerDImageURL** | ❌ | Optional. For MCQ/MSQ, each URL displays an image item labeled A..D before the question. |

### Type-specific rules

- **MCQ (Multiple Choice, one correct)**

  - Provide ≥ 2 unique options among `AnswerA..D`.
  - **Correct answer = `AnswerA`** (first non-blank).
  - Options are **shuffled** for students.
  - If < 2 options remain after de-dupe → falls back to **SA**.


## 4) Image support

- Add an `ImageURL` column header after `AnswerD`.
- For each question row, provide a **publicly accessible** image URL (e.g., from Google Drive with "Anyone with the link" sharing, or any public web URL).
- The script will fetch the image and insert it as an **Image item** directly before the question in the form.
- If the URL is invalid or unreachable, the image is skipped with a log warning (the question is still created).
- To show images for answer options, add `AnswerAImageURL..AnswerDImageURL`.
- In answer-image mode, the form displays labeled images (`A`, `B`, `C`, `D`) and the MCQ choices are those labels.
- **Note:** this is not embedded image-inside-choice rendering (a Forms API limitation); it is labeled image items plus selectable labels.

## 5) Answer handling

- **Uniqueness/cleanup**: Options are trimmed and compared case-insensitively; duplicates are removed.
- **Empty cells**: Fine; leave unused `AnswerC/D` blank.
- **Formulas**: Avoid formulas that evaluate to empty strings — prefer literal blanks.

## 6) Section behavior

- When the `Section` value changes, the script inserts a **new page** with:
  - Title: `"<section> Section — <total points in this section> pts total"`.
- Section totals = sum of `Points` in that section (non-numeric → 0).

## 7) What the script does automatically

- Creates a Form (from template if configured) in **Quiz mode**.
- Collects email; adds a required **Student Name** question.
- Links responses to the **same spreadsheet** (new "Form Responses" tab).
- Publishes the Form and retrieves the student link (`/forms/d/e/.../viewform`).
- Optionally sets **"Anyone with the link"** (if allowed by domain settings).
- Optionally **posts the form to Google Classroom** as a quiz assignment or class material.

## 8) Tips & gotchas

- Keep `Type` set to `MCQ` only.
- `Points` must be integer; non-numeric = 0.
- Avoid trailing spaces (`"0.3 "` → `"0.3"`).
- Fractions/percentages/decimals can be plain text: `2/3`, `12.5%`, `0.125`.
- `ImageURL` must be publicly accessible — private URLs will fail to fetch.

# Setup Guide

This project turns a Google Sheet of questions into a fully published Google Form quiz,
with optional **Google Classroom** integration.
We can run it as a **container-bound Apps Script** inside our Sheet (simplest method).

---

## 1) Prepare your Google Sheet

1. Create or open a Google Sheet (e.g. `Questions`).
2. Add rows according to [Questions Sheet Format](../sample/Questions.sheet-format.md).
   - Start with metadata rows (`FormTitle`, `FormDescription`, `LimitOneResponse`).
   - Leave a blank row.
   - Add the questions table (`Section, Question, Type, Points, AnswerA..D`).
   - Optionally add an `ImageURL` column for question images.

We can also import the provided sample CSV:
[sample/questions.sample.csv](../sample/math_assignment1_questions.csv)

## 2) Add the Apps Script

1. In your Sheet: go to **Extensions => Apps Script**.
2. In the editor:
   - Create or update the file named `Code.gs`.
   - Copy all contents from [src/Code.js](../src/Code.js) into it.
3. (Optional but recommended) Enable **Project Settings => Show "appsscript.json" manifest file**
   - Replace the contents with [src/appsscript.json](../src/appsscript.json) for scopes and advanced services.

## 3) Enable Drive API (optional)

This script needs the Drive API to adjust sharing and publish settings.

1. In the Script Editor:
   - Click the **+ Services** icon (left sidebar).
   - Enable **Drive API v2** (Advanced Google Services).
2. In the popup footer, click **Google Cloud Platform API dashboard**.
   - Ensure **Drive API** is enabled for your account/project.

## 4) Enable Classroom API (required for Classroom features)

To post quizzes to Google Classroom, you must enable the Classroom API:

1. In the Script Editor:
   - Click the **+ Services** icon (left sidebar).
   - Search for and enable **Google Classroom API v1**.
2. In the Google Cloud Console (linked to your Apps Script project):
   - Go to **APIs & Services => Library**.
   - Search for **Google Classroom API** and click **Enable**.
3. The first time you use a Classroom feature, Google will prompt you to authorize
   the additional Classroom scopes (courses, coursework, topics).

> **Note:** You must be a **teacher** in at least one Google Classroom course
> for the integration to work. The script lists only courses where you are a teacher.

## 5) Configure

At the top of `Code.gs` we'll find configuration options:

```javascript
const TEMPLATE_FORM_ID = "YOUR_TEMPLATE_FORM_ID"; // or '' for a fresh blank Form
const MAKE_PUBLIC_ANYONE_WITH_LINK = true; // true = allow external students
```

## 6) Using the Form Builder Menu

After reloading the spreadsheet, a **Form Builder** menu will appear:

| Menu Item                           | Description                                                            |
| ----------------------------------- | ---------------------------------------------------------------------- |
| **Create Form (Confirm)**           | Creates the Google Form quiz and links responses to the spreadsheet.   |
| **Create Form & Post to Classroom** | Creates the form, then opens a dialog to post it to Google Classroom.  |
| **Post Last Form to Classroom**     | Opens the Classroom dialog for the most recently created form.         |

### Classroom Dialog Options

When posting to Classroom, a dialog appears where you can configure:

- **Course** — select from your active Classroom courses
- **Post As** — Quiz Assignment (graded) or Class Material (ungraded)
- **Title** — pre-filled from the form title
- **Description** — optional instructions for students
- **Max Points** — for graded assignments (default: 100)
- **Due Date / Time** — optional deadline
- **Topic** — optional; organizes the post under a Classroom topic

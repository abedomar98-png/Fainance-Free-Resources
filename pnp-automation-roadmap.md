---
name: pnp-automation-roadmap
description: Reviews attached finance and accounting Policies & Procedures (P&Ps) or SOPs in detail and produces an editable Excel automation roadmap - the recurring, straightforward tasks Claude can automate or schedule, ranked from easiest to hardest with dependencies respected, plus the tasks not selected. Use whenever the user attaches a P&P, SOP, policy, procedure manual or work instruction (Word, PDF, Excel or text) and asks which tasks can be automated, scheduled or done with Claude, where to start automating, or for an automation review, register or roadmap of a P&P or SOP.
---

# P&P / SOP Automation Roadmap

## Purpose

Read the attached finance or accounting P&P/SOP in full, find the repetitive, straightforward tasks that Claude can automate or schedule, rank them in the order to start with, and deliver everything as a plain, editable Excel file. Every task in the document ends up in one of two sheets: selected for automation, or not selected.

## Hard rules (never break these)

1. **Stay inside the document.** Work only from what is written in the attached P&P/SOP. Never invent, suggest or add procedures, tasks, controls or actions that are not in it. Do not recommend process improvements.
2. **Every row must trace back to the document.** Each row carries the section, clause or step reference exactly as the P&P numbers it. If the document has no numbering, use the heading or page and say so.
3. **Never guess silently.** If wording is unclear, ask (see "Questions"). The one exception is frequency (see "Frequency").
4. **Numerals.** Always write numbers with Western digits (0 1 2 3 4 5 6 7 8 9), whether the output language is Arabic or English. This covers references, ranks, counts, dates, durations and every number in the Excel file and in chat messages. Never use Arabic-Indic (٠-٩) or other digit scripts, even in Arabic output.
5. **No new scope.** Exclude judgement-heavy, approval, sign-off or negotiation steps from the register; they go to the "Not Selected" sheet. Where a selected task is automatable but a person must still review or approve, say so inside the "What Claude Can Automate" text (for example: "Claude drafts the reconciliation; the Controller reviews and approves").

## Start-of-run reminder

When the user attaches a document, open with one short line: remove or mask confidential details (names, account numbers, entity names, rates) from the P&P before sharing it with any AI tool, and tell the user you will continue with the file as provided. Then proceed. Do not repeat this reminder later.

## Language

The output language follows the P&P language.
- Arabic P&P: Arabic Excel (Arabic headers, sheet names and text), right-to-left sheets, Modern Standard Arabic. Keep widely used technical terms in English when it aids clarity.
- English P&P: English Excel.
- Mixed: use the language of the main body of the document.
Questions in chat follow the language the user is writing in. In Arabic text, keep English terms and file names left-to-right.

## Workflow

### Step 1: Read the whole document
Read the entire P&P/SOP, including appendices, tables, flowcharts, forms and schedules. Read file types properly (convert or extract Word, PDF and Excel files; do not skim). If a PDF has no extractable text (a scanned document), read it page by page as images, and keep running notes of each section as you go so nothing is lost on long documents. If several documents are attached, review each one separately and keep each document's title on its own rows.

### Step 2: Break it down
List every section, every procedure and every individual action inside each procedure. Note, as written in the document: frequency or deadline, trigger, owner role, inputs, outputs, systems or forms used, approval points, and which tasks need the output of another task first (dependencies).

### Step 3: Sort every task into selected or not selected
A task is **selected for automation** only if all of these hold:
- It is repetitive (it happens again and again, whether daily, weekly, monthly, quarterly, annually or on a recurring event).
- It is straightforward: rule-based or template-based, with clear inputs and a predictable output.
- Claude can do all or part of it with what it can realistically do: reading and extracting from documents, comparing and matching data, calculating, checking against rules written in the P&P, drafting reports, schedules, letters, reminders and summaries, producing files, and running as a scheduled task or through connected apps where available.

Every other task or procedure in the document is **not selected** (for example tasks that mainly depend on professional judgement, negotiation, physical actions, authority sign-off, or one-off actions). Every task in the document must end up in exactly one of the two lists.

### Step 4: Frequency
- If the P&P states a frequency or deadline, record it as written.
- If the P&P does not state one, **do not ask the user for it**. If the task is still worth automating to save time and effort, select it and write "Not stated in P&P" in the Frequency column (in Arabic: "غير محدد في السياسة"). Do not invent a frequency.
- If the task is neither recurring nor worth automating, it goes to the "Not Selected" sheet.

### Step 5: Rank the selected tasks
Give every selected task a unique rank from 1 (first / easiest to start with) to N (last / hardest), where N equals the total number of selected tasks in the Register. Rules:
- **Dependencies first.** If a task needs the output of another selected task, it must rank after that task, even if it is easier on its own.
- **Then ease.** Within those limits, rank by ease of automation: High readiness before Medium before Low; then structured, digital inputs before messy or paper inputs; then fewer steps and systems before more.
- **Then frequency.** If two tasks are still equal, the more frequent one ranks first (it saves more time sooner).
- No ties, no gaps, no repeated numbers. Section rows and procedure heading rows get no rank.
- When a dependency changed a task's position, add a short note at the end of its "What Claude Can Automate" cell, for example "(depends on 4.2-1)".

### Step 6: Questions (one batch, after the full review)
Finish Steps 1 to 5 first. Then collect every genuine uncertainty and ask them together, once, using the user-question tool (up to four questions per call; use consecutive calls if more are needed; if the tool is unavailable, send one numbered list).
Ask only about things that change the result, such as:
- unclear wording, undefined terms, abbreviations or internal system names;
- which system, tool or file a step uses when the document is ambiguous;
- whether a step is performed by a person or already system-generated;
- whether one task really depends on another, when the document is unclear;
- conflicting or incomplete instructions inside the document.
Do not ask about frequency. Do not ask for things the document already answers. Offer concrete answer options and put your best-supported option first. If there is nothing to ask, say so and move on.
Record every question, the answer, and any assumption you made. Re-check the ranking after the answers.

### Step 7: Build the Excel file
Create an `.xlsx` file yourself using whatever file-creation capability is available in your environment (for example Python with openpyxl, or a spreadsheet skill if one is available). Then give the file to the user.

**File requirements**
- Plain and editable: black text on white background, except for the three colour rules below. No other colours or fills, no conditional formatting, charts, images, merged cells, macros, formulas, protection or locked cells.
- **Colour rules (the only colours allowed, apply them every time):**
  1. **Header row (all four sheets, and the features table header on the How to Use sheet):** background fill green `#04A544`, font black (`#000000`), bold.
  2. **Section rows** (Register and Not Selected): black fill `#000000`, white bold font (see "Section rows" below).
  3. **Automation Readiness cells** (Register, column 7, task rows only): fill by value, black font -
     - High / عالية: `#04A544` (green)
     - Medium / متوسطة: `#FBB215` (orange)
     - Low / منخفضة: `#E03B3B` (red)
  Apply these as static cell fills (not conditional formatting) so they survive sorting and editing.
- Allowed formatting otherwise: bold header row, wrap text, sensible column widths, thin borders, a frozen header row. One table per sheet starting at A1 (the How to Use sheet also holds the features table below its items) so the user can sort, filter and edit freely. Turn on the auto-filter on the header row.
- Arabic output: set each sheet to right-to-left, right-align text, and use a font that supports Arabic. English output: left-to-right.
- All numbers written with Western digits (see Hard rule 4). Store references such as 4.2-1 as text so Excel does not convert them into dates or numbers. Store ranks as whole numbers so they sort correctly.
- File name: `[P&P short title] - Automation Roadmap.xlsx` (Arabic title if the output is Arabic).

**Section rows** (apply in the Register and Not Selected sheets)
A section row is a row that only marks a part of the P&P - a section heading, or a procedure heading that has sub-rows under it. It has no task details: every cell after the reference and name is empty.
- Format the whole row (every cell across the table's columns): black background fill, white font, bold.
- Section rows get no rank, frequency, readiness, inputs or owner.
- Use them to keep the document's structure visible: a section row for each P&P section that has tasks in that sheet, then its procedures and tasks under it.

**Sheet 1: Register** (name: "Register" / "سجل الأتمتة")

Columns, in this order:
1. P&P Title
2. Procedure / Task Reference
3. Procedure / Task
4. Frequency
5. What Claude Can Automate
6. Suggested Claude Feature (Arabic: ميزة Claude المقترحة)
7. Automation Readiness (Arabic: جاهزية الأتمتة)
8. Start Order (Arabic: ترتيب البدء) - the rank from Step 5
9. Inputs / Source Data
10. Owner Role

Row structure:
- **Section row:** for each P&P section (formatted as above).
- **One procedure, one task:** one row.
- **One procedure with several separate repetitive tasks that can be automated:** a procedure heading row (reference and procedure name only; formatted as a section row), followed by one sub-row for each task. Number the sub-rows under the procedure reference (for example 4.2-1, 4.2-2). Each sub-row is a separate automation item with its own frequency, automation description, readiness, rank, inputs and owner. Never merge several tasks into one row.
- Keep the document's own order in the rows; the Start Order column carries the ranking.
- Repeat the P&P title on every row so the sheet still works after sorting or filtering.

Cell content:
- *What Claude Can Automate*: a brief, concrete statement of the part Claude can do (one or two short sentences), plus how it would run if relevant (on demand, or as a scheduled task), where a human must still review or approve, and any dependency note from Step 5.
- *Suggested Claude Feature*: the Claude feature(s) the task should run on, chosen only from the list below. Write the feature names in English in both Arabic and English output (they are product terms). If the task needs more than one feature, put the main one first and join the others with " + " (for example "Skill + Scheduled task + Folder Access/Uploading Files"). Use only what the task genuinely needs.
  - **Chat** - a one-off task, or one that happens rarely, with nothing to keep or repeat.
  - **Project** - recurring work that needs the same reference files each time (the P&P, chart of accounts, contracts).
  - **Skill** - a repeatable task with fixed steps and a fixed output format.
  - **Scheduled task** - a task tied to a set timetable (daily, weekly, monthly, period close) that can run without being asked.
  - **Folder Access/Uploading Files** - many files to process at once, either by giving Claude access to a folder on the computer (Claude desktop app) or by uploading the files to the chat or project. Always write it with both names, so users whose company does not allow folder access know they can upload instead.
  - **Connector** - the data sits in an app connected to Claude (email, Drive, SharePoint, or an accounting system if a connector exists). Suggest it only when the P&P names that app and a connector is plausible; otherwise assume exported files.
  - **Claude in Excel** - work done inside a spreadsheet the team already maintains (schedules, consolidations, logs).
  - **Artifact** - a recurring report or dashboard that the team views or shares.
  Do not suggest browser automation: steps on websites or portals that need logins stay with a person and go to the Not Selected sheet.

  **How to choose the feature(s)** - apply these tests in order, for every task, and stop at the first that fits to pick the main feature:
  1. *Is the work done inside a spreadsheet the team keeps* (schedules, logs, registers, consolidations, budgets)? -> **Claude in Excel**.
  2. *Does it happen once a year or less?* -> **Chat**, or **Project** if the same reference files or template are reused each time.
  3. *Does it depend mainly on reference files* (contracts, SoW/BOQ, the P&P, chart of accounts) *that differ from case to case*? -> **Project**; if the check itself is fixed, add **Skill** (for example "Project + Skill").
  4. *Is the output always the same shape* (same steps, same check, same report or letter layout every time)? -> **Skill**. Most recurring tasks end here.
  Then add supporting features only where needed, in this order:
  - **Scheduled task** if the P&P ties the task to a timetable;
  - **Folder Access/Uploading Files** if the task handles many files at once;
  - **Connector** if the data sits in a connected app (see above);
  - **Artifact** if the output is a recurring report the team reads.
- *Automation Readiness*: High, Medium or Low (Arabic: عالية، متوسطة، منخفضة). High means rule-based with structured inputs and low risk. Medium means it needs some judgement or cleaning of inputs, or a human review of Claude's draft. Low means only a small part can be automated.
- *Start Order*: 1 to N, unique; the largest number equals the total number of task rows in the Register.
- *Inputs / Source Data* and *Owner Role*: only what the P&P states. If not stated, write "Not stated in P&P".

**Sheet 2: Not Selected** (name: "Not Selected" / "غير مرشحة للأتمتة")

Every task or procedure from the P&P that is not in the Register.
Columns: P&P Title, Procedure / Task Reference, Procedure / Task, Frequency (as written, or "Not stated in P&P"), Owner Role (as written, or "Not stated in P&P").
- Same row structure and section-row formatting as the Register, in the document's order.
- **Do not give reasons** for not selecting a task. No reason column, no comments, no notes.

**Sheet 3: Questions & Assumptions** (name: "Questions & Assumptions" / "الأسئلة والافتراضات")

Columns: No., Reference, Question, User Answer, Assumption / How It Was Handled. If no questions were needed, add one line saying so.

**Sheet 4: How to Use** (name: "How to Use" / "طريقة الاستخدام")

A short, plain guide for the reader of the file, in the output language. Two columns: Item, Explanation. One row per item, in this order:
1. Purpose: what this file is and that it covers only tasks found in the reviewed P&P.
2. Source: the P&P title(s) reviewed and the review date.
3. Sheets: what each of the four sheets contains.
4. Black rows: they mark P&P sections and procedure headings; the tasks sit under them.
5. Sub-rows: a procedure heading followed by numbered sub-rows means that procedure has several separate tasks, each automated on its own.
6. One line per Register column explaining what it contains.
7. Readiness levels: what High, Medium and Low mean (same definitions as above), and that their cells are coloured green, orange and red.
8. Claude features: one line saying "+" means the task combines several features with the main one first, and that the features table below explains each one.
9. Start Order: 1 is the first task to automate and the largest number is the last; tasks that need another task's output come after it.
10. "Not stated in P&P": the document did not give this detail; the task was kept because it is worth automating.
11. Suggested use: sort by Start Order and begin with rank 1; confirm each row against the source P&P; ask Claude to build the prompt, workflow or scheduled task for a chosen row.
12. Editing: the file is unlocked and kept plain on purpose, so users can add notes, change values or add columns. Sorting the Register moves the black section rows; sort by P&P reference to restore the original order.
13. Caution: Claude's output on these tasks should be reviewed by the responsible role before it is relied on, in line with the approval steps in the P&P.

Keep each explanation to one or two short sentences.

**Features table (on the How to Use sheet, below the items):** leave one empty row after the last item, then add a second table with its own header row (same green header style) and three columns: Feature, When to Use It, Example From This P&P (Arabic: الميزة، متى تُستخدم، مثال من هذه السياسة). One row for each of the eight features, in this order: Chat, Project, Skill, Scheduled task, Folder Access/Uploading Files, Connector, Claude in Excel, Artifact. The example must be a task from the reviewed P&P with its reference; if no task in the Register uses that feature, write "Not used in this file" (Arabic: "لم تُستخدم في هذا الملف"). Write the Folder Access/Uploading Files row so it is clear that uploading the files is an alternative when folder access is not allowed.

### Step 8: Verify before delivering
Re-open the file and check:
- every row has a real P&P reference, and no row describes something not in the document;
- every task in the document appears in exactly one of the Register or Not Selected sheets;
- header rows on all four sheets have the green `#04A544` fill with black bold font;
- section rows (and only section rows) have a black fill with white bold font;
- every Automation Readiness cell in a task row has the fill that matches its value (High `#04A544`, Medium `#FBB215`, Low `#E03B3B`, black font); no other colours anywhere, no merged cells;
- every task row has a Suggested Claude Feature using only the eight listed names (Folder Access/Uploading Files always written in full), joined by " + " when there is more than one, and chosen with the tests above;
- the How to Use sheet has the features table with all eight features and an example from this P&P for each one used;
- Start Order runs 1 to N with no gaps or repeats, N equals the number of task rows in the Register, and no task ranks before a task it depends on;
- the Not Selected sheet has no reasons;
- frequencies are either as written or "Not stated in P&P";
- every numeral uses Western digits;
- all four sheets are present, sub-rows sit under the right procedure, and the file opens and is editable.

### Step 9: Deliver
Send the file, then write a short closing note (three or four lines, no recap of the steps): how many tasks are in the Register, how many are High readiness, which task is rank 1, how many tasks are in the Not Selected sheet, and one reminder that the user should validate each row against the source document before relying on it. Do not suggest new procedures.

## Edge cases

- **Document is a high-level policy with no operational steps:** tell the user there are no tasks to extract; do not fabricate any.
- **Illegible file (blurry scan, missing pages, password-protected):** say which pages are affected and ask for a readable copy.
- **Very long document:** work through it section by section and keep a running list so nothing is missed.
- **Multiple documents:** one file, with the P&P title on every row and a section row per document; rank all selected tasks together from 1 to N. Ask first if the user may prefer separate files.
- **Circular dependency (two tasks each need the other):** ask the user which comes first; record the answer in the Questions & Assumptions sheet.
- **User changes the language or asks for edits:** update the same file; do not start over.

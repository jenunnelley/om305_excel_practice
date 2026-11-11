# OM 305 Excel Practice

A self-paced practice site for OM 305. Students type real Excel formulas into a spreadsheet grid right in their browser and get instant, specific feedback. Nothing to install, no accounts, and nothing is sent anywhere.

## Put it online with GitHub Pages (about 10 minutes, free)

Same steps as the Python Practice Lab:

1. Go to **github.com** and sign in.
2. Click **+** (top right), then **New repository**. Name it something like `om305-excel-practice`, choose **Public**, and click **Create repository**.
3. Click **uploading an existing file**.
4. Drag in these files: `index.html`, `app.js`, `engine.js`, `problems.js`, `checker.js`, `style.css`, `README.md`.
   (The `dev` folder is only for testing. Don't upload it.)
5. Click **Commit changes**.
6. Go to **Settings**, then **Pages**. Under **Branch**, choose **main** and **/ (root)**, then click **Save**.
7. Wait 1–2 minutes and refresh. Your link appears at the top, like `https://YOUR-USERNAME.github.io/om305-excel-practice/`.
8. Open `blackboard-button.html`, change the link in it to yours, and paste it into a Blackboard item (HTML view).

## What's in it

**218 problems in 7 parts**, each broken into short levels. Every level mixes several business settings: outdoor retail, HR, shipping/3PL, an airline, a hospital, a restaurant, a hotel, a gym, a car dealership, a manufacturing plant, college athletics, and more.

| Part | Topics |
|---|---|
| 1. Counting & Conditional Math | SUM/AVERAGE/COUNT/COUNTA/MAX/MIN, absolute & mixed references, COUNTIFS, SUMIFS, AVERAGEIFS, wildcards, SUMPRODUCT |
| 2. Logical Functions | IF, AND, OR, nested IF, ISBLANK/ISNUMBER/ISTEXT/ISERROR, IFERROR, structured references in Excel Tables |
| 3. Lookup Functions | VLOOKUP exact & approximate, INDEX/MATCH, INDEX/MATCH/MATCH, lookups with IF/IFERROR, VLOOKUP into Tables |
| 4. Dates & Rounding | YEAR, MONTH, DAY, TODAY, ROUND (including years of service) |
| 5. Text Functions | LEFT, RIGHT, LEN, UPPER/LOWER/PROPER, TRIM, CONCATENATE/&, SUBSTITUTE, VALUE, SEARCH, TEXT |
| 6. Mix It Up: Combined Formulas | Nested formulas that cross topics: two functions, three or more, and "boss" formulas |
| 7. Exam-Style Practice | Five multi-step workbooks (HR, shipping/3PL, sales commissions, hospital patient log, restaurant tickets). Every step is required. |

**Every problem has fresh data.** Names and numbers are randomly generated, so "New data" gives unlimited practice, and two students rarely see the same numbers.

## How it works for students

- **Spreadsheet grid:** students type a formula for the highlighted cell and it's copied down (or across) for them, like dragging the fill handle. Sheet tabs, Excel Tables, and formatting look like Excel.
- **Excel-style typing help:** clicking or dragging across cells inserts references, Shift+click extends a range, **F4** cycles $ signs, **Tab** finishes a function name, and a tooltip shows the function's arguments. References in the formula are outlined in color.
- **Enter** calculates (free, any number of times). **Check answer** (Ctrl+Enter) grades it.
- **Answer checking:** the tool compares the student's results to the right answer in every filled cell. It also reruns their formula on two **hidden versions of the data**, so typing in numbers instead of cell references doesn't pass. Any correct formula counts (for example, `A2&" "&B2` and `CONCATENATE(A2," ",B2)` both work).
- **Feedback explains the mistake without giving the formula:** a missing $ ("your first cell works, but C3…"), FALSE vs. TRUE in VLOOKUP, `">=K1"` inside quotes, text that looks like a number, capitalization, a missing `""`, circular references, misspelled functions, unquoted text (#NAME?), and what each Excel error means.
- **Only class functions:** XLOOKUP, TEXTBEFORE, TEXTAFTER, IFS, SWITCH, CONCAT, TEXTJOIN, MAXIFS, and similar get a friendly "not part of OM 305" note. Some problems also require a specific function (for example, INDEX/MATCH problems don't accept VLOOKUP).
- **Testing out:** getting the first two problems in a level right on the first try, without hints, tests out of that level. Missing a problem 3 times or using 2 hints turns testing out off for that level. Locked levels offer a one-shot **jump-ahead** challenge.
- **Hints** point toward the answer. A model formula is shown only after the student gets it right.
- **⚡ Speed Drills:** 8 formulas against the clock in one category (Conditional Math, Logic, Lookups, Dates & Rounding, Text, Combined Formulas, or Mixed Bag). Every drill has fresh data. Skipping costs 15 seconds, and personal bests are saved.
- **My progress** shows a report students can copy and paste (for example, into a Blackboard submission) if you want proof of practice. It includes drill personal bests.

## Settings you can change

Open `app.js` on GitHub, click the pencil icon, and change the values near the top:

| Setting | What it does | Default |
|---|---|---|
| `LOCK_PARTS` | `true` = students must finish (or test out of) a part before the next opens | `false` |
| `JUMP_AHEAD` | Locked levels offer a one-shot jump-ahead challenge | `true` |
| `TEST_OUT_STREAK` | First-try answers in a row needed to test out of a level | `2` |
| `STRUGGLE_FAILS` / `STRUGGLE_HINTS` | Misses / hints on one problem before testing out turns off for that level | `3` / `2` |
| `NO_TEST_OUT_PARTS` | Parts where every problem is required | `[7]` |
| `DRILL_LENGTH` | Formulas per speed drill | `8` |
| `DRILL_SKIP_PENALTY` | Seconds added for skipping a drill formula | `15` |

To change which "extra" functions are accepted (COUNTIF, SUMIF, MID, FIND, ROUNDUP, and others that aren't the focus but aren't wrong), edit the `EXTRAS` list near the top of `engine.js`.

## Good to know

- **The formulas are checked against real Excel.** The calculator was tested against desktop Excel. Excel matched on every cell of 436 problem instances and 179 edge-case formulas (rounding, TEXT formats, wildcards, approximate lookups, and so on).
- **Progress lives in each student's browser.** It doesn't follow them to another computer. You only see progress if they send you their report.
- **Model answers are in the page's code.** A tech-savvy student could find them, so treat this as practice, not a test.
- **What it can't practice:** things that aren't formulas, like charts, PivotTables, Goal Seek/Solver, formatting, data validation, sorting/filtering, and Text to Columns. In real Excel, Ctrl+Shift+Arrow selects to the end of the data. In the practice grid, Shift+click does the same job.
- **TODAY():** problems that use TODAY() are built around the day the student practices, so they always make sense.
- **Browsers:** current Chrome, Edge, Firefox, and Safari on laptops. It works on phones, but typing formulas there is awkward.

## For future edits (optional)

The `dev` folder has a test page (`dev/test.html`, served by `python dev/devserver.py`) that runs every problem's model formula on several data sets, plus a script (`dev/excel_check.ps1` + `dev/compare.py`) that recalculates everything in desktop Excel and compares. Run it after adding problems to `problems.js`.

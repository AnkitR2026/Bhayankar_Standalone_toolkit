# Bhayankar Standalone Toolkit

**Put a Tableau workbook in. Get a Power BI project out.**
Fully offline: no internet, no Node.js, no AI. The only thing you install is **Python 3**. You also need **Power BI Desktop** to open the result.

Version: see `VERSION.txt`.

> [!IMPORTANT]
> **The output is a first draft, not a finished report.** A person still has to:
> - answer the `DECISION NEEDED` items,
> - fix the measures in the `Measures\_Needs fix` folders, and
> - check the numbers against Tableau.
>
> The tools never mark a workbook as signed off.

---

## Contents

1. [Quick start (first workbook in 10 minutes)](#quick-start-first-workbook-in-10-minutes)
2. [Set up (once per PC)](#1-set-up-once-per-pc)
3. [Prepare your inputs](#2-prepare-your-inputs)
4. [Migrate a workbook, step by step](#3-migrate-a-workbook-step-by-step)
5. [Input file reference (`sources.json`, `decisions.json`, `table_map.json`)](#4-input-file-reference)
6. [Verify in Power BI Desktop](#5-verify-in-power-bi-desktop)
7. [Test a Power BI project you already have](#6-test-a-power-bi-project-you-already-have)
8. [What you get (output files)](#7-what-you-get-output-files)
9. [All tools at a glance](#8-all-tools-at-a-glance)
10. [Troubleshooting](#9-troubleshooting)
11. [Known limitations](#10-known-limitations)
12. [Files in this folder, command reference](#11-files-in-this-folder)

---

## Quick start (first workbook in 10 minutes)

If the toolkit is already set up on this PC (see [section 1](#1-set-up-once-per-pc)):

| # | Do this | You should see |
|---|---|---|
| 1 | Copy `My Report.twbx` into `input\`. | |
| 2 | Double-click **`Run_Extractor.bat`**. | `1 workbook(s) → ...\output` |
| 3 | Open `output\My Report\bhayankar_digest.md` and read the **Flags** table. | A list of what needs a person. |
| 4 | Double-click **`Run_Build.bat`**. | `validate: succeeded (0 errors, 0 warnings)` |
| 5 | Open `output\My Report\pbip\My Report.pbip` in Power BI Desktop and click **Refresh**. | Pages with visuals and data. |
| 6 | Run `Run_Verify.bat "My Report"` from a command prompt in the toolkit folder. | `output\My Report\verify\VERIFY_REPORT.md` |

That's all you need for a first look. For a real migration, follow [section 3](#3-migrate-a-workbook-step-by-step).

> [!TIP]
> **Prefer a menu?** Double-click **`Run_Bhayankar.bat`**. It offers every tool by number (extract, build, workflow, verify, test, validate, self-check) and asks for the workbook name when one is needed.

> [!TIP]
> **Running a `.bat` with arguments:** open the toolkit folder in File Explorer, type `cmd` in the address bar and press Enter. Then type, for example, `Run_Verify.bat "My Report"`.
> A `.bat` that needs no arguments can just be double-clicked. Every window stays open at the end so you can read the result.

---

## 1. Set up (once per PC)

| # | Step | How to check it worked |
|---|---|---|
| 1 | **Copy the whole toolkit folder** to a short local path such as `C:\Work\Bhayankar`. Avoid OneDrive and other synced folders, because they can lock files during a build. Leave out `input\` and `output\` unless the receiver may see that client's data. | The folder contains `Run_SelfCheck.bat`, `resources\` and `tests\`. |
| 2 | **Install Python 3** from python.org. In the installer, tick ***Add python.exe to PATH***. | In a command prompt, `python --version` shows `Python 3.x` (tested on 3.14). |
| 3 | **Install Power BI Desktop** (tested with 2.156 and later). | It opens. |
| 4 | Double-click **`Run_SelfCheck.bat`**. | The last line reads `RESULT: SELF-CHECK PASSED`. |

If step 4 fails, don't use the toolkit on that PC yet. See [Troubleshooting](#9-troubleshooting).

**Optional:** `pip install tableauhyperapi` lets the extractor read the data packaged inside a `.twbx` (the `--extract-data` option, see [step 3.3](#33-extract)).

---

## 2. Prepare your inputs

### 2.1 What to collect from the Tableau side

| What | Required? | Where to get it | Why it matters |
|---|---|---|---|
| The workbook `.twb` or `.twbx` | **Yes** | Tableau Server/Cloud → *Download → Tableau Workbook*, or the author | Everything starts from this. |
| The published data source `.tds` / `.tdsx` | **When the workbook uses a published source.** The digest flag `PUBLISHED_DS` tells you. | Tableau Server → *Data Sources →* the source → *Download* | Joins, custom SQL and database table names live only in the `.tds`. Without it, joins are guessed. |
| Sample data as CSV, **or** the warehouse server / database / schema | Strongly recommended | A data extract from the source tables, or the DBA | Without data, the model loads nothing and nothing can be compared. |
| A screenshot of every dashboard, with its totals visible | Strongly recommended | Tableau, with the default filters applied | Your reference for the final number check. |
| Dashboard images and logos | Only if the dashboards show images | The workbook author, or inside the `.twbx` | Missing images are left out and the build log names them. |

### 2.2 How to lay out `input\`

```
input\
├── My Report.twbx                  ← the workbook (the file name becomes the Power BI project name)
├── MY PUBLISHED SOURCE.tds         ← published data source(s), if any
├── My Report_Data\                 ← sample CSVs, one per source table (optional)
│   ├── WORK_ORDERS.csv
│   └── TASKS.csv
├── Image\                          ← dashboard images, same names as in the workbook (.twb only)
│   └── logo.png
├── sources.json                    ← optional: where the data loads from
├── decisions.json                  ← optional: your answers to DECISION NEEDED items
└── table_map.json                  ← optional: Power BI table names
```

### 2.3 Naming rules (most first-run problems come from these)

| Item | Rule | Example |
|---|---|---|
| Workbook | Its file name, without the extension, is the **workbook name**. You'll use it in `decisions.json`, in `Run_Verify.bat "<name>"` and for the output folder. | `My Report.twbx` → `My Report` |
| `.tds` | Name it after the **published data source as Tableau calls it**. The extractor matches by that name. | Workbook uses source `WO Lifecycle` → `WO Lifecycle.tds` |
| Sample-data folder | `<workbook name>_Data`. Spaces may also become underscores. If the folder has this name, you don't need a `sources.json` for sample data. | `My Report_Data\` or `My_Report_Data\` |
| Sample CSV | `<SOURCE_TABLE>.csv`, named after the database table as the digest lists it. If the name is different, map it with `"file"` in `sources.json`. | `WORK_ORDERS.csv` |
| CSV content | UTF-8, comma-separated, **first row = column headers** using the **database column names**. Empty cells load as blank (null). | `WO_ID,STATUS,CREATED_DT` |
| Images | Same relative path and name as the workbook refers to them. A `.twbx` already contains its images. | `Image\logo.png` |

> [!NOTE]
> - Several workbooks can sit in `input\` together. The tools process each one.
> - **`sources.json` is shared by everything in `input\`.** If two workbooks need different data locations, migrate them one at a time, or give each one a `<workbook>_Data` folder and leave out `sources.json`.
> - Keep client data out of email and version control. `.gitignore` already excludes `input\` and `output\`.

---

## 3. Migrate a workbook, step by step

### 3.1 Collect
Gather everything in [2.1](#21-what-to-collect-from-the-tableau-side).

### 3.2 Place
Copy everything into `input\` as shown in [2.2](#22-how-to-lay-out-input).

### 3.3 Extract
Double-click **`Run_Extractor.bat`**. It writes `output\<workbook>\`, plus `output\inventory.csv` with one row per workbook.

Then open **`output\<workbook>\bhayankar_digest.md`** and read it in this order:
1. **Flags**: what the workbook uses that needs care (for example `PUBLISHED_DS`, `RLS`, `NAME_COLLISION`).
2. Every **`DECISION NEEDED`**.
3. Every item marked **`[low]`** (a low-confidence translation).

Large workbooks also get one file per dashboard: `bhayankar_digest_NN_<dashboard>.md`.

Options (from a command prompt):
- `Run_Extractor.bat "input\My Report.twbx" --show "Sales by Region"` prints one sheet, dashboard, calc, parameter or action.
- `Run_Extractor.bat "input\My Report.twbx" --extract-data` writes the `.twbx`'s own data to `output\My Report\extract_data\*.csv`. This needs `tableauhyperapi`. Copy the CSVs into `input\My Report_Data\`, or point `sources.json` at them.

### 3.4 Decide
Write your answers into **`input\decisions.json`** and say where the data is in **`input\sources.json`**. Both are explained with examples in [section 4](#4-input-file-reference). If you have nothing to decide yet, skip this step and come back to it.

### 3.5 Build
1. **Close the project in Power BI Desktop** if it is open.
2. Double-click **`Run_Build.bat`**. If the inputs changed, it extracts again by itself.
3. You should see `validate: succeeded (0 errors, 0 warnings); model check: 0 errors`. The validator checks the report files; the **model check** checks the semantic model for what makes Power BI Desktop refuse it (names that differ only in case, references to objects that do not exist, functions DAX does not have, circular measures, broken relationships). A model-check error makes the workflow verdict `BUILD FAILED`.
4. Work through **`output\<workbook>\bhayankar_build_log.md`**. It lists every measure that went to `_Needs fix`, every guessed join and every missing image. Lines starting with `DECISION NEEDED` are the ones a person must confirm.
5. Open **`output\<workbook>\DATA_REQUEST.md`**. It lists, for every table, the exact CSV file name and column headers the model reads (and the warehouse table). If you have a sample folder, it is checked: missing files or headers are flagged. Hand the header-only templates in `output\<workbook>\data_request\` to whoever extracts the data.

### 3.6 Open
Open `output\<workbook>\pbip\<workbook>.pbip` in Power BI Desktop and click **Refresh**.

### 3.7 Verify
Run `Run_Verify.bat "<workbook>"` ([section 5](#5-verify-in-power-bi-desktop)). Fix what it reports, then repeat **Build → Open → Verify** until it is clean.

### 3.8 Sign off (a person)
Compare the numbers with the screenshots and follow `docs\QA_CHECKLIST.md`. For anything to do with security, use `docs\SECURITY_RLS.md`.

> [!WARNING]
> **A rebuild overwrites hand edits made in Desktop.** Once you start finishing the report by hand, copy the `pbip` folder somewhere else and work there. You can also set `"pbip_folder"` in `decisions.json` so builds write to a different folder.

**Shortcut:** `Run_Workflow.bat` runs extract → gate → build → validate → test for every new or changed workbook. It stops where a person has to decide and writes `output\<workbook>\WORKFLOW_STATUS.md`. `Run_Watch.bat` does the same whenever a workbook lands in `input\`.

---

## 4. Input file reference

All three files are optional and live in `input\`. They are plain JSON. Edit them in Notepad or VS Code.

> [!TIP]
> **JSON rules that trip people up:**
> - Use double quotes `"…"` only.
> - No comma after the last item.
> - Write backslashes in paths twice: `"C:\\Data\\Extracts"`.
> - Use `true` / `false` without quotes.
>
> If a tool reports a JSON error, paste the file into any JSON checker.

### 4.1 `sources.json`: where the data loads from

**Option A: sample CSVs.** The simplest form is to have no `sources.json` at all and put the CSVs in `input\<workbook>_Data\`. Use the file only when you need more control:

```json
{
  "mode": "sample",
  "sample_folder": "My Report_Data",
  "tables": {
    "ORDERS": {
      "file": "orders_extract.csv",
      "where": { "ACTIVE": "Y" },
      "columns": { "FULL_NAME": "Name" },
      "text_replace": { "REGION": ["N. America", "North America"] }
    }
  }
}
```

| Key | Meaning |
|---|---|
| `mode` | `"sample"` = load CSVs. |
| `sample_folder` | Folder with the CSVs. Relative paths start at `input\`, so the setup works on any PC. A full path also works. |
| `tables` | Optional. One entry per **source table** (as named in the digest), and only for tables that need one. Every key inside is optional. |
| `file` | CSV file name, if it isn't `<SOURCE_TABLE>.csv`. Several tables may read the same CSV. |
| `where` | Keep only rows where a column equals a value. |
| `columns` | `{ "model column": "CSV header" }`, for when the CSV header is different from the column name. |
| `text_replace` | `{ "column": ["find", "replace with"] }`, a text fix applied while loading. |

Without a `.tds`, tables that read the same CSV are merged. Joins are inferred from CSV headers written as `X (TABLE)`. Both are logged as `DECISION NEEDED`, so check them.

**Option B: the warehouse.**

```json
{
  "mode": "warehouse",
  "warehouse": { "server": "myserver.database.windows.net", "database": "EDW", "schema": "dbo" }
}
```

Both options become Power Query parameters in the project, so you can switch later in Desktop (*Transform data → Edit parameters*) without rebuilding:

| Parameter | Meaning |
|---|---|
| `UseSampleData` | `true` loads the CSVs; `false` loads from the warehouse. |
| `SampleDataFolder` | Folder with the CSVs. |
| `WarehouseServer`, `WarehouseDatabase`, `WarehouseSchema` | The warehouse location. |

### 4.2 `decisions.json`: your answers, one block per workbook

The top-level keys are **workbook names**, exactly as in the file name without the extension.

```json
{
  "My Report": {
    "measure_names": { "Sales Amt": "Total Sales" },
    "slicer_defaults": {
      "Region": { "include": ["West", "East"] },
      "Status": { "exclude": ["Cancelled"] }
    },
    "preview_login": "jdoe@company.com",
    "pbip_folder": "C:\\PBI\\My Report",
    "workflow": {
      "rls_approved_by": "Jane Architect",
      "accept_guessed_joins": true,
      "tier_confirmed": "Medium"
    }
  }
}
```

| Key | Use it when | Value |
|---|---|---|
| `measure_names` | You want different measure names. | `{ "name the build gave": "new name" }` |
| `slicer_defaults` | A slicer must start with a selection (for example, the Tableau filter's default). | Keyed by the filter label or field name: `{ "include": [...] }` or `{ "exclude": [...] }` |
| `preview_login` | The workbook uses `USERNAME()`. The preview is built for this one login. By default, the first hard-coded login in the workbook is used. | A login |
| `pbip_folder` | The project should be written somewhere other than `output\<workbook>\pbip\`. | A folder path |
| `workflow` | Used only by `Run_Workflow.bat` / `Run_Watch.bat`. **Only a person writes these.** | See below |

`workflow` keys:
- `rls_approved_by`: name of the person who approved the security design. Required when the digest flags `RLS`, `PER_USER_COLUMN` or `HARDCODED_USERS`.
- `accept_guessed_joins`: `true` to build without the `.tds` when `PUBLISHED_DS` is flagged.
- `tier_confirmed`: the complexity tier you agreed on, kept as a record.

### 4.3 `table_map.json`: Power BI table names

```json
{ "WORK_ORDERS": "Work Order", "TASKS": "Task" }
```

The keys are source table names; the values are the table names you want in Power BI. The extractor applies the map, so **create it before you extract**. If you add it later, run `Run_Extractor.bat` (or `Run_Build.bat`) again. Leave this file out to keep the source names.

---

## 5. Verify in Power BI Desktop

`Run_Build.bat` validates the report **files**, but it can't see DAX or the data. `Run_Verify.bat` asks Desktop's own local engine. It uses the query library that ships with Desktop through Windows PowerShell, so there is nothing to install.

```
Run_Verify.bat "My Report" --open --refresh
```

| Option | What it does |
|---|---|
| *(none)* | Checks the project that is already open in Desktop. |
| `--open` | Opens the project in Desktop first and waits for it to load. |
| `--refresh` | Loads the data into the open session, like the Refresh button. |
| `--no-numbers` | Model checks only, with no comparison to Tableau. |
| `--list` | Shows which models are open, with their engine ports. |
| `--port N` | Uses a specific engine from `--list`. |

| Check | What it finds |
|---|---|
| Model health | Any measure, column or partition with an error. |
| Data loaded | Rows per table. If most tables are empty, refresh first. |
| Measures | Whether every measure evaluates, and which ones are slow. |
| Numbers | For each Tableau sheet, the Tableau logic is recomputed from the source data and compared with the engine. |

Results go to `output\<workbook>\verify\`: **`VERIFY_REPORT.md`** (read this), `test_report.xlsx`, and the engine's rows in `engine_exports\`. The verifier never changes the model or the project files.

**Limit:** a sheet is compared only when the query can apply all of its Tableau filters. Sheets whose filters became slicers or visual filters are listed under *"Not compared automatically"*. Check those by hand with the same selections. (On the Facility Metrics workbook, that is every sheet.)

---

## 6. Test a Power BI project you already have

`Run_PowerBI_Report_Test.bat` compares **any** Power BI project, including a hand-built one, with the Tableau workbook it replaces. It uses its own folders:

```
test_input\
└── My Report\                          ← one folder per report
    ├── My Report.twbx                  ← the Tableau workbook (same name as the project)
    ├── My Report.pbip                  ← the Power BI project …
    ├── My Report.SemanticModel\        ← … with its two folders
    ├── My Report.Report\
    ├── *.tds                           ← optional: published data sources
    ├── data\                           ← optional: source tables as <TABLE>.csv/.xlsx (enables data tests)
    ├── tableau_exports\<Sheet>.csv     ← optional: Tableau crosstab exports per sheet
    ├── pbi_exports\<Visual>.csv        ← optional: Power BI "Export data" per visual
    └── test_map.json                   ← optional: corrections to the automatic matching
```

On the first run it creates `test_input\` for you. The result is `test_output\My Report\test_report.xlsx` (Summary, Issues, Unit, Functional, Data, Mapping and more). A suggested `test_map_suggested.json` is written next to it. Copy it to `test_input\My Report\test_map.json` and correct it to fix any wrong matches.

To validate just the report files of any project offline: drag a `.pbip` onto **`Run_Validate.bat`**.

---

## 7. What you get (output files)

Everything is in `output\<workbook>\`:

| File | What it is | When to use it |
|---|---|---|
| `bhayankar_digest.md` | The workbook explained: logic, model, rebuild plan per page | **Read first.** |
| `bhayankar_digest_NN_<dashboard>.md` | Per-dashboard parts of the digest (large workbooks) | Instead of one huge digest |
| `bhayankar_measures.dax` | Draft measures | Paste into Desktop's DAX query view |
| `bhayankar_build_log.md` | What the build did and what it could not do | **After every build** |
| `DATA_REQUEST.md` + `data_request\*.csv` | Which CSV and columns each table needs (checked against your sample folder), plus header-only templates | Before you collect or refresh data |
| `pbip\<workbook>.pbip` | The Power BI project | Open in Desktop |
| `verify\VERIFY_REPORT.md` | Results from Desktop's engine | After `Run_Verify.bat` |
| `WORKFLOW_STATUS.md` | Where the workflow stopped and why, plus the **structural readiness %** (share of the tester's structure checks that pass) | After `Run_Workflow.bat` |
| `bhayankar_rebuild.json`, `bhayankar_digest.json` | Machine-readable spec | For tools, not people |

`output\inventory.csv` has one row per workbook, for planning a batch.

**Confidence labels** in the digest and the measure descriptions:

| Label | Meaning |
|---|---|
| `high` | Direct translation. |
| `medium` | Valid, but read the note. |
| `low` | A person has to fix it. |

**Long paths:** if a path would go over Windows' 260-character limit, the `.Report` / `.SemanticModel` folders get shorter names (the `.pbip` keeps the workbook name). If that is still too long, the project goes to `C:\PBI\<initials>`, and the build says so.

---

## 8. All tools at a glance

Each tool works on its own.

| Launcher | Use it to | Writes |
|---|---|---|
| `Run_Bhayankar.bat` | A menu for every tool below | (as the chosen tool) |
| `Run_SelfCheck.bat` | Test the toolkit itself (on a new PC, or after changing a `.py` file) | Screen |
| `Run_Extractor.bat` | Understand a workbook: digest, draft DAX, rebuild plan | `output\<workbook>\` |
| `Run_Build.bat` | Build the Power BI project and validate it offline (extracts first when needed) | `output\<workbook>\pbip\` |
| `Run_Verify.bat "<workbook>"` | Check the built project inside Power BI Desktop | `output\<workbook>\verify\` |
| `Run_Validate.bat "<file>.pbip"` | Validate any project's report files and semantic model offline | Screen |
| `Run_PowerBI_Report_Test.bat` | Compare any Power BI project with its Tableau workbook | `test_output\<Name>\` |
| `Run_Workflow.bat` | Extract, gate, build and test in one go. Options: `--no-build`, `--no-test`, `--force`, or a workbook name | `output\<workbook>\WORKFLOW_STATUS.md` |
| `Run_Watch.bat` | Same as Workflow, whenever a workbook lands in `input\` (Ctrl+C stops it) | Same |

**Moving the folders:** set the environment variables `BHAYANKAR_INPUT`, `BHAYANKAR_OUTPUT`, `BHAYANKAR_TEST_INPUT` and `BHAYANKAR_TEST_OUTPUT`.

---

## 9. Troubleshooting

| Problem | What to do |
|---|---|
| `'python' is not recognized` | Reinstall Python with *Add python.exe to PATH*, or use `py -3`. The `.bat` files try both. |
| `Run_SelfCheck.bat` fails | Check `python --version`, and check that the whole folder was copied (`tests\fixtures\`, `resources\`). |
| `No .twb/.twbx found` | The workbook isn't in `input\`, or `BHAYANKAR_INPUT` points somewhere else. |
| JSON error when reading `sources.json` / `decisions.json` | See the JSON rules in [section 4](#4-input-file-reference): no trailing commas, doubled backslashes. |
| My `decisions.json` changes don't apply | The top-level key must be the exact workbook name (file name without the extension). |
| Tables are empty after Refresh | Check that the `SampleDataFolder` parameter points to a folder that exists, and that each CSV is named `<SOURCE_TABLE>.csv` (or mapped with `"file"`). |
| Desktop: "problem with the definition content" | *Copy details* lists the DAX errors. Each one is a generator gap, so report it. `Run_Verify.bat` finds the same errors without the dialog. |
| Verify says `NOT VERIFIED` | The project isn't open in Desktop. Open it (or use `--open`), wait for it to load, and run again. |
| Verify says `NOT REFRESHED` | Click Refresh in Desktop, or use `--refresh`. |
| Verify: PowerShell bridge failed | Try `Run_Verify.bat --list`. If Desktop is installed in a non-standard folder, set `BHAYANKAR_ADOMD` to the full path of `Microsoft.PowerBI.AdomdClient.dll`. |
| Validate: `PBIR_SCHEMA_UNREACHABLE` | The project uses a newer schema version than the bundled ones. The other checks still ran. Add the schema under `resources\schemas\`. |
| Validate: `BHY_PATH_TOO_LONG` | Move the project to a short folder such as `C:\PBI\`. |
| "Can't determine relationships" | The fields come from unrelated tables (often Tableau blending). Add the relationship in Desktop. |
| Measures are blank | Look in `Measures\_Needs fix` and in the build log. |
| A logo is missing | Copy the dashboard images to `input\Image\`, using the names the workbook uses. |
| Permission error under OneDrive | Close the project in Desktop, let OneDrive finish syncing, and run again. Better: move the toolkit to a local folder. |

---

## 10. Known limitations

- **Published data source without its `.tds`:** joins are *guessed* from Tableau's renamed join columns (`WO Number (Custom SQL Query) #2` next to `WO Number`) and marked `DECISION NEEDED`. Tables that share only a name column (for example employee role tables) are left unrelated, with an orphan warning. Where a visual counts a "one side" table by a "many side" column (Tableau's joined rows), that relationship filters both directions, also marked `DECISION NEEDED`. Always confirm with the `.tds`.
- **Tableau relationships** (logical layer) become Power BI relationships: keys by field name, compound keys (one Power Query key column per table), keys on calculated columns. One side from Tableau cardinality "One", else sample data, else a key named after its table; otherwise many-to-many (Tableau's default; never drops rows) marked `DECISION NEEDED`. Inequality relationships are not built (`DECISION NEEDED`).
- **Blended data sources** (Tableau data blending): each active blend link (automatic same-name link, or a custom one from Data > Edit Blend Relationships) becomes a shared dimension table (Power Query: distinct keys of both sources) related to each blended table; a date link goes through the Date table. Visuals on a linked field use the shared dimension, a secondary-source dimension becomes an `<field> (ATTR)` measure (`*` when several values, as in Tableau), and blended visuals get a filter "`# <primary> Rows` is not blank" (Tableau shows only the primary's marks). **Not built:** links on calculated fields (one `DECISION NEEDED` line lists them) and links between tables that are already related. The dimension re-reads the linked tables on refresh. Compare totals per linked field with Tableau.
- **Calcs with the same name in several data sources** (different formulas) become one measure per data source: the first keeps the name, the others are `<calc> (<data source>)`, and each sheet uses the copy of its own data source. Check the names and the numbers.
- **Per-user logic** (`USERNAME()`) needs RLS. The build gives a preview for one login.
- **Table calculations** are written over the columns each sheet computes them along (Tableau "compute using": Table across / down, specific dimensions, or each cell): `RANK` / `RANK_DENSE` (RANKX), `WINDOW_SUM` / `AVG` / `MIN` / `MAX` / `COUNT` / `MEDIAN` without offsets, `TOTAL`, `SIZE`, `INDEX` (ROWNUMBER), `RUNNING_*` and `LOOKUP(x, n)` along one column, plus the quick table calcs `PERCENT_OF_TOTAL` and `RUNNING_TOTAL`. A sheet that computes the same calc along other columns gets a copy `<calc> (along <columns>)`. Still `_Needs fix`: `WINDOW_*` with offsets, `RUNNING_*` / `LOOKUP` along several columns, pane-relative "compute using", `PREVIOUS_VALUE`, `FIRST` / `LAST`, `RANK_PERCENTILE`. Ties: `RANK_UNIQUE` / `RANK_MODIFIED` share a rank in RANKX.
- **Sheet filters on parameter-driven calcs** (including a parameter that picks which true/false filter applies, e.g. `CASE [Period] WHEN "Last Week" THEN [Last Week] ...`) are folded into the visual's measures. A visual with no measure (a table of names) gets a hidden `<sheet> Rows` measure and a visual filter "is not blank".
- **Manual sorts** become a sort-by column: one order per column, so if two sheets order the same field differently, the first one wins (logged).
- **Filters that mix members and ranges** (`"APPR" UNION RANGE("CLOSE".."WAPPR")`) stay a `DECISION NEEDED`: list the real values.
- **Source connectors:** without a warehouse in `sources.json`, each table reads its Tableau source with the matching Power Query connector (Snowflake, SQL Server, Excel, text/CSV, Oracle, PostgreSQL, MySQL, Teradata, Redshift); credentials are asked on the first refresh. With `"warehouse": {"server": ...}` (or `"connector": "warehouse"`) every table reads the warehouse as before. Published sources without their `.tds`, and other types, read the warehouse (`DECISION NEEDED` for other types).
- **Data source filters** become Power Query row filters on their table; when that table filters others (one side of a relationship), or the field is a calc / relative date, the filter also goes on every page. Aggregate data source filters are `DECISION NEEDED`.
- **Value aliases** become a display column `<column> (alias)` used on axes, legends and rows; filters, slicers and DAX keep the source values.
- **Not built yet:** hierarchies, display folders, sorts by a measure on tables, numeric bins, show/hide zone bookmarks, week-start and fiscal-year settings. They are listed in the digest.
- **Filters** the extractor can't read stay as a note for a person.
- **Free-text parameters** are fixed to their default value.
- **Verify** needs Windows, Power BI Desktop and Windows PowerShell. The rest of the toolkit runs anywhere Python does.

---

## 11. Files in this folder

| File | What it does |
|---|---|
| `bhayankar_extractor.py` | Reads the workbook: digest, DAX drafts, rebuild spec |
| `bhayankar_dax.py` | Translates Tableau formulas to DAX |
| `bhayankar_report.py` | Rebuild spec (model, pages, visuals) |
| `bhayankar_build.py` | Writes the TMDL model and the PBIR report, then validates |
| `bhayankar_scaffold.py` | Writes the project shell (replaces the Node.js scaffold) |
| `bhayankar_validate.py` | Offline PBIR validator (replaces the Node.js validate) |
| `bhayankar_modelcheck.py` | Offline semantic-model check (names, references, DAX functions, circular measures, relationships) |
| `bhayankar_verify.py` | Checks a project in Desktop's engine |
| `bhayankar_test.py` + `_oracle`, `_pbip`, `_xlsx` | The report tester and its modules |
| `bhayankar_run.py`, `bhayankar_watch.py` | Workflow runner, folder watcher |
| `bhayankar_toolkit.py` | Name, version and folders shared by all tools |
| `bhayankar_visual_catalog.json` | Visual types and roles (Power BI CLI 0.4.0 catalog) |
| `resources\` | Bundled PBIR schemas, base theme, visual metadata snapshot |
| `tools\snapshot_pbir_metadata.mjs` | Maintenance only (needs Node.js): refreshes `resources\pbir_metadata.json` after a Power BI update |
| `tests\` | Regression tests and synthetic workbooks |
| `docs\` | Method, project log, QA checklist, security notes |
| `NOTICE.md` | Third-party material and licences |

### Command reference (for scripts and advanced use)

```
python bhayankar_extractor.py ["<workbook>" [--tds "<file>"] [--table-map "<file>"] [--show "<item>"] [--extract-data] [--split|--no-split]]
python bhayankar_build.py ["<workbook name>"]
python bhayankar_verify.py "<workbook name>" [--open] [--refresh] [--no-numbers] [--list] [--port N]
python bhayankar_validate.py "<file>.pbip" [--format text] [--no-schema]
python bhayankar_test.py ["<workbook>" "<project.pbip>" [--data <folder>] [--tds <file>] [--tableau-exports <folder>]
                         [--pbi-exports <folder>] [--map <test_map.json>] [--today YYYY-MM-DD] [--username <login>]]
python bhayankar_run.py scan | run ["<workbook name>"] [--no-build] [--no-test] [--force] | status "<workbook name>"
python -m unittest discover -s tests
```

### How this edition differs from the Kiro edition

This is the standalone edition of Bhayankar. The Kiro edition (`Bhayankar-Toolkit-Repo`) wraps the same engine in an AI workflow.

| | Kiro edition | Standalone edition |
|---|---|---|
| Project shell (`.pbip`, `.platform`, theme) | `powerbi-report-author scaffold` (Node.js) | `bhayankar_scaffold.py`, same files |
| Report validation | `powerbi-report-author validate` (Node.js, downloads schemas) | `bhayankar_validate.py`: same rules, schemas bundled in `resources\`, 28 test cases agree with the CLI |
| Check inside Power BI Desktop | By hand, or with the modeling MCP | `Run_Verify.bat` |
| Rebuilds | Random ids | Stable ids: the same workbook gives the same files |
| Needs | Python, Node.js, npm CLI, Kiro | **Python 3** (plus Power BI Desktop to open the result) |

`BHAYANKAR_CLI_CHECK=1` also runs the Microsoft CLI as a second opinion, if it happens to be installed. It is never required.

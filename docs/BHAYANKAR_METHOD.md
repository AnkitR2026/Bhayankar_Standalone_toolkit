# Bhayankar_Project: method

Bhayankar_Project is a separate test proof of concept. It reads the Kiro migration files (the TWB extraction spec, the steering files) but never changes them.

## 1. Rule

Never open a raw `.twb` in an AI context. Run the Bhayankar extractor and reason only over its digest.

The extractor does the following mechanically:
- builds an inventory of the workbook's parts;
- applies the used-only rule, so unused calcs, parameters and columns are dropped;
- resolves Tableau's internal names to readable captions;
- removes duplicates;
- detects unused (orphan) items;
- drafts DAX for every calc in use;
- writes a Power BI rebuild plan (model, pages and visuals), checked against the `powerbi-report-author` catalog.

It needs Python 3 and the standard library only. The agent spends its tokens only on review and decisions.

**Measured on 2026-10-19:**
- Employee Summary: 1,159,192 bytes became a 24 KB digest including the full rebuild plan (about 6k tokens, 98% smaller).
- Retail Performance Dashboard (60 sheets, 14 dashboards, 130 used calcs): 2.48 MB became one common file plus one file per dashboard.
- A 460 MB `.twbx` was processed in 2.6 seconds.
- Across the 15-workbook test set: 299 visuals mapped, 7 catalog-check failures left (all of them flagged decisions), and no crashes.

## 2. Inputs to collect before running

| # | Input | Why |
|---|---|---|
| 1 | The `.twb` / `.twbx` | The source of the logic. |
| 2 | The published data source `.tds` / `.tdsx` for every `PUBLISHED_DS` flag. Download it from Tableau Server (Data Sources → Download), or use REST / `tabcmd get`. | A workbook on a published source contains no joins, custom SQL or database type. Without the `.tds` the agent has to guess. In Employee Summary that led to a wrong Teradata assumption, truncated SQL and missing EMPLOYEE columns. |
| 3 | Screenshots of every dashboard with default filters, plus the totals shown | Layout, labels and colours, and the parity targets for QA. |
| 4 | Target platform and storage mode (for example, Fabric warehouse with Import or Direct Lake), plus warehouse and table names | Decides the Power Query source pattern. |
| 5 | Which naming and DAX rules apply (Fabric skills or Kiro steering) | Avoids renaming the model after the build. |
| 6 | Sample data whose keys link, or permission to generate synthetic data | Lets you verify in Desktop before the real source is connected. |
| 7 | PBIP target folder | Where the build goes. |

## 3. Run

```bat
rem from Desktop\Bhayankar_Files: put workbooks (+ their published .tds/.tdsx, optional table_map.json) in input\
python bhayankar_extractor.py
rem   scans input\ and writes output\<workbook>\ plus output\inventory.csv
Run_Extractor.bat
rem   same, by double-click
python bhayankar_extractor.py "<workbook>.twb|.twbx" --tds "<published>.tds"
rem   one workbook -> output\<name>
python bhayankar_extractor.py "<workbook>" --show "<sheet|dashboard|calc|action name>"
rem   raw XML of one item, pruned
python bhayankar_extractor.py --refresh-catalog
rem   re-pull visual roles after a CLI update
python -m unittest discover -s tests -v
rem   regression tests: run after any change
```

**How `.tds` files are matched in folder mode:** a `.tds` is attached only to the workbook whose published data source has the same name (friendly name or server name, ignoring punctuation). Files that match nothing are reported, never guessed.

**Table names:** if `input\table_map.json` exists (`{"SOURCE_TABLE": "Power BI table"}`), it's applied automatically, so the DAX uses the model's table names.

## 4. Outputs

Each workbook gets its own folder, `output\<workbook>\`. `output\inventory.csv` sits one level up and has one row per workbook.

| File | Use |
|---|---|
| `bhayankar_digest.md` | The agent's only input. Contains: counts, suggested tier, flags, data sources (connection, joins, relationships, custom SQL), used source columns with their table and field, parameters, measure header aliases, common filters (F1, F2…), dashboards (zone tree with % geometry), tabs, actions and orphans. Also contains the Power BI model and the rebuild plan for each page. |
| `bhayankar_digest_NN_<dashboard>.md` | Written when the digest exceeds 40 KB (override with `--split` / `--no-split`). One self-contained file per dashboard. To build one page, read `bhayankar_digest.md` plus that page's file. |
| `bhayankar_digest.json` | The extracted Tableau logic, structured. Each calc carries `dax`, `pbi_kind`, `confidence` and `dax_notes`. |
| `bhayankar_rebuild.json` | The Power BI rebuild spec, for generators and the planning step. **Model:** tables, columns to hide, calculated columns, measures (home table, formatString, displayFolder, description), relationships, date table, parameter tables and field-parameter candidates. **Pages:** size; visuals with catalog-checked type and roles, pixel position, title, sort, labels, colours, reference lines and visual filters; slicers with Tableau defaults and sync groups; page filters; drillthrough fields; tooltip pages; interactions; images, text and buttons. |
| `bhayankar_measures.dax` | The draft measures as one DAX query view script (`DEFINE MEASURE 'Table'[Name] = …` followed by `EVALUATE ROW(…)`). Run it in Desktop to test, then choose "update model". |
| `inventory.csv` | Per workbook: sizes, counts, low-confidence calcs, decisions needed, merged `.tds` files, suggested tier and flag codes. |

**Confidence levels.** Every measure, calculated column and visual is marked `high`, `medium` or `low`, with notes:
- `high`: a direct translation.
- `medium`: valid, but check the note. Examples: LOD filter scope, an assumed iterator, a split visual.
- `low`: needs a person. Examples: table calcs, mixed grain, an unmapped visual, a catalog-check failure.

**`DECISION NEEDED`.** Anything starting with this must be answered before the build. For example, a Tableau filter that mixes listed values with an alphabetical range is never collapsed to a guess.

## 5. Checked against the Fabric skills (2026-10-19)

The rebuild plan follows the rules of the skills the build uses, so generated work doesn't have to be redone.

| Skill rule | How the extractor applies it |
|---|---|
| powerbi-report-cli: take visual types and roles only from the `powerbi-report-author` catalog, never from memory | `bhayankar_visual_catalog.json` is a snapshot of `catalog list` / `catalog describe` (57 types, CLI 0.4.0). Every visual is checked for role names, Grouping/Measure kinds, required roles and maximum fields. Failures print as `CATALOG CHECK FAILED` and set confidence to low. |
| Use modern types only (`cardVisual`, `tableEx`, `pivotTable`, `azureMap`); never `card`, `multiRowCard`, `table`, `matrix`, `map` or `filledMap` | Only modern types are emitted; legacy types fail the check. |
| A `tableEx` with both dimensions and measures can render no rows | Text tables with measures become `pivotTable` (stepped layout off, subtotals off). Dimension-only tables stay `tableEx`. |
| Slicer types: Dropdown, Between and Relative use `slicer`; lists use `listSlicer`; tiles use `advancedSlicerVisual`; search uses `textSlicer` | Tableau quick-filter modes map to these. A Start/End date-parameter pair becomes one Between slicer. |
| Drillthrough is `pageBinding` plus drillthrough filters; tooltip pages have page type Tooltip | Filter actions that target another dashboard become drillthrough, using the source sheets' dimensions. Help sheets and viz-in-tooltip become tooltip pages. |
| Per-series colours use `dataPoint.fill` with literal hex, not a theme index | Tableau palettes (data-source and sheet level) are passed through as value → hex maps. |
| semantic-model-authoring: use explicit measures and hide the aggregated base column | Every Tableau shelf aggregation (such as `SUM(X)`) becomes an explicit measure, and its base column is marked hidden. |
| Every visible measure has a `formatString` | Converted from Tableau formats, or set by data type. Text measures get none. |
| Measures live on their fact tables, not in one measures table | Each measure's home table is the table of the columns it aggregates. |
| DAX: `DIVIDE`, `VAR _x`, `REMOVEFILTERS` rather than `ALL`, `'Table'[Column]` and `[Measure]`, never `1 - x/y` | Written that way. `x/y ± 1` is rewritten as `VAR _num`, `VAR _den` and one `DIVIDE`. |
| Naming: business names, `#` prefix for counts, no abbreviations, no Fact/Dim | Names are suggested (`# …`), abbreviations are listed for the glossary, and table names are suggested. Real names come from `table_map.json`. |
| Date table built in Power Query, single-direction relationships, no composite keys | A date table is requested whenever date parts are used. Relationships are OneDirection. Composite keys are flagged. |
| Use field parameters, not SWITCH, for "view by" and metric selectors | A `CASE [Parameter]` that returns fields becomes the `field parameter` kind. |
| Calculated columns can't react to slicers | Row-level calcs that read a parameter are flagged and replaced by a measure filter (`COUNTROWS(FILTER(…)) > 0`) on each visual. |

**Observed differences from the Kiro steering files.** These are notes only; the Kiro files are not changed.
- The Kiro TWB extraction spec (Section 13) and `naming-standards.md` ask for one `_Measures` table and Dim/Fact-style names. The skills say measures live on their fact tables and names should be business names.
- The Employee Summary PBIP built on 2026-10-18 keeps every measure in one `Report Measures` table, which the modeling guidelines advise against.

## 6. How the outputs feed the build

1. **Planning** (powerbi-report-cli): the page section of the digest is the draft page plan. Answer every `DECISION NEEDED` and low-confidence item, then lock `_brief/report-spec.md` with the user.
2. **Model** (semantic-model-authoring, modeling MCP):
   1. Build the tables, relationships, date table and parameter tables from the model section.
   2. Paste `bhayankar_measures.dax` into DAX query view to test the measures, and fix the `medium` and `low` ones.
   3. Create the measures with their formatString, displayFolder and description from `bhayankar_rebuild.json`.
3. **Report** (authoring): a Node generator reads `bhayankar_rebuild.json` and writes PBIR. Still run `powerbi-report-author validate`, then reload and take screenshots after each batch. The catalog check prevents wrong roles, but it doesn't replace validation.
4. **QA:** check parity against the totals shown in the Tableau screenshots (input 3).

## 7. What the script handles, so the agent doesn't have to

- **File formats:**
  - unzips `.twbx` and reads only the inner `.twb`;
  - normalises Tableau 2020.2+ feature-flag tags (`_.fcp.X.true...tag`);
  - parses a truncated workbook in recovery mode and flags it `TRUNCATED_FILE`.
- **Used-only rule:**
  - starts from worksheet dependencies, shelves, filters, dashboard zones, action URLs and data source filters;
  - follows those through formulas and groups;
  - names unused (orphan) items once and never analyses them.
- **Internal names become captions everywhere.** For example:
  - `[Calculation_123]`, `[Parameter 1]` and `[sum:X:qk]` become their captions;
  - `pcto:sum:X` becomes `PERCENT_OF_TOTAL(SUM(X))`;
  - `__tableau_internal_object_id__` becomes `COUNT(rows of T)`.
  - Stale references show as `(deleted field …)`.
- **Deduplication:**
  - repeated `datasource-dependencies` blocks;
  - identical worksheet-local calcs;
  - identical dual-axis panes;
  - filters shared by three or more sheets (given F-IDs);
  - quick filters folded into one line.
- **Dropped entirely:** thumbnails, metadata records for unused columns, window state, zone styling, fonts, and revision/manifest data.

## 8. Flags

Each flag row becomes either an open question or a build decision.

| Group | Flags | Notes |
|---|---|---|
| File and source | `TRUNCATED_FILE`, `PUBLISHED_DS`, `CUSTOM_SQL`, `EXTRACT`, `BLENDING`, `BLEND_NO_LINK`, `BLEND_LINK_CALC`, `BLEND_LINK_TYPE`, `KEY_NAME_MISMATCH` | `BLENDING` names the primary and the active links per sheet; the build makes them shared dimensions. |
| Calculations | `LOD_FIXED`, `LOD_INCLUDE`, `LOD_EXCLUDE`, `TABLE_CALC`, `QUICK_TABLE_CALC`, `UNSUPPORTED_FN`, `DUPLICATE_LOGIC` | |
| Filters | `DATE_PARAM_VS_DATETIME`, `STRING_RANGE_FILTER`, `CONTEXT_FILTER`, `TOP_N_FILTER` | `DATE_PARAM_VS_DATETIME`: is the end date exclusive? `STRING_RANGE_FILTER`: an alphabetical range on a text field, such as a status range. |
| Visuals | `REFLINE_CROSS_AXIS`, `DUAL_AXIS`, `MULTI_MEASURE_AXES`, `MEASURE_VALUES`, `VIZ_IN_TOOLTIP` | |
| Actions | `FILTER_ACTION`, `URL_ACTION`, `NAVIGATION_ACTION`, `PARAMETER_ACTION`, `SET_ACTION` | |
| Other | `SET`, `CLUSTER`, `RLS` | |

## 9. When the digest isn't enough

Use `--show "<name>"` for the one item in question, for example a table calc's "compute using" or a zone's formatting. Don't open the TWB. If the same kind of detail keeps needing `--show`, extend `bhayankar_extractor.py` instead and note it in `docs/PROJECT_LOG.md`.

**Limits:**
- The suggested tier is only a starting point; a person confirms it.
- "Source table.field" comes from Tableau metadata. For a published source it shows the published field names until a `.tds` supplies the database columns and joins.
- DAX drafts aren't run by the extractor. Test `bhayankar_measures.dax` in Desktop.
- Visual inference covers standard marks. Tableau drawing tricks (trig or path charts, shape grids, viz extensions) come out as `custom`, `shape` or unmapped, with a note.

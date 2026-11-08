# QA checklist (before hand-off)

The tools build and check a project; they never sign one off. Work through this list for every migrated report.

1. **Build is clean.** `Run_Build.bat` says `validate: succeeded (0 errors, 0 warnings)`; `bhayankar_build_log.md` has no open `DECISION NEEDED`.
2. **Engine check.** Open the project in Power BI Desktop, click Refresh, run `Run_Verify.bat "<workbook>"`. Expect `ENGINE CHECKS PASSED`: no model errors, every measure evaluates, tables have rows.
3. **`_Needs fix` is empty or explained.** Each measure in a `Measures\_Needs fix` folder is rewritten by hand, or the report page that uses it is removed.
4. **Numbers.** For every page, set the same filters as the Tableau screenshot and compare the totals and a few rows. Always do this for the items under "Compare with Tableau" in `WORKFLOW_STATUS.md` (JOINED_GRAIN, PHYSICAL_JOIN, NAME_COLLISION, BLENDING, DATE_PARAM_VS_DATETIME).
5. **Filters and slicers.** Defaults match Tableau; a slicer that Tableau scoped to some sheets changes only those visuals.
6. **Navigation.** Drillthrough pages open from the right visuals and have a Back button; page order is Tableau's tab order.
7. **Security.** Anything flagged `RLS`, `PER_USER_COLUMN` or `HARDCODED_USERS` follows `SECURITY_RLS.md` and is approved by the architect.
8. **Look.** Page size, fonts, colours, number formats and logos match the screenshots; no visual shows an error in Desktop.
9. **Data source.** `UseSampleData` is false for production and the warehouse parameters are set; refresh succeeds against the real source.
10. **Hand-off.** Keep a copy of the finished `pbip` folder outside `output\`: a rebuild overwrites hand edits.

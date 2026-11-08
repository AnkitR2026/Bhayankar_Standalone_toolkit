# Notices

The Bhayankar Standalone Toolkit is the user's own tool. It bundles some Microsoft material so it can work offline.

| Item | Where | Source | Licence |
|---|---|---|---|
| PBIR / PBIP JSON Schemas | `resources\schemas\` | developer.microsoft.com/json-schemas/fabric (the `$schema` URLs Power BI writes) | published by Microsoft for validating Power BI project files |
| Visual metadata (types, roles, formatting objects) | `resources\pbir_metadata.json` | snapshot made with `tools\snapshot_pbir_metadata.mjs` from `@microsoft/powerbi-report-authoring-cli` 0.4.0 and `@microsoft/powerbi-core-visual-schema` 0.1.1 | MIT, Copyright (c) Microsoft Corporation |
| Base theme `CY26SU10.json` | `resources\themes\` | written by `powerbi-report-author scaffold` 0.4.0 | MIT, Copyright (c) Microsoft Corporation |
| Validation rules re-implemented in Python | `bhayankar_validate.py` | behaviour of `powerbi-report-author validate` 0.4.0 | MIT, Copyright (c) Microsoft Corporation |

MIT licence text: https://opensource.org/license/mit — "Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction ... The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software."

`bhayankar_verify.py` loads `Microsoft.PowerBI.AdomdClient.dll` from the user's own Power BI Desktop installation at run time. It is not bundled.

Client data: `input\` and `output\` can contain client workbooks, extracts and built reports. Do not share them outside the project's data rules.

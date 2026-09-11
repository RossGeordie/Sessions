# parkrun_Participation - 2026

- **File:** `parkrun_Participation - 2026.json`
- **Modified:** 2026-03-05T22:15:18.5504998+00:00
- **Output format:** Csv
- **Shared queries:** 5

## 1. Data Sources

| # | Type | Connection / Entity | Used by queries |
|---|------|---------------------|-----------------|
| 1 | **SharePoint.Files** | `https://«redacted-tenant».sharepoint.com/sites/JesmondDeneparkrun/` | `SharePointFiles_All` |
| 2 | **SharePoint.Tables** | `https://«redacted-tenant».sharepoint.com/sites/JesmondDeneparkrun/` | `TrackedEvents_Base` |
| 3 | **parameter** | `binary/parameter sample` | `Parameter` |

## 2. Transformation Steps (per query)

### `SharePointFiles_All`

_source: `SharePoint.Files` — `https://«redacted-tenant».sharepoint.com/sites/JesmondDeneparkrun/` · no load flag_

| # | Step name | Function |
|---|-----------|----------|
| 1 | Source | `SharePoint.Files` |
| 2 | Filtered Rows | `Table.SelectRows` |
| 3 | Removed other columns | `Table.SelectColumns` |
| 4 | Split column by delimiter | `Table.SplitColumn` |
| 5 | Filtered Rows1 | `Table.SelectRows` |
| 6 | Merged Queries | `Table.NestedJoin` |
| 7 | Expanded TrackedEvents_Base | `Table.ExpandTableColumn` |

### `TrackedEvents_Base`

_source: `SharePoint.Tables` — `https://«redacted-tenant».sharepoint.com/sites/JesmondDeneparkrun/` · no load flag_

| # | Step name | Function |
|---|-----------|----------|
| 1 | Source | `SharePoint.Tables` |
| 2 | Navigation 1 | — |
| 3 | Removed other columns | `Table.SelectColumns` |
| 4 | Changed column type | `Table.TransformColumnTypes` |

### `SharePointFiles_Participation_OLD`

_source: (references another query) · no load flag_

| # | Step name | Function |
|---|-----------|----------|
| 1 | Source | — |
| 2 | Filtered Rows | `Table.SelectRows` |
| 3 | Removed Other Columns | `Table.SelectColumns` |
| 4 | GetEventNumber | `Table.TransformColumns` |
| 5 | Renamed Columns | `Table.RenameColumns` |
| 6 | Added Custom1 | `Table.AddColumn` |
| 7 | Added Custom | `Table.AddColumn` |
| 8 | Expanded DataTables | `Table.ExpandTableColumn` |
| 9 | Expanded Data | `Table.ExpandTableColumn` |
| 10 | Removed columns 1 | `Table.RemoveColumns` |
| 11 | Renamed Expanded Columns | `Table.RenameColumns` |
| 12 | Inserted Text After Delimiter | `Table.AddColumn` |
| 13 | Split Column by Delimiter | `Table.SplitColumn` |
| 14 | Inserted Text Before Delimiter | `Table.AddColumn` |
| 15 | Removed Columns | `Table.RemoveColumns` |
| 16 | Split Column by Delimiter1 | `Table.SplitColumn` |
| 17 | Added Custom club1 | `Table.AddColumn` |
| 18 | Replaced Errors | `Table.ReplaceErrorValues` |
| 19 | Renamed Columns2 | `Table.RenameColumns` |
| 20 | Added Custom club2 | `Table.AddColumn` |
| 21 | Replaced Errors1 | `Table.ReplaceErrorValues` |
| 22 | Removed Columns1 | `Table.RemoveColumns` |
| 23 | Duplicated Column | `Table.DuplicateColumn` |
| 24 | Split Column by Delimiter2 | `Table.SplitColumn` |
| 25 | Changed Type | `Table.TransformColumnTypes` |
| 26 | Inserted Hour | `Table.AddColumn` |
| 27 | Inserted Minute | `Table.AddColumn` |
| 28 | Inserted Second | `Table.AddColumn` |
| 29 | Added Custom2 | `Table.AddColumn` |
| 30 | Changed Type1 | `Table.TransformColumnTypes` |
| 31 | Removed Columns2 | `Table.RemoveColumns` |
| 32 | Inserted Minute1 | `Table.AddColumn` |
| 33 | Inserted Hour1 | `Table.AddColumn` |
| 34 | Added Custom3 | `Table.AddColumn` |
| 35 | Changed Type2 | `Table.TransformColumnTypes` |
| 36 | Removed Columns3 | `Table.RemoveColumns` |
| 37 | Extracted Text Before Delimiter | `Table.TransformColumns` |
| 38 | Changed Type3 | `Table.TransformColumnTypes` |
| 39 | Divided Column | `Table.TransformColumns` |
| 40 | Renamed Columns3 | `Table.RenameColumns` |
| 41 | Changed Type4 | `Table.TransformColumnTypes` |
| 42 | Replaced value | `Table.ReplaceValue` |
| 43 | Transform columns | `Table.TransformColumnTypes` |
| 44 | Replace errors | `Table.ReplaceErrorValues` |

### `SharePointFiles_Participation`

_source: (references another query) · → loads to output_

| # | Step name | Function |
|---|-----------|----------|
| 1 | Source | — |
| 2 | Filtered Rows | `Table.SelectRows` |
| 3 | Removed Other Columns | `Table.SelectColumns` |
| 4 | GetEventNumber | `Table.TransformColumns` |
| 5 | Renamed Columns | `Table.RenameColumns` |
| 6 | Added Custom1 | `Table.AddColumn` |
| 7 | Filtered hidden files | `Table.SelectRows` |
| 8 | Transform columns | `Table.TransformColumnTypes` |
| 9 | Replace errors | `Table.ReplaceErrorValues` |
| 10 | Invoke custom function | `Table.AddColumn` |
| 11 | Removed columns | `Table.RemoveColumns` |
| 12 | Expanded table column | `Table.ExpandTableColumn` |
| 13 | Changed column type | `Table.TransformColumnTypes` |
| 14 | Source | — |
| 15 | Filtered Rows | `Table.SelectRows` |
| 16 | Removed Other Columns | `Table.SelectColumns` |
| 17 | GetEventNumber | `Table.TransformColumns` |
| 18 | Renamed Columns | `Table.RenameColumns` |
| 19 | Added Custom1 | `Table.AddColumn` |
| 20 | Filtered hidden files | `Table.SelectRows` |
| 21 | Navigation | — |

### `Parameter`

_source: `parameter` — `binary/parameter sample` · no load flag_

| # | Step name | Function |
|---|-----------|----------|
| 1 | Parameter | — |
| 2 | Source | `Excel.Workbook` |
| 3 | Navigation | — |
| 4 | Renamed Expanded Columns | `Table.RenameColumns` |
| 5 | Inserted Text After Delimiter | `Table.AddColumn` |
| 6 | Split Column by Delimiter | `Table.SplitColumn` |
| 7 | Inserted Text Before Delimiter | `Table.AddColumn` |
| 8 | Removed Columns | `Table.RemoveColumns` |
| 9 | Parkruns to Whole Number | `Table.TransformColumnTypes` |
| 10 | Replace parkrun counts for invalid parkrunID's | `Table.ReplaceErrorValues` |
| 11 | Split Column by Delimiter1 | `Table.SplitColumn` |
| 12 | Added Custom club1 | `Table.AddColumn` |
| 13 | Replaced Errors | `Table.ReplaceErrorValues` |
| 14 | Renamed Columns2 | `Table.RenameColumns` |
| 15 | Added Custom club2 | `Table.AddColumn` |
| 16 | Replaced Errors1 | `Table.ReplaceErrorValues` |
| 17 | Removed Columns1 | `Table.RemoveColumns` |
| 18 | Duplicated Column | `Table.DuplicateColumn` |
| 19 | Split Column by Delimiter2 | `Table.SplitColumn` |
| 20 | Changed column type | `Table.TransformColumnTypes` |
| 21 | Inserted Hour | `Table.AddColumn` |
| 22 | Inserted Minute | `Table.AddColumn` |
| 23 | Inserted Second | `Table.AddColumn` |
| 24 | Added Custom2 | `Table.AddColumn` |
| 25 | Changed Type1 | `Table.TransformColumnTypes` |
| 26 | Removed Columns2 | `Table.RemoveColumns` |
| 27 | Inserted Minute1 | `Table.AddColumn` |
| 28 | Inserted Hour1 | `Table.AddColumn` |
| 29 | Added Custom3 | `Table.AddColumn` |
| 30 | Removed Columns3 | `Table.RemoveColumns` |
| 31 | Extracted Text Before Delimiter | `Table.TransformColumns` |
| 32 | Changed column type 1 | `Table.TransformColumnTypes` |
| 33 | Replace Age Grade % with null | `Table.ReplaceErrorValues` |
| 34 | Divided Column | `Table.TransformColumns` |
| 35 | Renamed Columns3 | `Table.RenameColumns` |
| 36 | Replaced value | `Table.ReplaceValue` |
| 37 | Removed columns 1 | `Table.RemoveColumns` |
| 38 | Source | `Excel.Workbook` |
| 39 | Navigation | — |
| 40 | Renamed Expanded Columns | `Table.RenameColumns` |
| 41 | Inserted Text After Delimiter | `Table.AddColumn` |
| 42 | Split Column by Delimiter | `Table.SplitColumn` |
| 43 | Inserted Text Before Delimiter | `Table.AddColumn` |
| 44 | Removed Columns | `Table.RemoveColumns` |
| 45 | Parkruns to Whole Number | `Table.TransformColumnTypes` |
| 46 | Replace parkrun counts for invalid parkrunID's | `Table.ReplaceErrorValues` |
| 47 | Split Column by Delimiter1 | `Table.SplitColumn` |
| 48 | Added Custom club1 | `Table.AddColumn` |
| 49 | Replaced Errors | `Table.ReplaceErrorValues` |
| 50 | Renamed Columns2 | `Table.RenameColumns` |
| 51 | Added Custom club2 | `Table.AddColumn` |
| 52 | Replaced Errors1 | `Table.ReplaceErrorValues` |
| 53 | Removed Columns1 | `Table.RemoveColumns` |
| 54 | Duplicated Column | `Table.DuplicateColumn` |
| 55 | Split Column by Delimiter2 | `Table.SplitColumn` |
| 56 | Changed column type | `Table.TransformColumnTypes` |
| 57 | Inserted Hour | `Table.AddColumn` |
| 58 | Inserted Minute | `Table.AddColumn` |
| 59 | Inserted Second | `Table.AddColumn` |
| 60 | Added Custom2 | `Table.AddColumn` |
| 61 | Changed Type1 | `Table.TransformColumnTypes` |
| 62 | Removed Columns2 | `Table.RemoveColumns` |
| 63 | Inserted Minute1 | `Table.AddColumn` |
| 64 | Inserted Hour1 | `Table.AddColumn` |
| 65 | Added Custom3 | `Table.AddColumn` |
| 66 | Removed Columns3 | `Table.RemoveColumns` |
| 67 | Extracted Text Before Delimiter | `Table.TransformColumns` |
| 68 | Changed column type 1 | `Table.TransformColumnTypes` |
| 69 | Replace Age Grade % with null | `Table.ReplaceErrorValues` |
| 70 | Divided Column | `Table.TransformColumns` |
| 71 | Renamed Columns3 | `Table.RenameColumns` |
| 72 | Replaced value | `Table.ReplaceValue` |
| 73 | Removed columns 1 | `Table.RemoveColumns` |
| 74 | Source | — |
| 75 | Kept errors | `Table.SelectRowsWithErrors` |

## 3. Output Tables

### `SharePointFiles_Participation`

- **Refresh location:** `SharePointFiles_Participation.csv`
- **Columns: 17**

| Column | Type |
|--------|------|
| Event | int64 |
| LocationID | int64 |
| EKey | string |
| Category | string |
| FinishPosition | int64 |
| Participant | string |
| LINK | string |
| Gender | string |
| Age Grade | double |
| Club | string |
| Note | string |
| ParkrunID | string |
| parkruns | int64 |
| Volunteer Club | int64 |
| parkun Club | int64 |
| Time | time |
| FinishMinute | int64 |

## 4. Lineage — output ← sources

SharePoint.Tables: `https://«redacted-tenant».sharepoint.com/sites/JesmondDeneparkrun/` ← `SharePointFiles_Participation`  

## 5. Connection Overrides

- {"path": "https://«redacted-tenant».sharepoint.com/sites/JesmondDeneparkrun", "kind": "SharePoint", "provider": "CdsA", "authenticationKind": null, "environmentName": null, "apiName": null, "connectionName": "{\"kind\":\"SharePoint\",\"path\":\"https://«redacted-tenant».sharepoint.com/sites/JesmondDeneparkrun\"}", "audience": null}

---
_Generated by pqdoc from `parkrun_Participation - 2026.json`_

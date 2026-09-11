# parkrun_Volunteers+History - 2026

- **File:** `parkrun_Volunteers+History - 2026.json`
- **Modified:** 2026-03-14T16:54:19.209981+00:00
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

### `SharePointFiles_Volunteers`

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
| 8 | Filtered hidden files 1 | `Table.SelectRows` |
| 9 | Invoke custom function | `Table.AddColumn` |
| 10 | Removed errors 1 | `Table.RemoveRowsWithErrors` |
| 11 | Removed columns | `Table.RemoveColumns` |
| 12 | Changed column type | `Table.TransformColumnTypes` |
| 13 | Removed errors | `Table.RemoveRowsWithErrors` |
| 14 | Expanded Transform file | `Table.ExpandTableColumn` |
| 15 | Changed column type 1 | `Table.TransformColumnTypes` |
| 16 | Source | — |
| 17 | Filtered Rows | `Table.SelectRows` |
| 18 | Sorted rows | `Table.Sort` |
| 19 | Removed Other Columns | `Table.SelectColumns` |
| 20 | GetEventNumber | `Table.TransformColumns` |
| 21 | Renamed Columns | `Table.RenameColumns` |
| 22 | Added Custom1 | `Table.AddColumn` |
| 23 | Filtered hidden files | `Table.SelectRows` |
| 24 | Filtered hidden files 1 | `Table.SelectRows` |
| 25 | Navigation | — |

### `Parameter`

_source: `parameter` — `binary/parameter sample` · no load flag_

| # | Step name | Function |
|---|-----------|----------|
| 1 | Parameter | — |
| 2 | Source | `Excel.Workbook` |
| 3 | Navigation | — |
| 4 | Changed column type | `Table.TransformColumnTypes` |
| 5 | Split column by delimiter 1 | `Table.SplitColumn` |
| 6 | Split column by delimiter 2 | `Table.SplitColumn` |
| 7 | Split column by delimiter | `Table.ExpandListColumn` |
| 8 | Split column by delimiter 3 | `Table.SplitColumn` |
| 9 | Removed columns | `Table.RemoveColumns` |
| 10 | Removed columns 1 | `Table.RemoveColumns` |
| 11 | Renamed columns | `Table.RenameColumns` |
| 12 | Changed column type 1 | `Table.TransformColumnTypes` |
| 13 | Replaced errors | `Table.ReplaceErrorValues` |
| 14 | Source | `Excel.Workbook` |
| 15 | Navigation | — |
| 16 | Changed column type | `Table.TransformColumnTypes` |
| 17 | Split column by delimiter 1 | `Table.SplitColumn` |
| 18 | Split column by delimiter 2 | `Table.SplitColumn` |
| 19 | Split column by delimiter | `Table.ExpandListColumn` |
| 20 | Split column by delimiter 3 | `Table.SplitColumn` |
| 21 | Removed columns | `Table.RemoveColumns` |
| 22 | Removed columns 1 | `Table.RemoveColumns` |
| 23 | Renamed columns | `Table.RenameColumns` |
| 24 | Changed column type 1 | `Table.TransformColumnTypes` |
| 25 | Replaced errors | `Table.ReplaceErrorValues` |

### `SharePointFiles_EventHistory`

_source: (references another query) · → loads to output_

| # | Step name | Function |
|---|-----------|----------|
| 1 | Source | — |
| 2 | Filtered Rows | `Table.SelectRows` |
| 3 | Removed Other Columns | `Table.SelectColumns` |
| 4 | Renamed Columns | `Table.RenameColumns` |
| 5 | Added Custom | `Table.AddColumn` |
| 6 | Expanded DataTables | `Table.ExpandTableColumn` |
| 7 | Removed columns 1 | `Table.RemoveColumns` |
| 8 | Expanded Data | `Table.ExpandTableColumn` |
| 9 | Changed Type | `Table.TransformColumnTypes` |
| 10 | Removed Columns | `Table.RemoveColumns` |
| 11 | Split Column by Positions | `Table.SplitColumn` |
| 12 | Renamed_Columns | `Table.RenameColumns` |
| 13 | Changed Type1 | `Table.TransformColumnTypes` |
| 14 | Renamed Columns2 | `Table.RenameColumns` |
| 15 | Removed Columns1 | `Table.RemoveColumns` |
| 16 | Added Custom Finishers | `Table.AddColumn` |
| 17 | Added Custom1 | `Table.AddColumn` |
| 18 | Split Column by Character Transition | `Table.SplitColumn` |
| 19 | Merged Columns | `Table.CombineColumns` |
| 20 | Renamed Columns3 | `Table.RenameColumns` |
| 21 | Split Column by Character Transition1 | `Table.SplitColumn` |
| 22 | Merged Columns1 | `Table.CombineColumns` |
| 23 | Renamed Columns4 | `Table.RenameColumns` |
| 24 | Changed Type2 | `Table.TransformColumnTypes` |
| 25 | Removed Columns2 | `Table.RemoveColumns` |
| 26 | Inserted Year | `Table.AddColumn` |
| 27 | Added Custom2 | `Table.AddColumn` |
| 28 | Added Custom3 | `Table.AddColumn` |
| 29 | Removed Columns3 | `Table.RemoveColumns` |
| 30 | Renamed Columns5 | `Table.RenameColumns` |
| 31 | Changed Type3 | `Table.TransformColumnTypes` |
| 32 | Added custom 2 | `Table.TransformColumnTypes` |

## 3. Output Tables

### `SharePointFiles_Volunteers`

- **Refresh location:** `SharePointFiles_Volunteers.csv`
- **Columns: 10**

| Column | Type |
|--------|------|
| Event | int64 |
| LocationID | int64 |
| EKey | string |
| Volunteer | string |
| ParkrunID | int64 |
| Completed | int64 |
| Volunteer Club | string |
| parkrun Club | string |
| Role | string |
| Club | string |

### `SharePointFiles_EventHistory`

- **Refresh location:** `SharePointFiles_EventHistory.csv`
- **Columns: 11**

| Column | Type |
|--------|------|
| LocationID | int64 |
| Event | int64 |
| Event Date | date |
| MaleFF | string |
| MaleFFTime | string |
| FemaleFF | string |
| FemaleFFTime | string |
| Year | int64 |
| FInishers | int64 |
| Volunteers | int64 |
| EKey | string |

## 4. Lineage — output ← sources

SharePoint.Tables: `https://«redacted-tenant».sharepoint.com/sites/JesmondDeneparkrun/` ← `SharePointFiles_Volunteers`  
SharePoint.Tables: `https://«redacted-tenant».sharepoint.com/sites/JesmondDeneparkrun/` ← `SharePointFiles_EventHistory`  

## 5. Connection Overrides

- {"path": "https://«redacted-tenant».sharepoint.com/sites/JesmondDeneparkrun", "kind": "SharePoint", "provider": "CdsA", "authenticationKind": null, "environmentName": null, "apiName": null, "connectionName": "{\"kind\":\"SharePoint\",\"path\":\"https://«redacted-tenant».sharepoint.com/sites/JesmondDeneparkrun\"}", "audience": null}

---
_Generated by pqdoc from `parkrun_Volunteers+History - 2026.json`_

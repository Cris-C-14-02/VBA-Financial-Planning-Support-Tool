# VBA source map

The root `.xlsm` is the runnable application and source of truth. This folder is a readable snapshot extracted from its embedded VBA project, with UTF-8 encoding and LF line endings. Extraction does not execute macros.

| Path | Contents | Editing contract |
| --- | --- | --- |
| `modules/*.bas` | Standard modules, including original `Attribute VB_Name` | Names match the workbook project |
| `forms/*.frm.txt` | UserForm code and attributes | Code only; designer layout, controls and `.frx` resources remain in the workbook |
| `host/*.cls.txt` | Worksheet and ThisWorkbook code | Includes empty host components for completeness; retain the existing host identity |
| `manifest.json` | Workbook hash, component paths and source hashes | Regenerate with the extraction script |
| `legacy/` | Original loose repository snippets | Historical references; may differ from the embedded project |

Do not bulk-import the text snapshots into Excel. To edit a form or sheet event, use its existing component in the VBA editor. For full UserForm exports, use Excel's Export File command and keep each `.frm` together with its `.frx` when present.

## Synchronize after an Excel change

Install the optional extraction dependency, then run from the repository root:

```powershell
python -m pip install -r scripts/requirements.txt
python scripts/extract_vba.py "Cris Chen - Individual Milestone 4 - Prototype.xlsm"
python scripts/extract_vba.py "Cris Chen - Individual Milestone 4 - Prototype.xlsm" --check
```

Save and close the workbook before extracting. Review both the binary change and text diff. Editing this snapshot alone does not update the workbook. `--check` verifies source synchronization and the workbook hash; it does not compile VBA or test Excel behavior.

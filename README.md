[README.md](https://github.com/user-attachments/files/28913748/README.md)
# Excel_Macros_Alim
Standard format style for numbers, percentage, multiples and colors (hard-coded, direct cross sheet, formulas, etc)
# Alimzhan Format Macros (v3)

Excel VBA macros for fast equity-research formatting: number/percent/multiple formats, IB-style font color-coding (inputs blue, cross-sheet links green, plug-ins purple, calculations black), and one-click chart formatting.

> These are **Excel VBA macros**, not a web add-in. They live inside a macro-enabled workbook (`.xlsm`) or your Personal Macro Workbook. GitHub here is just a place to **download the code** — colleagues copy it into their own Excel. Macros cannot be "served" or auto-loaded from a website the way the Formula Navigator add-in is.

## Files

| File | What it is |
|---|---|
| `AlimzhanFormatMacros_v3.bas` | The macro module (`Module1`). Import this into the VBA editor. |
| `ThisWorkbook_v3.txt` | One line for the `ThisWorkbook` object so shortcuts register on open. |
| `МАКРОСЫ-как-обновить.md` | Step-by-step install guide (in Russian). |

## Shortcuts

| Shortcut | Action |
|---|---|
| Ctrl+Shift+Q | Number format `#,##0` (accounting) |
| Ctrl+Shift+T | Percentage `0.0%` |
| Ctrl+Shift+R | Multiples `0.0x` |
| Ctrl+Shift+A | Color-code the selected range by cell type |
| Ctrl+Shift+Z | Restore original font colors (last color-code only) |
| Ctrl+Shift+J | Format active chart to 7×4 inches |
| Ctrl+Shift+M | Format active chart to 5×4 inches |

Color-coding legend: **blue** = hardcoded inputs (numbers, dates, TRUE/FALSE, number-as-text), **green** = direct single-cell link to another sheet (`=Sheet!A1`), **purple** = FactSet/CapIQ/Bloomberg functions (`FDS(`, `CIQ(`, `BDP(`…), **black** = all other formulas. Plain text cells are left untouched.

## Install

1. Open your workbook → **Alt+F11** (VBA editor).
2. Right-click **Modules** → **Import File…** → choose `AlimzhanFormatMacros_v3.bas`. (Or open `Module1`, select all, paste the code.)
3. Double-click **ThisWorkbook** in the tree → make sure `Workbook_Open` calls `SetupShortcuts` (see `ThisWorkbook_v3.txt`).
4. Save the workbook as **.xlsm** (macro-enabled).
5. Reopen the workbook, **or** run `SetupShortcuts` once (Alt+F8 → SetupShortcuts → Run). Shortcuts are now live.

To use the macros in **every** workbook (not just one), import the module into your **Personal Macro Workbook** (`PERSONAL.XLSB`) instead — then they're always available.

## What's new in v3

- Esc no longer interrupts a running macro (`Application.EnableCancelKey = xlDisabled`) — fixes the "Code execution has been interrupted" error.
- `ScreenUpdating` is always restored, even if a macro is interrupted mid-run (screen no longer "freezes").
- Key binding rewritten via `SetupShortcuts` (clears old bindings first, so reopening never leaves stale shortcuts).
- Color-coding logic, number formats, chart sizing, and plug-in detection are unchanged from v2.

## Notes

- **No undo:** font-color changes from color-coding are normal cell edits and *can* be undone with Ctrl+Z; `Ctrl+Shift+Z` restores the colors saved by the **most recent** color-code run only.
- These macros are independent of the **Formula Navigator** add-in (which uses Ctrl+Shift+C / D / E) — no shortcut conflict.

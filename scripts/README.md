# Stage 5 PDF → Excel Extractor

One-off pipeline that turned 1284 single-page "Stage 5 — CCJs and searches"
underwriting PDFs (in `Stage 5a.rar`, `Stage 5b.rar`, `Stage 5c.rar` at repo root)
into a single Excel sheet of structured data.

This README exists so a future Claude session can re-run, extend, or debug the
pipeline without re-discovering the format from scratch.

## What it does

Reads every `*.pdf` in a directory, pulls 6 fields out of each, and writes them
to an `.xlsx` with all columns formatted as real numbers (Score formatted `0.0`).

Output columns (sheet `Stage5 CCJs`):

| Column | Type  | Source field in PDF                                          |
|--------|-------|--------------------------------------------------------------|
| Lead ID                 | int   | `ACCOUNT NO:` line                                |
| # unsatisfied CCJs      | int   | `# unsatisfied CCJs (from credit agency report)`  |
| # satisfied CCJs        | int   | `# satisfied CCJs (from credit agency report)`    |
| # credit searches 3m    | int   | `# credit searches last 3 months ...`             |
| # credit searches 12m   | int   | `# credit searches last 12 months ...`            |
| Score                   | float | `Resulting score to enter into "5 - Overall ..."` |

A second sheet `Errors` lists any file where one or more fields failed to match
(e.g. agent left a field blank in the source PDF) — the row is still written to
the main sheet with `None` for the missing cell(s).

## PDF format (consistent across all 1284)

Single-page, text-based (no OCR needed). Verbatim layout:

```
CUSTOMER NAME: <name>
ACCOUNT NO: <digits>
AGENT NAME: <name>
DATE:
# unsatisfied CCJs (from credit agency report) <int>
# satisfied CCJs (from credit agency report) <int>
# credit searches last 3 months (from credit agency report) <int>
# credit searches last 12 months (from credit agency report) <int>
Resulting score to enter into "5 - Overall grading - all factors considered" <decimal>
```

Notes:
- `DATE:` is always blank — not extracted.
- Score range observed: `0.0` to `2.9`.
- Filename pattern is `5 - CCJs and searches - <Name> - <AccountNo>[ - 00-01-1900].pdf`
  (the trailing date suffix is inconsistent). **Don't parse Lead ID from the
  filename** — always use `ACCOUNT NO:` from inside the PDF.

## Setup

```bash
# inside the repo root
sudo apt-get install -y unrar                              # if not present
pip install --upgrade --ignore-installed cffi cryptography # fixes pdfminer on this image
pip install pdfplumber openpyxl

# extract PDFs once
mkdir -p /tmp/pdfs
unrar x -o+ "Stage 5a.rar" /tmp/pdfs/
unrar x -o+ "Stage 5b.rar" /tmp/pdfs/
unrar x -o+ "Stage 5c.rar" /tmp/pdfs/
```

## Run

```bash
# smoke test
python3 scripts/extract_stage5.py /tmp/pdfs /tmp/test50.xlsx --limit 50

# full batch
python3 scripts/extract_stage5.py /tmp/pdfs output_stage5_all.xlsx
```

Takes ~30 seconds for the full 1284 on this container.

## Last run result (2026-05-26)

- **1284 PDFs** processed. The archives contain 1281 files with a `.pdf`
  extension plus **3 PDFs that have no extension at all** (`Ashley Mitchell`,
  `Daniel Sigley `, `Sharon McAleer`). The script detects PDFs by magic bytes
  (`%PDF`) instead of file extension so all 1284 are picked up.
- **1283 fully clean** rows.
- **1 partial**: `Simon Scott - 223600570` — agent left "# unsatisfied CCJs"
  blank in the source PDF. The other 5 fields extracted correctly; the
  unsatisfied cell is `None` in the output. Logged in the `Errors` sheet.
- Output: `output_stage5_all.xlsx`

## Implementation notes (why the code looks like it does)

- **`re.MULTILINE` anchoring** (`^...\s*$`). Each label phrase is unique on its
  line, so anchored patterns are unambiguous. In particular:
  - `^#\s*satisfied CCJs` does **not** match the `# unsatisfied CCJs` line
    because `^` resets at every `\n` and the line starting with `# unsatisfied`
    has `un` between `#\s*` and `satisfied`.
  - The `3 months` and `12 months` patterns can't cross-match for the same
    reason (anchored at start of line).
- **`\d+(?:\.\d+)?`** for the score covers `0.0`, `0.1`, `1.0`, `2.9`, and any
  hypothetical integer score — `float()` handles both forms.
- **Per-file try/except** wraps `pdfplumber.open` so one corrupt PDF cannot
  abort the batch (verified — see review #4 in the original session).
- **`pdfplumber.extract_text()` returning `None`** (e.g. scanned/image-only PDF)
  is defensively replaced with `""`, so all 6 fields would simply be `None` and
  the file would land in the Errors sheet. No OCR fallback is implemented; if
  you ever hit a scanned PDF, add Tesseract as a fallback.
- **Numeric cell types**: `int` for Lead ID and counts, `float` for Score. The
  `number_format = "0.0"` on the Score column means Excel renders whole-number
  scores as `1.0` not `1` (openpyxl may serialise `1.0` as `1` in XML — the
  format string still controls display).

## If you need to re-run with different rules

The extraction logic is a single dict of compiled regexes at the top of
`extract_stage5.py` (`PATTERNS`). Add a new key, add a header in `main()`,
and append a column to the `ws.append([...])` call.

If you ever see new errors in the Errors sheet, inspect the offending PDF
with:

```bash
python3 -c "
import pdfplumber
with pdfplumber.open('/tmp/pdfs/<filename>') as pdf:
    print(pdf.pages[0].extract_text())
"
```

…then either adjust the regex or treat it as genuine missing source data
(the Simon Scott case).

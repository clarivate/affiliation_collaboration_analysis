# Web of Science Affiliation Collaboration Analysis

A Python command-line tool that retrieves records from the Web of Science Expanded API and identifies authors affiliated with a target organization.

The script creates an Excel workbook containing record-level citation, author, address, and collaboration details. It saves the workbook after every retrieved page so completed work is preserved if a later API request fails.

## Features

- Searches Web of Science Core Collection records using a user-supplied query.
- Identifies authors connected to a target affiliation.
- Counts matching authors and distinct matching addresses.
- Flags records with at least two matching authors.
- Flags collaborations where every author matches the target affiliation.
- Extracts matching author names, full addresses, suborganizations, and author-to-address relationships.
- Includes Web of Science Platform and Core Collection citation counts.
- Checks the target organization and query size before retrieval.
- Automatically retries failed pages using smaller page sizes.
- Gradually increases the page size again after successful retrievals.
- Saves results after every page.
- Logs unretrievable records in the output workbook instead of stopping the entire run.

## Repository Files

```text
affiliation_collaboration_analysis.py
wosesrclient_robust.py
requirements.txt
README.md
.env
```

`wosesrclient_robust.py` is a required companion module and must be located in the same directory as the analysis script, unless it is installed as an importable Python module.

Do not commit `.env` or any other file containing an API key.

## Requirements

- Python 3.9 or newer
- A Web of Science Expanded API key
- Access to the Web of Science Expanded API
- The companion `wosesrclient_robust.py` module

## Installation

Clone or download the repository, then create a virtual environment.

### Windows

```powershell
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

### macOS or Linux

```bash
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## API Key Setup

Create a file named `.env` in the project directory:

```dotenv
EXPANDED_APIKEY=your_api_key_here
```

The script reads the key from the `EXPANDED_APIKEY` environment variable.

## Basic Usage

Run the script with its configured defaults:

```bash
python affiliation_collaboration_analysis.py
```

Run a custom query and target affiliation:

```bash
python affiliation_collaboration_analysis.py \
  --query "OG=(University of Pittsburgh) AND PY=2025" \
  --affiliation "University of Pittsburgh"
```

On Windows PowerShell, the same command can be entered on one line:

```powershell
python affiliation_collaboration_analysis.py --query "OG=(University of Pittsburgh) AND PY=2025" --affiliation "University of Pittsburgh"
```

Specify an output filename:

```bash
python affiliation_collaboration_analysis.py \
  --query "OG=(University of Pittsburgh)" \
  --affiliation "University of Pittsburgh" \
  --output pitt_affiliation_analysis.xlsx
```

If `--output` is omitted, the script creates a timestamped `.xlsx` filename using the configured output prefix.

## Command-Line Options

| Option | Description |
|---|---|
| `-q`, `--query` | Web of Science search query. |
| `-a`, `--affiliation` | Target organization name used for address matching. |
| `-o`, `--output` | Output Excel filename. |
| `--count` | Preferred number of records requested per API page. |
| `--yes` | Automatically continue past preflight warnings. |
| `--skip-affiliation-check` | Skip the preliminary `OG=(affiliation)` validation query. |
| `--skip-main-query-count-check` | Skip the preliminary main-query size check. |
| `--affiliation-warning-threshold` | Warn when the affiliation preflight returns fewer than this number of records. |
| `--main-query-warning-threshold` | Warn when the main query returns this number of records or more. |

Display the complete built-in help:

```bash
python affiliation_collaboration_analysis.py --help
```

## Default Configuration

The default query, affiliation, output prefix, thresholds, and page-recovery settings are defined near the top of the script:

```python
DEFAULT_QUERY = "OG=(Trinity College)"
DEFAULT_AFFILIATION = "Trinity College"
OUTPUT_PREFIX = "TrinColl"

COUNT = 50
AFFILIATION_WARNING_THRESHOLD = 500
MAIN_QUERY_WARNING_THRESHOLD = 25000

ADAPTIVE_STEP_UP_AFTER_PAGES = 5
PAGE_SIZE_FALLBACKS = [50, 25, 10, 5, 1]
```

These defaults can be changed in the script or overridden with command-line options where available.

## Excel Output

The workbook contains an **Affiliation Analysis** worksheet with run metadata followed by these columns:

| Column | Description |
|---|---|
| UT | Web of Science accession number. |
| Title | Item title. |
| WOS Platform Cites | Citation count associated with `coll_id="WOK"`. |
| Core Collection Cites | Citation count associated with `coll_id="WOS"`. |
| Total Authors | Total author count reported for the record. |
| Matching Authors | Number of distinct authors connected to matching addresses. |
| Distinct Matching Addresses | Number of distinct matching address records. |
| >=2 Matching Authors | `YES` when at least two authors match the target affiliation. |
| Internal Only Collab | `YES` when all authors match the target affiliation. |
| Matching Author Names | Names of authors connected to matching addresses. |
| Matching Full Addresses | Full matching addresses with address identifiers. |
| Known Suborganizations | Departments or other suborganizations found in matching addresses. |
| Author Address Matches | Author-to-address relationships. |
| Error | Record-processing or retrieval error details. |

The worksheet includes filters, frozen headers, wrapped text, and adjusted column widths.

## Preflight Checks

By default, the script performs two checks before retrieving the full result set:

1. It runs `OG=(target affiliation)` and warns when the result count is below the affiliation threshold. This can reveal a misspelled or unexpected organization name.
2. It checks the size of the main query and warns when the result count reaches the large-query threshold.

Use `--yes` for unattended runs where warnings should be accepted automatically.

## Page Recovery and Saving

The preferred page size is controlled by `--count`. When a page fails, the script retries the same starting record with progressively smaller page sizes. After several successful pages, it attempts to increase the page size again.

The workbook is saved after every processed page. Because page-level saving already reports completed progress, the script does not include a separate record-level progress interval option.

If even a one-record request fails, the script writes a retrieval-error row, advances to the next record, and continues.

## Matching Method

Organization names are normalized for case and whitespace, then compared using exact matching. The script first checks preferred organization names and then checks all organization names attached to the address.

The value supplied through `--affiliation` should therefore match the organization form used in Web of Science.

## Important Notes

- The script uses API Expanded record quota and does not use the Short Record option.
- The script does not create an output file when the query returns zero records.
- The script stops when the query exceeds the configured 100,000-record maximum.
- A workbook may still be created when records are retrieved but no matching authors are found. This can help troubleshoot organization-name or address-matching issues.
- API behavior, entitlements, and available fields depend on the user's Web of Science subscription and API access.

# OCLC Record Matcher

This project provides tools for processing bibliographic records and matching them with OCLC data. It includes scripts for extracting data from MARC files, analyzing MARC field usage, and matching records against the WorldCat Metadata API—including **Excel, CSV, and TSV** inputs, **direct MARC** matching, and optional **combined MARCXML** export for matched OCLC numbers.

## Features

### OCLC Record Matching (`oclc_record_matcher.py`)
- **Tabular input**: Excel (`.xlsx`, `.xls`), UTF-8 **CSV**, or UTF-8 **TSV** with ISBN columns (same column detection rules as Excel)
- **MARC input**: `.mrc` / `.marc` files are converted via the bundled extractor to a temporary workbook, then matched like other inputs
- **Multi-ISBN support**: Automatically detects and processes multiple ISBN columns (for example XML ISBN, HC ISBN, PB ISBN, ePub ISBN, ePDF ISBN)
- **OR query optimization**: Combines all ISBNs from the same row in a single brief-bibs API call
- **Alternative search**: When no ISBN is available, searches using title, author, publisher, and publication date
- **Format-based search**: Maps format types to appropriate OCLC API parameters
- **LCSH detection (optional)**: Off by default for fewer API calls; pass **`--lcsh`** to fill `hasLCSHSubjects` via full bib JSON (`GET /worldcat/bibs/{oclcNumber}`)
- **MARCXML export**: Optional combined **MARCXML** file for distinct matched OCLC numbers using `GET /worldcat/manage/bibs/{oclcNumber}` (`--marcxml-output`)
- **MARC-only runs**: If you pass `--marcxml-output` and **omit** `-o`, no Excel file is written; you only get the combined MARCXML (matching still runs in memory)
- **Comprehensive logging**: Detailed API request/response logging with configurable verbosity

### MARC Data Extraction (`marc_extractor.py`)
- **MARC Field Extraction**: Extracts data from standard MARC fields:
  - `020$a` - ISBN
  - `245$a` + `245$b` - Title (normalized)
  - `100$a` or `110$a` - Author
  - `260$b` or `264$b` - Publisher
  - `260$c` or `264$c` - Publication Date (normalized to 4-digit year)
  - `300` - Physical Description
- **Format Detection**: Automatically determines format based on MARC leader and control fields
- **Data Normalization**: Cleans and normalizes extracted data

### MARC Field Analysis (`marc_field_analyzer.py`)
- **Field Usage Analysis**: Analyzes MARC files to determine most common fields and subfields
- **Control Field Analysis**: Examines leader positions and control field usage
- **Statistical Reports**: Generates Excel reports with field frequency data

### OCLC API Integration
- **WorldCat Metadata API**: Uses the official WorldCat Metadata API with OAuth 2.0 authentication
- **Secure credential management**: Environment variables via `python-dotenv` for API keys and secrets
- **OAuth 2.0 authentication**: Automatic token management with client credentials flow
- **Smart parameter mapping**: Maps format types to `itemType` or `itemSubType` for brief-bibs search
- **Rate limiting**: Configurable delay between requests (`API_RATE_LIMIT_DELAY`), applied to search and MARCXML retrieval
- **Error handling**: Automatic token refresh on `401` responses; resilient handling of per-record failures during export

## Files

### Main Scripts
- `oclc_record_matcher.py` - OCLC API matching for Excel, CSV, TSV, or MARC input; optional combined MARCXML export
- `marc_extractor.py` - Extract MARC data to Excel format
- `marc_field_analyzer.py` - Analyze MARC field usage and frequency

### Documentation
- `README.md` - This file
- `MARC_EXTRACTOR_README.md` - Documentation for MARC extraction functionality
- `MARC_FIELD_ANALYZER_README.md` - Documentation for MARC field analysis

### Sample Data
- `sampleData/recordsToMatch.xlsx` - Sample Excel file with ISBNs
- `sampleData/MLN-cataloging-RFP-vendor-sample-batch.mrc` - Sample MARC file
- `sampleData/testRecords.xlsx` - Smaller test Excel file

### Configuration
- `pyproject.toml` - Project metadata and dependencies
- `requirements.txt` - Legacy Python dependencies (for reference)
- `.env.example` - Example environment variables file (copy to `.env` and fill in your credentials)
- `.env` - Your actual environment variables (not tracked in git)
- `pyRecordMatch.code-workspace` - VS Code workspace configuration

## Setup

1. **Install uv (if not already installed):**
   ```bash
   curl -LsSf https://astral.sh/uv/install.sh | sh
   ```

2. **Setup and activate virtual environment:**
   ```bash
   uv venv my_venv
   source my_venv/bin/activate
   ```

3. **Install project dependencies:**
   ```bash
   uv pip install -e .
   ```
   
   Or for development with additional tools:
   ```bash
   uv pip install -e ".[dev]"
   ```

4. **Configure WorldCat Metadata API credentials:**
   
   a. **Get API credentials:**
      - Visit the [OCLC Developer Network](https://www.oclc.org/developer/api/oclc-apis/worldcat-metadata-api.en.html)
      - Register for a WorldCat Metadata API key and secret
   
   b. **Set up environment variables:**
      ```bash
      # Copy the example file
      cp .env.example .env
      
      # Edit .env and add your credentials
      # Required: OCLC_API_KEY and OCLC_API_SECRET
      ```
   
   c. **Required environment variables:**
      - `OCLC_API_KEY` - Your OCLC API key (required)
      - `OCLC_API_SECRET` - Your OCLC API secret (required)
   
   d. **Optional environment variables:**
      - `OCLC_API_BASE_URL` - API base URL (default: `https://metadata.api.oclc.org`)
      - `OCLC_OAUTH_TOKEN_URL` - OAuth token URL (default: `https://oauth.oclc.org/token`)
      - `API_TIMEOUT` - Request timeout in seconds (default: `30`)
      - `API_RATE_LIMIT_DELAY` - Delay between requests in seconds (default: `0.5`)
      - `LOG_LEVEL` - Logging level: DEBUG, INFO, WARNING, ERROR (default: `INFO`)
      - `LOG_FILE` - Log file path (default: `oclc_matcher.log`)
      - `API_LOGGING` - Enable detailed API logging: true/false (default: `true`)

5. **Verify input file structure:**
   - **Excel**: ISBN columns are detected when the header contains `ISBN` (case-insensitive substring)
   - **CSV / TSV**: UTF-8 (optional BOM); same header rules as Excel; comma for `.csv`, tab for `.tsv`
   - **MARC (`.mrc` / `.marc`)**: Standard fields are extracted to columns via `marc_extractor` before matching

## Usage

For **`oclc_record_matcher.py`**, see **Examples** in the [Command-line options](#command-line-options) section below for copy-paste commands grouped by switches.

### Workflow Overview

Common paths:

1. **From MARC**: Extract to Excel with `marc_extractor.py`, then match with `oclc_record_matcher.py`, **or** run the matcher directly on the `.mrc` file (it extracts to a temporary workbook).
2. **From Excel / CSV / TSV**: Run `oclc_record_matcher.py` on the file; default output is an Excel workbook with `matchingOCLCNumber` and related columns.
3. **MARCXML only**: Run the matcher with `--marcxml-output records.xml` and **omit** `-o` to skip writing an `.xlsx` file while still downloading full MARCXML for matched OCLC numbers.

### Step 1: MARC Data Extraction

**Extract MARC data to Excel:**
```bash
python marc_extractor.py -i sampleData/sample-batch.mrc -o extracted_data.xlsx
```

**Analyze MARC field usage:**
```bash
python marc_field_analyzer.py -i sampleData/sample-batch.mrc -o field_analysis.xlsx
```

### Step 2: OCLC Record Matching

**Process Excel (default sample input):**
```bash
python oclc_record_matcher.py -i sampleData/recordsToMatch.xlsx -o matched_output.xlsx
```

**Process CSV or TSV (UTF-8):**
```bash
python oclc_record_matcher.py -i my_records.csv -o matched_output.xlsx
python oclc_record_matcher.py -i my_records.tsv -o matched_output.xlsx
```

**Excel plus combined MARCXML for matched OCLC numbers:**
```bash
python oclc_record_matcher.py -i sampleData/recordsToMatch.xlsx -o matched_output.xlsx --marcxml-output matched_bibs.xml
```

**MARCXML only (no Excel output):**
```bash
python oclc_record_matcher.py -i my_records.csv --marcxml-output matched_bibs.xml
```

**Use default input (sample Excel):**
```bash
python oclc_record_matcher.py
```

### Complete Workflow Example

**From MARC to OCLC-matched data:**
```bash
# Option A: Two-step (explicit Excel in the middle)
python marc_extractor.py -i sampleData/MLN-cataloging-RFP-vendor-sample-batch.mrc -o marc_data.xlsx
python oclc_record_matcher.py -i marc_data.xlsx -o final_matched_data.xlsx

# Option B: One-step (matcher extracts MARC to a temp workbook internally)
python oclc_record_matcher.py -i sampleData/MLN-cataloging-RFP-vendor-sample-batch.mrc -o final_matched_data.xlsx
```

### Command-Line Options

#### OCLC Record Matcher (`oclc_record_matcher.py`)

| Option | Description | Default |
|--------|-------------|---------|
| `-i, --input` | Input path: `.xlsx` / `.xls`, `.csv`, `.tsv`, or `.mrc` / `.marc` | `sampleData/recordsToMatch.xlsx` |
| `-o, --output` | Output Excel (`.xlsx`). Optional when `--marcxml-output` is set; omit `-o` for MARCXML-only | `<input_stem>_with_oclc.xlsx` |
| `--marcxml-output FILE` | After matching, write one combined MARCXML file (manage bibs) for distinct matched OCLC numbers | (disabled) |
| `--lcsh` | After each match, call GET `/worldcat/bibs/{oclcNumber}` to detect LCSH and fill `hasLCSHSubjects` | Off (fewer API calls) |
| `--no-backup` | Skip creating backup of input file | Create backup |
| `--log-level` | Set logging level (DEBUG, INFO, WARNING, ERROR) | `INFO` |
| `--log-file` | Custom log file path | `oclc_matcher.log` |
| `--no-api-logging` | Disable detailed API request/response logging | Enable API logging |
| `-h, --help` | Show help message | - |

### Examples: `oclc_record_matcher.py`

Use `python` or `uv run python` (from the project root with the venv active).

#### Input and output (`-i`, `-o`)

```bash
# Explicit Excel input and output paths
python oclc_record_matcher.py -i sampleData/recordsToMatch.xlsx -o results/matched.xlsx

# CSV or TSV (UTF-8); output is always an Excel workbook when -o is set
python oclc_record_matcher.py -i data/titles.csv -o data/titles_with_oclc.xlsx
python oclc_record_matcher.py -i data/titles.tsv -o data/titles_with_oclc.xlsx

# Binary MARC: matcher extracts to a temp sheet, then writes your -o Excel
python oclc_record_matcher.py -i sampleData/MLN-cataloging-RFP-vendor-sample-batch.mrc -o matched_from_marc.xlsx

# Use default sample input; output defaults to sampleData/recordsToMatch_with_oclc.xlsx
python oclc_record_matcher.py
```

#### MARCXML export (`--marcxml-output`)

Requires API key scopes that allow **manage bibs** / **view MARC** (see [API documentation](#api-documentation) below).

```bash
# Excel results plus one combined MARCXML file for distinct matched OCLC numbers
python oclc_record_matcher.py -i sampleData/recordsToMatch.xlsx -o matched.xlsx --marcxml-output exports/bibs.xml

# Same, from a UTF-8 CSV
python oclc_record_matcher.py -i data/books.csv -o data/books_with_oclc.xlsx --marcxml-output exports/books.xml
```

#### LCSH column (`--lcsh`)

By default **`hasLCSHSubjects`** is left empty. Pass **`--lcsh`** to run an extra bib lookup per match and populate true/false (requires **`WorldCatMetadataAPI:read_bibs`** or equivalent for `GET /worldcat/bibs/{oclcNumber}`).

```bash
python oclc_record_matcher.py -i sampleData/recordsToMatch.xlsx -o matched.xlsx --lcsh
```

#### MARCXML only (omit `-o`)

No `.xlsx` is written; matching still runs in memory, then MARCXML is fetched.

```bash
python oclc_record_matcher.py -i data/books.csv --marcxml-output exports/books_only.xml
python oclc_record_matcher.py -i sampleData/MLN-cataloging-RFP-vendor-sample-batch.mrc --marcxml-output exports/from_marc.xml
```

#### Backups (`--no-backup`)

By default the matcher copies the **input** file to `INPUT.backup_YYYYMMDD_HHMMSS` before processing.

```bash
# Skip creating that backup (useful for large inputs or CI)
python oclc_record_matcher.py -i big_list.csv -o big_list_with_oclc.xlsx --no-backup
```

#### Logging (`--log-level`, `--log-file`, `--no-api-logging`)

```bash
# Verbose application logging (row details, timing, etc.)
python oclc_record_matcher.py -i input.xlsx -o output.xlsx --log-level DEBUG

# Write logs to a specific file (still mirrors to console unless you rely on your shell)
python oclc_record_matcher.py -i input.xlsx -o output.xlsx --log-file logs/matcher-run.log

# Quieter HTTP: turn off per-request URL/body logging from the client
python oclc_record_matcher.py -i input.xlsx -o output.xlsx --no-api-logging

# Typical “deep debug” combo
python oclc_record_matcher.py -i input.xlsx -o output.xlsx --log-level DEBUG --log-file logs/debug.log
```

#### Combined switches

```bash
# MARCXML + no input backup + quieter API trace + custom log file
python oclc_record_matcher.py -i data/list.xlsx -o data/list_with_oclc.xlsx \
  --marcxml-output data/list_bibs.xml --no-backup --no-api-logging --log-file logs/oclc.log

# MARCXML-only run with DEBUG and no backup
python oclc_record_matcher.py -i data/list.csv --marcxml-output data/list.xml \
  --no-backup --log-level DEBUG
```

#### Help (`-h`, `--help`)

```bash
python oclc_record_matcher.py -h
```

#### MARC Extractor (`marc_extractor.py`)

| Option | Description | Default |
|--------|-------------|---------|
| `-i, --input` | Input MARC file path | Required |
| `-o, --output` | Output Excel file path | Required |
| `-h, --help` | Show help message | - |

```bash
python marc_extractor.py -i sampleData/MLN-cataloging-RFP-vendor-sample-batch.mrc -o extracted.xlsx
python marc_extractor.py -h
```

#### MARC Field Analyzer (`marc_field_analyzer.py`)

| Option | Description | Default |
|--------|-------------|---------|
| `-i, --input` | Input MARC file path | Required |
| `-o, --output` | Output Excel file path | Required |
| `-h, --help` | Show help message | - |

```bash
python marc_field_analyzer.py -i sampleData/MLN-cataloging-RFP-vendor-sample-batch.mrc -o analysis.xlsx
python marc_field_analyzer.py -h
```

## OCLC Search Methods

### ISBN Search
When ISBNs are available, the script uses optimized OR queries:
- **Query Structure**: `q=bn:{ISBN1} OR bn:{ISBN2} OR bn:{ISBN3}`
- **Format Parameters**: Automatically maps format types to `itemSubType` or `itemType`
- **Efficiency**: One API call per row, combining all available ISBNs

### Alternative Search
When no ISBNs are available, the script searches using bibliographic data:
- **Query Structure**: `q=te:{Title} AND au:{Author} AND pb:{Publisher}`
- **Date Filtering**: `datePublished={Publication Date}` (4-digit year)
- **Flexible Fields**: Gracefully handles missing fields

### Format Mapping
The script intelligently maps format types to OCLC API parameters:

| Format | Parameter Type | Value |
|--------|----------------|-------|
| `book-digital` | `itemSubType` | `book-digital` |
| `book-largeprint` | `itemSubType` | `book-largeprint` |
| `book-print` | `itemSubType` | `book-print` |
| `video` | `itemType` | `video` |
| `audiobook` | `itemType` | `audiobook` |
| `music` | `itemType` | `music` |

## Output

### OCLC Record Matcher Output

**Excel (when `-o` is used or defaulted):** A new workbook with:

- All original columns from the input sheet
- `matchingOCLCNumber` — OCLC number when a match is found
- `hasLCSHSubjects` — when **`--lcsh`** is set: whether LCSH-style subjects were detected; otherwise the column is left empty (default)
- `Other Identifier` — propagated when present in the source column mapping
- Empty cells / missing values where no match was found

**MARCXML (when `--marcxml-output` is set):** One UTF-8 XML document whose root is a [MARC 21 slim](http://www.loc.gov/MARC21/slim) `collection` containing one `record` per successfully retrieved OCLC number (order follows first appearance among matches; duplicates are collapsed). Failed lookups are skipped with log warnings.

If you omit `-o` but pass `--marcxml-output`, **no** Excel file is produced; only the MARCXML file (plus logs).

### MARC Extractor Output
The MARC extractor creates an Excel file with:
- `ISBN` - Extracted from 020$a
- `Title` - Combined from 245$a + 245$b (normalized)
- `Author` - From 100$a or 110$a
- `Publisher` - From 260$b or 264$b
- `Publication Date` - From 260$c or 264$c (normalized to 4-digit year)
- `Physical Description` - From 300 field
- `Format` - Determined from LDR 06 + 008 23 logic

### MARC Field Analyzer Output
The field analyzer creates an Excel report with:
- Field frequency statistics
- Most common MARC fields and subfields
- Control field analysis
- Leader position analysis

## Logging and Monitoring

### OCLC Record Matcher Logging
- **API request details**: URL, parameters, headers, query structure
- **Response details**: Status code, headers, response size, content (truncated where large)
- **Error logging**: Detailed error information with response bodies when available
- **Statistics**: API usage, success rates, optional LCSH counts when **`--lcsh`** is used, and manage-bib MARCXML fetch counts when **`--marcxml-output`** is used
- **Progress tracking**: Real-time progress with ETA during row processing
- **Row-by-row details**: Per-record processing messages at verbose log levels

### Log Levels
- **INFO**: Standard processing information
- **DEBUG**: Detailed debugging information
- **WARNING**: Non-critical issues
- **ERROR**: Critical errors

## Performance Features

### Optimized API Usage
- **Same-Row OR Processing**: Combines all ISBNs from the same row in single API calls
- **Rate Limiting**: Built-in delays to respect API rate limits
- **Error Recovery**: Continues processing even when individual API calls fail
- **Statistics Tracking**: Monitors API usage and success rates

### MARC Processing
- **Efficient Extraction**: Processes MARC records with minimal memory usage
- **Field Detection**: Automatically identifies and extracts relevant MARC fields
- **Format Logic**: Complex format determination based on MARC leader and control fields
- **Data Validation**: Validates and normalizes extracted data

## Error Handling

- **Network Issues**: Handles timeouts and connection problems gracefully
- **API Errors**: Detailed logging of API errors with response information
- **Data Validation**: Skips invalid or empty data with appropriate warnings
- **File Operations**: Handles file access and permission issues
- **MARC Processing**: Handles malformed MARC records gracefully

## Troubleshooting

### Common Issues

1. **Import Errors**: Make sure all dependencies are installed:
   ```bash
   uv pip install -e .
   ```

2. **File format issues**: 
   - Excel: Check column names (ISBN columns must include `ISBN` in the header)
   - CSV / TSV: Use UTF-8 encoding; ensure the correct extension (`.csv` vs `.tsv`) for delimiter detection
   - MARC: Ensure valid MARC21 binary (`.mrc` / `.marc`)

3. **Network Issues**: The OCLC matcher includes timeout handling and will log any connection problems

4. **API authentication and scopes**: 
   - Ensure `OCLC_API_KEY` and `OCLC_API_SECRET` are set in your `.env` file
   - Verify your credentials at the [OCLC Developer Network](https://www.oclc.org/developer/api/oclc-apis/worldcat-metadata-api.en.html)
   - Matching needs **`WorldCatMetadataAPI:match_bibs`** (and related search scopes) on the WSKey
   - **`--lcsh`** adds **`GET /worldcat/bibs/{oclcNumber}`**; ensure **`WorldCatMetadataAPI:read_bibs`** (or equivalent) is enabled if you use that flag
   - **`--marcxml-output`** calls **`GET /worldcat/manage/bibs/{oclcNumber}`** with `Accept: application/marcxml+xml`; the key needs **`WorldCatMetadataAPI:manage_bibs`** and/or **`WorldCatMetadataAPI:view_marc_bib`** per the [OpenAPI security requirements](https://developer.api.oclc.org/docs/wc-metadata/openapi-external-prod.yaml)

### Testing

**Test MARC field analysis:**
```bash
python marc_field_analyzer.py -i sampleData/MLN-cataloging-RFP-vendor-sample-batch.mrc -o test_analysis.xlsx
```

**Test MARC extraction:**
```bash
python marc_extractor.py -i sampleData/MLN-cataloging-RFP-vendor-sample-batch.mrc -o test_extraction.xlsx
```

**Test OCLC matching:**
```bash
python oclc_record_matcher.py -i sampleData/testRecords.xlsx -o test_matching.xlsx
```

## API Documentation

This project uses the **WorldCat Metadata API** as documented in the [OpenAPI specification](https://developer.api.oclc.org/docs/wc-metadata/openapi-external-prod.yaml).

### Key Endpoints

| Operation | HTTP | Path | Role in this project |
|-----------|------|------|----------------------|
| Brief search | `GET` | `/worldcat/search/brief-bibs` | Primary ISBN / title-author matching; returns `oclcNumber` among other brief fields |
| Full bib (JSON) | `GET` | `/worldcat/bibs/{oclcNumber}` | Optional LCSH-style subject check when you pass **`--lcsh`** (`Accept: application/json`) |
| Manage bib (MARCXML) | `GET` | `/worldcat/manage/bibs/{oclcNumber}` | Optional export: full bibliographic record as MARCXML (`Accept: application/marcxml+xml`) |

- **Authentication**: OAuth 2.0 client credentials
- **Base URL**: `https://metadata.api.oclc.org` (override with `OCLC_API_BASE_URL` if needed)
- **OAuth token URL**: `https://oauth.oclc.org/token`

The matcher favors **brief-bibs** for discovery because it is efficient for high-volume matching. Full **MARCXML** for holdings or cataloging workflows is fetched only when you pass **`--marcxml-output`**, using the **manage bibs** read operation described in the [OpenAPI specification](https://developer.api.oclc.org/docs/wc-metadata/openapi-external-prod.yaml).

### Authentication

The API uses OAuth 2.0 with client credentials:
- Access tokens are automatically obtained and refreshed
- Tokens are refreshed automatically on 401 Unauthorized responses
- Credentials are securely managed via environment variables

For more information, see the [WorldCat Metadata API documentation](https://www.oclc.org/developer/api/oclc-apis/worldcat-metadata-api.en.html).

## Development

### Using uv for Package Management

This project uses `uv` for fast Python package management. Key benefits:
- **Fast installation**: Significantly faster than pip
- **Reliable dependency resolution**: Better conflict resolution
- **Virtual environment management**: Automatic venv creation and management
- **Lock file support**: Reproducible builds with `uv.lock`

### Development Setup

1. **Create virtual environment:**
   ```bash
   uv venv
   source .venv/bin/activate  # On Windows: .venv\Scripts\activate
   ```

2. **Install in development mode:**
   ```bash
   uv pip install -e ".[dev]"
   ```

3. **Run scripts:**
   ```bash
   python oclc_record_matcher.py -i sampleData/recordsToMatch.xlsx -o output.xlsx
   ```

### Alternative Installation Methods

**Using pip (if uv is not available):**
```bash
pip install -r requirements.txt
```

**Using uv with specific Python version:**
```bash
uv pip install --python 3.12 -e .
```

## License

This project is licensed under the Apache License 2.0.

See the [LICENSE](LICENSE) file for the full license text.

Copyright 2024

Licensed under the Apache License, Version 2.0 (the "License");
you may not use this file except in compliance with the License.
You may obtain a copy of the License at

    http://www.apache.org/licenses/LICENSE-2.0

Unless required by applicable law or agreed to in writing, software
distributed under the License is distributed on an "AS IS" BASIS,
WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
See the License for the specific language governing permissions and
limitations under the License.
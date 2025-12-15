# OCLC Record Matcher

This project provides tools for processing bibliographic records and matching them with OCLC data. It includes scripts for extracting data from MARC files, analyzing MARC field usage, and matching records against the WorldCat Metadata API.

## Features

### OCLC Record Matching (`oclc_record_matcher.py`)
- **Excel File Processing**: Processes Excel files with ISBN columns
- **Multi-ISBN Support**: Automatically detects and processes multiple ISBN columns (XML ISBN, HC ISBN, PB ISBN, ePub ISBN, ePDF ISBN)
- **OR Query Optimization**: Combines all ISBNs from the same row in single API calls
- **Alternative Search**: When no ISBN is available, searches using title, author, publisher, and publication date
- **Format-Based Search**: Maps format types to appropriate OCLC API parameters
- **LCSH Detection**: Identifies Library of Congress Subject Headings in matched records
- **Comprehensive Logging**: Detailed API request/response logging with configurable verbosity

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
- **Secure Credential Management**: Uses environment variables via `python-dotenv` for API keys and secrets
- **OAuth 2.0 Authentication**: Automatic token management with client credentials flow
- **Smart Parameter Mapping**: Automatically maps format types to `itemType` or `itemSubType` parameters
- **Rate Limiting**: Built-in delays to respect API rate limits (configurable)
- **Error Handling**: Comprehensive error handling with automatic token refresh on 401 errors

## Files

### Main Scripts
- `oclc_record_matcher.py` - OCLC API matching for Excel files with ISBNs
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

2. **Install project dependencies:**
   ```bash
   uv pip install -e .
   ```
   
   Or for development with additional tools:
   ```bash
   uv pip install -e ".[dev]"
   ```

3. **Configure WorldCat Metadata API credentials:**
   
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

4. **Verify input file structure:**
   - **Excel files**: Script automatically detects ISBN columns by name
   - **MARC files**: Script automatically extracts standard MARC fields

## Usage

### Workflow Overview

The typical workflow involves two main steps:

1. **Extract MARC data to Excel** (if starting with MARC files)
2. **Match records with OCLC API** (for Excel files with ISBNs)

### Step 1: MARC Data Extraction

**Extract MARC data to Excel:**
```bash
python3 marc_extractor.py -i sampleData/MLN-cataloging-RFP-vendor-sample-batch.mrc -o extracted_data.xlsx
```

**Analyze MARC field usage:**
```bash
python3 marc_field_analyzer.py -i sampleData/MLN-cataloging-RFP-vendor-sample-batch.mrc -o field_analysis.xlsx
```

### Step 2: OCLC Record Matching

**Process Excel file with ISBNs:**
```bash
python3 oclc_record_matcher.py -i sampleData/recordsToMatch.xlsx -o matched_output.xlsx
```

**Use default files:**
```bash
python3 oclc_record_matcher.py
```

### Complete Workflow Example

**From MARC to OCLC-matched data:**
```bash
# Step 1: Extract MARC data
python3 marc_extractor.py -i sampleData/MLN-cataloging-RFP-vendor-sample-batch.mrc -o marc_data.xlsx

# Step 2: Match with OCLC API
python3 oclc_record_matcher.py -i marc_data.xlsx -o final_matched_data.xlsx
```

### Advanced Options

**Disable detailed API logging:**
```bash
python3 oclc_record_matcher.py -i input.xlsx -o output.xlsx --no-api-logging
```

**Process without creating backup:**
```bash
python3 oclc_record_matcher.py -i input.xlsx -o output.xlsx --no-backup
```

**Use different log level:**
```bash
python3 oclc_record_matcher.py -i input.xlsx -o output.xlsx --log-level DEBUG
```

### Command-Line Options

#### OCLC Record Matcher (`oclc_record_matcher.py`)

| Option | Description | Default |
|--------|-------------|---------|
| `-i, --input` | Input Excel file path | `sampleData/recordsToMatch.xlsx` |
| `-o, --output` | Output Excel file path | `input_file_with_oclc.xlsx` |
| `--no-backup` | Skip creating backup of input file | Create backup |
| `--log-level` | Set logging level (DEBUG, INFO, WARNING, ERROR) | `INFO` |
| `--log-file` | Custom log file path | `oclc_matcher.log` |
| `--no-api-logging` | Disable detailed API request/response logging | Enable API logging |
| `-h, --help` | Show help message | - |

#### MARC Extractor (`marc_extractor.py`)

| Option | Description | Default |
|--------|-------------|---------|
| `-i, --input` | Input MARC file path | Required |
| `-o, --output` | Output Excel file path | Required |
| `-h, --help` | Show help message | - |

#### MARC Field Analyzer (`marc_field_analyzer.py`)

| Option | Description | Default |
|--------|-------------|---------|
| `-i, --input` | Input MARC file path | Required |
| `-o, --output` | Output Excel file path | Required |
| `-h, --help` | Show help message | - |

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
The OCLC matcher creates a new Excel file with:
- All original data from the input file
- New `matchingOCLCNumber` column containing OCLC numbers for matched records
- New `hasLCSHSubjects` column indicating Library of Congress Subject Headings presence
- `None` values for records that couldn't be matched

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
- **API Request Details**: URL, parameters, headers, query structure
- **Response Details**: Status code, headers, response size, content (truncated)
- **Error Logging**: Detailed error information with response content
- **Statistics**: API usage statistics and success rates
- **Progress Tracking**: Real-time progress updates with ETA calculations
- **Row-by-Row Details**: Detailed processing information for each record

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

2. **File Format Issues**: 
   - Excel files: Check column names and file format
   - MARC files: Ensure proper MARC21 format

3. **Network Issues**: The OCLC matcher includes timeout handling and will log any connection problems

4. **API Authentication Issues**: 
   - Ensure `OCLC_API_KEY` and `OCLC_API_SECRET` are set in your `.env` file
   - Verify your credentials are correct at the [OCLC Developer Network](https://www.oclc.org/developer/api/oclc-apis/worldcat-metadata-api.en.html)
   - Check that your API key has the required scopes: `WorldCatMetadataAPI:read_bibs` and `WorldCatMetadataAPI:match_bibs`

### Testing

**Test MARC field analysis:**
```bash
python3 marc_field_analyzer.py -i sampleData/MLN-cataloging-RFP-vendor-sample-batch.mrc -o test_analysis.xlsx
```

**Test MARC extraction:**
```bash
python3 marc_extractor.py -i sampleData/MLN-cataloging-RFP-vendor-sample-batch.mrc -o test_extraction.xlsx
```

**Test OCLC matching:**
```bash
python3 oclc_record_matcher.py -i sampleData/testRecords.xlsx -o test_matching.xlsx
```

## API Documentation

This project uses the **WorldCat Metadata API** as documented in the [OpenAPI specification](https://developer.api.oclc.org/docs/wc-metadata/openapi-external-prod.yaml).

### Key Endpoints

- `GET /worldcat/search/bibs` - Search for bibliographic records
- **Authentication**: OAuth 2.0 Client Credentials flow
- **Base URL**: `https://metadata.api.oclc.org`
- **OAuth Token URL**: `https://oauth.oclc.org/token`

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
   python3 oclc_record_matcher.py -i sampleData/recordsToMatch.xlsx -o output.xlsx
   ```

### Alternative Installation Methods

**Using pip (if uv is not available):**
```bash
pip3 install -r requirements.txt
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
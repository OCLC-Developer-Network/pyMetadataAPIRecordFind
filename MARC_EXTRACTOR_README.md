# MARC Data Extractor

A Python script that extracts specific fields from MARC21 bibliographic records and exports them to an Excel spreadsheet.

## Features

- Extracts ISBN/ISSN, Title, Author, Publisher, and Publication Date from MARC records
- Handles both personal (100) and corporate (110) authors
- Supports both older (260) and newer (264) publication fields
- Creates formatted Excel output with auto-adjusted column widths
- Comprehensive logging and error handling
- Command-line interface with flexible options

## Extracted Fields

| MARC Field | Subfield | Description | Excel Column |
|------------|----------|-------------|--------------|
| 020 | $a | ISBN | ISBN |
| 245 | $a + $b | Title (combined) | Title (normalized, trailing punctuation removed) |
| 100 | $a + $d | Personal Author | Author |
| 110 | $a + $b | Corporate Author | Author |
| 260 | $b | Publisher (older format) | Publisher |
| 264 | $b | Publisher (newer format) | Publisher |
| 260 | $c | Publication Date (older format) | Publication Date (normalized to 4-digit year) |
| 264 | $c | Publication Date (newer format) | Publication Date (normalized to 4-digit year) |
| 300 | All subfields | Physical Description | Physical Description |
| LDR 06 + 008 23 | Combined logic | Format determination | Format |

## Title Normalization

The Title column is normalized by stripping trailing punctuation:

- **Strips trailing punctuation**: Removes /, :, ;, ., ,, =, + and other common MARC punctuation
- **Preserves internal punctuation**: Keeps colons and other punctuation within the title
- **Handles various formats**: 
  - "The Great Book /" → "The Great Book"
  - "Amazing Story: A Tale" → "Amazing Story: A Tale" (colon preserved)
  - "Fantastic Novel;" → "Fantastic Novel"
  - "Wonderful Book." → "Wonderful Book"
- **Maintains readability**: Titles remain clean and consistent for analysis

## Publication Date Normalization

The Publication Date column is normalized to 4-digit year format:

- **Strips trailing punctuation**: Removes periods, commas, semicolons, colons
- **Extracts 4-digit years**: Finds years in the 1900s and 2000s range
- **Handles various formats**: 
  - "2023" → "2023"
  - "2023." → "2023" 
  - "20231128" → "2023"
  - "2023-11-28" → "2023"
  - "Nov. 2023" → "2023"
  - "23" → "2023" (assumes 20xx)
- **Returns empty string**: For dates without recognizable year patterns

## Format Logic

The Format column is determined by combining LDR position 06 and 008 position 23 (extracted internally but not displayed as separate columns):

- **LDR 06 = "g"** → Format: `video`
- **LDR 06 = "i"** → Format: `audiobook`
- **LDR 06 = "j"** → Format: `music`
- **LDR 06 = "a" + 008 23 = "d"** → Format: `book-largeprint`
- **LDR 06 = "a" + 008 23 = "s"** → Format: `book-digital`
- **LDR 06 = "a" + 008 23 = any other value** → Format: `book-print`

## Requirements

- Python 3.7+
- pymarc
- openpyxl

## Installation

1. Install required dependencies:
```bash
pip install -r requirements.txt
```

Or install individually:
```bash
pip install pymarc openpyxl
```

## Usage

### Basic Usage

```bash
python marc_extractor.py -i input.mrc -o output.xlsx
```

### Advanced Usage

```bash
# With debug logging
python marc_extractor.py -i input.mrc -o output.xlsx --log-level DEBUG

# With log file
python marc_extractor.py -i input.mrc -o output.xlsx --log-file extraction.log

# Full example
python marc_extractor.py -i sampleData/MLN-cataloging-RFP-vendor-sample-batch.mrc -o extracted_data.xlsx --log-level INFO --log-file marc_extraction.log
```

### Command Line Options

| Option | Description | Required |
|--------|-------------|----------|
| `-i, --input` | Input MARC file (.mrc) | Yes |
| `-o, --output` | Output Excel file (.xlsx) | Yes |
| `--log-level` | Logging level (DEBUG, INFO, WARNING, ERROR) | No (default: INFO) |
| `--log-file` | Log file path | No |

## Example Output

The script creates an Excel file with the following columns:

| ISBN | Title | Author | Publisher | Publication Date | Physical Description | Format |
|------|-------|--------|-----------|------------------|---------------------|--------|
| 9781797155593 | Onlookers | Beattie, Ann/ Ryan, Allyson (NRT) | Blackstone Pub | 2023 | Compact Disc | audiobook |
| 9780735241305 | The Mystery Guest: A Maid Novel | Prose, Nita | Viking | 20231128 |  | book-print |
| 9781797176161 | Harbor Lights | Burke, James Lee | Blackstone Pub | 2024 |  | book-print |

## Sample Results

Based on the test with `MLN-cataloging-RFP-vendor-sample-batch.mrc`:

- **Total records processed**: 250
- **Records with ISBNs**: 187 (74.8%)
- **Records with Titles**: 250 (100.0%)
- **Records with Authors**: 199 (79.6%)
- **Records with Publishers**: 193 (77.2%)
- **Records with Dates**: 191 (76.4%)
- **Records with Physical Descriptions**: 125 (50.0%)
  - Examples: "1 videodisc (ca. 107 min.) : sd., col. ; 4 3/4 in.", "Compact Disc"
- **Records with Format values**: 250 (100.0%)
  - Format "audiobook": 29 records
  - Format "book-largeprint": 8 records
  - Format "book-print": 163 records
  - Format "music": 18 records
  - Format "video": 32 records

## Error Handling

The script includes comprehensive error handling:

- Invalid MARC records are logged and skipped
- Missing fields are handled gracefully (empty strings)
- File I/O errors are caught and reported
- Progress is logged every 100 records

## License

Apache License 2.0

## Author

AI Assistant

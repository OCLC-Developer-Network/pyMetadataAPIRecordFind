# MARC Field Analyzer

A Python script that analyzes MARC21 bibliographic files to determine the most common fields, subfields, and leader positions.

## Features

- Counts occurrences of all data fields (010 and above)
- Counts occurrences of all control fields (001-009)
- Analyzes subfield usage within data fields
- Analyzes leader position character distributions
- Creates comprehensive Excel output with multiple sheets
- Provides detailed console summary

## Requirements

- Python 3.7+
- pymarc
- openpyxl

## Installation

1. Install required dependencies:
```bash
pip install pymarc openpyxl
```

## Usage

### Basic Usage
```bash
python marc_field_analyzer.py -i input.mrc -o analysis.xlsx
```

### With Logging Options
```bash
python marc_field_analyzer.py -i input.mrc -o analysis.xlsx --log-level DEBUG
python marc_field_analyzer.py -i input.mrc -o analysis.xlsx --log-file analysis.log
```

### Command Line Options

| Option | Description | Required |
|--------|-------------|----------|
| `-i, --input` | Input MARC file (.mrc) | Yes |
| `-o, --output` | Output Excel file (.xlsx) | Yes |
| `--log-level` | Logging level (DEBUG, INFO, WARNING, ERROR) | No (default: INFO) |
| `--log-file` | Log file path | No |

## Output

The script creates an Excel file with the following sheets:

1. **Summary** - Overall statistics
2. **Field Counts** - Data field occurrences (sorted by frequency)
3. **Control Fields** - Control field occurrences
4. **Subfield Counts** - Subfield usage within data fields
5. **Leader Analysis** - Character distribution by leader position

## Example Results

Based on analysis of `MLN-cataloging-RFP-vendor-sample-batch.mrc` (250 records):

### Most Common Data Fields
1. **020** (ISBN) - 261 occurrences (10.52%)
2. **245** (Title) - 250 occurrences (10.08%)
3. **907** (Local field) - 250 occurrences (10.08%)
4. **998** (Local field) - 250 occurrences (10.08%)
5. **260** (Publication info) - 214 occurrences (8.63%)
6. **500** (General note) - 192 occurrences (7.74%)
7. **100** (Personal author) - 190 occurrences (7.66%)
8. **300** (Physical description) - 125 occurrences (5.04%)
9. **250** (Edition) - 98 occurrences (3.95%)
10. **024** (Other standard number) - 81 occurrences (3.26%)

### Control Fields
1. **008** (Fixed-length data) - 250 occurrences (35.36%)
2. **001** (Control number) - 198 occurrences (28.01%)
3. **005** (Date/time) - 105 occurrences (14.85%)
4. **007** (Physical description) - 82 occurrences (11.60%)
5. **003** (Control number identifier) - 72 occurrences (10.18%)

### Key Insights
- Every record has a 008 control field (100% coverage)
- Most records have ISBN (020) - 104.4% (some have multiple)
- Every record has a title (245) - 100% coverage
- Most records have publication info (260) - 85.6% coverage
- Most records have an author (100) - 76% coverage
- Local fields 907 and 998 appear in every record (vendor-specific)
- Physical description (300) appears in 50% of records

### Leader Analysis
- **Position 6** (Type of record): 171 records are "a" (Language material)
- **Position 6**: Also includes "g" (Projected medium), "i" (Sound recording), "j" (Music)
- **Position 23** (Form of material): Mostly "0" (Not specified)

## Field Definitions

| Field | Description |
|-------|-------------|
| 001 | Control number |
| 003 | Control number identifier |
| 005 | Date and time of latest transaction |
| 007 | Physical description fixed field |
| 008 | Fixed-length data elements |
| 010 | Library of Congress control number |
| 020 | International Standard Book Number |
| 024 | Other standard identifier |
| 040 | Cataloging source |
| 050 | Library of Congress call number |
| 082 | Dewey Decimal classification number |
| 100 | Personal name (main entry) |
| 110 | Corporate name (main entry) |
| 245 | Title statement |
| 246 | Varying form of title |
| 250 | Edition statement |
| 260 | Publication, distribution, etc. |
| 264 | Production, publication, distribution, manufacture, and copyright notice |
| 300 | Physical description |
| 306 | Playing time |
| 336 | Content type |
| 337 | Media type |
| 338 | Carrier type |
| 490 | Series statement |
| 500 | General note |
| 505 | Formatted contents note |
| 508 | Creation/production credits note |
| 511 | Participant or performer note |
| 520 | Summary, etc. |
| 521 | Target audience note |
| 538 | System details note |
| 546 | Language note |
| 600 | Subject added entry - Personal name |
| 650 | Subject added entry - Topical term |
| 700 | Added entry - Personal name |
| 710 | Added entry - Corporate name |
| 730 | Added entry - Uniform title |
| 830 | Series added entry - Uniform title |
| 907 | Local field (vendor-specific) |
| 998 | Local field (vendor-specific) |

## License

Apache License 2.0

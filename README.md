# XLSX Processor Utility

A Python utility to read financial transaction data from an XLSX file and generate calculated statistics based on predefined formulas.

## Features

- **Read SRT Data**: Processes transaction data from source sheets
- **Generate Statistics**: Calculates monthly financial metrics including:
  - End of Day balances
  - Credit/Debit summaries by group (BT, EXP, ZIH, DBT, etc.)
  - Loan payments and receipts
  - Net cash flows and ECS calculations
- **Export Results**: Generates formatted XLSX output files

## Installation

1. Ensure Python 3.7+ is installed
2. Install required dependencies:

```bash
pip3 install -r requirements.txt
```

## Usage

### Command Line Interface

```bash
python3 process_xlsx.py input_file.xlsx output_file.xlsx [options]
```

#### Options:
- `--start-date`: Start date for calculations (YYYY-MM-DD format, default: 2025-02-01)
- `--months`: Number of months to calculate (default: 6)
- `--srt-sheet`: Name of the SRT sheet (default: SRT)
- `--stat-sheet`: Name of the STAT sheet (default: STAT)

#### Examples:

```bash
# Basic usage
python3 process_xlsx.py "Client Stat.xlsx" "output.xlsx"

# Custom date range
python3 process_xlsx.py "Client Stat.xlsx" "output.xlsx" --start-date "2025-01-01" --months 12

# Custom sheet names
python3 process_xlsx.py "data.xlsx" "results.xlsx" --srt-sheet "Transactions" --stat-sheet "Summary"
```

### Python API

```python
from datetime import datetime
from xlsx_processor import XLSXProcessor

# Initialize processor
processor = XLSXProcessor()

# Read source data
processor.read_srt_data("input_file.xlsx")

# Generate statistics
start_date = datetime(2025, 2, 1)
processor.generate_stat_data(start_date, num_months=6)

# Write output
processor.write_output_file("output_file.xlsx")
```

## Data Structure

### SRT Sheet (Input)
Expected columns:
- `#`: Record number
- `GRP`: Group identifier (BT, EXP, ZIH, DBT, ecs, ecs pvt)
- `Date`: Transaction date
- `C2`: Transaction description/code
- `Debit`: Debit amount
- `Credit`: Credit amount
- `Balance`: Running balance
- `C1`: Transaction category
- `First Level Classification`: Transaction classification
- `ITA`: Additional field

### STAT Sheet (Output)
Calculated metrics include:
- **EOD monthly balance**: End of day balance for each month
- **Credit/Debit (BT)**: Business transaction summaries
- **Expense**: Monthly expense totals
- **ZIH Cr/Dr**: ZIH transaction summaries
- **DBT Cr/Dr**: DBT transaction summaries
- **Monthly Loan payments**: ECS bank and private loan payments
- **Loan received**: ECS loan receipts
- **Net calculations**: Various derived financial metrics

## Calculation Logic

The utility implements the following key formulas:

1. **Monthly Summations**: Uses SUMIFS-like logic to filter by:
   - Group (GRP column)
   - Date ranges (monthly boundaries)
   - Additional conditions (e.g., excluding RTN records)

2. **Date Handling**: 
   - Calculates month-end dates using calendar logic
   - Handles month boundaries properly for filtering

3. **Derived Metrics**:
   - Net ZIH: ZIH Credit - ZIH Debit
   - Cash Flow: Credit - Debit - Expense
   - Net ECS: ECS Credit - ECS Debit

## Files

- `xlsx_processor.py`: Main processor class with calculation logic
- `process_xlsx.py`: Command-line interface
- `requirements.txt`: Python dependencies
- `README.md`: This documentation

## Error Handling

The utility includes comprehensive error handling for:
- File not found errors
- Invalid sheet names
- Data type conversion issues
- Missing columns

## Testing

Test the utility with the provided sample file:

```bash
python3 process_xlsx.py "Client Stat.xlsx" "test_output.xlsx"
```

This will process the sample data and generate a test output file to verify functionality.

## Dependencies

- `pandas>=2.0.0`: Data manipulation and analysis
- `openpyxl>=3.1.0`: Excel file reading/writing
- `python-dateutil>=2.8.0`: Date manipulation utilities

## License

This utility is provided as-is for financial data processing purposes.

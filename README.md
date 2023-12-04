# Joiner

## Overview

The Joiner is a Python script designed to accumulate phone numbers from multiple files, format them uniformly, remove duplicates, and support various file formats including .csv, .xls, and .xlsx. This script simplifies the process of consolidating phone numbers from different sources into a clean and organized list.

## Features

- **Accumulation**: Gather phone numbers from multiple files.
- **Formatting**: Ensure uniform formatting of phone numbers.
- **Duplicate Removal**: Remove duplicate phone numbers from the accumulated list.
- **File Format Support**: Accepts .csv, .xls, and .xlsx file formats.

## Prerequisites

- Python 3.x
- Required Python packages: click, openpyxl, xlrd (install them using `pip install click openpyxl xlrd`)

## Usage

1. Clone this repository:

   ```bash
   git clone https://github.com/Jlgtri/joiner.git
   ```

2. Navigate to the project directory:

   ```bash
   cd joiner
   ```

3. Place your input files (in .csv, .xls, or .xlsx format) in the `input_files` directory.

4. Run the script:

   ```bash
   python joiner.py
   ```

5. The processed output will be stored in the `output` directory.

## Configuration

Usage: joiner.py [OPTIONS] [INPUT]...

Options:

- -o, --output FILE The path to output data to. Defaults to numbers.txt.
- -c, --column TEXT The column (e.g. A, B, AB) to process from table.
- --all / --no-all If all suggested columns should be used from table (>90% valid).
- --sort / --no-sort Whether export data should be sorted.
- -l, --logging TEXT The logging level used to print information.
- --help Show this message and exit.

## Example

Suppose you have the following input files:

- `contacts.csv`:

  ```csv
  Name,Phone
  John Doe,123-456-7890
  Jane Smith,9876543210
  ```

- `more_contacts.xlsx`:

  | Name        | Phone        |
  | ----------- | ------------ |
  | Alice Lee   | 555-1234     |
  | Bob Johnson | 987-654-3210 |

After running the script, the accumulated and formatted output in `numbers.txt` might look like:

```txt
1234567890
9876543210
5551234
```

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

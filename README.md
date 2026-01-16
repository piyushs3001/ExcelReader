# Excel Importer

A lightweight PHP package to import Excel files (.xlsx and .xls) with data type support.

[![PHP Version](https://img.shields.io/badge/php-%3E%3D7.2-blue.svg)](https://php.net)
[![License](https://img.shields.io/badge/license-MIT-green.svg)](LICENSE)

## Features

- **Read .xlsx files** (Office Open XML format - Excel 2007+)
- **Read .xls files** (BIFF format - Excel 97-2003)
- **Data type detection**: strings, numbers, dates, booleans
- **Date conversion**: Automatic Excel date to PHP DateTime conversion
- **No external dependencies**: Pure PHP implementation
- **PHP 7.2+ compatible**
- **Memory efficient**: Reads files on demand
- **Simple API**: Easy to use fluent interface

## Requirements

- PHP >= 7.2
- ext-zip (for .xlsx files)
- ext-xml or ext-simplexml (for XML parsing)

## Installation

Install via Composer:

```bash
composer require piyush/excel-importer
```

## Quick Start

```php
<?php

require 'vendor/autoload.php';

use Piyush\ExcelImporter\ExcelReader;

// Load an Excel file (auto-detects format)
$excel = ExcelReader::load('/path/to/file.xlsx');

// Get all sheet names
$sheetNames = $excel->getSheetNames();
print_r($sheetNames);

// Get a specific sheet by name or index
$sheet = $excel->getSheet('Sheet1');
// or
$sheet = $excel->getSheet(0);

// Iterate through rows
foreach ($sheet->getRows() as $rowIndex => $row) {
    foreach ($row->getCells() as $colIndex => $cell) {
        $value = $cell->getValue();        // mixed (string, int, float, DateTime, bool, null)
        $type = $cell->getType();          // 'string', 'number', 'date', 'boolean', 'empty'
        $formatted = $cell->getFormattedValue();  // Always returns string

        echo "Row {$rowIndex}, Col {$colIndex}: {$formatted} ({$type})\n";
    }
}

// Quick export to array
$data = $sheet->toArray();
print_r($data);
```

## API Reference

### ExcelReader (Main Entry Point)

```php
use Piyush\ExcelImporter\ExcelReader;

// Load entire file
$workbook = ExcelReader::load('/path/to/file.xlsx');

// Load only specific sheets
$workbook = ExcelReader::loadSheets('/path/to/file.xlsx', ['Sheet1', 'Data']);

// Get sheet names without loading data
$names = ExcelReader::getSheetNames('/path/to/file.xlsx');

// Check if file can be read
$canRead = ExcelReader::canRead('/path/to/file.xlsx');
```

### Workbook

```php
// Get sheet by name
$sheet = $workbook->getSheet('Sheet1');

// Get sheet by index (0-based)
$sheet = $workbook->getSheet(0);

// Get first sheet
$sheet = $workbook->getFirstSheet();

// Get all sheets
$sheets = $workbook->getSheets();

// Get sheet names
$names = $workbook->getSheetNames();

// Get sheet count
$count = $workbook->getSheetCount();
// or
$count = count($workbook);

// Check if sheet exists
$exists = $workbook->hasSheet('Sheet1');

// Iterate through sheets
foreach ($workbook as $sheet) {
    echo $sheet->getName();
}
```

### Worksheet

```php
// Get sheet name and index
$name = $sheet->getName();
$index = $sheet->getIndex();

// Get a specific row (1-based index)
$row = $sheet->getRow(1);

// Get all rows
$rows = $sheet->getRows();

// Get a specific cell by coordinate
$cell = $sheet->getCell('A1');
$cell = $sheet->getCell('B2');

// Get dimensions
$highestRow = $sheet->getHighestRow();
$highestCol = $sheet->getHighestColumn();        // Returns index
$highestColLetter = $sheet->getHighestColumnLetter();  // Returns letter (e.g., 'D')

// Convert to 2D array
$data = $sheet->toArray();

// Convert to associative array (first row as headers)
$data = $sheet->toAssociativeArray();

// Iterate through rows
foreach ($sheet as $row) {
    // ...
}
```

### Row

```php
// Get row index (1-based)
$index = $row->getRowIndex();

// Get cell by column index (0-based)
$cell = $row->getCell(0);  // First column

// Get cell by column letter
$cell = $row->getCellByColumn('A');

// Get all cells
$cells = $row->getCells();

// Convert to array
$values = $row->toArray();

// Check if row is empty
$isEmpty = $row->isEmpty();

// Iterate through cells
foreach ($row as $cell) {
    // ...
}
```

### Cell

```php
// Get value (mixed type)
$value = $cell->getValue();

// Get type ('string', 'number', 'date', 'boolean', 'empty', 'formula')
$type = $cell->getType();

// Get formatted value (always string)
$formatted = $cell->getFormattedValue();

// Get coordinate
$coord = $cell->getCoordinate();  // e.g., 'A1'

// Type checking
$cell->isEmpty();      // true if empty
$cell->isString();     // true if string
$cell->isNumber();     // true if number
$cell->isDate();       // true if date
$cell->isBoolean();    // true if boolean
```

## Cell Types

| Type | Description | PHP Value Type |
|------|-------------|----------------|
| `string` | Text values | `string` |
| `number` | Integer or float | `int` or `float` |
| `date` | Excel date | `DateTime` |
| `boolean` | TRUE/FALSE | `bool` |
| `empty` | Blank cell | `null` |
| `formula` | Calculated value | `mixed` |

## Working with Dates

Excel stores dates as numeric values (days since 1900-01-01). This package automatically converts them to PHP `DateTime` objects when the cell has a date format.

```php
$cell = $sheet->getCell('A1');

if ($cell->isDate()) {
    $dateTime = $cell->getValue();  // DateTime object
    echo $dateTime->format('Y-m-d'); // 2024-01-15
}
```

You can also use the DateHelper directly:

```php
use Piyush\ExcelImporter\Helpers\DateHelper;

// Convert Excel date to DateTime
$dateTime = DateHelper::excelToDateTime(44941);  // 2023-01-15

// Convert to timestamp
$timestamp = DateHelper::excelToTimestamp(44941);

// Convert to formatted string
$formatted = DateHelper::excelToFormatted(44941, 'Y-m-d');
```

## Working with Headers

Convert rows to associative arrays using the first row as headers:

```php
$data = $sheet->toAssociativeArray();

// Result:
// [
//     ['Name' => 'John', 'Email' => 'john@example.com', 'Age' => 30],
//     ['Name' => 'Jane', 'Email' => 'jane@example.com', 'Age' => 25],
// ]

foreach ($data as $record) {
    echo $record['Name'];
    echo $record['Email'];
}
```

## Example: Import to Database

```php
use Piyush\ExcelImporter\ExcelReader;

$excel = ExcelReader::load('users.xlsx');
$sheet = $excel->getFirstSheet();

// Skip header row
$rows = $sheet->toAssociativeArray();

foreach ($rows as $row) {
    $db->insert('users', [
        'name' => $row['Name'],
        'email' => $row['Email'],
        'created_at' => $row['Join Date']->format('Y-m-d H:i:s')
    ]);
}
```

## Limitations

- **Read-only**: This package only reads Excel files, no writing/export capability
- **No formula calculation**: Only cached/calculated values are read, not the formulas themselves
- **Basic merged cell support**: Merged cells return value only in the top-left cell
- **No image/chart extraction**: Only cell data is extracted
- **No password protection support**: Cannot read encrypted files

## License

MIT License. See [LICENSE](LICENSE) for details.

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## Support

If you encounter any issues, please [open an issue](https://github.com/piyush/excel-importer/issues) on GitHub.

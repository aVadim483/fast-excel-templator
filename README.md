# FastExcelTemplator

This is a part of the **FastExcelPhp Project**

This library **FastExcelTemplator** can generate Excel-compatible spreadsheets in XLSX format (Office 2007+) from XLSX templates, 
very quickly and with minimal memory usage.

## Introduction

This library is designed to be lightweight, super-fast and requires minimal memory usage.

**Features**

* Supports XLSX format only (Office 2007+) with multiple worksheets
* Transfers from templates to target spreadsheets styles, images, notes
* Replaces the entire cell values and substrings
* You can use any row from a template as row template to insert and replace a row with new values
* The library can read styling options of cells - formatting patterns, colors, borders, fonts, etc.

## Installation

Use `composer` to install **FastExcelTemplator** into your project:

```
composer require avadim/fast-excel-templator
```

Also, you can download package and include autoload file of the library:
```php
require 'path/to/fast-excel-templator/src/autoload.php';
```

## Usage

You can find more examples in */demo* folder



## Excel

Excel::template($template, $output);

* fill(array $replacement) - Replaces the entire cell value for all sheets
* replace(array $replacement) - Replaces a substring in a cell for all sheets
* save()

## Sheet

* fill(array $replacement) - Replaces the entire cell value for the sheet
* replace(array $replacement) - Replaces a substring in a cell for the sheet
* getRowTemplate($rowNumber) - Gets template from the row
* insertRow($rowNumber, $rowTemplate, ?array $cellData = [])
* replaceRow($rowNumber, $rowTemplate, ?array $cellData = [])
* insertRowAfterLast($rowTemplate, ?array $cellData = [])

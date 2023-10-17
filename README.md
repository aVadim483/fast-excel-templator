# FastExcelTemplator

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

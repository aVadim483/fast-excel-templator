# FastExcelTemplator

## Excel

Excel::template($template, $output);

* fill() - Replaces the entire cell value for all sheets
* replace() - Replaces a substring in a cell for all sheets
* save()

## Sheet

* fill() - Replaces the entire cell value for the sheet
* replace() - Replaces a substring in a cell for the sheet
* getRowTemplate($rowNumber) - Gets template from the row
* insertRow()
* replaceRow()
* insertRowAfterLast()

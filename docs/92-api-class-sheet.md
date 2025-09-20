# Class \avadim\FastExcelTemplator\Sheet

---

* [__construct()](#__construct)
* [isActive()](#isactive)
* [cloneRow()](#clonerow)
* [fill()](#fill) -- Replacement for the entire cell value
* [firstCol()](#firstcol)
* [firstRow()](#firstrow)
* [hasDrawings()](#hasdrawings)
* [hasImage()](#hasimage) -- Returns TRUE if the cell contains an image
* [isHidden()](#ishidden)
* [getImageBlob()](#getimageblob) -- Returns an image from the cell as a blob (if exists) or null
* [insertRow()](#insertrow)
* [nextRow()](#nextrow) -- Read cell values row by row, returns either an array of values or an array of arrays
* [replace()](#replace) -- Replacement for substrings in a cell
* [replaceRow()](#replacerow)
* [getRowHeight()](#getrowheight) -- Returns row height for a specific row number.
* [rows()](#rows) -- Transfers all rows from template to output
* [getRowTemplate()](#getrowtemplate)
* [getRowTemplates()](#getrowtemplates)
* [skipRows()](#skiprows) -- Skip rows from template
* [skipRowsUntil()](#skiprowsuntil) -- Skip rows from template
* [transferRows()](#transferrows) -- Transfers rows from template to output
* [transferRowsUntil()](#transferrowsuntil) -- Transfers rows from template to output
* [isVisible()](#isvisible)
* [writeCell()](#writecell) -- Write value to the current cell and move a pointer to the next cell in the row
* [writeRow()](#writerow) -- Write values to the current row

---

## __construct()

---

```php
public function __construct($sheetName, $sheetId, $file, $path, $excel)
```


### Parameters

* `$sheetName`
* `$sheetId`
* `$file`
* `$path`
* `$excel`

---

## isActive()

---

```php
public function isActive(): bool
```


### Parameters

_None_

---

## cloneRow()

---

```php
public function cloneRow(?array $cellData = [])
```


### Parameters

* `array|null $cellData`

---

## fill()

---

```php
public function fill(array $params): Sheet
```
_Replacement for the entire cell value_

### Parameters

* `array $params`

---

## firstCol()

---

```php
public function firstCol(): string
```


### Parameters

_None_

---

## firstRow()

---

```php
public function firstRow(): int
```


### Parameters

_None_

---

## hasDrawings()

---

```php
public function hasDrawings(): bool
```


### Parameters

_None_

---

## hasImage()

---

```php
public function hasImage(string $cell): bool
```
_Returns TRUE if the cell contains an image_

### Parameters

* `string $cell`

---

## isHidden()

---

```php
public function isHidden(): bool
```


### Parameters

_None_

---

## getImageBlob()

---

```php
public function getImageBlob(string $cell): ?string
```
_Returns an image from the cell as a blob (if exists) or null_

### Parameters

* `string $cell`

---

## insertRow()

---

```php
public function insertRow($row, ?array $cellData = []): Sheet
```


### Parameters

* `array|RowTemplateCollection|RowTemplate $row`
* `array|null $cellData`

---

## nextRow()

---

```php
public function nextRow($columnKeys, ?int $resultMode = null, 
                        ?bool $styleIdxInclude = null, 
                        ?int $rowLimit = 0): ?Generator
```
_Read cell values row by row, returns either an array of values or an array of arraysnextRow(..., ...) : <rowNum> => \[<colNum1> => <value1>, <colNum2> => <value2>, ...]nextRow(..., ..., true) : <rowNum> => \[<colNum1> => \['v' => <value1>, 's' => <style1>], <colNum2> => \['v' => <value2>, 's' => <style2>], ...]_

### Parameters

* `array|bool|int|null $columnKeys`
* `int|null $resultMode`
* `bool|null $styleIdxInclude`
* `int|null $rowLimit`

---

## replace()

---

```php
public function replace(array $params): Sheet
```
_Replacement for substrings in a cell_

### Parameters

* `array $params`

---

## replaceRow()

---

```php
public function replaceRow($row, ?array $cellData = [])
```


### Parameters

* `mixed $row`
* `array|null $cellData`

---

## getRowHeight()

---

```php
public function getRowHeight(int $rowNumber): ?float
```
_Returns row height for a specific row number._

### Parameters

* `int $rowNumber`

---

## rows()

---

```php
public function rows($callback): Sheet
```
_Transfers all rows from template to output_

### Parameters

* `mixed $callback` -- function ($rowNum, $rowData)

---

## getRowTemplate()

---

```php
public function getRowTemplate(int $rowNumber, 
                               ?bool $savePointerPosition = false): RowTemplateCollection
```


### Parameters

* `int $rowNumber`
* `bool|null $savePointerPosition`

---

## getRowTemplates()

---

```php
public function getRowTemplates(int $rowNumberMin, int $rowNumberMax, 
                                ?bool $savePointerPosition = false): RowTemplateCollection
```


### Parameters

* `int $rowNumberMin`
* `int $rowNumberMax`
* `bool|null $savePointerPosition`

---

## skipRows()

---

```php
public function skipRows(?int $countRows = null): Sheet
```
_Skip rows from template_

### Parameters

* `int|null $countRows` -- Number of rows

---

## skipRowsUntil()

---

```php
public function skipRowsUntil(?int $maxRowNum = null): Sheet
```
_Skip rows from template_

### Parameters

* `int|null $maxRowNum` -- Max row of template

---

## transferRows()

---

```php
public function transferRows(?int $countRows = null, $callback): Sheet
```
_Transfers rows from template to output_

### Parameters

* `int|null $countRows` -- Number of rows
* `$callback`

---

## transferRowsUntil()

---

```php
public function transferRowsUntil(?int $maxRowNum = null, $callback): Sheet
```
_Transfers rows from template to output_

### Parameters

* `int|null $maxRowNum` -- Max row of template
* `$callback`

---

## isVisible()

---

```php
public function isVisible(): bool
```


### Parameters

_None_

---

## writeCell()

---

```php
public function writeCell($value, ?array $styles = null): Sheet
```
_Write value to the current cell and move a pointer to the next cell in the row_

### Parameters

* `mixed $value`
* `array|null $styles`

---

## writeRow()

---

```php
public function writeRow(array $rowValues = [], ?array $rowStyle = null, 
                         ?array $cellStyles = null): Sheet
```
_Write values to the current row_

### Parameters

* `array $rowValues` -- Values of cells
* `array|null $rowStyle` -- Style applied to the entire row
* `array|null $cellStyles` -- Styles of specified cells in the row

---


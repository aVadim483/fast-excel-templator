# Class \avadim\FastExcelTemplator\RowTemplate

---

* [__construct()](#__construct) – RowTemplate constructor
* [addCell()](#addcell) – Add a cell to the row
* [appendCell()](#appendcell) – Append a source cell to the end of the row
* [attribute()](#attribute) – Source row attribute
* [cells()](#cells) – Get all cells of the row
* [cloneCell()](#clonecell) – Clone a cell
* [current()](#current) – Return the current element
* [getAttributes()](#getattributes) – All source row attributes
* [getSheetTemplate()](#getsheettemplate) – Get sheet template
* [getValue()](#getvalue) – Returns value of cell
* [getValues()](#getvalues) – Get values of all cells of the row
* [key()](#key) – Return the key of the current element
* [next()](#next) – Move forward to next element
* [removeCell()](#removecell) – Mark a cell as removed
* [removeCells()](#removecells) – Mark multiple cells as removed
* [rewind()](#rewind) – Rewind the Iterator to the first element
* [rowHeight()](#rowheight) – Source row height
* [rowNumber()](#rownumber) – Source row number
* [setAttributes()](#setattributes) – Set row attributes
* [setCellStyle()](#setcellstyle) – Set cell style
* [setRowStyle()](#setrowstyle) – Set row style
* [setSheetTemplate()](#setsheettemplate) – Set sheet template
* [setValue()](#setvalue) – Set cell value
* [setValues()](#setvalues) – Set multiple cell values
* [valid()](#valid) – Checks if current position is valid
* [withValue()](#withvalue) – Assign a value to the last added cell
* [withValues()](#withvalues) – Assign values to the last added cells

---

## __construct()

---

```php
public function __construct(?array $cellData = [])
```
_RowTemplate constructor_

### Parameters

* `array|null $cellData`

---

## addCell()

---

```php
public function addCell(string $colLetter, $value, $style): RowTemplate
```
_Add a cell to the row_

### Parameters

* `string $colLetter`
* `mixed $value`
* `mixed|null $style`

---

### Examples

```php
$row->addCell('G', $value);
```


---

## appendCell()

---

```php
public function appendCell(?string $colSource = null, 
                           ?int $number = null): RowTemplate
```
_Append a source cell to the end of the row_

### Parameters

* `string|null $colSource`
* `int|null $number`

---

### Examples

```php
$row->appendCell('E'); // Append cell like E to the end of the row
$row->appendCell('E', 2); // Append two cells like E to the end of the row
```


---

## attribute()

---

```php
public function attribute($name): ?string
```
_Source row attribute_

### Parameters

* `string $name`

---

## cells()

---

```php
public function cells(): array
```
_Get all cells of the row_

### Parameters

_None_

---

## cloneCell()

---

```php
public function cloneCell(string $colSource, $colTarget, 
                          ?bool $checkMerge = false): RowTemplate
```
_Clone a cell_

### Parameters

* `string $colSource`
* `string|string[] $colTarget`
* `bool|null $checkMerge`

---

## current()

---

```php
public function current(): mixed
```
_Return the current element_

### Parameters

_None_

---

## getAttributes()

---

```php
public function getAttributes(): array
```
_All source row attributes_

### Parameters

_None_

---

## getSheetTemplate()

---

```php
public function getSheetTemplate(): ?SheetTemplate
```
_Get sheet template_

### Parameters

_None_

---

## getValue()

---

```php
public function getValue(string $colLetter): mixed|null
```
_Returns value of cell_

### Parameters

* `string $colLetter`

---

## getValues()

---

```php
public function getValues(): array
```
_Get values of all cells of the row_

### Parameters

_None_

---

## key()

---

```php
public function key(): mixed
```
_Return the key of the current element_

### Parameters

_None_

---

## next()

---

```php
public function next(): void
```
_Move forward to next element_

### Parameters

_None_

---

## removeCell()

---

```php
public function removeCell(string $col): RowTemplate
```
_Mark a cell as removed_

### Parameters

* `string $col`

---

### Examples

```php
$row->removeCell('G');
```


---

## removeCells()

---

```php
public function removeCells(array $cols): RowTemplate
```
_Mark multiple cells as removed_

### Parameters

* `string[] $cols`

---

### Examples

```php
$row->removeCells(['G', 'H']);
```


---

## rewind()

---

```php
public function rewind(): void
```
_Rewind the Iterator to the first element_

### Parameters

_None_

---

## rowHeight()

---

```php
public function rowHeight(): ?string
```
_Source row height_

### Parameters

_None_

---

## rowNumber()

---

```php
public function rowNumber(): ?int
```
_Source row number_

### Parameters

_None_

---

## setAttributes()

---

```php
public function setAttributes(array $attributes): RowTemplate
```
_Set row attributes_

### Parameters

* `array $attributes`

---

## setCellStyle()

---

```php
public function setCellStyle(string $colLetter, $style): RowTemplate
```
_Set cell style_

### Parameters

* `string $colLetter`
* `mixed $style`

---

## setRowStyle()

---

```php
public function setRowStyle($style): RowTemplate
```
_Set row style_

### Parameters

* `mixed $style`

---

## setSheetTemplate()

---

```php
public function setSheetTemplate(avadim\FastExcelTemplator\SheetTemplate $sheetTemplate): RowTemplate
```
_Set sheet template_

### Parameters

* `SheetTemplate $sheetTemplate`

---

## setValue()

---

```php
public function setValue(string $colLetter, $value): RowTemplate
```
_Set cell value_

### Parameters

* `string $colLetter`
* `mixed $value`

---

### Examples

```php
$row->setValue('G', '=SUM(A1:A10)');
```


---

## setValues()

---

```php
public function setValues(array $values): RowTemplate
```
_Set multiple cell values_

### Parameters

* `array $values`

---

### Examples

```php
$row->setValues(['G' => '=SUM(A1:A10)', 'H' => 'value']);
```


---

## valid()

---

```php
public function valid(): bool
```
_Checks if current position is valid_

### Parameters

_None_

---

## withValue()

---

```php
public function withValue($value): RowTemplate
```
_Assign a value to the last added cell_

### Parameters

* `mixed $value`

---

### Examples

```php
$row->appendCell('E')->withValue('value');
```


---

## withValues()

---

```php
public function withValues(array $values): RowTemplate
```
_Assign values to the last added cells_

### Parameters

* `array $values`

---

### Examples

```php
$row->appendCell('E', 2)->withValues(['value1', 'value2']);
```


---


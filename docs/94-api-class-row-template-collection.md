# Class \avadim\FastExcelTemplator\RowTemplateCollection

---

* [__construct()](#__construct) – RowTemplateCollection constructor
* [addRowTemplate()](#addrowtemplate) – Add row template to collection
* [cloneCell()](#clonecell) – Clone cell in all row templates of the collection
* [current()](#current) – Return the current element
* [delRowTemplate()](#delrowtemplate) – Delete row template from collection
* [key()](#key) – Return the key of the current element
* [next()](#next) – Move forward to next element, wrapping around at the end
* [rewind()](#rewind) – Rewind the Iterator to the first element
* [setSheet()](#setsheet) – Set sheet template
* [valid()](#valid) – Checks if current position is valid

---

## __construct()

---

```php
public function __construct(?array $rowData = [], 
                            ?SheetTemplate $sheet = null)
```
_RowTemplateCollection constructor_

### Parameters

* `array|null $rowData`
* `SheetTemplate|null $sheet` – Sheet the templates were captured from; some operations (cloneCell) need it, it can also be set later by setSheet()

---

## addRowTemplate()

---

```php
public function addRowTemplate($row, ?int $rowNum = 0): void
```
_Add row template to collection_

### Parameters

* `RowTemplate $row`
* `int|null $rowNum`

---

## cloneCell()

---

```php
public function cloneCell(string $colSource, $colTarget, 
                          ?bool $checkMerge = false): RowTemplateCollection
```
_Clone cell in all row templates of the collection_

### Parameters

* `string $colSource`
* `string|array $colTarget`
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

## delRowTemplate()

---

```php
public function delRowTemplate(int $rowNum): void
```
_Delete row template from collection_

### Parameters

* `int $rowNum`

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
public function next(): mixed
```
_Move forward to next element, wrapping around at the end_

_The wrap-around is what insertRow() relies on: calling next() repeatedly cycles throughthe captured templates (e.g. two rows for zebra striping) as long as rows are inserted._

### Parameters

_None_

---

## rewind()

---

```php
public function rewind(): mixed
```
_Rewind the Iterator to the first element_

### Parameters

_None_

---

## setSheet()

---

```php
public function setSheet($sheet)
```
_Set sheet template_

### Parameters

* `SheetTemplate $sheet`

---

## valid()

---

```php
public function valid(): bool
```
_Checks if current position is valid_

_Since next() deliberately wraps around, the position alone can never end a foreach.The counter reset by rewind() is what makes a foreach walk the collection exactly once._

### Parameters

_None_

---


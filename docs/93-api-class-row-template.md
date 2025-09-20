# Class \avadim\FastExcelTemplator\RowTemplate

---

* [__construct()](#__construct)
* [addCell()](#addcell)
* [appendCell()](#appendcell)
* [attribute()](#attribute) -- Source row attribute
* [getAttributes()](#getattributes) -- All source row attributes
* [setAttributes()](#setattributes)
* [cells()](#cells)
* [cloneCell()](#clonecell)
* [current()](#current)
* [key()](#key)
* [next()](#next)
* [removeCell()](#removecell)
* [removeCells()](#removecells)
* [rewind()](#rewind)
* [rowHeight()](#rowheight) -- Source row height
* [rowNumber()](#rownumber) -- Source row number
* [valid()](#valid)
* [getValue()](#getvalue)
* [setValue()](#setvalue)
* [getValues()](#getvalues)
* [setValues()](#setvalues)
* [withValue()](#withvalue)
* [withValues()](#withvalues)

---

## __construct()

---

```php
public function __construct(?array $cellData = [])
```


### Parameters

* `$cellData`

---

## addCell()

---

```php
public function addCell(string $colLetter, $cell): RowTemplate
```


### Parameters

* `string $colLetter`
* `$cell`

---

## appendCell()

---

```php
public function appendCell(?string $colSource = null, 
                           ?int $number = null): RowTemplate
```


### Parameters

* `string|null $colSource`
* `int|null $number`

---

## attribute()

---

```php
public function attribute($name): ?string
```
_Source row attribute_

### Parameters

* `$name`

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

## setAttributes()

---

```php
public function setAttributes(array $attributes): RowTemplate
```


### Parameters

* `array $attributes`

---

## cells()

---

```php
public function cells(): array
```


### Parameters

_None_

---

## cloneCell()

---

```php
public function cloneCell(string $colSource, $colTarget, 
                          ?bool $checkMerge = false): RowTemplate
```


### Parameters

* `string $colSource`
* `string|string[] $colTarget`
* `bool|null $checkMerge`

---

## current()

---

```php
public function current()
```


### Parameters

_None_

---

## key()

---

```php
public function key()
```


### Parameters

_None_

---

## next()

---

```php
public function next()
```


### Parameters

_None_

---

## removeCell()

---

```php
public function removeCell(string $col): RowTemplate
```


### Parameters

* `$col`

---

## removeCells()

---

```php
public function removeCells(array $cols): RowTemplate
```


### Parameters

* `string[] $cols`

---

## rewind()

---

```php
public function rewind()
```


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

## valid()

---

```php
public function valid(): bool
```


### Parameters

_None_

---

## getValue()

---

```php
public function getValue(string $colLetter): mixed|null
```


### Parameters

* `string $colLetter`

---

## setValue()

---

```php
public function setValue(string $colLetter, $value): RowTemplate
```


### Parameters

* `string $colLetter`
* `mixed $value`

---

## getValues()

---

```php
public function getValues(): array
```


### Parameters

_None_

---

## setValues()

---

```php
public function setValues(array $values): RowTemplate
```


### Parameters

* `array $values`

---

## withValue()

---

```php
public function withValue($value): RowTemplate
```


### Parameters

* `$value`

---

## withValues()

---

```php
public function withValues(array $values): RowTemplate
```


### Parameters

* `array $values`

---


# Class \avadim\FastExcelTemplator\Excel

---

* [__construct()](#__construct) -- Excel constructor
* [colLetter()](#colletter) -- Convert column number to letter
* [colNum()](#colnum) -- Converts an alphabetic column index to a numeric
* [createReader()](#createreader)
* [createSheet()](#createsheet)
* [setTempDir()](#settempdir) -- Set dir for temporary files
* [template()](#template)
* [getCompleteStyleByIdx()](#getcompletestylebyidx)
* [countImages()](#countimages) -- Returns the total count of images in the workbook
* [getDateFormat()](#getdateformat)
* [setDateFormat()](#setdateformat)
* [getDateFormatPattern()](#getdateformatpattern)
* [dateFormatter()](#dateformatter) -- Sets custom date formatter
* [getDateFormatter()](#getdateformatter)
* [getDefinedNames()](#getdefinednames) -- Returns defined names of workbook
* [download()](#download) -- Download generated file to client (send to browser)
* [fill()](#fill) -- Set replacements of entire cell values for the sheet
* [getFirstSheet()](#getfirstsheet) -- Returns the first sheet as default
* [formatDate()](#formatdate)
* [getFormatPattern()](#getformatpattern)
* [hasDrawings()](#hasdrawings) -- Returns TRUE if the workbook contains an any draw objects (not images only)
* [hasExtraImages()](#hasextraimages)
* [hasImages()](#hasimages) -- Returns TRUE if any sheet contains an image object
* [getImageList()](#getimagelist) -- Returns the list of images from the workbook
* [mediaImageFiles()](#mediaimagefiles)
* [metadataImage()](#metadataimage)
* [output()](#output) -- Alias of download()
* [replace()](#replace) -- Set replacements of any occurring substrings
* [save()](#save) -- Save generated XLSX-file
* [selectFirstSheet()](#selectfirstsheet) -- Selects the first sheet as default
* [selectSheet()](#selectsheet) -- Selects default sheet by name
* [selectSheetById()](#selectsheetbyid) -- Selects default sheet by ID
* [sharedString()](#sharedstring) -- Returns string array by index
* [sheet()](#sheet)
* [getSheet()](#getsheet) -- Returns a sheet by name
* [getSheetById()](#getsheetbyid) -- Returns a sheet by ID
* [getSheetNames()](#getsheetnames) -- Returns names array of all sheets
* [sheets()](#sheets) -- Array of all sheets
* [styleByIdx()](#stylebyidx) -- Returns style array by style Idx
* [timestamp()](#timestamp) -- Convert date to timestamp

---

## __construct()

---

```php
public function __construct(string $templateFile, ?string $outputFile = null, 
                            ?array $options = [])
```
_Excel constructor_

### Parameters

* `string $templateFile`
* `string|null $outputFile`
* `array|null $options`

---

## colLetter()

---

```php
public static function colLetter(int $colNumber): string
```
_Convert column number to letter_

### Parameters

* `int $colNumber` -- ONE based

---

## colNum()

---

```php
public static function colNum(string $colLetter): int
```
_Converts an alphabetic column index to a numeric_

### Parameters

* `string $colLetter`

---

## createReader()

---

```php
public static function createReader(string $file, 
                                    ?array $parserProperties = []): Reader
```


### Parameters

* `$file`
* `$parserProperties`

---

## createSheet()

---

```php
public static function createSheet(string $sheetName, $sheetId, $file, $path, 
                                   $excel): Sheet
```


### Parameters

* `string $sheetName`
* `$sheetId`
* `$file`
* `$path`
* `$excel`

---

## setTempDir()

---

```php
public static function setTempDir($tempDir)
```
_Set dir for temporary files_

### Parameters

* `$tempDir`

---

## template()

---

```php
public static function template(string $templateFile, 
                                ?string $outputFile = null, 
                                ?array $options = []): Excel
```


### Parameters

* `string $templateFile`
* `string|null $outputFile`
* `array|null $options`

---

## getCompleteStyleByIdx()

---

```php
public function getCompleteStyleByIdx(int $styleIdx, 
                                      ?bool $flat = false): array
```


### Parameters

* `int $styleIdx`
* `bool|null $flat`

---

## countImages()

---

```php
public function countImages(): int
```
_Returns the total count of images in the workbook_

### Parameters

_None_

---

## getDateFormat()

---

```php
public function getDateFormat(): ?string
```


### Parameters

_None_

---

## setDateFormat()

---

```php
public function setDateFormat(string $dateFormat): avadim\FastExcelReader\Excel
```


### Parameters

* `string $dateFormat`

---

## getDateFormatPattern()

---

```php
public function getDateFormatPattern(int $styleIdx): ?string
```


### Parameters

* `int $styleIdx`

---

## dateFormatter()

---

```php
public function dateFormatter($formatter): avadim\FastExcelReader\Excel
```
_Sets custom date formatter_

### Parameters

* `\Closure|callable|string|bool $formatter`

---

## getDateFormatter()

---

```php
public function getDateFormatter(): callable|\Closure|bool|null
```


### Parameters

_None_

---

## getDefinedNames()

---

```php
public function getDefinedNames(): array
```
_Returns defined names of workbook_

### Parameters

_None_

---

## download()

---

```php
public function download(?string $name = null)
```
_Download generated file to client (send to browser)_

### Parameters

* `string|null $name`

---

## fill()

---

```php
public function fill(array $params): Excel
```
_Set replacements of entire cell values for the sheet_

### Parameters

* `array $params`

---

## getFirstSheet()

---

```php
public function getFirstSheet(?string $areaRange = null, 
                              ?bool $firstRowKeys = false): avadim\FastExcelReader\Sheet
```
_Returns the first sheet as default_

### Parameters

* `string|null $areaRange`
* `bool|null $firstRowKeys`

---

## formatDate()

---

```php
public function formatDate($value, $format, $styleIdx): false|mixed|string
```


### Parameters

* `$value`
* `$format`
* `$styleIdx`

---

## getFormatPattern()

---

```php
public function getFormatPattern(int $styleIdx): mixed|string
```


### Parameters

* `int $styleIdx`

---

## hasDrawings()

---

```php
public function hasDrawings(): bool
```
_Returns TRUE if the workbook contains an any draw objects (not images only)_

### Parameters

_None_

---

## hasExtraImages()

---

```php
public function hasExtraImages(): bool
```


### Parameters

_None_

---

## hasImages()

---

```php
public function hasImages(): bool
```
_Returns TRUE if any sheet contains an image object_

### Parameters

_None_

---

## getImageList()

---

```php
public function getImageList(): array
```
_Returns the list of images from the workbook_

### Parameters

_None_

---

## mediaImageFiles()

---

```php
public function mediaImageFiles(): array
```


### Parameters

_None_

---

## metadataImage()

---

```php
public function metadataImage(int $vmIndex): ?string
```


### Parameters

* `int $vmIndex`

---

## output()

---

```php
public function output(?string $name = null): void
```
_Alias of download()_

### Parameters

* `string|null $name`

---

## replace()

---

```php
public function replace(array $params): Excel
```
_Set replacements of any occurring substrings_

### Parameters

* `array $params`

---

## save()

---

```php
public function save(?string $fileName = null, ?bool $overWrite = true): bool
```
_Save generated XLSX-file_

### Parameters

* `string|null $fileName`
* `bool|null $overWrite`

---

## selectFirstSheet()

---

```php
public function selectFirstSheet(?string $areaRange = null, 
                                 ?bool $firstRowKeys = false): avadim\FastExcelReader\Sheet
```
_Selects the first sheet as default_

### Parameters

* `string|null $areaRange`
* `bool|null $firstRowKeys`

---

## selectSheet()

---

```php
public function selectSheet(string $name, ?string $areaRange = null, 
                            ?bool $firstRowKeys = false): avadim\FastExcelReader\Sheet
```
_Selects default sheet by name_

### Parameters

* `string $name`
* `string|null $areaRange`
* `bool|null $firstRowKeys`

---

## selectSheetById()

---

```php
public function selectSheetById(int $sheetId, ?string $areaRange = null, 
                                ?bool $firstRowKeys = false): avadim\FastExcelReader\Sheet
```
_Selects default sheet by ID_

### Parameters

* `int $sheetId`
* `string|null $areaRange`
* `bool|null $firstRowKeys`

---

## sharedString()

---

```php
public function sharedString($stringId): ?string
```
_Returns string array by index_

### Parameters

* `$stringId`

---

## sheet()

---

```php
public function sheet(?string $name = null): ?Sheet
```


### Parameters

* `string|null $name`

---

## getSheet()

---

```php
public function getSheet(?string $name = null, ?string $areaRange = null, 
                         ?bool $firstRowKeys = false): Sheet
```
_Returns a sheet by name_

### Parameters

* `string|null $name`
* `string|null $areaRange`
* `bool|null $firstRowKeys`

---

## getSheetById()

---

```php
public function getSheetById(int $sheetId, ?string $areaRange = null, 
                             ?bool $firstRowKeys = false): avadim\FastExcelReader\Sheet
```
_Returns a sheet by ID_

### Parameters

* `int $sheetId`
* `string|null $areaRange`
* `bool|null $firstRowKeys`

---

## getSheetNames()

---

```php
public function getSheetNames(): array
```
_Returns names array of all sheets_

### Parameters

_None_

---

## sheets()

---

```php
public function sheets(): array
```
_Array of all sheets_

### Parameters

_None_

---

## styleByIdx()

---

```php
public function styleByIdx($styleIdx): array
```
_Returns style array by style Idx_

### Parameters

* `$styleIdx`

---

## timestamp()

---

```php
public function timestamp($excelDateTime): int
```
_Convert date to timestamp_

### Parameters

* `$excelDateTime`

---


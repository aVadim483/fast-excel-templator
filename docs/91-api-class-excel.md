# Class \avadim\FastExcelTemplator\Excel

---

* [__construct()](#__construct) – Excel constructor
* [colLetter()](#colletter) – Convert column number to letter
* [colNum()](#colnum) – Converts an alphabetic column index to a numeric
* [createReader()](#createreader) – Create internal reader
* [createSheet()](#createsheet) – Create sheet template instance
* [isXls()](#isxls) – TRUE if the file starts with the OLE2 compound file signature
* [isXlsx()](#isxlsx) – TRUE if the file starts with the ZIP local file header signature
* [open()](#open) – Open a spreadsheet, choosing the reader by the file signature
* [openCsv()](#opencsv) – Open CSV file
* [openStream()](#openstream) – Open a spreadsheet from an open stream resource
* [openString()](#openstring) – Open a spreadsheet held in a string, choosing the reader by its signature
* [openXls()](#openxls) – Open an XLS (Excel 97-2003, BIFF8) file
* [setTempDir()](#settempdir) – Set dir for temporary files
* [template()](#template)
* [validate()](#validate) – Validate XLSX file
* [countExtraImages()](#countextraimages) – Count "extra" images (images that are in the media folder but not in the drawings)
* [countImages()](#countimages) – Returns the total count of images in the workbook
* [countSheets()](#countsheets) – Returns the number of sheets in the workbook
* [dateFormatter()](#dateformatter) – Set custom date formatter
* [download()](#download) – Download generated file to client (send to browser)
* [fill()](#fill) – Set replacements of entire cell values for the sheet
* [formatDate()](#formatdate) – Format date value
* [from()](#from) – Set top left of read area
* [getCompleteStyleByIdx()](#getcompletestylebyidx) – Get complete style by style index
* [getDateFormat()](#getdateformat) – Get current date format
* [getDateFormatPattern()](#getdateformatpattern) – Get PHP date format pattern by style index
* [getDateFormatter()](#getdateformatter) – Get date formatter
* [getDefinedNames()](#getdefinednames) – Get defined names of workbook
* [getFirstSheet()](#getfirstsheet) – Returns the first sheet as default
* [getFormatPattern()](#getformatpattern) – Get format pattern by style index
* [getImageList()](#getimagelist) – Get the list of images from the workbook
* [getProperties()](#getproperties) – Get the document properties of the workbook
* [getSheet()](#getsheet) – Returns a sheet by name
* [getSheetById()](#getsheetbyid) – Returns a sheet by ID
* [getSheetNames()](#getsheetnames) – Get names array of all sheets
* [hasDrawings()](#hasdrawings) – Returns TRUE if the workbook contains an any draw objects (not images only)
* [hasExtraImages()](#hasextraimages) – Returns TRUE if there are any "extra" images
* [hasImages()](#hasimages) – Returns TRUE if any sheet contains an image object
* [hiddenSheets()](#hiddensheets) – Array of hidden sheets only
* [innerFileList()](#innerfilelist) – Get list of inner files in XLSX
* [mediaImageFiles()](#mediaimagefiles) – Get list of media image files in the workbook
* [metadataImage()](#metadataimage) – Get image file name from metadata by index
* [output()](#output) – Alias of download()
* [outputFile()](#outputfile) – Name of the output file, or an empty string if it was not specified yet
* [readCallback()](#readcallback) – Reads cell values and passes them to a callback function
* [readCells()](#readcells) – Returns the values of all cells as array
* [readCellStyles()](#readcellstyles) – Returns the styles of all cells as array
* [readCellsWithStyles()](#readcellswithstyles) – Returns the values and styles of all cells as array
* [readColumns()](#readcolumns) – Returns cell values as a two-dimensional array from default sheet [col][row]
* [readColumnsWithStyles()](#readcolumnswithstyles) – Returns cell values and styles as a two-dimensional array from default sheet [col][row]
* [readRows()](#readrows) – Returns cell values as a two-dimensional array from default sheet [row][col]
* [readRowsWithStyles()](#readrowswithstyles) – Returns cell values and styles as a two-dimensional array from default sheet [row][col]
* [readStyles()](#readstyles) – Read all workbook styles
* [replace()](#replace) – Set replacements of any occurring substrings
* [save()](#save) – Save generated XLSX-file
* [selectFirstSheet()](#selectfirstsheet) – Selects the first sheet as default
* [selectSheet()](#selectsheet) – Selects default sheet by name
* [selectSheetById()](#selectsheetbyid) – Selects default sheet by ID
* [setDateFormat()](#setdateformat) – Set date format for reading
* [setReadArea()](#setreadarea) – Set top left and right bottom of read area
* [sharedString()](#sharedstring) – Get string by index
* [sheet()](#sheet)
* [sheetExists()](#sheetexists) – Returns TRUE if a sheet with the given name exists
* [sheets()](#sheets) – Array of all sheets
* [stat()](#stat) – Returns statistics of the workbook: per-sheet breakdown and totals
* [styleByIdx()](#stylebyidx) – Get style array by style index
* [templateFile()](#templatefile)
* [timestamp()](#timestamp) – Convert date to timestamp
* [useLocaleFormats()](#uselocaleformats) – Opt in to locale-dependent patterns for the built-in date codes (numFmtId 14-22).
* [visibleSheets()](#visiblesheets) – Array of visible sheets only

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

* `int $colNumber` – ONE based

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
_Create internal reader_

### Parameters

* `string $file`
* `array|null $parserProperties`

---

## createSheet()

---

```php
public static function createSheet(string $sheetName, $sheetId, $file, $path, 
                                   $excel): SheetTemplate
```
_Create sheet template instance_

### Parameters

* `string $sheetName`
* `int|string $sheetId`
* `string $file`
* `string $path`
* `Excel $excel`

---

## isXls()

---

```php
public static function isXls(string $file): bool
```
_TRUE if the file starts with the OLE2 compound file signature_

### Parameters

* `string $file`

---

## isXlsx()

---

```php
public static function isXlsx(string $file): bool
```
_TRUE if the file starts with the ZIP local file header signatureAn XLSX package is a ZIP archive, so it begins with "PK\x03\x04". This iswhat lets open() tell a real XLSX apart from plain text, which is thenread as CSV._

### Parameters

* `string $file`

---

## open()

---

```php
public static function open(string $file, 
                            $options): avadim\FastExcelReader\AbstractBook
```
_Open a spreadsheet, choosing the reader by the file signatureThe OLE2 magic number is a legacy XLS workbook, a ZIP container is XLSX,and anything else is treated as delimited text (CSV). The file extensionis not consulted, because it is often wrong on files arriving from othersystems. Pass $options\['format'] = 'csv' to force the CSV reader, and anyCsvOptions keys (delimiter, enclosure, encoding, ...) to configure it._

### Parameters

* `string $file`
* `CsvOptions|array|null $options`

---

## openCsv()

---

```php
public static function openCsv(string $file, 
                               $options): avadim\FastExcelReader\Csv\CsvReader
```
_Open CSV file_

### Parameters

* `string $file`
* `CsvOptions|array|null $options`

---

## openStream()

---

```php
public static function openStream($stream, 
                                  $options): avadim\FastExcelReader\AbstractBook
```
_Open a spreadsheet from an open stream resourceThe stream is copied (from its current position, without seeking, sonon-rewindable streams such as HTTP wrappers work) into a temporary fileand then opened like open(). This is the entry point for URLs(fopen('https://...')), php://memory and Flysystem/S3 read streams. Thecaller keeps ownership of the stream; it is not closed here. The temporaryfile is removed on script shutdown._

### Parameters

* `resource $stream` – An open, readable stream resource
* `CsvOptions|array|null $options` – Same options as open()

---

## openString()

---

```php
public static function openString(string $content, 
                                  $options): avadim\FastExcelReader\AbstractBook
```
_Open a spreadsheet held in a string, choosing the reader by its signatureThe content is written to a temporary file and then opened exactly likeopen() - format is detected from the bytes, not from any file name, so anXLSX/XLS/CSV payload each read back the same as its on-disk counterpart.The temporary file is removed on script shutdown. Handy for content comingfrom a database blob, an HTTP response body or an S3/Flysystem read._

### Parameters

* `string $content` – Raw bytes of the workbook
* `CsvOptions|array|null $options` – Same options as open()

---

## openXls()

---

```php
public static function openXls(string $file): avadim\FastExcelReader\Xls\XlsBook
```
_Open an XLS (Excel 97-2003, BIFF8) file_

### Parameters

* `string $file`

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

## validate()

---

```php
public static function validate(string $file, ?array &$errors = []): bool
```
_Validate XLSX file_

### Parameters

* `string $file`
* `array|null $errors`

---

## countExtraImages()

---

```php
public function countExtraImages(): int
```
_Count "extra" images (images that are in the media folder but not in the drawings)_

### Parameters

_None_

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

## countSheets()

---

```php
public function countSheets(): int
```
_Returns the number of sheets in the workbook_

### Parameters

_None_

---

## dateFormatter()

---

```php
public function dateFormatter($formatter): avadim\FastExcelReader\AbstractBook
```
_Set custom date formatter_

### Parameters

* `\Closure|callable|string|bool|null $formatter`

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

## formatDate()

---

```php
public function formatDate($value, $format, $styleIdx): false|mixed|string
```
_Format date value_

### Parameters

* `mixed $value`
* `string|null $format`
* `int|null $styleIdx`

---

## from()

---

```php
public function from(string $topLeftCell, 
                     ?bool $firstRowKeys = false): avadim\FastExcelReader\AbstractSheet
```
_Set top left of read area_

### Parameters

* `string $topLeftCell`
* `bool|null $firstRowKeys`

---

## getCompleteStyleByIdx()

---

```php
public function getCompleteStyleByIdx(int $styleIdx, 
                                      ?bool $flat = false): array
```
_Get complete style by style index_

### Parameters

* `int $styleIdx`
* `bool|null $flat`

---

## getDateFormat()

---

```php
public function getDateFormat(): ?string
```
_Get current date format_

### Parameters

_None_

---

## getDateFormatPattern()

---

```php
public function getDateFormatPattern(int $styleIdx): ?string
```
_Get PHP date format pattern by style index_

### Parameters

* `int $styleIdx`

---

## getDateFormatter()

---

```php
public function getDateFormatter(): callable|\Closure|bool|null
```
_Get date formatter_

### Parameters

_None_

---

## getDefinedNames()

---

```php
public function getDefinedNames(): array
```
_Get defined names of workbook_

### Parameters

_None_

---

## getFirstSheet()

---

```php
public function getFirstSheet(?string $areaRange = null, 
                              ?bool $firstRowKeys = false): avadim\FastExcelReader\AbstractSheet
```
_Returns the first sheet as default_

### Parameters

* `string|null $areaRange`
* `bool|null $firstRowKeys`

---

## getFormatPattern()

---

```php
public function getFormatPattern(int $styleIdx): mixed|string
```
_Get format pattern by style index_

### Parameters

* `int $styleIdx`

---

## getImageList()

---

```php
public function getImageList(): array
```
_Get the list of images from the workbook_

### Parameters

_None_

---

## getProperties()

---

```php
public function getProperties(): array
```
_Get the document properties of the workbookReads the core properties (docProps/core.xml) and the extended,application properties (docProps/app.xml) into a single associative arraywith normalised keys - 'creator', 'lastModifiedBy', 'created', 'modified','title', 'subject', 'description', 'keywords', 'category', 'revision','application', 'company', 'manager', ... Only the properties present in thefile are returned; a workbook without a docProps part returns an emptyarray. The result is read on demand and cached._

### Parameters

_None_

---

## getSheet()

---

```php
public function getSheet(?string $name = null, ?string $areaRange = null, 
                         ?bool $firstRowKeys = false): SheetTemplate
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
                             ?bool $firstRowKeys = false): avadim\FastExcelReader\AbstractSheet
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
_Get names array of all sheets_

### Parameters

_None_

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
_Returns TRUE if there are any "extra" images_

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

## hiddenSheets()

---

```php
public function hiddenSheets(): array
```
_Array of hidden sheets only_

### Parameters

_None_

---

## innerFileList()

---

```php
public function innerFileList(): array
```
_Get list of inner files in XLSX_

### Parameters

_None_

---

## mediaImageFiles()

---

```php
public function mediaImageFiles(): array
```
_Get list of media image files in the workbook_

### Parameters

_None_

---

## metadataImage()

---

```php
public function metadataImage(int $vmIndex): ?string
```
_Get image file name from metadata by index_

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

## outputFile()

---

```php
public function outputFile(): string
```
_Name of the output file, or an empty string if it was not specified yet_

### Parameters

_None_

---

## readCallback()

---

```php
public function readCallback(callable $callback, ?int $resultMode = null, 
                             ?bool $styleIdxInclude = null)
```
_Reads cell values and passes them to a callback function_

### Parameters

* `callback $callback`
* `int|null $resultMode`
* `bool|null $styleIdxInclude`

---

## readCells()

---

```php
public function readCells(): array
```
_Returns the values of all cells as array_

### Parameters

_None_

---

## readCellStyles()

---

```php
public function readCellStyles(?bool $flat = false): array
```
_Returns the styles of all cells as array_

### Parameters

* `bool|null $flat`

---

## readCellsWithStyles()

---

```php
public function readCellsWithStyles(): array
```
_Returns the values and styles of all cells as array_

### Parameters

_None_

---

## readColumns()

---

```php
public function readColumns($columnKeys, ?int $resultMode = null): array
```
_Returns cell values as a two-dimensional array from default sheet \[col]\[row]_

### Parameters

* `array|bool|int|null $columnKeys`
* `int|null $resultMode`

---

## readColumnsWithStyles()

---

```php
public function readColumnsWithStyles($columnKeys, 
                                      ?int $resultMode = null): array
```
_Returns cell values and styles as a two-dimensional array from default sheet \[col]\[row]_

### Parameters

* `array|bool|int|null $columnKeys`
* `int|null $resultMode`

---

## readRows()

---

```php
public function readRows($columnKeys, ?int $resultMode = null, 
                         ?bool $styleIdxInclude = null): array
```
_Returns cell values as a two-dimensional array from default sheet \[row]\[col]readRows()readRows(true)readRows(false, Excel::KEYS_ZERO_BASED)readRows(Excel::KEYS_ZERO_BASED | Excel::KEYS_RELATIVE)_

### Parameters

* `array|bool|int|null $columnKeys`
* `int|null $resultMode`
* `bool|null $styleIdxInclude`

---

## readRowsWithStyles()

---

```php
public function readRowsWithStyles($columnKeys, 
                                   ?int $resultMode = null): array
```
_Returns cell values and styles as a two-dimensional array from default sheet \[row]\[col]_

### Parameters

* `array|bool|int|null $columnKeys`
* `int|null $resultMode`

---

## readStyles()

---

```php
public function readStyles(): array
```
_Read all workbook styles_

### Parameters

_None_

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
                                 ?bool $firstRowKeys = false): avadim\FastExcelReader\AbstractSheet
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
                            ?bool $firstRowKeys = false): avadim\FastExcelReader\AbstractSheet
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
                                ?bool $firstRowKeys = false): avadim\FastExcelReader\AbstractSheet
```
_Selects default sheet by ID_

### Parameters

* `int $sheetId`
* `string|null $areaRange`
* `bool|null $firstRowKeys`

---

## setDateFormat()

---

```php
public function setDateFormat(string $dateFormat): avadim\FastExcelReader\AbstractBook
```
_Set date format for reading_

### Parameters

* `string $dateFormat`

---

## setReadArea()

---

```php
public function setReadArea(string $areaRange, 
                            ?bool $firstRowKeys = false): avadim\FastExcelReader\AbstractSheet
```
_Set top left and right bottom of read area_

### Parameters

* `string $areaRange`
* `bool|null $firstRowKeys`

---

## sharedString()

---

```php
public function sharedString($stringId): ?string
```
_Get string by index_

### Parameters

* `int $stringId`

---

## sheet()

---

```php
public function sheet(?string $name = null): ?SheetTemplate
```


### Parameters

* `string|null $name`

---

## sheetExists()

---

```php
public function sheetExists(string $name): bool
```
_Returns TRUE if a sheet with the given name exists_

### Parameters

* `string $name`

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

## stat()

---

```php
public function stat(): array
```
_Returns statistics of the workbook: per-sheet breakdown and totals\['sheets' => \['<sheetName>' => \['rows' => \[...], 'cols' => \[...], 'cells' => \['total' => int, 'filled' => int]],...],'total' => \['sheets'  => int,   // number of sheets'visible' => int,   // number of visible sheets'hidden'  => int,   // number of hidden sheets'rows'    => int,   // sum of actual rows over all sheets'cells'   => \['total' => int, 'filled' => int],],]Note: scans every sheet fully (see Sheet::stat()); expensive on large workbooks._

### Parameters

_None_

---

## styleByIdx()

---

```php
public function styleByIdx($styleIdx): array
```
_Get style array by style index_

### Parameters

* `int $styleIdx`

---

## templateFile()

---

```php
public function templateFile(): string
```


### Parameters

_None_

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

## useLocaleFormats()

---

```php
public function useLocaleFormats(?string $locale = null): avadim\FastExcelReader\AbstractBook
```
_Opt in to locale-dependent patterns for the built-in date codes (numFmtId 14-22).By default these codes resolve to fixed, deterministic patterns, so the same file yields the same output on any server. Call this to render them with ICU locale patterns instead (the pre-4.x behaviour, but now explicit and with a chosen locale). Requires ext-intl._

### Parameters

* `string|null $locale` – ICU locale name (e.g. 'ru_RU'); NULL uses the process default locale

---

## visibleSheets()

---

```php
public function visibleSheets(): array
```
_Array of visible sheets only_

### Parameters

_None_

---


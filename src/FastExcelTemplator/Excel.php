<?php

namespace avadim\FastExcelTemplator;

use avadim\FastExcelHelper\Helper;
use \avadim\FastExcelReader\Excel as ExcelReader;

/**
 * Class Excel
 *
 * @package avadim\FastExcelTemplator
 * @phpstan-consistent-constructor
 */
class Excel extends ExcelReader
{
    /**
     * MIME type of the generated files
     */
    public const MIME_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

    /**
     * @deprecated Assigned but never read, and a second workbook overwrites the first one.
     *             Keep a reference to the object returned by template() instead. To be removed in 4.0
     */
    public static Excel $instance;

    public ExcelWriter $excelWriter;

    public StyleManager $styleManager;

    protected string $templateFile;

    protected string $outputFile = '';

    /** @var SheetTemplate[] */
    protected array $sheets = [];


    /**
     * Excel constructor
     *
     * @param string $templateFile
     * @param string|null $outputFile
     * @param array|null $options
     */
    public function __construct(string $templateFile, ?string $outputFile = null, ?array $options = [])
    {
        parent::__construct($templateFile);
        self::$instance = $this;

        $this->templateFile = $templateFile;
        $options['style_manager'] = StyleManager::class;
        $this->excelWriter = new ExcelWriter($options);
        $this->styleManager = $this->excelWriter->styleManager;
        $this->styleManager->excelTemplator = $this;
        $this->styleManager->loadStyles();

        if ($outputFile) {
            $this->excelWriter->setFileName($outputFile);
            $this->outputFile = $outputFile;
        }
        $this->_importSheets();
        $this->excelWriter->setTemplateSheets($this->sheets);
    }

    /**
     * @return string
     */
    public function templateFile(): string
    {
        return $this->templateFile;
    }

    /**
     * Name of the output file, or an empty string if it was not specified yet
     *
     * @return string
     */
    public function outputFile(): string
    {
        return $this->outputFile;
    }

    /**
     * Return sheet/code/message diagnostics from the latest archive generation, including strict rejections; preliminary path errors do not refresh the report.
     *
     * @return array Diagnostic entries; an empty array means no metadata limitation was reported
     */
    public function metadataDiagnostics(): array
    {
        return $this->excelWriter->metadataDiagnostics();
    }

    /**
     * @return void
     */
    protected function _importSheets()
    {
        foreach ($this->sheets as $sheetId => $sheet) {
            $sheet->sheetWriter = $this->excelWriter->makeSheet($sheet->name());
            $sheet->sheetWriter->setTemplatePath($sheet->path());
            $this->xmlReader->openZip($sheet->path());
            $sheetViewsAttributes = [];
            while ($this->xmlReader->read()) {
                if ($this->xmlReader->nodeType === \XMLReader::ELEMENT && $this->xmlReader->name === 'sheetData') {
                    break;
                }
                if ($this->xmlReader->nodeType === \XMLReader::ELEMENT && $this->xmlReader->name === 'sheetView') {
                    // <sheetView tabSelected="1" view="pageBreakPreview" zoomScaleNormal="100" zoomScaleSheetLayoutView="100" workbookViewId="0">
                    $attributes = $this->xmlReader->getAllAttributes();
                    if ($attributes) {
                        $sheetViewsAttributes['_attr'] = $attributes;
                    }
                }
                if ($this->xmlReader->nodeType === \XMLReader::ELEMENT && $this->xmlReader->name === 'pane' && $sheetViewsAttributes) {
                    // <pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/>
                    $attributes = $this->xmlReader->getAllAttributes();
                    $sheetViewsAttributes['_items'][] = [
                        '_tag' => 'pane',
                        '_attr' => $attributes ?: [],
                    ];
                    if (isset($attributes['ySplit'], $attributes['xSplit'])) {
                        $sheet->sheetWriter->setFreezeRows((int)$attributes['ySplit']);
                        $sheet->sheetWriter->setFreezeColumns((int)$attributes['xSplit']);
                    }
                    elseif (isset($attributes['ySplit'])) {
                        $sheet->sheetWriter->setFreezeRows((int)$attributes['ySplit']);
                    }
                    elseif (isset($attributes['xSplit'])) {
                        $sheet->sheetWriter->setFreezeColumns((int)$attributes['xSplit']);
                    }
                }
                if ($this->xmlReader->nodeType === \XMLReader::ELEMENT && $this->xmlReader->name === 'selection' && $sheetViewsAttributes) {
                    // <selection pane="bottomLeft" activeCell="D1048557" sqref="D1048557"/>
                    $attributes = $this->xmlReader->getAllAttributes();
                    $sheetViewsAttributes['_items'][] = [
                        '_tag' => 'selection',
                        '_attr' => $attributes ?: [],
                    ];
                }
                elseif ($this->xmlReader->nodeType === \XMLReader::END_ELEMENT && $this->xmlReader->name === 'sheetViews') {
                    $sheet->sheetWriter->_setSheetViewsAttributes($sheetViewsAttributes);
                }
                if ($this->xmlReader->nodeType === \XMLReader::ELEMENT && $this->xmlReader->name === 'sheetFormatPr') {
                    // <sheetFormatPr defaultColWidth="9.1015625" defaultRowHeight="12.9" x14ac:dyDescent="0.5"/>
                    $attributes = $this->xmlReader->getAllAttributes();
                    if ($attributes) {
                        $sheet->sheetWriter->_setSheetFormatPrAttributes($attributes);
                    }
                }
                elseif ($this->xmlReader->nodeType === \XMLReader::ELEMENT && $this->xmlReader->name === 'col') {
                    // <col min="1" max="1" width="20.83203125" customWidth="1"/>
                    $attributes = $this->xmlReader->getAllAttributes();
                    if (isset($attributes['min'])) {
                        $colIdx = $attributes['min'] - 1;
                        $sheet->sheetWriter->_setColAttributes($colIdx, $attributes);
                    }
                }
                elseif ($this->xmlReader->nodeType === \XMLReader::END_ELEMENT && $this->xmlReader->name === 'cols') {
                    break;
                }
            }

            $this->xmlReader->close();
        }

        foreach ($this->sheets as $sheetId => $sheet) {
            foreach ($sheet->sheetWriter->getColAttributes() as $colIdx => $attributes) {
                if (!empty($attributes['style'])) {
                    $style = $this->getCompleteStyleByIdx($attributes['style'], true);
                    if ($style) {
                        $sheet->sheetWriter->setColStyle(Helper::colLetter($colIdx + 1), $style);
                    }
                }
            }
        }
    }

    /**
     * Open an XLSX template and create its output writer.
     *
     * @param string $templateFile Existing XLSX template
     * @param string|null $outputFile Output path; may also be supplied to save()
     * @param array|null $options Writer options; strict_metadata (bool, default false) rejects detected unsupported metadata transformations at save time
     *
     * @return Excel
     */
    public static function template(string $templateFile, ?string $outputFile = null, ?array $options = []): Excel
    {
        // 'static' rather than 'self', so that a subclass of Excel gets an instance of itself
        return new static($templateFile, $outputFile, $options);
    }

    /**
     * Create sheet template instance
     *
     * @param string $sheetName
     * @param int|string $sheetId
     * @param string $file
     * @param string $path
     * @param Excel $excel
     *
     * @return SheetTemplate
     */
    public static function createSheet(string $sheetName, $sheetId, $file, $path, $excel): SheetTemplate
    {
        return new SheetTemplate($sheetName, $sheetId, $file, $path, $excel);
    }

    /**
     * Create internal reader
     *
     * @param string $file
     * @param array|null $parserProperties
     *
     * @return Reader
     */
    public static function createReader(string $file, ?array $parserProperties = []): Reader
    {
        return new Reader($file, $parserProperties);
    }

    /**
     * Set dir for temporary files
     *
     * @param $tempDir
     */
    public static function setTempDir($tempDir)
    {
        ExcelWriter::setTempDir($tempDir);
    }

    /**
     * @param string|null $name
     *
     * @return \avadim\FastExcelReader\Sheet|SheetTemplate|null
     */
    public function sheet(?string $name = null): ?SheetTemplate
    {
        return parent::sheet($name);
    }

    /**
     * Returns a sheet by name
     *
     * @param string|null $name
     * @param string|null $areaRange
     * @param bool|null $firstRowKeys
     *
     * @return \avadim\FastExcelReader\Sheet|SheetTemplate|null
     */
    public function getSheet(?string $name = null, ?string $areaRange = null, ?bool $firstRowKeys = false): SheetTemplate
    {
        return parent::getSheet($name, $areaRange, $firstRowKeys);
    }

    /**
     * Array of all sheets
     *
     * @return SheetTemplate[]
     */
    public function sheets(): array
    {
        return $this->sheets;
    }

    /**
     * Set replacements of entire cell values for the sheet
     *
     * @param array $params
     *
     * @return $this
     */
    public function fill(array $params): Excel
    {
        foreach ($this->sheets as $sheet) {
            $sheet->fill($params);
        }

        return $this;
    }

    /**
     * Set replacements of any occurring substrings
     *
     * @param array $params
     *
     * @return $this
     */
    public function replace(array $params): Excel
    {
        foreach ($this->sheets as $sheet) {
            $sheet->replace($params);
        }
        return $this;
    }

    /** Resolve file symlinks explicitly: Windows realpath() may retain the link's own path. */
    private function resolveSavePath(string $path): ?string
    {
        $visited = [];
        while (true) {
            clearstatcache(true, $path);
            $resolved = realpath($path);
            if ($resolved === false) {
                return null;
            }
            if (!is_link($resolved)) {
                return $resolved;
            }
            $key = DIRECTORY_SEPARATOR === '\\' ? strtolower($resolved) : $resolved;
            if (isset($visited[$key]) || count($visited) >= 40) {
                throw new Exception('Cannot resolve symbolic link "' . $path . '"');
            }
            $visited[$key] = true;
            $target = readlink($resolved);
            if ($target === false) {
                throw new Exception('Cannot resolve symbolic link "' . $resolved . '"');
            }
            if (DIRECTORY_SEPARATOR === '\\') {
                $target = str_replace('/', '\\', $target);
                // A rooted path without a drive is relative to the link's drive.
                if (substr($target, 0, 1) === '\\' && substr($target, 0, 2) !== '\\\\'
                    && preg_match('/^[A-Za-z]:/', $resolved)) {
                    $target = substr($resolved, 0, 2) . $target;
                }
            }
            $path = preg_match('~^(?:[A-Za-z]:[\\\\/]|[\\\\/])~', $target)
                ? $target : dirname($resolved) . DIRECTORY_SEPARATOR . $target;
        }
    }

    /**
     * Save XLSX after metadata checks; publish only the completed workbook, preserving the template and an existing destination on failure.
     *
     * @param string|null $fileName Output path; null uses the configured writer filename
     * @param bool|null $overWrite False refuses existing destinations, including files created during generation
     *
     * @return bool
     * @throws Exception On invalid output paths, publication failure or a strict metadata rejection
     */
    public function save(?string $fileName = null, ?bool $overWrite = true): bool
    {
        if (!$fileName) {
            $fileName = $this->excelWriter->getFileName();
        }
        $sourcePath = $this->resolveSavePath($this->file);
        $targetPath = $this->resolveSavePath($fileName);
        if ($sourcePath !== null && $targetPath !== null
            && (DIRECTORY_SEPARATOR === '\\' ? strcasecmp($sourcePath, $targetPath) === 0 : $sourcePath === $targetPath)) {
            throw new Exception('Cannot overwrite the template file "' . $fileName . '"');
        }
        if (!$overWrite && (file_exists($fileName) || is_link($fileName))) {
            throw new Exception('Cannot overwrite existing file "' . $fileName . '"');
        }
        $directory = realpath(dirname($fileName));
        if ($directory === false || !is_dir($directory)) {
            throw new Exception('Output directory does not exist for "' . $fileName . '"');
        }
        $temporaryFile = @tempnam($directory, 'fxt-');
        if ($temporaryFile === false) {
            throw new Exception('Cannot create temporary output for "' . $fileName . '"');
        }
        try {
            // tempnam() can fall back to the system directory; publishing must stay on one filesystem.
            if (realpath(dirname($temporaryFile)) !== $directory) {
                throw new Exception('Cannot create temporary output in "' . $directory . '"');
            }
            foreach ($this->sheets as $sheet) {
                $sheet->saveSheet();
            }
            $this->excelWriter->saved = false;
            if (!$this->excelWriter->replaceSheetsAndSave($this->file, $temporaryFile, $this->styleManager->loaded)) {
                return false;
            }
            $this->excelWriter->saved = false;
            // rename replaces an existing file without deleting it first. link publishes exclusively,
            // so overWrite=false also protects a destination created while generation was in progress.
            $published = $overWrite ? @rename($temporaryFile, $fileName) : @link($temporaryFile, $fileName);
            if (!$published) {
                throw new Exception('Cannot write output file "' . $fileName . '"');
            }
            $this->excelWriter->saved = true;
            return true;
        }
        finally {
            if (is_file($temporaryFile)) {
                @unlink($temporaryFile);
            }
        }
    }

    /**
     * Headers sent by download() before the body of the file
     *
     * @param string $name Name of the file offered to the client
     * @param int $fileSize Size of the file in bytes
     *
     * @return string[]
     */
    protected function downloadHeaders(string $name, int $fileSize): array
    {
        return [
            'Cache-Control: max-age=0',
            'Content-Type: ' . self::MIME_TYPE,
            'Content-Disposition: attachment; filename="' . $name . '"',
            'Content-Length: ' . $fileSize,
        ];
    }

    /**
     * Sends the headers of the response
     *
     * @param string[] $headers
     *
     * @return void
     */
    protected function sendHeaders(array $headers)
    {
        foreach ($headers as $header) {
            header($header);
        }
    }

    /**
     * Whether the response headers were already sent, so no file can be sent anymore
     *
     * Under the CLI SAPI there is no response to spoil: headers_sent() turns true with the
     * first byte of ordinary output, so the check is limited to the SAPIs where it means
     * something.
     *
     * @param string|null $file Name of the file that started the output
     * @param int|null $line Line that started the output
     *
     * @return bool
     */
    protected function headersSent(?string &$file = null, ?int &$line = null): bool
    {
        return PHP_SAPI !== 'cli' && headers_sent($file, $line);
    }

    /**
     * Download generated file to client (send to browser)
     *
     * @param string|null $name
     */
    public function download(?string $name = null)
    {
        if ($this->headersSent($file, $line)) {
            throw new Exception('Cannot send the file, the headers were already sent in "' . $file . '" on line ' . $line);
        }

        $tmpFile = $this->excelWriter->writer->makeTempFileName(uniqid('xlsx_writer_'));
        try {
            if (!$this->save($tmpFile)) {
                throw new Exception('Cannot generate download file');
            }
            if (!$name) {
                $name = basename($tmpFile) . '.xlsx';
            }
            else {
                $name = basename($name);
                if (strtolower(pathinfo($name, PATHINFO_EXTENSION)) !== 'xlsx') {
                    $name .= '.xlsx';
                }
            }

            $this->sendHeaders($this->downloadHeaders($name, (int)filesize($tmpFile)));
            if (readfile($tmpFile) === false) {
                throw new Exception('Cannot read generated file "' . $tmpFile . '"');
            }
        }
        finally {
            if (is_file($tmpFile)) {
                @unlink($tmpFile);
            }
        }
    }

    /**
     * Alias of download()
     *
     * @param string|null $name
     *
     * @return void
     */
    public function output(?string $name = null)
    {
        $this->download($name);
    }

}

// EOF

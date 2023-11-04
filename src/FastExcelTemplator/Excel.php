<?php

namespace avadim\FastExcelTemplator;

use \avadim\FastExcelReader\Excel as ExcelReader;

/**
 * Class Excel
 *
 * @package avadim\FastExcelTemplator
 */
class Excel extends ExcelReader
{
    public static Excel $instance;

    public $excelWriter;

    protected string $templateFile;

    /** @var Sheet[] */
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
        $this->excelWriter = new ExcelWriter($options);
        if ($outputFile) {
            $this->excelWriter->setFileName($outputFile);
        }
        $this->_importSheets();
    }

    /**
     * @return void
     */
    protected function _importSheets()
    {
        foreach ($this->sheets as $sheetId => $sheet) {
            $sheet->sheetWriter = $this->excelWriter->makeSheet($sheet->name());
            $this->xmlReader->openZip($sheet->path());
            while ($this->xmlReader->read()) {
                if ($this->xmlReader->nodeType === \XMLReader::ELEMENT && $this->xmlReader->name === 'col') {
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
    }

    /**
     * @param string $templateFile
     * @param string|null $outputFile
     * @param array|null $options
     *
     * @return Excel
     */
    public static function template(string $templateFile, ?string $outputFile = null, ?array $options = []): Excel
    {
        return new self($templateFile, $outputFile, $options);
    }

    /**
     * @param string $sheetName
     * @param $sheetId
     * @param $file
     * @param $path
     * @param $excel
     *
     * @return Sheet
     */
    public static function createSheet(string $sheetName, $sheetId, $file, $path, $excel): Sheet
    {
        return new Sheet($sheetName, $sheetId, $file, $path, $excel);
    }

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
     * @return \avadim\FastExcelReader\Sheet|Sheet|null
     */
    public function sheet(?string $name = null): ?Sheet
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
     * @return \avadim\FastExcelReader\Sheet|Sheet|null
     */
    public function getSheet(?string $name = null, ?string $areaRange = null, ?bool $firstRowKeys = false): Sheet
    {
        return parent::getSheet($name, $areaRange, $firstRowKeys);
    }

    /**
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

    /**
     * Save generated XLSX-file
     *
     * @param string|null $fileName
     * @param bool|null $overWrite
     *
     * @return bool
     */
    public function save(?string $fileName = null, ?bool $overWrite = true): bool
    {
        if (!$fileName) {
            $fileName = $this->excelWriter->getFileName();
        }
        if (is_file($fileName) && $overWrite) {
            @unlink($fileName);
        }
        foreach ($this->sheets as $sheet) {
            $sheet->transferRows();
        }

        return $this->excelWriter->replaceSheets($this->file, $fileName);
    }

    /**
     * Download generated file to client (send to browser)
     *
     * @param string|null $name
     */
    public function download(string $name = null)
    {

        $this->excelWriter->download($name);
    }

    /**
     * Alias of download()
     *
     * @param string|null $name
     *
     * @return void
     */
    public function output(string $name = null)
    {
        $this->download($name);
    }

}

// EOF
<?php

namespace avadim\FastExcelTemplator;

use avadim\FastExcelWriter\Exceptions\ExceptionFile;

class ExcelWriter extends \avadim\FastExcelWriter\Excel
{
    /** @var SheetTemplate[] */
    protected array $templateSheets = [];
    protected bool $strictMetadata = false;
    protected array $metadataDiagnostics = [];

    /** @internal */
    public function setTemplateSheets(array $sheets): void
    {
        $this->templateSheets = $sheets;
    }

    public function metadataDiagnostics(): array
    {
        return $this->metadataDiagnostics;
    }

    public function __construct($options = null)
    {
        $options = $options ?? [];
        $this->strictMetadata = (bool)($options['strict_metadata'] ?? false);
        unset($options['strict_metadata']);
        $options['writer_class'] = $options['writer_class'] ?? TemplateWriter::class;
        parent::__construct($options);
    }

    public function __destruct()
    {
        // The base writer auto-saves in its destructor. A templator must only save explicitly:
        // after a failed save this would overwrite the destination with a newly built workbook.
    }

    /**
     * Create sheet writer instance
     *
     * @param string $sheetName
     *
     * @return SheetWriter
     */
    public static function createSheet(string $sheetName): SheetWriter
    {
        return new SheetWriter($sheetName);
    }

    /**
     * Make new sheet
     *
     * @param string|null $sheetName
     *
     * @return \avadim\FastExcelWriter\Sheet|SheetWriter
     */
    public function makeSheet(?string $sheetName = null): SheetWriter
    {
        return parent::makeSheet($sheetName);
    }

    /**
     * Replace sheets in template and save result
     *
     * @param string $inputFile
     * @param string $outputFile
     * @param bool|null $replaceStyles
     *
     * @return bool
     */
    public function replaceSheetsAndSave(string $inputFile, string $outputFile, ?bool $replaceStyles = false): bool
    {
        $relationShips = [
            'rel_id' => ['workbook' => 0],
        ];

        $this->saved = false;
        $this->metadataDiagnostics = [];
        if (!$this->writer->_replaceSheets($inputFile, $outputFile, $relationShips)) {
            return false;
        }
        if ($replaceStyles && !$this->writer->_replaceStyles($outputFile)) {
            ExceptionFile::throwNew('Unable to replace styles in "%s"', $outputFile);
        }
        $zip = new \ZipArchive();
        $result = $zip->open($outputFile);
        if ($result !== true) {
            ExceptionFile::throwNew('Unable to open zip "%s". Error code "%s"', $outputFile, $result);
        }
        $completed = false;
        $temporaryFiles = [];
        $metadata = new WorksheetMetadata($this->strictMetadata);
        try {
            if ($replaceStyles) {
                $original = new \ZipArchive();
                if ($original->open($inputFile) !== true) { throw new Exception('Cannot read template style metadata'); }
                try { $sourceStyles = $original->getFromName('xl/styles.xml'); }
                finally { $original->close(); }
                $generatedStyles = $zip->getFromName('xl/styles.xml');
                if ($sourceStyles === false || $generatedStyles === false
                    || !$zip->addFromString('xl/styles.xml', WorksheetMetadata::preserveStyleMetadata($sourceStyles, $generatedStyles))) {
                    throw new Exception('Cannot preserve style metadata');
                }
            }
            foreach ($this->templateSheets as $sheet) {
                $temporary = $metadata->rewrite($inputFile, $outputFile, $sheet->path(), $sheet->hasStructuralChanges(), dirname($outputFile));
                $temporaryFiles[] = $temporary;
                if (!$zip->addFile($temporary, $sheet->path())) {
                    ExceptionFile::throwNew('Unable to preserve metadata for "%s"', $sheet->path());
                }
            }
            if ($zip->statName('xl/calcChain.xml') !== false) {
                if (!$zip->deleteName('xl/calcChain.xml')) {
                    ExceptionFile::throwNew('Unable to remove calculation chain from "%s"', $outputFile);
                }
                foreach (['[Content_Types].xml', 'xl/_rels/workbook.xml.rels'] as $entry) {
                    $str = $zip->getFromName($entry);
                    if ($str === false) {
                        ExceptionFile::throwNew('Unable to read entry "%s"', $entry);
                    }
                    $str = preg_replace('#<[^>]+calcChain.xml[^>]+>#i', '', $str);
                    if ($str === null || !$zip->addFromString($entry, $str)) {
                        ExceptionFile::throwNew('Unable to update entry "%s"', $entry);
                    }
                }
            }
            $completed = true;
        }
        finally {
            $this->metadataDiagnostics = $metadata->diagnostics();
            try { $closed = $zip->close(); }
            finally {
                foreach ($temporaryFiles as $temporary) { @unlink($temporary); }
            }
            if ($completed && !$closed) {
                ExceptionFile::throwNew('Unable to write zip "%s"', $outputFile);
            }
        }
        $this->saved = true;
        return true;
    }
}

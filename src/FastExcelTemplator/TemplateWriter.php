<?php

namespace avadim\FastExcelTemplator;

use avadim\FastExcelWriter\Exceptions\ExceptionFile;

/**
 * Checked archive replacement for templators. The base Writer's replacement paths do not
 * check entry writes or close(), unlike its normal save path. Keep this adapter until the
 * supported dependency versions provide those guarantees.
 *
 * @internal
 */
class TemplateWriter extends \avadim\FastExcelWriter\Writer\Writer
{
    protected function createArchive(): \ZipArchive
    {
        return new \ZipArchive();
    }

    private function openArchive(string $outputFile): void
    {
        $this->zip = $this->createArchive();
        $result = $this->zip->open($outputFile);
        if ($result !== true) {
            ExceptionFile::throwNew('Unable to open zip "%s". Error code "%s"', $outputFile, $result);
        }
    }

    public function _replaceSheets(string $inputFile, string $outputFile, array $relationShips): bool
    {
        if (!@copy($inputFile, $outputFile)) {
            ExceptionFile::throwNew('Cannot write file "%s"', $outputFile);
        }
        $this->openArchive($outputFile);
        $completed = false;
        try {
            $entryList = [];
            for ($i = 0; $i < $this->zip->numFiles; $i++) {
                $entryList[$i] = $this->zip->getNameIndex($i);
            }
            $sheets = $this->excel->getSheets();
            $relationShips = [];
            if (!$this->_writeSheetsFiles($sheets, $relationShips)) {
                ExceptionFile::throwNew('Unable to generate sheets for "%s"', $outputFile);
            }
            foreach ($entryList as $index => $name) {
                if (strpos($name, 'xl/worksheets/sheet') === 0 && !empty($this->tempFiles['zip'][$name])) {
                    if (PHP_VERSION_ID >= 80300) {
                        $written = $this->zip->replaceFile($this->tempFiles['zip'][$name], $index);
                    } else {
                        $written = $this->zip->addFile($this->tempFiles['zip'][$name], $name);
                    }
                    if (!$written) {
                        ExceptionFile::throwNew('Unable to replace entry "%s"', $name);
                    }
                }
            }
            $completed = true;
        }
        finally {
            $closed = $this->zip->close();
            if ($completed && !$closed) {
                ExceptionFile::throwNew('Unable to write zip "%s"', $outputFile);
            }
        }
        return true;
    }

    public function _replaceStyles(string $outputFile): bool
    {
        $this->openArchive($outputFile);
        $completed = false;
        try {
            $stylesFile = $this->_writeStylesXML();
            if (!$stylesFile || !$this->zip->addFile($stylesFile, 'xl/styles.xml')) {
                ExceptionFile::throwNew('Unable to replace styles in "%s"', $outputFile);
            }
            $completed = true;
        }
        finally {
            $closed = $this->zip->close();
            if ($completed && !$closed) {
                ExceptionFile::throwNew('Unable to write zip "%s"', $outputFile);
            }
        }
        return true;
    }
}

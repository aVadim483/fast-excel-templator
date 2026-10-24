<?php

declare(strict_types=1);
namespace avadim\FastExcelTemplator;

use avadim\FastExcelWriter\Exceptions\ExceptionFile;
use PHPUnit\Framework\TestCase;

final class ArchiveFailureTest extends TestCase
{
    private function writer(\ZipArchive $archive): ArchiveTestWriter
    {
        $writer = new ArchiveTestWriter();
        $writer->archive = $archive;
        return $writer;
    }

    public function testNumericZipOpenErrorIsRejected(): void
    {
        $zip = new FaultArchive();
        $zip->openResult = \ZipArchive::ER_NOZIP;
        $this->expectException(ExceptionFile::class);
        $this->expectExceptionMessage('Error code');
        try { $this->writer($zip)->_replaceStyles('output.xlsx'); }
        finally {
            self::assertSame(1, $zip->openCalls);
            self::assertSame(0, $zip->writeCalls);
            self::assertSame(0, $zip->closeCalls);
        }
    }

    public function testFailedEntryWriteClosesArchiveAndThrows(): void
    {
        $zip = new FaultArchive();
        $zip->writeResult = false;
        $this->expectException(ExceptionFile::class);
        $this->expectExceptionMessage('replace styles');
        try { $this->writer($zip)->_replaceStyles('output.xlsx'); }
        finally {
            self::assertSame(1, $zip->writeCalls);
            self::assertSame(1, $zip->closeCalls);
        }
    }

    public function testFailedArchiveCloseIsNotReportedAsSuccess(): void
    {
        $zip = new FaultArchive();
        $zip->closeResult = false;
        $this->expectException(ExceptionFile::class);
        $this->expectExceptionMessage('write zip');
        try { $this->writer($zip)->_replaceStyles('output.xlsx'); }
        finally { self::assertSame(1, $zip->closeCalls); }
    }

    public function testExceptionDuringGenerationClosesArchive(): void
    {
        $zip = new FaultArchive();
        $writer = $this->writer($zip);
        $writer->failGeneration = true;
        $this->expectException(\RuntimeException::class);
        $this->expectExceptionMessage('generation failure');
        try { $writer->_replaceStyles('output.xlsx'); }
        finally { self::assertSame(1, $zip->closeCalls); }
    }
}

/** A small I/O fault boundary that does not depend on PHPUnit mocking internal ZIP constants. */
class FaultArchive extends \ZipArchive
{
    public $openResult = true;
    public bool $writeResult = true;
    public bool $closeResult = true;
    public int $openCalls = 0;
    public int $writeCalls = 0;
    public int $closeCalls = 0;

    #[\ReturnTypeWillChange]
    public function open($filename, $flags = 0)
    {
        $this->openCalls++;
        return $this->openResult;
    }

    #[\ReturnTypeWillChange]
    public function addFile($filepath, $entryname = '', $start = 0, $length = 0, $flags = 0)
    {
        $this->writeCalls++;
        return $this->writeResult;
    }

    #[\ReturnTypeWillChange]
    public function close()
    {
        $this->closeCalls++;
        return $this->closeResult;
    }
}

class ArchiveTestWriter extends TemplateWriter
{
    public \ZipArchive $archive;
    public bool $failGeneration = false;

    protected function createArchive(): \ZipArchive
    {
        return $this->archive;
    }

    protected function _writeStylesXML()
    {
        if ($this->failGeneration) {
            throw new \RuntimeException('generation failure');
        }
        return 'styles.xml';
    }
}

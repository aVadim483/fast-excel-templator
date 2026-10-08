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
        $zip = $this->createMock(\ZipArchive::class);
        $zip->expects(self::once())->method('open')->willReturn(\ZipArchive::ER_NOZIP);
        $zip->expects(self::never())->method('addFile');
        $zip->expects(self::never())->method('close');
        $this->expectException(ExceptionFile::class);
        $this->expectExceptionMessage('Error code');
        $this->writer($zip)->_replaceStyles('output.xlsx');
    }

    public function testFailedEntryWriteClosesArchiveAndThrows(): void
    {
        $zip = $this->createMock(\ZipArchive::class);
        $zip->method('open')->willReturn(true);
        $zip->expects(self::once())->method('addFile')->willReturn(false);
        $zip->expects(self::once())->method('close')->willReturn(true);
        $this->expectException(ExceptionFile::class);
        $this->expectExceptionMessage('replace styles');
        $this->writer($zip)->_replaceStyles('output.xlsx');
    }

    public function testFailedArchiveCloseIsNotReportedAsSuccess(): void
    {
        $zip = $this->createMock(\ZipArchive::class);
        $zip->method('open')->willReturn(true);
        $zip->method('addFile')->willReturn(true);
        $zip->expects(self::once())->method('close')->willReturn(false);
        $this->expectException(ExceptionFile::class);
        $this->expectExceptionMessage('write zip');
        $this->writer($zip)->_replaceStyles('output.xlsx');
    }

    public function testExceptionDuringGenerationClosesArchive(): void
    {
        $zip = $this->createMock(\ZipArchive::class);
        $zip->method('open')->willReturn(true);
        $zip->expects(self::once())->method('close')->willReturn(true);
        $writer = $this->writer($zip);
        $writer->failGeneration = true;
        $this->expectException(\RuntimeException::class);
        $this->expectExceptionMessage('generation failure');
        $writer->_replaceStyles('output.xlsx');
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

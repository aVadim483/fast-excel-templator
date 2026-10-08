<?php

declare(strict_types=1);
namespace avadim\FastExcelTemplator;

use avadim\FastExcelWriter\Excel as Writer;
use PHPUnit\Framework\TestCase;

final class SafeSaveTest extends TestCase
{
    private string $dir;
    private string $tpl;
    private string $out;

    protected function setUp(): void
    {
        $this->dir = sys_get_temp_dir() . '/fxt-save-' . uniqid();
        mkdir($this->dir);
        $this->tpl = $this->dir . '/tpl.xlsx';
        $this->out = $this->dir . '/out.xlsx';
        $writer = Writer::create(['Sheet1']);
        $writer->sheet()->writeRow(['original']);
        $writer->save($this->tpl);
    }

    protected function tearDown(): void
    {
        gc_collect_cycles();
        foreach (glob($this->dir . '/*') as $file) {
            if (is_dir($file)) { rmdir($file); } else { unlink($file); }
        }
        rmdir($this->dir);
    }

    private function workbook(): Excel
    {
        $excel = Excel::template($this->tpl, $this->out);
        $excel->sheet()->fill(['original' => 'updated'])->transferRows();
        return $excel;
    }

    public function testOverwriteFalsePreservesExistingFile(): void
    {
        file_put_contents($this->out, 'existing');
        $excel = $this->workbook();
        try {
            $excel->save(null, false);
            self::fail('Expected refusal to overwrite');
        } catch (Exception $e) {
            self::assertStringContainsString('existing file', $e->getMessage());
        }
        self::assertSame('existing', file_get_contents($this->out));
        self::assertFalse($excel->excelWriter->saved);
    }

    public function testTemplateCannotBeOverwritten(): void
    {
        $original = file_get_contents($this->tpl);
        $excel = $this->workbook();
        try {
            $excel->save($this->dir . '/./tpl.xlsx');
            self::fail('Expected refusal to overwrite template');
        } catch (Exception $e) {
            self::assertStringContainsString('template', $e->getMessage());
        }
        self::assertSame($original, file_get_contents($this->tpl));
    }

    public function testGenerationFailurePreservesDestinationAndCleansTemporaryFile(): void
    {
        file_put_contents($this->out, 'existing');
        $excel = $this->workbook();
        $excel->excelWriter->writer = new FailingSaveWriter();
        try {
            $excel->save();
            self::fail('Expected generation failure');
        } catch (\RuntimeException $e) {
            self::assertSame('injected write failure', $e->getMessage());
        }
        self::assertSame('existing', file_get_contents($this->out));
        self::assertSame([], glob($this->dir . '/fxt-*'));
        self::assertFalse($excel->excelWriter->saved);
    }

    public function testSaveWithAndWithoutOverwriteProducesReadableWorkbook(): void
    {
        self::assertTrue($this->workbook()->save(null, false));
        self::assertTrue($this->workbook()->save());
        $rows = \avadim\FastExcelReader\Excel::open($this->out)->readRows();
        self::assertSame([1 => ['A' => 'updated']], $rows);
        self::assertSame([], glob($this->dir . '/fxt-*'));
    }

    public function testDestinationCreatedDuringGenerationIsNotOverwritten(): void
    {
        $writer = new RacingSaveWriter();
        $writer->destination = $this->out;
        $excel = Excel::template($this->tpl, $this->out, ['writer_class' => $writer]);
        $excel->sheet()->transferRows();
        try {
            $excel->save(null, false);
            self::fail('Expected exclusive publication failure');
        } catch (Exception $e) {
            self::assertStringContainsString('Cannot write output file', $e->getMessage());
        }
        self::assertSame('concurrent file', file_get_contents($this->out));
        self::assertFalse($excel->excelWriter->saved);
        self::assertSame([], glob($this->dir . '/fxt-*'));
    }

    public function testLockedDestinationIsPreservedOnWindows(): void
    {
        if (DIRECTORY_SEPARATOR !== '\\') {
            self::markTestSkipped('Windows file sharing semantics');
        }
        file_put_contents($this->out, 'existing');
        $handle = fopen($this->out, 'rb');
        $excel = $this->workbook();
        try {
            try {
                $excel->save();
                self::fail('Expected locked destination failure');
            } catch (Exception $e) {
                self::assertStringContainsString('Cannot write output file', $e->getMessage());
            }
            self::assertSame('existing', file_get_contents($this->out));
            self::assertFalse($excel->excelWriter->saved);
            self::assertSame([], glob($this->dir . '/fxt-*'));
        } finally {
            fclose($handle);
        }
    }

    public function testPublicationFailureCleansTemporaryFile(): void
    {
        mkdir($this->out);
        $excel = $this->workbook();
        try {
            $excel->save();
            self::fail('Expected publication failure');
        } catch (Exception $e) {
            self::assertStringContainsString('Cannot write output file', $e->getMessage());
        }
        self::assertDirectoryExists($this->out);
        self::assertFalse($excel->excelWriter->saved);
        self::assertSame([], glob($this->dir . '/fxt-*'));
    }

    public function testFailedSaveDoesNotAutoSaveDuringDestruction(): void
    {
        $excel = $this->workbook();
        $excel->excelWriter->writer = new FailingSaveWriter();
        try { $excel->save(); } catch (\RuntimeException $e) {}
        $writer = $excel->excelWriter;
        $writer->__destruct();
        self::assertFileDoesNotExist($this->out);
    }
}

class FailingSaveWriter extends \avadim\FastExcelWriter\Writer\Writer
{
    public function _replaceSheets(string $inputFile, string $outputFile, array $relationShips): bool
    {
        file_put_contents($outputFile, 'partial');
        throw new \RuntimeException('injected write failure');
    }
}

class RacingSaveWriter extends TemplateWriter
{
    public string $destination;

    public function _replaceSheets(string $inputFile, string $outputFile, array $relationShips): bool
    {
        file_put_contents($this->destination, 'concurrent file');
        return parent::_replaceSheets($inputFile, $outputFile, $relationShips);
    }
}

<?php

declare(strict_types=1);

namespace avadim\FastExcelTemplator;

use avadim\FastExcelWriter\Excel as Writer;
use PHPUnit\Framework\TestCase;

/**
 * Two workbooks of different locales, alive at the same time, must not borrow each other's
 * settings. The style manager of the writer used to keep a static instance overwritten by
 * every constructor, and its static helpers read the default font and the locale from it,
 * so the workbook created last won: a 'ru' workbook wrote the euro sign. Fixed in the
 * writer 6.16.1, where the helpers take an explicit context.
 *
 * The templator reaches that code through RowTemplate::setCellStyle() / setRowStyle().
 */
final class LocaleIsolationTest extends TestCase
{
    private string $tpl;
    private string $outFirst;
    private string $outSecond;

    protected function setUp(): void
    {
        $name = preg_replace('/\W/', '', $this->getName());
        $dir = sys_get_temp_dir() . DIRECTORY_SEPARATOR;
        $this->tpl = $dir . 'fxt-locale-' . $name . '-tpl.xlsx';
        $this->outFirst = $dir . 'fxt-locale-' . $name . '-ru.xlsx';
        $this->outSecond = $dir . 'fxt-locale-' . $name . '-de.xlsx';
        foreach ([$this->tpl, $this->outFirst, $this->outSecond] as $file) {
            @unlink($file);
        }

        $writer = Writer::create(['Sheet1']);
        $writer->sheet()->writeRow([1234.5, 'text']);
        $writer->save($this->tpl);
    }

    protected function tearDown(): void
    {
        gc_collect_cycles();
        foreach ([$this->tpl, $this->outFirst, $this->outSecond] as $file) {
            @unlink($file);
        }
    }

    private function formatOf(string $file): ?string
    {
        $cells = \avadim\FastExcelReader\Excel::open($file)->readCellsWithStyles();

        return $cells['A1']['s']['format']['format-pattern'] ?? null;
    }

    /**
     * The workbook created first keeps its own currency even though a second one,
     * of another locale, was created after it and is still alive
     */
    public function testCurrencyOfOneWorkbookDoesNotLeakIntoAnother()
    {
        $ru = Excel::template($this->tpl, $this->outFirst, ['locale' => 'ru']);
        $de = Excel::template($this->tpl, $this->outSecond, ['locale' => 'de']);

        $money = function ($sourceRowNum, $targetRowNum, RowTemplate $row) {
            $row->setCellStyle('A', ['format' => '@money']);

            return $row;
        };

        $ru->sheet()->rows($money);
        $ru->save();

        $de->sheet()->rows($money);
        $de->save();

        self::assertStringContainsString('₽', (string)$this->formatOf($this->outFirst));
        self::assertStringNotContainsString('€', (string)$this->formatOf($this->outFirst));

        self::assertStringContainsString('€', (string)$this->formatOf($this->outSecond));
        self::assertStringNotContainsString('₽', (string)$this->formatOf($this->outSecond));
    }

    /**
     * The order in which the workbooks are filled must not matter either
     */
    public function testTheOrderOfFillingDoesNotMatter()
    {
        $ru = Excel::template($this->tpl, $this->outFirst, ['locale' => 'ru']);
        $de = Excel::template($this->tpl, $this->outSecond, ['locale' => 'de']);

        $money = function ($sourceRowNum, $targetRowNum, RowTemplate $row) {
            $row->setCellStyle('A', ['format' => '@money']);

            return $row;
        };

        // the one created last is filled first this time
        $de->sheet()->rows($money);
        $de->save();

        $ru->sheet()->rows($money);
        $ru->save();

        self::assertStringContainsString('₽', (string)$this->formatOf($this->outFirst));
        self::assertStringContainsString('€', (string)$this->formatOf($this->outSecond));
    }
}

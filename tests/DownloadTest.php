<?php

declare(strict_types=1);

namespace avadim\FastExcelTemplator;

use avadim\FastExcelWriter\Excel as Writer;
use PHPUnit\Framework\TestCase;

/**
 * Excel::download() and its alias output(): the generated file is streamed to the client
 * behind headers that describe it. The headers used to announce the file as the old .xls
 * ('application/vnd.ms-excel'), carried no length, and were sent even when the response had
 * already started.
 *
 * The headers cannot be observed under the CLI SAPI, so the test drives a subclass that
 * records what download() asked to send instead of calling header() for real.
 */
final class DownloadTest extends TestCase
{
    private string $tpl;

    protected function setUp(): void
    {
        $name = preg_replace('/\W/', '', $this->getName());
        $this->tpl = sys_get_temp_dir() . DIRECTORY_SEPARATOR . 'fxt-download-' . $name . '-tpl.xlsx';
        @unlink($this->tpl);

        $writer = Writer::create(['Sheet1']);
        $writer->sheet()->writeRow(['h1', 'h2']);
        $writer->sheet()->writeRow([1, 2]);
        $writer->save($this->tpl);
    }

    protected function tearDown(): void
    {
        @unlink($this->tpl);
    }

    /**
     * Runs download() (or output()) and returns [$headers, $body]
     */
    private function capture(string $method, ?string $name = null): array
    {
        /** @var DownloadTestExcel $excel */
        $excel = DownloadTestExcel::template($this->tpl);
        $excel->sheet()->transferRows();

        ob_start();
        try {
            $excel->$method($name);
        } finally {
            $body = (string)ob_get_clean();
        }

        return [$excel->sentHeaders, $body];
    }

    private function header(array $headers, string $name): string
    {
        foreach ($headers as $header) {
            if (stripos($header, $name . ':') === 0) {
                return trim(substr($header, strlen($name) + 1));
            }
        }

        self::fail('Header "' . $name . '" was not sent: ' . implode(' | ', $headers));
    }

    public function testDownloadSendsTheXlsxMimeType()
    {
        [$headers] = $this->capture('download', 'report.xlsx');

        self::assertSame(
            'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            $this->header($headers, 'Content-Type')
        );
        self::assertSame(Excel::MIME_TYPE, $this->header($headers, 'Content-Type'));
    }

    public function testDownloadSendsTheRealLengthOfTheBody()
    {
        [$headers, $body] = $this->capture('download', 'report.xlsx');

        self::assertSame((string)strlen($body), $this->header($headers, 'Content-Length'));
        self::assertNotSame('0', $this->header($headers, 'Content-Length'));
    }

    public function testDownloadedBodyIsTheGeneratedWorkbook()
    {
        [, $body] = $this->capture('download', 'report.xlsx');

        // a zip container, and a readable workbook inside it
        self::assertSame('PK', substr($body, 0, 2));

        self::assertSame(
            [1 => ['A' => 'h1', 'B' => 'h2'], 2 => ['A' => 1, 'B' => 2]],
            $this->rowsOf($body)
        );
    }

    /**
     * The name is reduced to its base name (a caller must not be able to send a path)
     * and gets the .xlsx extension when it has none
     */
    public function testFileNameIsNormalized()
    {
        [$headers] = $this->capture('download', 'reports/2026/report');

        self::assertSame('attachment; filename="report.xlsx"', $this->header($headers, 'Content-Disposition'));
    }

    public function testFileNameKeepsTheExtensionItAlreadyHas()
    {
        [$headers] = $this->capture('download', 'REPORT.XLSX');

        self::assertSame('attachment; filename="REPORT.XLSX"', $this->header($headers, 'Content-Disposition'));
    }

    public function testDefaultFileNameIsAnXlsx()
    {
        [$headers] = $this->capture('download');

        self::assertMatchesRegularExpression(
            '/^attachment; filename="[^"]+\.xlsx"$/',
            $this->header($headers, 'Content-Disposition')
        );
    }

    /**
     * output() is documented as an alias, so it must send the same headers and the same
     * workbook. The bytes themselves are not comparable: the writer stamps every sheet with
     * a fresh xr:uid, so two runs over the same template deflate to slightly different files.
     */
    public function testOutputIsAnAliasOfDownload()
    {
        [$downloadHeaders, $downloadBody] = $this->capture('download', 'report.xlsx');
        [$outputHeaders, $outputBody] = $this->capture('output', 'report.xlsx');

        $withoutLength = function (array $headers): array {
            return array_values(array_filter($headers, function (string $header): bool {
                return stripos($header, 'Content-Length:') !== 0;
            }));
        };

        self::assertSame($withoutLength($downloadHeaders), $withoutLength($outputHeaders));
        self::assertSame($this->rowsOf($downloadBody), $this->rowsOf($outputBody));
        self::assertNotEmpty($this->rowsOf($outputBody));
    }

    /**
     * Reads back a workbook received as a string of bytes
     */
    private function rowsOf(string $body): array
    {
        $file = sys_get_temp_dir() . DIRECTORY_SEPARATOR . 'fxt-download-' . md5($body) . '.xlsx';
        file_put_contents($file, $body);
        $rows = \avadim\FastExcelReader\Excel::open($file)->readRows();
        @unlink($file);

        return $rows;
    }

    /**
     * Once the response has started, the file can only arrive corrupted — say so
     * instead of appending a workbook to whatever was printed before
     */
    public function testDownloadThrowsWhenTheHeadersWereAlreadySent()
    {
        /** @var DownloadTestExcel $excel */
        $excel = DownloadTestExcel::template($this->tpl);
        $excel->pretendHeadersSent = true;

        $this->expectException(Exception::class);
        $this->expectExceptionMessage('the headers were already sent');

        $excel->download('report.xlsx');
    }

    public function testNothingIsSentWhenTheHeadersWereAlreadySent()
    {
        /** @var DownloadTestExcel $excel */
        $excel = DownloadTestExcel::template($this->tpl);
        $excel->pretendHeadersSent = true;

        ob_start();
        try {
            $excel->download('report.xlsx');
        }
        catch (Exception $e) {
            // expected, asserted in the test above
        }
        $body = (string)ob_get_clean();

        self::assertSame('', $body);
        self::assertSame([], $excel->sentHeaders);
    }
}

/**
 * Collects the headers instead of sending them, and can pretend the response has already started.
 */
class DownloadTestExcel extends Excel
{
    /** @var string[] */
    public array $sentHeaders = [];

    public bool $pretendHeadersSent = false;

    protected function sendHeaders(array $headers)
    {
        $this->sentHeaders = $headers;
    }

    protected function headersSent(?string &$file = null, ?int &$line = null): bool
    {
        if ($this->pretendHeadersSent) {
            $file = 'index.php';
            $line = 1;

            return true;
        }

        return parent::headersSent($file, $line);
    }
}

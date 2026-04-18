<?php

namespace LaravelEnso\Excel\Tests\Features;

require_once __DIR__.'/../TestCase.php';

use Illuminate\Support\Facades\Storage;
use LaravelEnso\Excel\Services\ExcelExport;
use LaravelEnso\Excel\Tests\Fixtures\BasicExport;
use LaravelEnso\Excel\Tests\Fixtures\CustomFolderExport;
use LaravelEnso\Excel\Tests\TestCase;
use OpenSpout\Reader\XLSX\Reader;
use PHPUnit\Framework\Attributes\Test;
use Symfony\Component\HttpFoundation\BinaryFileResponse;

class ExcelExportTest extends TestCase
{
    #[Test]
    public function saves_export_to_default_temp_folder(): void
    {
        $path = (new ExcelExport(new BasicExport()))->save();

        $this->assertSame(
            Storage::path('temp/basic-export.xlsx'),
            $path
        );
        $this->assertFileExists($path);
    }

    #[Test]
    public function saves_export_to_custom_folder_when_exporter_implements_saves_to_disk(): void
    {
        $path = (new ExcelExport(new CustomFolderExport()))->save();

        $this->assertSame(
            Storage::path('exports/reports/custom-folder-export.xlsx'),
            $path
        );
        $this->assertFileExists($path);
    }

    #[Test]
    public function creates_missing_storage_directory_before_writing(): void
    {
        Storage::deleteDirectory('exports/reports');

        $path = (new ExcelExport(new CustomFolderExport()))->save();

        $this->assertDirectoryExists(dirname($path));
        $this->assertFileExists($path);
    }

    #[Test]
    public function returns_download_response_for_inline_exports(): void
    {
        $response = (new ExcelExport(new BasicExport()))->inline();

        $this->assertInstanceOf(BinaryFileResponse::class, $response);
        $this->assertSame('basic-export.xlsx', $response->getFile()->getFilename());
        $this->assertSame(
            'attachment; filename=basic-export.xlsx',
            $response->headers->get('content-disposition')
        );
    }

    #[Test]
    public function writes_all_declared_sheets_headings_and_rows(): void
    {
        $path = (new ExcelExport(new BasicExport()))->save();

        $reader = new Reader();
        $reader->open($path);

        $sheets = array_values(iterator_to_array($reader->getSheetIterator()));

        $this->assertCount(2, $sheets);
        $this->assertSame('stock', $sheets[0]->getName());
        $this->assertSame('history', $sheets[1]->getName());

        $stockRows = array_values(array_map(
            fn ($row) => $row->toArray(),
            iterator_to_array($sheets[0]->getRowIterator())
        ));
        $historyRows = array_values(array_map(
            fn ($row) => $row->toArray(),
            iterator_to_array($sheets[1]->getRowIterator())
        ));

        $reader->close();

        $this->assertSame([
            ['Product', 'Quantity'],
            ['Panel', 12],
            ['Inverter', 4],
        ], $stockRows);

        $this->assertSame([
            ['Date', 'Product', 'Quantity'],
            ['2026-04-18', 'Panel', 12],
        ], $historyRows);
    }
}

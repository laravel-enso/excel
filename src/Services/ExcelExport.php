<?php

namespace LaravelEnso\Excel\Services;

use Illuminate\Support\Collection;
use Illuminate\Support\Facades\Response;
use Illuminate\Support\Facades\Storage;
use LaravelEnso\Excel\Contracts\ExportsExcel;
use LaravelEnso\Excel\Contracts\SavesToDisk;
use OpenSpout\Common\Entity\Row;
use OpenSpout\Writer\XLSX\Writer;
use Symfony\Component\HttpFoundation\BinaryFileResponse;

class ExcelExport
{
    private Writer $writer;

    public function __construct(private ExportsExcel $exporter)
    {
        $this->writer = new Writer();
    }

    public function inline(): BinaryFileResponse
    {
        $this->handle();
        $args = [$this->path(), $this->exporter->filename()];

        return Response::download(...$args)->deleteFileAfterSend();
    }

    public function save(): string
    {
        $this->handle();

        return $this->path();
    }

    private function handle(): void
    {
        $this->writer->openToFile($this->path());

        Collection::wrap($this->exporter->sheets())
            ->each(fn ($sheet, $index) => $this
                ->sheet($sheet, $index)
                ->heading($sheet)
                ->rows($sheet));

        $this->writer->close();
    }

    private function sheet(string $sheet, int $index): self
    {
        if ($index > 0) {
            $this->writer->addNewSheetAndMakeItCurrent();
        }

        $this->writer->getCurrentSheet()->setName($sheet);

        return $this;
    }

    private function heading(string $sheet): self
    {
        $this->writer->addRow($this->row($this->exporter->heading($sheet)));

        return $this;
    }

    private function rows(string $sheet): self
    {
        foreach ($this->exporter->rows($sheet) as $row) {
            $this->writer->addRow($this->row($row));
        }

        return $this;
    }

    private function row(array $data): Row
    {
        return Row::fromValues($data);
    }

    private function path(): string
    {
        $folder = $this->exporter instanceof SavesToDisk
            ? $this->exporter->folder()
            : 'temp';

        if (! Storage::exists($folder)) {
            Storage::makeDirectory($folder);
        }

        return Storage::path("{$folder}/{$this->exporter->filename()}");
    }
}

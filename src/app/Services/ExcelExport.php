<?php

namespace LaravelEnso\Excel\app\Exports;

use Box\Spout\Writer\Common\Creator\Style\StyleBuilder;
use Box\Spout\Writer\Common\Creator\WriterEntityFactory;
use LaravelEnso\Excel\app\Contracts\ExportsExcel;
use LaravelEnso\Excel\app\Contracts\SavesToDisk;
use UnexpectedValueException;

class ExcelExport
{
    private $writer;
    private $exporter;
    private $inline;

    public function __construct(ExportsExcel $exporter)
    {
        $this->exporter = $exporter;
        $this->inline = true;
    }

    public function inline()
    {
        $this->handle();

        return $this->exporter->filename();
    }

    public function save()
    {
        if (! $this->exporter instanceof SavesToDisk) {
            throw new UnexpectedValueException('User must implement SavesToDisk interface.');
        }

        $this->inline = false;

        $this->handle();

        return $this->filePath();
    }

    private function handle()
    {
        $this->writer()
            ->addHeading()
            ->addRows();

        $this->writer->close();
    }

    private function writer()
    {
        $defaultStyle = (new StyleBuilder())
            ->setShouldWrapText(false)
            ->build();

        $this->writer = WriterEntityFactory::createXLSXWriter();

        $this->writer->setDefaultRowStyle($defaultStyle);

        if ($this->inline) {
            $this->writer->openToBrowser($this->exporter->filename());

            return $this;
        }

        $this->writer->openToFile($this->filePath());

        return $this;
    }

    private function addHeading()
    {
        $this->writer->addRow($this->row($this->exporter->heading()));

        return $this;
    }

    private function addRows()
    {
        collect($this->exporter->rows())->each(fn($row) => (
            $this->writer->addRow($this->row($row))
        ));

        return $this;
    }

    private function row($data)
    {
        return WriterEntityFactory::createRowFromArray($data);
    }

    private function filePath()
    {
        return $this->exporter->path()
            .DIRECTORY_SEPARATOR
            .$this->exporter->filename();
    }
}

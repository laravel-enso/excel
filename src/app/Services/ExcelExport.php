<?php
/**
 * Created with luv for rolcrispa.
 * User: mihai
 * Date: 10/1/19
 * Time: 5:59 PM
 */

namespace LaravelEnso\Excel\app\Exports;

use UnexpectedValueException;
use LaravelEnso\Excel\app\Contracts\SavesToDisk;
use LaravelEnso\Excel\app\Contracts\ExportsExcel;
use Box\Spout\Writer\Common\Creator\Style\StyleBuilder;
use Box\Spout\Writer\Common\Creator\WriterEntityFactory;

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
    }

    public function save()
    {
        if (! $this->exporter instanceof SavesToDisk) {
            throw new UnexpectedValueException('User must implement SavesToDisk interface.');
        }

        $this->inline = false;

        $this->handle();
    }

    private function handle()
    {
        $this->setWriter()
            ->addHeading()
            ->addRows();

        $this->writer->close();
    }

    private function setWriter()
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

        $this->writer->openToFile(
            $this->exporter->filePath()
            .DIRECTORY_SEPARATOR
            .$this->exporter->filename()
        );

        return $this;
    }

    private function addHeading()
    {
        $this->writer->addRow($this->row($this->exporter->heading()));

        return $this;
    }

    private function addRows()
    {
        collect($this->exporter->rows())->each(function ($row) {
            $this->writer->addRow($this->row($row));
        });

        return $this;
    }

    private function row($data)
    {
        return WriterEntityFactory::createRowFromArray($data);
    }
}

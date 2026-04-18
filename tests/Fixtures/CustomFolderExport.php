<?php

namespace LaravelEnso\Excel\Tests\Fixtures;

use LaravelEnso\Excel\Contracts\SavesToDisk;

class CustomFolderExport extends BasicExport implements SavesToDisk
{
    public function filename(): string
    {
        return 'custom-folder-export.xlsx';
    }

    public function folder(): string
    {
        return 'exports/reports';
    }
}

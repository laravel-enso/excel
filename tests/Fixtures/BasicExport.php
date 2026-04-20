<?php

namespace LaravelEnso\Excel\Tests\Fixtures;

use LaravelEnso\Excel\Contracts\ExportsExcel;

class BasicExport implements ExportsExcel
{
    public function filename(): string
    {
        return 'basic-export.xlsx';
    }

    public function heading(string $sheet): array
    {
        return match ($sheet) {
            'stock'   => ['Product', 'Quantity'],
            'history' => ['Date', 'Product', 'Quantity'],
        };
    }

    public function rows(string $sheet): array
    {
        return match ($sheet) {
            'stock' => [
                ['Panel', 12],
                ['Inverter', 4],
            ],
            'history' => [
                ['2026-04-18', 'Panel', 12],
            ],
        };
    }

    public function sheets(): array
    {
        return ['stock', 'history'];
    }
}

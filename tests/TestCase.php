<?php

namespace LaravelEnso\Excel\Tests;

use Illuminate\Support\Facades\File;
use Illuminate\Support\Facades\Storage;
use Tests\TestCase as BaseTestCase;

abstract class TestCase extends BaseTestCase
{
    protected function setUp(): void
    {
        parent::setUp();

        $this->loadFixtures();
        Storage::deleteDirectory('temp');
        Storage::deleteDirectory('exports');
    }

    protected function tearDown(): void
    {
        Storage::deleteDirectory('temp');
        Storage::deleteDirectory('exports');

        parent::tearDown();
    }

    protected function loadFixtures(): void
    {
        foreach (File::allFiles(__DIR__.'/Fixtures') as $file) {
            if ($file->getExtension() === 'php') {
                require_once $file->getRealPath();
            }
        }
    }
}

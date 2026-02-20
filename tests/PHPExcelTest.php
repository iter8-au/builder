<?php

namespace Builder\Tests;

use Builder\Builder;
use Builder\Builders\PhpSpreadsheet;
use PHPUnit\Framework\TestCase;

class PHPExcelTest extends TestCase
{
    public function test_builder_is_correct_builder(): void
    {
        // Arrange
        $builder = $this->makeBuilder();

        // Assert
        $this->assertInstanceOf(PhpSpreadsheet::class, $builder->getBuilder());
    }

    public function test_can_create_single_sheet_spreadsheet(): void
    {
        // Arrange
        $builder = $this->makeBuilder();

        // Act
        $builder->setSheetTitles('PHPExcel Test');
        $builder->setData(
            [
                [
                    'Column 1' => 'column_1',
                    'Column 2' => 'column_2',
                    'Column 3' => 'column_3',
                ],
                [
                    'Column 1' => '1',
                    'Column 2' => 'Two',
                    'Column 3' => '333'
                ],
                [
                    'Column 1' => 'One',
                    'Column 2' => '2',
                    'Column 3' => 'Three x 3'
                ],
            ]
        );
        $builder->generateExcel();

        $generatedExcelFile = $builder->getTempName();

        // Assert
        $this->assertFileExists($generatedExcelFile);
        $this->assertGreaterThan(3000, stat($generatedExcelFile)['size']);
    }

    public function test_can_create_multi_sheet_spreadsheet(): void
    {
        // Arrange
        $builder = $this->makeBuilder();

        // Act
        $builder->setSheetTitles(
            [
                'Sheet 1 of 2',
                'Sheet 2 of 2',
            ]
        );
        $builder->setSheets(
            [
                [
                    [
                        'Column 1' => 'Row 1',
                        'Column 2' => 'Sheet 1',
                    ],
                ],
                [
                    [
                        'Column 1' => 'Row 2',
                        'Column 2' => 'Sheet 2',
                    ],
                    [
                        'Column 1' => 'Row 3',
                        'Column 2' => 'Sheet 2',
                    ],
                ],
            ]
        );
        $builder->generateExcel();

        $generatedExcelFile = $builder->getTempName();

        // Assert
        $this->assertFileExists($generatedExcelFile);
        $this->assertGreaterThan(3000, stat($generatedExcelFile)['size']);
    }

    private function makeBuilder(): Builder
    {
        return new Builder(
            $this->makePhpSpreadsheet(),
            $this->getCacheDir()
        );
    }

    private function makePhpSpreadsheet(): PhpSpreadsheet
    {
        return new PhpSpreadsheet();
    }

    public static function getCacheDir(): string
    {
        return __DIR__ . '/cache/phpexcel';
    }

    /**
     * Create the cache folder required for testing.
     */
    public static function setUpBeforeClass(): void
    {
        if (is_dir(self::getCacheDir()) === false) {
            mkdir(self::getCacheDir());
        }
    }

    /**
     * Remove the cache folder required for testing.
     */
    public static function tearDownAfterClass(): void
    {
        if (is_dir(self::getCacheDir()) === true) {
            self::removeTempFilesAndDirectory(self::getCacheDir());
        }
    }

    // @see https://stackoverflow.com/questions/3349753/delete-directory-with-files-in-it
    public static function removeTempFilesAndDirectory(string $directory): void
    {
        $it = new \RecursiveDirectoryIterator($directory, \RecursiveDirectoryIterator::SKIP_DOTS);
        $files = new \RecursiveIteratorIterator(
            $it,
            \RecursiveIteratorIterator::CHILD_FIRST
        );
        foreach($files as $file) {
            if ($file->getExtension() === 'tmp') {
                unlink($file->getPathname());
            }
        }
        rmdir($directory);
    }
}

<?php

namespace Tests;

use Box\Spout\Common\Helper\FileSystemHelper;
use Box\Spout\Common\Type;
use Box\Spout\Reader\ReaderFactory;
use Box\Spout\Writer\Style\Color;
use Box\Spout\Writer\Style\Style;
use Box\Spout\Writer\Style\StyleBuilder;
use Builder\Builders\SpoutBuilder;
use Builder\Interfaces\BuilderInterface;
use PHPUnit\Framework\TestCase;

class SpoutTest extends TestCase
{
    public function test_builder_is_correct_builder()
    {
        // Arrange
        $app = new Application();
        $app['builder.default'] = 'spout';
        $app['builder.cache_dir'] = __DIR__.'/cache/spout';

        // Act
        $app->register(new BuilderServiceProvider());

        // Assert
        $this->assertInstanceOf(SpoutBuilder::class, $app['builder']->getBuilder());
    }

    public function test_row_style_is_parsed_correctly()
    {
        // Arrange
        $app = new Application();
        $app->register(new BuilderServiceProvider(), [
            'builder.default' => 'spout',
            'builder.cache_dir' => $this->getCacheDir(),
        ]);

        $styleArray = [
            'font' => [
                'color' => [
                    'rgb' => BuilderInterface::COLOUR_BLACK_RGB,
                ],
                'bold' => true,
            ],
        ];

        $styleBuilder = new StyleBuilder();
        $styleBuilder->setFontColor(Color::toARGB(BuilderInterface::COLOUR_BLACK_RGB))
                     ->setFontBold();

        // Act
        /** @var \Box\Spout\Writer\Style\Style $style */
        $style = $app['builder']->getBuilder()->buildRowStyle($styleArray);
        $styleBuilderStyle = $styleBuilder->build();

        // Assert
        $this->assertInstanceOf(Style::class, $style);
        $this->assertSame($style->serialize(), $styleBuilderStyle->serialize());
    }

    public function test_can_create_single_sheet_spreadsheet()
    {
        // Arrange
        $app = new Application();
        $app->register(new BuilderServiceProvider(), [
            'builder.default' => 'spout',
            'builder.cache_dir' => $this->getCacheDir(),
        ]);
        $reader = ReaderFactory::create(Type::XLSX);

        // Act
        $app['builder']->setSheetTitles('Spout Test');
        $app['builder']->setHeaders([
            'Column 1',
            'Column 2',
            'Column 3',
        ]);
        $app['builder']->setData(
            [
                ['column_1', 'column_2', 'column_3'],
                ['1', 'Two', '333'],
                ['One', '2', 'Three x 3'],
            ]
        );
        $app['builder']->generateExcel();

        $generatedExcelFile = $app['builder']->getTempName();
        $reader->open($generatedExcelFile);

        // Need to use `iterator_to_array` as it's the only way to coerce Spout to read the spreadsheet into memory
        // unless you manually do a foreach over the rows.
        $sheets = iterator_to_array($reader->getSheetIterator());
        // Sheets array is *NOT* zero-based when fetched from the iterator.
        $sheetData = iterator_to_array($sheets[1]->getRowIterator());
        $headers = array_shift($sheetData);
        $rows = $sheetData;
        $row1 = $rows[0];
        $row3 = $rows[2];

        // Assert
        $this->assertFileExists($generatedExcelFile);
        $this->assertGreaterThan(3000, stat($generatedExcelFile)['size']);

        $this->assertCount(3, $headers, sprintf('Headers row should have 3 values, "%d" supplied.', \count($headers)));
        $this->assertEquals('Column 1', $headers[0]);
        $this->assertEquals('Column 2', $headers[1]);
        $this->assertEquals('Column 3', $headers[2]);

        $this->assertCount(3, $rows, sprintf('Rows should have 3 rows, "%d" supplied.', \count($rows)));

        $this->assertCount(3, $row1, sprintf('Row 1 should have 3 values, "%d" supplied.', \count($row1)));
        $this->assertEquals('column_1', $row1[0]);
        $this->assertEquals('column_2', $row1[1]);
        $this->assertEquals('column_3', $row1[2]);

        $this->assertCount(3, $row3, sprintf('Row 3 should have 3 values, "%d" supplied.', \count($row3)));
        $this->assertEquals('One', $row3[0]);
        $this->assertEquals('2', $row3[1]);
        $this->assertEquals('Three x 3', $row3[2]);
    }

    public function test_can_create_multi_sheet_spreadsheet()
    {
        // Arrange
        $app = new Application();
        $app->register(new BuilderServiceProvider(), [
            'builder.default' => 'spout',
            'builder.cache_dir' => $this->getCacheDir(),
        ]);

        // Act
        $app['builder']->setSheetTitles(
            [
                'Sheet 1 of 2',
                'Sheet 2 of 2',
            ]
        );
        $app['builder']->setSheets(
            [
                [
                    'headers' => [
                        'Column 1',
                        'Column 2',
                     ],
                    'rows' => [
                        [
                            'Row 1',
                            'Sheet 1',
                        ],
                    ],
                ],
                [
                    'headers' => [
                        'Column 1',
                        'Column 2',
                    ],
                    'rows' => [
                        [
                            'Row 2',
                            'Sheet 2',
                        ],
                        [
                            'Row 3',
                            'Sheet 2',
                        ],
                    ],
                ],
            ]
        );
        $app['builder']->generateExcel();

        $generatedExcelFile = $app['builder']->getTempName();

        // Assert
        $this->assertFileExists($generatedExcelFile);
        $this->assertGreaterThan(3000, stat($generatedExcelFile)['size']);
    }

    /**
     * @return string
     */
    private function getCacheDir()
    {
        return __DIR__.'/cache/spout';
    }

    /**
     * Create the cache folder required for testing.
     */
    public static function setUpBeforeClass(): void
    {
        $fileSystemHelper = new FileSystemHelper(__DIR__.'/cache');

        if (false === is_dir(__DIR__.'/cache/spout')) {
            $fileSystemHelper->createFolder(__DIR__.'/cache', 'spout');
        }
    }

    /**
     * Remove the cache folder required for testing.
     */
    public static function tearDownAfterClass(): void
    {
        $fileSystemHelper = new FileSystemHelper(__DIR__.'/cache');

        if (true === is_dir(__DIR__.'/cache/spout')) {
            $fileSystemHelper->deleteFolderRecursively(__DIR__.'/cache/spout');
        }
    }
}

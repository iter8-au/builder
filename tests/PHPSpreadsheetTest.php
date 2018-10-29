<?php

namespace Tests\Builder;

use Box\Spout\Common\Helper\FileSystemHelper;
use Builder\Builders\PhpSpreadsheet;
use Builder\Interfaces\BuilderTestInterface;
use Builder\Provider\Silex\BuilderServiceProvider;
use PhpOffice\PhpSpreadsheet\Reader\Xlsx;
use PHPUnit\Framework\TestCase;
use Silex\Application;

/**
 * Class PHPExcelTest
 */
class PHPSpreadsheetTest extends TestCase implements BuilderTestInterface
{
    /**
     * @test
     */
    public function builder_is_correct_builder()
    {
        // Arrange
        $app = new Application();
        $app['builder.default']   = 'phpspreadsheet';
        $app['builder.cache_dir'] = __DIR__ . '/cache/phpspreadsheet';

        // Act
        $app->register(new BuilderServiceProvider());

        // Assert
        $this->assertInstanceOf(PhpSpreadsheet::class, $app['builder']->getBuilder());
    }

    /**
     * @test
     */
    public function can_create_single_sheet_spreadsheet()
    {
        // Arrange
        $app = new Application();
        $app->register(new BuilderServiceProvider(), [
            'builder.default'   => 'phpspreadsheet',
            'builder.cache_dir' => $this->getCacheDir(),
        ]);
        $reader = new Xlsx();
        $reader->setReadDataOnly(true);

        // Act
        $app['builder']->setSheetTitles('PHPExcel Test');
        $app['builder']->setHeaders(
            [
                'Column 1',
                'Column 2',
                'Column 3',
            ]
        );
        $app['builder']->setData(
            [
                [
                    'column_1',
                    'column_2',
                    'column_3',
                ],
                [
                    '1',
                    'Two',
                    '333'
                ],
                [
                    'One',
                    '2',
                    'Three x 3'
                ],
            ]
        );
        $app['builder']->generateExcel();

        $generatedExcelFile = $app['builder']->getTempName();
        $sheet = $reader->load($generatedExcelFile);
        $sheetData = $sheet->getActiveSheet()->toArray(null, true, true, true);
        $headers = array_shift($sheetData);
        $rows = $sheetData;
        $row1 = $rows[0];
        $row3 = $rows[2];

        // Assert
        $this->assertFileExists($generatedExcelFile);
        $this->assertGreaterThan(3000, stat($generatedExcelFile)['size']);

        $this->assertCount(3, $headers, sprintf('Headers row should have 3 values, "%d" supplied.', count($headers)));
        $this->assertEquals('Column 1', $headers['A']);
        $this->assertEquals('Column 2', $headers['B']);
        $this->assertEquals('Column 3', $headers['C']);

        $this->assertCount(3, $rows, sprintf('Rows should have 3 rows, "%d" supplied.', count($rows)));

        $this->assertCount(3, $row1, sprintf('Row 1 should have 3 values, "%d" supplied.', count($row1)));
        $this->assertEquals('column_1', $row1['A']);
        $this->assertEquals('column_2', $row1['B']);
        $this->assertEquals('column_3', $row1['C']);

        $this->assertCount(3, $row3, sprintf('Row 3 should have 3 values, "%d" supplied.', count($row3)));
        $this->assertEquals('One', $row3['A']);
        $this->assertEquals('2', $row3['B']);
        $this->assertEquals('Three x 3', $row3['C']);
    }

    /**
     * @test
     */
    public function can_create_multi_sheet_spreadsheet()
    {
        // Arrange
        $app = new Application();
        $app->register(new BuilderServiceProvider(), [
            'builder.default'   => 'phpspreadsheet',
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
                        ]
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
                    ]
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
        return __DIR__ . '/cache/phpspreadsheet';
    }

    /**
     * Create the cache folder required for testing.
     */
    public static function setUpBeforeClass()
    {
        $fileSystemHelper = new FileSystemHelper(__DIR__ . '/cache');

        if (is_dir(__DIR__ . '/cache/phpspreadsheet') === false) {
            $fileSystemHelper->createFolder(__DIR__ . '/cache', 'phpspreadsheet');
        }
    }

    /**
     * Remove the cache folder required for testing.
     */
    public static function tearDownAfterClass()
    {
        $fileSystemHelper = new FileSystemHelper(__DIR__ . '/cache');

        if (is_dir(__DIR__ . '/cache/phpspreadsheet') === true) {
            $fileSystemHelper->deleteFolderRecursively(__DIR__ . '/cache/phpspreadsheet');
        }
    }
}

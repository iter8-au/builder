<?php

namespace Tests\Builder;

use Box\Spout\Common\Helper\FileSystemHelper;
use Builder\Builders\PhpSpreadsheet;
use Builder\Interfaces\BuilderTestInterface;
use Builder\Provider\Silex\BuilderServiceProvider;
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

        // Assert
        $this->assertFileExists($generatedExcelFile);
        $this->assertGreaterThan(3000, stat($generatedExcelFile)['size']);
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

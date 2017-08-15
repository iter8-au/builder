<?php

namespace Builder\Tests;

use Builder\Interfaces\BuilderTestInterface;
use Builder\Provider\Silex\BuilderServiceProvider;
use Silex\Application;
use PHPUnit_Framework_TestCase;
use Box\Spout\Common\Helper\FileSystemHelper;

/**
 * Class PHPExcelTest
 * @package Builder\Tests
 */
class PHPExcelTest extends PHPUnit_Framework_TestCase implements BuilderTestInterface
{
    /**
     * @test
     */
    public function builder_is_correct_builder()
    {
        // Arrange
        $app = new Application();
        $app['builder.default']   = 'phpexcel';
        $app['builder.cache_dir'] = __DIR__ . '/cache/phpexcel';

        // Act
        $app->register(new BuilderServiceProvider());

        // Assert
        $this->assertInstanceOf('Builder\Builders\PHPExcelBuilder', $app['builder']->getBuilder());
    }

    /**
     * @test
     */
    public function can_create_single_sheet_spreadsheet()
    {
        // Arrange
        $app = new Application();
        $app->register(new BuilderServiceProvider(), [
            'builder.default'   => 'phpexcel',
            'builder.cache_dir' => $this->getCacheDir(),
        ]);

        // Act
        $app['builder']->setSheetTitles('PHPExcel Test');
        $app['builder']->setData(
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
            'builder.default'   => 'phpexcel',
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
        return __DIR__ . '/cache/phpexcel';
    }

    /**
     * Create the cache folder required for testing.
     */
    public static function setUpBeforeClass()
    {
        $fileSystemHelper = new FileSystemHelper(__DIR__ . '/cache');

        if (is_dir(__DIR__ . '/cache/phpexcel') === false) {
            $fileSystemHelper->createFolder(__DIR__ . '/cache', 'phpexcel');
        }
    }

    /**
     * Remove the cache folder required for testing.
     */
    public static function tearDownAfterClass()
    {
        $fileSystemHelper = new FileSystemHelper(__DIR__ . '/cache');

        if (is_dir(__DIR__ . '/cache/phpexcel') === true) {
            $fileSystemHelper->deleteFolderRecursively(__DIR__ . '/cache/phpexcel');
        }
    }
}

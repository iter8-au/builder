<?php

namespace Builder\Tests;

use Builder\Interfaces\BuilderTestInterface;
use Builder\Provider\BuilderServiceProvider;
use Silex\Application;
use PHPUnit_Framework_TestCase;

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
        $app['builder.driver'] = 'phpexcel';
        $app['builder.cache_dir'] = __DIR__ . '/cache/phpexcel';

        // Act
        $app->register(new BuilderServiceProvider());

        // Assert
        $this->assertInstanceOf('Builder\PHPExcelBuilder', $app['builder']->getBuilder());
    }

    public function can_create_single_sheet_spreadsheet()
    {
        // Arrange
        $app = new Application();
        $app->register(new BuilderServiceProvider(), [
            'builder.driver'    => 'phpexcel',
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
                    'column_1' => '1',
                    'column_2' => 'Two',
                    'column_3' => '333'
                ],
                [
                    'column_1' => 'One',
                    'column_2' => '2',
                    'column_3' => 'Three x 3'
                ],
            ]
        );
        $app['builder']->generateExcel();

        $generatedExcelFile = $app['builder']->getTempName();

        // Assert
        $this->assertFileExists($generatedExcelFile);
        $this->assertGreaterThan(3000, stat($generatedExcelFile)['size']);
    }

    public function can_create_multi_sheet_spreadsheet()
    {
        //
    }

    /**
     * @return string
     */
    private function getCacheDir()
    {
        return __DIR__ . '/cache/phpexcel';
    }
}

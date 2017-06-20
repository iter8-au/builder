<?php

namespace Builder\Tests;

use Builder\Interfaces\BuilderInterface;
use Builder\Interfaces\BuilderTestInterface;
use Builder\Provider\BuilderServiceProvider;
use Silex\Application;
use Box\Spout\Common\Type;
use PHPUnit_Framework_TestCase;
use Box\Spout\Writer\Style\Color;
use Box\Spout\Reader\ReaderFactory;
use Box\Spout\Writer\Style\StyleBuilder;
use Box\Spout\Common\Helper\FileSystemHelper;

/**
 * Class SpoutTest
 * @package Builder\Tests
 */
class SpoutTest extends PHPUnit_Framework_TestCase implements BuilderTestInterface
{
    /**
     * @test
     */
    public function builder_is_correct_builder()
    {
        // Arrange
        $app = new Application();
        $app['builder.driver'] = 'spout';
        $app['builder.cache_dir'] = __DIR__ . '/cache/spout';

        // Act
        $app->register(new BuilderServiceProvider());

        // Assert
        $this->assertInstanceOf('Builder\SpoutBuilder', $app['builder']->getBuilder());
    }

    /**
     * @test
     */
    public function row_style_is_parsed_correctly()
    {
        // Arrange
        $app = new Application();
        $app->register(new BuilderServiceProvider(), [
            'builder.driver'    => 'spout',
            'builder.cache_dir' => $this->getCacheDir(),
        ]);

        $styleArray = [
            'font' => [
                'color' => [
                    'rgb' => BuilderInterface::COLOUR_BLACK_RGB,
                ],
                'bold' => true,
            ],
            'fill' => [
                'color' => [
                    'rgb' => BuilderInterface::COLOUR_WHITE_RGB,
                ],
            ],
        ];

        $styleBuilder = new StyleBuilder();
        $styleBuilder->setFontColor(Color::toARGB(BuilderInterface::COLOUR_BLACK_RGB))
                     ->setFontBold()
                     ->setBackgroundColor(Color::toARGB(BuilderInterface::COLOUR_WHITE_RGB));

        // Act
        /** @var \Box\Spout\Writer\Style\Style $style */
        $style = $app['builder']->getBuilder()->buildRowStyle($styleArray);
        $styleBuilderStyle = $styleBuilder->build();

        // Assert
        $this->assertInstanceOf('Box\Spout\Writer\Style\Style', $style);
        $this->assertSame($style->serialize(), $styleBuilderStyle->serialize());
    }

    /**
     * @test
     *
     * TODO: Open the file with the Reader and verify a row/column value.
     */
    public function can_create_single_sheet_spreadsheet()
    {
        // Arrange
        $app = new Application();
        $app->register(new BuilderServiceProvider(), [
            'builder.driver'    => 'spout',
            'builder.cache_dir' => $this->getCacheDir(),
        ]);
        // $reader = ReaderFactory::create(Type::XLSX);

        // Act
        $app['builder']->setSheetTitles('Spout Test');
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
//        $reader->open($generatedExcel);
//
//        $sheets = $reader->getSheetIterator();
//
//        $sheets->rewind();
//
//        $sheet = $sheets->current()->getRowIterator()->next()->current();
//        die(var_dump($sheet));
    }

    /**
     * @test
     */
    public function can_create_multi_sheet_spreadsheet()
    {
        // Arrange
        $app = new Application();
        $app->register(new BuilderServiceProvider(), [
            'builder.driver'    => 'spout',
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
                        'Column 1' => 'column_1',
                        'Column 2' => 'column_2',
                    ],
                    [
                        'Row 1',
                        'Sheet 1',
                    ],
                ],
                [
                    [
                        'Column 1' => 'column_1',
                        'Column 2' => 'column_2',
                    ],
                    [
                        'Row 2',
                        'Sheet 2',
                    ],
                    [
                        'Row 3',
                        'Sheet 2',
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
        return __DIR__ . '/cache/spout';
    }

    /**
     * Create the cache folder required for testing.
     */
    public static function setUpBeforeClass()
    {
        $fileSystemHelper = new FileSystemHelper(__DIR__ . '/cache');

        if (is_dir(__DIR__ . '/cache/spout') === false) {
            $fileSystemHelper->createFolder(__DIR__ . '/cache', 'spout');
        }
    }

    /**
     * Remove the cache folder required for testing.
     */
    public static function tearDownAfterClass()
    {
        $fileSystemHelper = new FileSystemHelper(__DIR__ . '/cache');

        if (is_dir(__DIR__ . '/cache/spout') === true) {
            $fileSystemHelper->deleteFolderRecursively(__DIR__ . '/cache/spout');
        }
    }
}

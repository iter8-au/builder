<?php

namespace Tests;

use Box\Spout\Common\Helper\FileSystemHelper;
use Box\Spout\Writer\Style\Color;
use Box\Spout\Writer\Style\Style;
use Box\Spout\Writer\Style\StyleBuilder;
use Builder\Builder;
use Builder\Builders\SpoutBuilder;
use Builder\Interfaces\BuilderInterface;
use Builder\Provider\Silex\BuilderServiceProvider;
use PHPUnit\Framework\TestCase;
use Silex\Application;
use Tests\Interfaces\BuilderTestInterface;

/**
 * Class SpoutTest
 */
class SpoutTest extends TestCase implements BuilderTestInterface
{
    /**
     * @test
     */
    public function builder_is_correct_builder()
    {
        // Arrange
        $app = new Application();
        $app['builder.default']   = 'spout';
        $app['builder.cache_dir'] = __DIR__ . '/cache/spout';

        // Act
        $app->register(new BuilderServiceProvider());

        // Assert
        $this->assertInstanceOf(SpoutBuilder::class, $app['builder']->getBuilder());
    }

    /**
     * @test
     */
    public function row_style_is_parsed_correctly()
    {
        // Arrange
        $app = new Application();
        $app->register(new BuilderServiceProvider(), [
            'builder.default'   => 'spout',
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
            'builder.default' => 'spout',
            'builder.cache_dir' => $this->getCacheDir(),
        ]);
        // $reader = ReaderFactory::create(Type::XLSX);
        /** @var Builder $builder */
        $builder = $app['builder'];

        // Act
        $builder->setFilename('test_name');
        $builder->setSheetTitles('Spout Test');
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

        $generatedExcelFile = $builder->getFilename();

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
            'builder.default'   => 'spout',
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
            //$fileSystemHelper->deleteFolderRecursively(__DIR__ . '/cache/spout');
        }
    }
}

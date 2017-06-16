<?php

namespace Builder\Tests;

use Builder\SpoutBuilder;
use Builder\BuilderInterface;
use Builder\Provider\BuilderServiceProvider;
use Silex\Application;
use Box\Spout\Writer\Style\Color;
use Box\Spout\Writer\Style\Style;
use Box\Spout\Writer\Style\StyleBuilder;
use Box\Spout\Common\Helper\FileSystemHelper;

/**
 * Class SpoutTest
 * @package Builder\Tests
 */
class SpoutTest extends \PHPUnit_Framework_TestCase
{
    /**
     * @test
     */
    public function builder_is_spout_builder()
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

        $fileSystemHelper->createFolder(__DIR__ . '/cache', 'spout');
    }

    /**
     * Remove the cache folder required for testing.
     */
    public static function tearDownAfterClass()
    {
        $fileSystemHelper = new FileSystemHelper(__DIR__ . '/cache');

        $fileSystemHelper->deleteFolderRecursively(__DIR__ . '/cache/spout');
    }
}

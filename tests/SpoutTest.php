<?php

namespace Builder\Tests;

use Box\Spout\Writer\Style\Color;
use Box\Spout\Writer\Style\StyleBuilder;
use Builder\SpoutBuilder;
use Builder\BuilderInterface;
use Builder\Provider\BuilderServiceProvider;
use Box\Spout\Writer\Style\Style;
use Silex\Application;

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
        $app = new Application();
        $app['builder.driver'] = 'spout';
        $app['builder.cache_dir'] = __DIR__ . '/cache';

        $app->register(new BuilderServiceProvider());

        $this->assertInstanceOf('Builder\SpoutBuilder', $app['builder']->getBuilder());
    }

    /**
     * @test
     */
    public function row_style_is_parsed_correctly()
    {
        $app = new Application();
        $app->register(new BuilderServiceProvider(), [
            'builder.driver' => 'spout',
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
        /** @var \Box\Spout\Writer\Style\Style $style */
        $style = $app['builder']->getBuilder()->buildRowStyle($styleArray);

        $styleBuilder = new StyleBuilder();
        $styleBuilder->setFontColor(Color::toARGB(BuilderInterface::COLOUR_BLACK_RGB))
                     ->setFontBold()
                     ->setBackgroundColor(Color::toARGB(BuilderInterface::COLOUR_WHITE_RGB));
        $styleBuilderStyle = $styleBuilder->build();

        $this->assertInstanceOf('Box\Spout\Writer\Style\Style', $style);
        $this->assertSame($style->serialize(), $styleBuilderStyle->serialize());
    }

    /**
     * @return string
     */
    private function getCacheDir()
    {
        return __DIR__ . '/cache';
    }
}

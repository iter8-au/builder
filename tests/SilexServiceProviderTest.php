<?php

namespace Builder\Tests;

use Builder\Provider\Silex\BuilderServiceProvider;
use Silex\Application;
use PHPUnit_Framework_TestCase;
use Box\Spout\Common\Helper\FileSystemHelper;

/**
 * Class SilexServiceProviderTest
 * @package Builder\Tests
 */
class SilexServiceProviderTest extends PHPUnit_Framework_TestCase
{
    /**
     * @test
     */
    public function defaults_to_correct_builder()
    {
        // Arrange
        $app = new Application();

        // Act
        $app->register(new BuilderServiceProvider());
        // Even though we hard-code 'phpspreadsheet' as the default in the class, we're fetching it here too just to make sure.
        $defaultDriverClass = BuilderServiceProvider::mapBuilderToClass($app['builder.default']);

        // Assert
        $this->assertInstanceOf($defaultDriverClass, $app['builder']->getBuilder());
    }

    /**
     * @test
     */
    public function spout_builder_is_available_when_phpspreadsheet_is_default()
    {
        // Arrange
        $app = new Application();

        // Act
        $app->register(new BuilderServiceProvider(), [
            'builder.default'   => 'phpspreadsheet',
            'builder.cache_dir' => $this->getCacheDir(),
        ]);

        // Assert
        $this->assertInstanceOf('Builder\Builders\PhpSpreadsheet', $app['builder']->getBuilder());
        $this->assertInstanceOf('Builder\Builders\SpoutBuilder', $app['builders']['spout']->getBuilder());
    }

    /**
     * @test
     */
    public function phpspreadsheet_builder_is_available_when_spout_is_default()
    {
        // Arrange
        $app = new Application();

        // Act
        $app->register(new BuilderServiceProvider(), [
            'builder.default'   => 'spout',
            'builder.cache_dir' => $this->getCacheDir(),
        ]);

        // Assert
        $this->assertInstanceOf('Builder\Builders\SpoutBuilder', $app['builder']->getBuilder());
        $this->assertInstanceOf('Builder\Builders\PhpSpreadsheet', $app['builders']['phpspreadsheet']->getBuilder());
    }

    /**
     * @return string
     */
    private function getCacheDir()
    {
        return __DIR__ . '/cache/builder';
    }

    /**
     * Create the cache folder required for testing.
     */
    public static function setUpBeforeClass()
    {
        $fileSystemHelper = new FileSystemHelper(__DIR__ . '/cache');

        if (is_dir(__DIR__ . '/cache/builder') === false) {
            $fileSystemHelper->createFolder(__DIR__ . '/cache', 'builder');
        }
    }

    /**
     * Remove the cache folder required for testing.
     */
    public static function tearDownAfterClass()
    {
        $fileSystemHelper = new FileSystemHelper(__DIR__ . '/cache');

        if (is_dir(__DIR__ . '/cache/builder') === true) {
            $fileSystemHelper->deleteFolderRecursively(__DIR__ . '/cache/builder');
        }
    }
}

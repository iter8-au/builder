<?php

namespace Builder\Tests;

use Builder\Provider\BuilderServiceProvider;
use Silex\Application;

/**
 * Class PHPExcelTest
 * @package Builder\Tests
 */
class PHPExcelTest extends \PHPUnit_Framework_TestCase
{
    /**
     * @test
     */
    public function builder_is_phpexcel_builder()
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
}

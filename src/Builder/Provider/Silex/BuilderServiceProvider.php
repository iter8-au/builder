<?php

namespace Builder\Provider\Silex;

use Builder\Builder;
use Builder\Builders\SpoutBuilder;
use Builder\Builders\PHPExcelBuilder;
use Pimple\Container;
use Pimple\ServiceProviderInterface;

/**
 * Class BuilderServiceProvider
 */
class BuilderServiceProvider implements ServiceProviderInterface
{
    private $builderMappings = [
        'spout'    => SpoutBuilder::class,
        'phpexcel' => PHPExcelBuilder::class,
    ];

    public function register(Container $app)
    {
        $app['builder.default_options'] = [
            'default'   => 'phpexcel',
            'cache_dir' => __DIR__ . '/../../../../../../../../cache/builder',
        ];

        // Merge default builder option.
        $app['builder.default'] = isset($app['builder.default'])
                                ? $app['builder.default']
                                : $app['builder.default_options']['default'];

        // Merge cache_dir option.
        $app['builder.cache_dir'] = isset($app['builder.cache_dir'])
                                  ? $app['builder.cache_dir']
                                  : $app['builder.default_options']['cache_dir'];

        $app['builders'] = function ($app) {
            $cacheDir = $app['builder.cache_dir'];

            // Initialise new Pimple container.
            $builders = new Container();

            foreach ($this->builderMappings as $builderName => $builderClassMapping) {
                $builders[$builderName] = function () use ($builderClassMapping, $cacheDir) {
                    return new Builder(
                        new $builderClassMapping(),
                        $cacheDir
                    );
                };
            }

            return $builders;
        };

        $app['builder'] = function ($app) {
            $builders = $app['builders'];

            return $builders[$app['builder.default']];
        };
    }

    /**
     * @param  string $builderName
     *
     * @return string Fully qualified class name (FQCN) for the specified Builder class.
     */
    public static function mapBuilderToClass($builderName)
    {
        return (new self)->builderMappings[$builderName];
    }
}
<?php

namespace Builder\Provider\Silex;

use Builder\Builder;
use Pimple;
use Silex\Application;
use Silex\ServiceProviderInterface;

/**
 * Class BuilderServiceProvider
 * @package Builder\Provider\Silex
 */
class BuilderServiceProvider implements ServiceProviderInterface
{
    private $builderMappings = [
        'spout'    => 'Builder\Builders\SpoutBuilder',
        'phpexcel' => 'Builder\Builders\PHPExcelBuilder',
    ];

    public function register(Application $app)
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

        $app['builders'] = $app->share(function ($app) {
            $cacheDir = $app['builder.cache_dir'];

            // Initialise new Pimple container.
            $builders = new Pimple();

            foreach ($this->builderMappings as $builderName => $builderClassMapping) {
                $builders[$builderName] = $builders->share(function () use ($builderClassMapping, $cacheDir) {
                    return new Builder(
                        new $builderClassMapping(),
                        $cacheDir
                    );
                });
            }

            return $builders;
        });

        $app['builder'] = $app->share(function ($app) {
            $builders = $app['builders'];

            return $builders[$app['builder.default']];
        });
    }

    public function boot(Application $app)
    {
        //
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

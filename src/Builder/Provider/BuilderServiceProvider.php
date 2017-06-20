<?php

namespace Builder\Provider;

use Builder\Builder;
use Builder\Builders\SpoutBuilder;
use Builder\Builders\PHPExcelBuilder;
use Silex\Application;
use Silex\ServiceProviderInterface;

/**
 * Class BuilderServiceProvider
 * @package Builder\Provider
 */
class BuilderServiceProvider implements ServiceProviderInterface
{
    public function register(Application $app)
    {
        $app['builder.default_options'] = [
            'driver'    => 'phpexcel',
            'cache_dir' => __DIR__ . '/../../../../../../../../cache/builder',
        ];

        $app['builder'] = $app->share(function ($app) {
            $cacheDir = !empty($app['builder.cache_dir'])
                      ? $app['builder.cache_dir']
                      : $app['builder.default_options']['cache_dir'];

            switch ($app['builder.driver']) {
                case 'spout':
                    $builder = new Builder(
                        new SpoutBuilder(),
                        $cacheDir
                    );

                    break;

                case 'phpexcel':
                default:
                    $builder = new Builder(
                        new PHPExcelBuilder(),
                        $cacheDir
                    );

                    break;
            }

            return $builder;
        });
    }

    public function boot(Application $app)
    {
        //
    }
}

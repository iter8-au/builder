<?php

namespace Builder\Provider;

use Builder\Builder;
use Builder\SpoutBuilder;
use Builder\PHPExcelBuilder;
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
        $app['builder'] = $app->share(function ($app) {
            switch ($app['builder.driver']) {
                case 'spout':
                    $builder = new Builder(
                        new SpoutBuilder(),
                        $app['builder.cache_dir']
                    );

                    break;

                case 'phpexcel':
                default:
                    $builder = new Builder(
                        new PHPExcelBuilder(),
                        $app['builder.cache_dir']
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

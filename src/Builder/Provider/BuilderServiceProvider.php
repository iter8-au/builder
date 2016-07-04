<?php

namespace Builder\Provider;

use Builder\Builder;
use Silex\Application;
use Silex\ServiceProviderInterface;

class BuilderServiceProvider implements ServiceProviderInterface
{
    public function register(Application $app)
    {
        $app['builder'] = $app->share(function ($app) {
            $builder = new Builder(
                new \PHPExcel(),
                $app['builder.cache_dir']
            );
            
            return $builder;
        });
    }

    public function boot(Application $app)
    {
    }
}

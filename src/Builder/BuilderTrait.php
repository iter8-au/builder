<?php

namespace Builder;

/**
 * Trait BuilderTrait
 * @package Builder
 */
trait BuilderTrait
{
    /**
     * @var string
     */
    private $cacheDir;

    /**
     * @return string
     */
    public function getCacheName()
    {
        return sprintf(
            '%s/%s.xlsx',
            $this->cacheDir,
            md5(microtime() . mt_rand(0, mt_getrandmax()))
        );
    }

    /**
     * @param  string $cacheDir
     *
     * @return $this
     */
    public function setCacheDir($cacheDir)
    {
        $this->cacheDir = $cacheDir;

        return $this;
    }
}

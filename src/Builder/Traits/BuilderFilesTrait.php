<?php

namespace Builder\Traits;

/**
 * Trait BuilderFilesTrait
 * @package Builder\Traits
 */
trait BuilderFilesTrait
{
    /**
     * @var string
     */
    private $cacheDir;

    /**
     * @var string
     */
    private $cacheName;

    /**
     * Path to the temporary file.
     *
     * @return string
     */
    public function getTempName()
    {
        return sprintf(
            '%s.tmp',
            $this->getCacheName()
        );
    }

    /**
     * @return string
     */
    public function getCacheName()
    {
        if (empty($this->cacheName)) {
            return $this->cacheName = sprintf(
                '%s/%s.xlsx',
                $this->cacheDir,
                md5(microtime() . mt_rand(0, mt_getrandmax()))
            );
        }

        return $this->cacheName;
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

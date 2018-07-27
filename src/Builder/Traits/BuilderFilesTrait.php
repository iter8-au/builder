<?php

declare(strict_types=1);

namespace Builder\Traits;

/**
 * Trait BuilderFilesTrait
 */
trait BuilderFilesTrait
{
    /**
     * @var string
     */
    private $cacheDir;

    /**
     * @var null|string
     */
    private $cacheName;

    /**
     * {@inheritdoc}
     */
    public function getTempName(): string
    {
        return sprintf(
            '%s.tmp',
            $this->getCacheName()
        );
    }

    /**
     * {@inheritdoc}
     */
    public function getCacheName(): string
    {
        if (null === $this->cacheName) {
            $this->cacheName = sprintf(
                '%s/%s.xlsx',
                $this->cacheDir,
                md5(microtime() . random_int(0, mt_getrandmax()))
            );
        }

        return $this->cacheName;
    }

    /**
     * {@inheritdoc}
     */
    public function setCacheDir(string $cacheDir)
    {
        $this->cacheDir = $cacheDir;

        return $this;
    }
}

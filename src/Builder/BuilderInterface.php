<?php

namespace Builder;

/**
 * Interface BuilderInterface
 * @package Builder
 */
interface BuilderInterface
{
    const ALIGNMENT_LEFT   = 'ALIGNMENT_LEFT';
    const ALIGNMENT_CENTRE = 'ALIGNMENT_CENTRE';
    const ALIGNMENT_RIGHT  = 'ALIGNMENT_RIGHT';

    const FILL_SOLID = 'FILL_SOLID';
    const FILL_NONE  = 'FILL_NONE';

    const COLOUR_BLACK_RGB = '000000';
    const COLOUR_WHITE_RGB = 'FFFFFF';

    /**
     * @return void
     */
    public function initialise();

    /**
     * @param  string $cacheDir
     *
     * @return $this
     */
    public function setCacheDir($cacheDir);

    /**
     * @return string
     */
    public function getCacheName();

    /**
     * @param  string|null $creator
     *
     * @return $this
     */
    public function setCreator($creator);

    /**
     * @param  string|null $lastModifiedBy
     *
     * @return $this
     */
    public function setLastModifiedBy($lastModifiedBy);

    /**
     * @param  string|null $title
     *
     * @return $this
     */
    public function setTitle($title);

    /**
     * @param  string|null $subject
     *
     * @return $this
     */
    public function setSubject($subject);

    /**
     * @param  string|null $description
     *
     * @return $this
     */
    public function setDescription($description);

    /**
     * @param  int $sheetIndex
     *
     * @return $this
     */
    public function setActiveSheetIndex($sheetIndex);

    /**
     * @param  array $style
     *
     * @return mixed
     */
    public function buildRowStyle(array $style);

    /**
     * @param  array $columns
     * @param  mixed|null $style
     */
    public function buildHeaderRow($columns, $style = null);
}

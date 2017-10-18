<?php

namespace Builder\Interfaces;

/**
 * Interface BuilderInterface
 */
interface BuilderInterface
{
    const ALIGNMENT_LEFT   = 'ALIGNMENT_LEFT';
    const ALIGNMENT_CENTRE = 'ALIGNMENT_CENTRE';
    const ALIGNMENT_RIGHT  = 'ALIGNMENT_RIGHT';

    const FILL_SOLID       = 'FILL_SOLID';
    const FILL_NONE        = 'FILL_NONE';

    const COLOUR_BLACK_RGB = '000000';
    const COLOUR_WHITE_RGB = 'FFFFFF';

    /**
     * @return void
     */
    public function initialise();

    /**
     * @param  string $cacheDir
     *
     * @return BuilderInterface
     */
    public function setCacheDir($cacheDir);

    /**
     * Path to the temporary file.
     *
     * @return string
     */
    public function getTempName();

    /**
     * @return string
     */
    public function getCacheName();

    /**
     * @param  string|null $creator
     *
     * @return BuilderInterface
     */
    public function setCreator($creator);

    /**
     * @param  string|null $lastModifiedBy
     *
     * @return BuilderInterface
     */
    public function setLastModifiedBy($lastModifiedBy);

    /**
     * @param  string|null $title
     *
     * @return BuilderInterface
     */
    public function setTitle($title);

    /**
     * @param  string|null $subject
     *
     * @return BuilderInterface
     */
    public function setSubject($subject);

    /**
     * @param  string|null $description
     *
     * @return BuilderInterface
     */
    public function setDescription($description);

    /**
     * @param  int $sheetIndex
     *
     * @return BuilderInterface
     */
    public function setActiveSheetIndex($sheetIndex);

    /**
     * @param  string $title
     *
     * @return BuilderInterface
     */
    public function setSheetTitle($title);

    /**
     * @return void
     */
    public function createNewSheet();

    /**
     * @param  array $style
     *
     * @return mixed Depending on the builder this can return an array or a specific class.
     */
    public function buildRowStyle(array $style);

    /**
     * @param  array      $columns
     * @param  mixed|null $style
     *
     * @return void
     */
    public function buildHeaderRow($columns, $style = null);

    /**
     * @param  array      $row
     * @param  mixed|null $style
     * @param  int        $rowIndex
     *
     * @return void
     */
    public function buildRow($row, $style = null, $rowIndex = 1);

    /**
     * @param  array      $rows
     * @param  mixed|null $style
     *
     * @return void
     */
    public function buildRows($rows, $style = null);

    /**
     * @param  array    $columns
     * @param  array    $widths
     * @param  int|null $sheet
     *
     * @return void
     */
    public function applyColumnWidths(array $columns, array $widths, $sheet = null);

    /**
     * @param  array    $columns
     * @param  int|null $sheet
     *
     * @return void
     */
    public function autoSizeColumns(array $columns, $sheet = null);

    /**
     * @param  string $type
     *
     * @return void
     */
    public function closeAndWrite($type = '');
}
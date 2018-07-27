<?php

declare(strict_types=1);

namespace Builder\Interfaces;

/**
 * Interface BuilderInterface
 */
interface BuilderInterface
{
    public const ALIGNMENT_LEFT   = 'ALIGNMENT_LEFT';
    public const ALIGNMENT_CENTRE = 'ALIGNMENT_CENTRE';
    public const ALIGNMENT_RIGHT  = 'ALIGNMENT_RIGHT';

    public const FILL_SOLID = 'FILL_SOLID';
    public const FILL_NONE  = 'FILL_NONE';

    public const COLOUR_BLACK_RGB = '000000';
    public const COLOUR_WHITE_RGB = 'FFFFFF';

    /**
     * @return void
     */
    public function initialise(): void;

    /**
     * @param string|null $creator
     *
     * @return static
     */
    public function setCreator(?string $creator = null);

    /**
     * @param null|string $lastModifiedBy
     *
     * @return static
     */
    public function setLastModifiedBy(?string $lastModifiedBy = null);

    /**
     * @param null|string $title
     *
     * @return static
     */
    public function setTitle(?string $title = null);

    /**
     * @param null|string $subject
     *
     * @return static
     */
    public function setSubject(?string $subject = null);

    /**
     * @param null|string $description
     *
     * @return static
     */
    public function setDescription(?string $description = null);

    /**
     * @param int $sheetIndex
     *
     * @return static
     */
    public function setActiveSheetIndex(int $sheetIndex = 1);

    /**
     * @param string $title
     *
     * @return static
     */
    public function setSheetTitle(string $title);

    /**
     * Creates a new worksheet and sets it as the current worksheet.
     *
     * @return void
     */
    public function createNewSheet(): void;

    /**
     * @param array $style
     *
     * @return mixed Depending on the builder this can return an array or a specific class.
     */
    public function buildRowStyle(array $style);

    /**
     * @param array      $columns
     * @param mixed|null $style
     *
     * @return void
     */
    public function buildHeaderRow(
        array $columns,
        $style = null
    ): void;

    /**
     * @param array      $row
     * @param mixed|null $style
     * @param int        $rowIndex
     *
     * @return void
     */
    public function buildRow(
        array $row,
        $style = null,
        $rowIndex = 1
    ): void;

    /**
     * @param array      $rows
     * @param mixed|null $style
     *
     * @return void
     */
    public function buildRows(
        array $rows,
        $style = null
    ): void;

    /**
     * @param array    $columns
     * @param array    $widths
     * @param int|null $sheet
     *
     * @return void
     */
    public function applyColumnWidths(
        array $columns,
        array $widths,
        $sheet = null
    ): void;

    /**
     * @param array    $columns
     * @param int|null $sheet
     *
     * @return void
     */
    public function autoSizeColumns(
        array $columns,
        $sheet = null
    ): void;

    /**
     * @param string $type Type of writer we should use.  (Defaults to Xlsx file type for PhpSpreadsheet).
     *
     * @return void
     */
    public function closeAndWrite(string $type = ''): void;

    // BuilderFilesTrait

    /**
     * Path and name of the temporary cache file.
     *
     * @return string
     */
    public function getTempName(): string;

    /**
     * Path and name of the cache file.
     *
     * @return string
     */
    public function getCacheName(): string;

    /**
     * Directory path of where cache and temporary files should be stored.
     *
     * @param string $cacheDir
     *
     * @return static
     */
    public function setCacheDir(string $cacheDir);
}

<?php

declare(strict_types=1);

namespace Iter8\Builder\Interfaces;

interface BuilderInterface
{
    public const ALIGNMENT_LEFT = 'ALIGNMENT_LEFT';
    public const ALIGNMENT_CENTRE = 'ALIGNMENT_CENTRE';
    public const ALIGNMENT_RIGHT = 'ALIGNMENT_RIGHT';

    public const FILL_SOLID = 'FILL_SOLID';
    public const FILL_NONE = 'FILL_NONE';

    public const COLOUR_BLACK_RGB = '000000';
    public const COLOUR_WHITE_RGB = 'FFFFFF';

    public function initialise(): void;

    /**
     * @return static
     */
    public function setCreator(?string $creator = null);

    /**
     * @return static
     */
    public function setLastModifiedBy(?string $lastModifiedBy = null);

    /**
     * @return static
     */
    public function setTitle(?string $title = null);

    /**
     * @return static
     */
    public function setSubject(?string $subject = null);

    /**
     * @return static
     */
    public function setDescription(?string $description = null);

    /**
     * @return static
     */
    public function setActiveSheetIndex(int $sheetIndex = 1);

    /**
     * @return static
     */
    public function setSheetTitle(string $title);

    /**
     * Creates a new worksheet and sets it as the current worksheet.
     */
    public function createNewSheet(): void;

    /**
     * @return mixed depending on the builder this can return an array or a specific class
     */
    public function buildRowStyle(array $style);

    /**
     * @param mixed|null $style
     */
    public function buildHeaderRow(
        array $columns,
        $style = null
    ): void;

    /**
     * @param mixed|null $style
     * @param int        $rowIndex
     */
    public function buildRow(
        array $row,
        $style = null,
        $rowIndex = 1
    ): void;

    /**
     * @param mixed|null $style
     */
    public function buildRows(
        array $rows,
        $style = null
    ): void;

    /**
     * @param int|null $sheet
     */
    public function applyColumnWidths(
        array $columns,
        array $widths,
        $sheet = null
    ): void;

    /**
     * @param int|null $sheet
     */
    public function autoSizeColumns(
        array $columns,
        $sheet = null
    ): void;

    /**
     * @param string $type Type of writer we should use.  (Defaults to Xlsx file type for PhpSpreadsheet).
     */
    public function closeAndWrite(string $type = ''): void;

    // BuilderFilesTrait

    /**
     * Path and name of the temporary cache file.
     */
    public function getTempName(): string;

    /**
     * Path and name of the cache file.
     */
    public function getCacheName(): string;

    /**
     * Directory path of where cache and temporary files should be stored.
     *
     * @return static
     */
    public function setCacheDir(string $cacheDir);
}

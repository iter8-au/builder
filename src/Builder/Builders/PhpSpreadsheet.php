<?php

declare(strict_types=1);

namespace Iter8\Builder\Builders;

use Iter8\Builder\Interfaces\BuilderInterface;
use Iter8\Builder\Traits\BuilderFilesTrait;
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Fill;

/**
 * Class PhpSpreadsheet.
 */
class PhpSpreadsheet implements BuilderInterface
{
    use BuilderFilesTrait;

    /**
     * @var Spreadsheet
     */
    private $builder;

    /**
     * PhpSpreadsheet constructor.
     */
    public function __construct()
    {
        $this->builder = new Spreadsheet();
    }

    /**
     * {@inheritdoc}
     */
    public function initialise(): void
    {
        // No initialisation required for the PhpSpreadsheet library.
    }

    /**
     * {@inheritdoc}
     */
    public function setCreator(?string $creator = null)
    {
        $this->builder->getProperties()->setCreator($creator);

        return $this;
    }

    /**
     * {@inheritdoc}
     */
    public function setLastModifiedBy(?string $lastModifiedBy = null)
    {
        $this->builder->getProperties()->setLastModifiedBy($lastModifiedBy);

        return $this;
    }

    /**
     * {@inheritdoc}
     */
    public function setTitle(?string $title = null)
    {
        $this->builder->getProperties()->setTitle($title);

        return $this;
    }

    /**
     * {@inheritdoc}
     */
    public function setSubject(?string $subject = null)
    {
        $this->builder->getProperties()->setSubject($subject);

        return $this;
    }

    /**
     * {@inheritdoc}
     */
    public function setDescription(?string $description = null)
    {
        $this->builder->getProperties()->setDescription($description);

        return $this;
    }

    /**
     * {@inheritdoc}
     */
    public function setActiveSheetIndex(int $sheetIndex = 1)
    {
        $this->builder->setActiveSheetIndex($sheetIndex);

        return $this;
    }

    /**
     * {@inheritdoc}
     */
    public function setSheetTitle(string $title)
    {
        $this->builder->getActiveSheet()->setTitle($title);

        return $this;
    }

    /**
     * {@inheritdoc}
     */
    public function createNewSheet(): void
    {
        $newSheet = $this->builder->createSheet();

        $newSheetIndex = $this->builder->getIndex($newSheet);

        $this->setActiveSheetIndex($newSheetIndex);
    }

    /**
     * http://stackoverflow.com/questions/12918586/phpexcel-specific-cell-formatting-from-style-object.
     *
     * @return array
     */
    public function buildRowStyle(array $style)
    {
        $finalStyleArray = [];
        $defaultStyleArray = [
            'alignment' => [
                'horizontal' => Alignment::HORIZONTAL_LEFT,
            ],
            'font' => [
                'color' => [
                    'rgb' => BuilderInterface::COLOUR_BLACK_RGB,
                ],
                'bold' => true,
            ],
            'fill' => [
                'type' => Fill::FILL_SOLID,
                'color' => [
                    'rgb' => '000000',
                ],
            ],
        ];

        if (!\array_key_exists('alignment', $style)) {
            $finalStyleArray['alignment']['horizontal'] = $defaultStyleArray['alignment']['horizontal'];
        } else {
            switch ($style['alignment']) {
                case 'right':
                    $finalStyleArray['alignment']['horizontal'] = Alignment::HORIZONTAL_RIGHT;

                    break;

                case 'centre':
                case 'center':
                    $finalStyleArray['alignment']['horizontal'] = Alignment::HORIZONTAL_CENTER;

                    break;

                case 'left':
                default:
                    $finalStyleArray['alignment']['horizontal'] = Alignment::HORIZONTAL_LEFT;

                    break;
            }
        }

        if (!\array_key_exists('font', $style)) {
            $finalStyleArray['font'] = $defaultStyleArray['font'];
        } else {
            $finalStyleArray['font'] = $style['font'];
        }

        return $finalStyleArray;
    }

    /**
     * {@inheritdoc}
     */
    public function buildHeaderRow(
        array $columns,
        $style = null
    ): void {
        // The row needs to start at 1 at the beginning of execution.
        // The top left corner of the sheet is actually position (col = 0, row = 1).
        $row = 1;
        $column = 1;

        foreach ($columns as $columnName) {
            $this->builder->getActiveSheet()->setCellValueByColumnAndRow($column, $row, $columnName);

            if (\is_array($style)) {
                $this->builder
                     ->getActiveSheet()
                     ->getStyleByColumnAndRow($column, $row)
                     ->applyFromArray($style);
            }

            ++$column;
        }

        return;
    }

    /**
     * {@inheritdoc}
     */
    public function buildRow(
        array $row,
        $style = null,
        $rowIndex = 1
    ): void {
        $columnIndex = 1;

        foreach ($row as $column) {
            $this->builder->getActiveSheet()->setCellValueByColumnAndRow($columnIndex, $rowIndex, $column);

            ++$columnIndex;
        }

        return;
    }

    /**
     * http://stackoverflow.com/questions/2584954/phpexcel-how-to-set-cell-value-dynamically.
     *
     * {@inheritdoc}
     */
    public function buildRows(
        array $rows,
        $style = null
    ): void {
        // The row needs to start at 1 at the beginning of execution.
        // The top left corner of the sheet is actually position (col = 1, row = 1).
        $rowIndex = 1;

        // If we have a header row then we need to bump the row index down one,
        // otherwise we'll overwrite the header (not ideal).
        if ($this->builder->getActiveSheet()->cellExistsByColumnAndRow(1, $rowIndex)) {
            $rowIndex = 2;
        }

        foreach ($rows as $row) {
            $this->buildRow($row, $style, $rowIndex);

            ++$rowIndex;
        }

        return;
    }

    /**
     * {@inheritdoc}
     */
    public function applyColumnWidths(
        array $columns,
        array $widths,
        $sheet = null
    ): void {
        if (null !== $sheet) {
            $this->builder->setActiveSheetIndex($sheet);
        }

        // Loop through all of our column values -  we only set values for columns that we actually have.
        foreach ($widths as $columnKey => $columnWidth) {
            $this->builder->getActiveSheet()->getColumnDimension($columns[$columnKey])->setWidth($columnWidth);
        }

        return;
    }

    /**
     * {@inheritdoc}
     */
    public function autoSizeColumns(
        array $columns,
        $sheet = null
    ): void {
        if (null !== $sheet) {
            $this->builder->setActiveSheetIndex($sheet);
        }

        $columnCount = \count($columns);

        for ($columnIndex = 1; $columnIndex <= $columnCount; ++$columnIndex) {
            $this->builder->getActiveSheet()->getColumnDimensionByColumn($columnIndex)->setAutoSize(true);
        }

        return;
    }

    /**
     * {@inheritdoc}
     */
    public function closeAndWrite(string $type = 'Xlsx'): void
    {
        $writer = IOFactory::createWriter($this->builder, $type);

        $writer->save($this->getTempName());
    }
}

<?php

namespace Builder;

use Builder\Traits\BuilderFilesTrait;
use Builder\Interfaces\BuilderInterface;
use PHPExcel;
use PHPExcel_IOFactory;
use PHPExcel_Style_Fill;
use PHPExcel_Style_Alignment;

/**
 * Class PHPExcelBuilder
 * @package Builder
 */
class PHPExcelBuilder implements BuilderInterface
{
    use BuilderFilesTrait;

    /**
     * @var \PHPExcel
     */
    private $builder;

    /**
     * PHPExcelBuilder constructor.
     */
    public function __construct()
    {
        $this->builder = new PHPExcel();
    }

    /**
     * @return void
     */
    public function initialise()
    {
        // No initialisation required for the PHPExcel library.
    }

    /**
     * @param  string|null $creator
     *
     * @return $this
     */
    public function setCreator($creator)
    {
        $this->builder->getProperties()->setCreator($creator);

        return $this;
    }

    /**
     * @param  string|null $lastModifiedBy
     *
     * @return $this
     */
    public function setLastModifiedBy($lastModifiedBy)
    {
        $this->builder->getProperties()->setLastModifiedBy($lastModifiedBy);

        return $this;
    }

    /**
     * @param  string|null $title
     *
     * @return $this
     */
    public function setTitle($title)
    {
        $this->builder->getProperties()->setTitle($title);

        return $this;
    }

    /**
     * @param  string|null $subject
     *
     * @return $this
     */
    public function setSubject($subject)
    {
        $this->builder->getProperties()->setSubject($subject);

        return $this;
    }

    /**
     * @param  string|null $description
     *
     * @return $this
     */
    public function setDescription($description)
    {
        $this->builder->getProperties()->setDescription($description);

        return $this;
    }

    /**
     * @param  int $sheetIndex
     *
     * @return $this
     *
     * @throws \PHPExcel_Exception
     */
    public function setActiveSheetIndex($sheetIndex)
    {
        $this->builder->setActiveSheetIndex($sheetIndex);

        return $this;
    }

    /**
     * @param  string $title
     *
     * @return $this
     *
     * @throws \PHPExcel_Exception
     */
    public function setSheetTitle($title)
    {
        $this->builder->getActiveSheet()->setTitle($title);

        return $this;
    }

    /**
     * @return void
     *
     * @throws \PHPExcel_Exception
     */
    public function createNewSheet()
    {
        $this->builder->createSheet();
    }

    /**
     * http://stackoverflow.com/questions/12918586/phpexcel-specific-cell-formatting-from-style-object
     *
     * @param  array $style
     *
     * @return array
     */
    public function buildRowStyle(array $style)
    {
        $finalStyleArray   = [];
        $defaultStyleArray = [
            'alignment' => [
                'horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_LEFT,
            ],
            'font'      => [
                'color' => [
                    'rgb' => 'FFFFFF',
                ],
                'bold' => true,
            ],
            'fill'      => [
                'type'  => PHPExcel_Style_Fill::FILL_SOLID,
                'color' => [
                    'rgb' => '000000',
                ],
            ],
        ];

        if (!array_key_exists('alignment', $style)) {
            $finalStyleArray['alignment']['horizontal'] = $defaultStyleArray['alignment']['horizontal'];
        } else {
            switch ($style['alignment']) {
                case 'right':
                    $finalStyleArray['alignment']['horizontal'] = PHPExcel_Style_Alignment::HORIZONTAL_RIGHT;

                    break;

                case 'centre':
                case 'center':
                    $finalStyleArray['alignment']['horizontal'] = PHPExcel_Style_Alignment::HORIZONTAL_CENTER;

                    break;

                case 'left':
                default:
                    $finalStyleArray['alignment']['horizontal'] = PHPExcel_Style_Alignment::HORIZONTAL_LEFT;

                    break;
            }
        }

        if (!array_key_exists('font', $style)) {
            $finalStyleArray['font'] = $defaultStyleArray['font'];
        } else {
            $finalStyleArray['font'] = $style['font'];
        }

        if (!array_key_exists('fill', $style)) {
            $finalStyleArray['style'] = $defaultStyleArray['fill'];
        } else {
            $fill = $style['fill'];

            switch ($fill['type']) {
                case 'none':
                    $fill['type'] = PHPExcel_Style_Fill::FILL_NONE;

                    break;

                case 'solid':
                default:
                    $fill['type'] = PHPExcel_Style_Fill::FILL_SOLID;

                    break;
            }

            $finalStyleArray['fill'] = $fill;
        }

        return $finalStyleArray;
    }

    /**
     * @param  array $columns
     * @param  mixed $style
     *
     * @return void
     *
     * @throws \PHPExcel_Exception
     */
    public function buildHeaderRow($columns, $style = null)
    {
        // The row needs to start at 1 at the beginning of execution.
        // The top left corner of the sheet is actually position (col = 0, row = 1).
        $row    = 1;
        $column = 0;

        foreach (array_keys($columns) as $key) {
            $this->builder->getActiveSheet()->setCellValueByColumnAndRow($column, $row, $key);

            if (is_array($style)) {
                $this->builder
                     ->getActiveSheet()
                     ->getStyleByColumnAndRow($column, $row)
                     ->applyFromArray($style);
            }

            $column++;
        }
    }

    /**
     * @param  array      $row
     * @param  mixed|null $style
     * @param  int        $rowIndex
     *
     * @return void
     *
     * @throws \PHPExcel_Exception
     */
    public function buildRow($row, $style = null, $rowIndex = 1)
    {
        $columnIndex = 0;

        foreach ($row as $column) {
            $this->builder->getActiveSheet()->setCellValueByColumnAndRow($columnIndex, $rowIndex, $column);

            $columnIndex++;
        }
    }

    /**
     * http://stackoverflow.com/questions/2584954/phpexcel-how-to-set-cell-value-dynamically
     *
     * @param  array      $rows
     * @param  mixed|null $style
     *
     * @return void
     *
     * @throws \PHPExcel_Exception
     */
    public function buildRows($rows, $style = null)
    {
        // The row needs to start at 1 at the beginning of execution.
        // The top left corner of the sheet is actually position (col = 0, row = 1).
        $rowIndex = 1;

        // If we have a header row then we need to bump the row index down one,
        // otherwise we'll overwrite the header (not ideal).
        if ($this->builder->getActiveSheet()->cellExistsByColumnAndRow(0, $rowIndex)) {
            $rowIndex = 2;
        }

        foreach ($rows as $row) {
            $this->buildRow($row, $style, $rowIndex);

            $rowIndex++;
        }
    }

    /**
     * @param  array    $columns
     * @param  array    $widths
     * @param  int|null $sheet
     *
     * @return void
     *
     * @throws \PHPExcel_Exception
     */
    public function applyColumnWidths(array $columns, array $widths, $sheet = null)
    {
        if ($sheet !== null) {
            $this->builder->setActiveSheetIndex($sheet);
        }

        // Loop through all of our column values -  we only set values for columns that we actually have.
        foreach ($widths as $columnKey => $columnWidth) {
            $this->builder->getActiveSheet()->getColumnDimension($columns[$columnKey])->setWidth($columnWidth);
        }
    }

    /**
     * @param  array    $columns
     * @param  int|null $sheet
     *
     * @return void
     *
     * @throws \PHPExcel_Exception
     */
    public function autoSizeColumns(array $columns, $sheet = null)
    {
        if ($sheet !== null) {
            $this->builder->setActiveSheetIndex($sheet);
        }

        $columnCount = count($columns);

        for ($columnIndex = 0; $columnIndex <= $columnCount; $columnIndex++) {
            $this->builder->getActiveSheet()->getColumnDimensionByColumn($columnIndex)->setAutoSize(true);
        }
    }

    /**
     * @param  string $type Type of writer we should use.  Defaults to Excel2007 file type.
     *
     * @return void
     *
     * @throws \PHPExcel_Reader_Exception
     * @throws \PHPExcel_Writer_Exception
     */
    public function closeAndWrite($type = 'Excel2007')
    {
        $writer = PHPExcel_IOFactory::createWriter($this->builder, $type);

        $writer->save($this->getTempName());
    }
}

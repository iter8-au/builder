<?php

namespace Builder;

use PHPExcel;
use PHPExcel_Style_Fill;
use PHPExcel_Style_Alignment;

/**
 * Class PHPExcelBuilder
 * @package Builder
 */
class PHPExcelBuilder implements BuilderInterface
{
    /**
     * @var PHPExcel
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
     * http://stackoverflow.com/questions/12918586/phpexcel-specific-cell-formatting-from-style-object
     *
     * @param  array $style
     *
     * @return array
     */
    public function buildRowStyle(array $style)
    {
        $finalStyleArray = [];
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
}

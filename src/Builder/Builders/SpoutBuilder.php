<?php

namespace Builder\Builders;

use Builder\Traits\BuilderFilesTrait;
use Builder\Interfaces\BuilderInterface;
use Builder\Traits\InitialisationStateTrait;
use Box\Spout\Common\Type;
use Box\Spout\Writer\Style\Color;
use Box\Spout\Writer\Style\Style;
use Box\Spout\Writer\WriterFactory;
use Box\Spout\Writer\Style\StyleBuilder;
use Box\Spout\Writer\Exception\SheetNotFoundException;
use Box\Spout\Common\Exception\UnsupportedTypeException;

/**
 * Class SpoutBuilder
 * @package Builder\Builders
 */
class SpoutBuilder implements BuilderInterface
{
    use BuilderFilesTrait, InitialisationStateTrait;

    /**
     * @var \Box\Spout\Writer\AbstractMultiSheetsWriter
     */
    private $writer;

    /**
     * SpoutBuilder constructor.
     */
    public function __construct()
    {
        try {
            $this->writer = WriterFactory::create(Type::XLSX);
        } catch (UnsupportedTypeException $e) {
            // This will never happen.
        }
    }

    /**
     * @return void
     *
     * @throws \Box\Spout\Common\Exception\IOException
     */
    public function initialise()
    {
        $this->writer->openToFile($this->getTempName());

        $this->setAsInitialised();
    }

    /**
     * @param  string|null $creator
     *
     * @return $this
     */
    public function setCreator($creator)
    {
        return $this;
    }

    /**
     * @param  string|null $lastModifiedBy
     *
     * @return $this
     */
    public function setLastModifiedBy($lastModifiedBy)
    {
        return $this;
    }

    /**
     * @param  string|null $title
     *
     * @return $this
     */
    public function setTitle($title)
    {
        return $this;
    }

    /**
     * @param  string|null $subject
     *
     * @return $this
     */
    public function setSubject($subject)
    {
        return $this;
    }

    /**
     * @param  string|null $description
     *
     * @return $this
     */
    public function setDescription($description)
    {
        return $this;
    }

    /**
     * @param  int $sheetIndex
     *
     * @return $this
     *
     * @throws \Box\Spout\Writer\Exception\SheetNotFoundException
     * @throws \Box\Spout\Writer\Exception\WriterNotOpenedException
     */
    public function setActiveSheetIndex($sheetIndex)
    {
        $sheets = $this->writer->getSheets();

        if (!isset($sheets[$sheetIndex])) {
            throw new SheetNotFoundException(
                'The given sheet does not exist in the workbook.'
            );
        }

        $this->writer->setCurrentSheet($sheets[$sheetIndex]);

        return $this;
    }

    /**
     * @param  string $title
     *
     * @return $this
     *
     * @throws \Box\Spout\Common\Exception\IOException
     * @throws \Box\Spout\Writer\Exception\WriterNotOpenedException
     * @throws \Box\Spout\Writer\Exception\InvalidSheetNameException
     */
    public function setSheetTitle($title)
    {
        if ($this->isNotInitialised()) {
            $this->initialise();
        }

        $this->writer->getCurrentSheet()->setName($title);

        return $this;
    }

    /**
     * @return void
     *
     * @throws \Box\Spout\Writer\Exception\WriterNotOpenedException
     */
    public function createNewSheet()
    {
        $this->writer->addNewSheetAndMakeItCurrent();
    }

    /**
     * @param  array $style
     *
     * @return \Box\Spout\Writer\Style\Style
     */
    public function buildRowStyle(array $style)
    {
        $finalStyle = new StyleBuilder();

        // Spout doesn't offer support for changing column alignment.

        if (array_key_exists('font', $style)) {
            if (array_key_exists('color', $style['font'])) {
                $finalStyle->setFontColor(
                    Color::toARGB($style['font']['color']['rgb'])
                );
            }

            if (array_key_exists('bold', $style['font']) && ($style['font']['bold'] === true)) {
                $finalStyle->setFontBold();
            }
        }

        if (array_key_exists('fill', $style) && array_key_exists('color', $style['fill'])) {
            $finalStyle->setBackgroundColor(
                Color::toARGB($style['fill']['color']['rgb'])
            );
        }

        return $finalStyle->build();
    }

    /**
     * @param array $columns
     * @param mixed|null $style
     *
     * @throws \Box\Spout\Common\Exception\IOException
     * @throws \Box\Spout\Common\Exception\InvalidArgumentException
     * @throws \Box\Spout\Writer\Exception\WriterNotOpenedException
     */
    public function buildHeaderRow($columns, $style = null)
    {
        if ($this->isNotInitialised()) {
            $this->initialise();
        }

        $keys = array_keys($columns);

        if ($style instanceof Style) {
            $this->writer->addRowWithStyle($keys, $style);
        } else {
            $this->writer->addRow($keys);
        }
    }

    /**
     * @param  array      $row
     * @param  mixed|null $style
     * @param  int        $rowIndex
     *
     * @return void
     *
     * @throws \Box\Spout\Common\Exception\IOException
     * @throws \Box\Spout\Common\Exception\InvalidArgumentException
     * @throws \Box\Spout\Writer\Exception\WriterNotOpenedException
     */
    public function buildRow($row, $style = null, $rowIndex = 1)
    {
        if ($this->isNotInitialised()) {
            $this->initialise();
        }

        if ($style instanceof Style) {
            $this->writer->addRowWithStyle($row, $style);
        } else {
            $this->writer->addRow($row);
        }
    }

    /**
     * @param  array      $rows
     * @param  mixed|null $style
     *
     * @return void
     *
     * @throws \Box\Spout\Common\Exception\IOException
     * @throws \Box\Spout\Common\Exception\InvalidArgumentException
     * @throws \Box\Spout\Writer\Exception\WriterNotOpenedException
     */
    public function buildRows($rows, $style = null)
    {
        if ($this->isNotInitialised()) {
            $this->initialise();
        }

        if ($style instanceof Style) {
            $this->writer->addRowsWithStyle($rows, $style);
        } else {
            $this->writer->addRows($rows);
        }
    }

    /**
     * @param  array    $columns
     * @param  array    $widths
     * @param  int|null $sheet
     *
     * @return void
     */
    public function applyColumnWidths(array $columns, array $widths, $sheet = null)
    {
        // Spout doesn't support setting fixed column widths yet.
    }

    /**
     * @param  array $columns
     * @param  int|null $sheet
     *
     * @return void
     */
    public function autoSizeColumns(array $columns, $sheet = null)
    {
        // Spout doesn't support auto-sizing of columns yet.
    }

    /**
     * @param  string $type
     *
     * @return void
     */
    public function closeAndWrite($type = '')
    {
        $this->writer->close();
    }
}

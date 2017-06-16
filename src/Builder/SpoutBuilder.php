<?php

namespace Builder;

use Box\Spout\Common\Type;
use Box\Spout\Writer\Style\Color;
use Box\Spout\Writer\Style\Style;
use Box\Spout\Writer\WriterFactory;
use Box\Spout\Writer\Style\StyleBuilder;

/**
 * Class SpoutBuilder
 * @package Builder
 */
class SpoutBuilder implements BuilderInterface
{
    use BuilderTrait;

    /**
     * @var \Box\Spout\Writer\WriterInterface
     */
    private $writer;

    /**
     * SpoutBuilder constructor.
     */
    public function __construct()
    {
        $this->writer = new WriterFactory(Type::XLSX);
    }

    /**
     * @return void
     *
     * @throws \Box\Spout\Common\Exception\IOException
     */
    public function initialise()
    {
        $this->writer->openToFile($this->getCacheName());
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
     */
    public function setActiveSheetIndex($sheetIndex)
    {
        return $this;
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
        $keys = array_keys($columns);

        if ($style instanceof Style) {
            $this->writer->addRowWithStyle($keys, $style);
        } else {
            $this->writer->addRow($keys);
        }
    }
}

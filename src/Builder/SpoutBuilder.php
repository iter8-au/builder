<?php

namespace Builder;

use Box\Spout\Writer\Style\Color;
use Box\Spout\Writer\Style\StyleBuilder;
use Box\Spout\Writer\XLSX\Writer as XLSXWriter;

/**
 * Class SpoutBuilder
 * @package Builder
 */
class SpoutBuilder implements BuilderInterface
{
    /**
     * @var \Box\Spout\Writer\XLSX\Writer
     */
    private $writer;

    public function __construct()
    {
        $this->writer = new XLSXWriter();
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
}

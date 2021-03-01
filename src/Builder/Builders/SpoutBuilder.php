<?php

declare(strict_types=1);

namespace Iter8\Builder\Builders;

use Box\Spout\Common\Entity\Style\Color;
use Box\Spout\Common\Entity\Style\Style;
use Box\Spout\Writer\Common\Creator\Style\StyleBuilder;
use Box\Spout\Writer\Common\Creator\WriterEntityFactory;
use Box\Spout\Writer\Exception\SheetNotFoundException;
use Box\Spout\Writer\WriterMultiSheetsAbstract;
use Iter8\Builder\Interfaces\BuilderInterface;
use Iter8\Builder\Traits\BuilderFilesTrait;
use Iter8\Builder\Traits\InitialisationStateTrait;

/**
 * Class SpoutBuilder.
 */
class SpoutBuilder implements BuilderInterface
{
    use BuilderFilesTrait;
    use InitialisationStateTrait;

    /**
     * @var WriterMultiSheetsAbstract
     */
    private $writer;

    /**
     * SpoutBuilder constructor.
     */
    public function __construct()
    {
        $this->writer = WriterEntityFactory::createXLSXWriter();
    }

    /**
     * {@inheritdoc}
     */
    public function initialise(): void
    {
        $this->writer->openToFile($this->getTempName());

        $this->setAsInitialised();
    }

    /**
     * {@inheritdoc}
     */
    public function setCreator(?string $creator = null)
    {
        // Spout does not implement this ability.

        return $this;
    }

    /**
     * {@inheritdoc}
     */
    public function setLastModifiedBy(?string $lastModifiedBy = null)
    {
        // Spout does not implement this ability.

        return $this;
    }

    /**
     * {@inheritdoc}
     */
    public function setTitle(?string $title = null)
    {
        // Spout does not implement this ability.

        return $this;
    }

    /**
     * {@inheritdoc}
     */
    public function setSubject(?string $subject = null)
    {
        // Spout does not implement this ability.

        return $this;
    }

    /**
     * {@inheritdoc}
     */
    public function setDescription(?string $description = null)
    {
        // Spout does not implement this ability.

        return $this;
    }

    /**
     * {@inheritdoc}
     */
    public function setActiveSheetIndex(int $sheetIndex = 1)
    {
        $sheets = $this->writer->getSheets();

        if (!isset($sheets[$sheetIndex])) {
            throw new SheetNotFoundException(sprintf('The given sheet "%d" does not exist in the workbook.', $sheetIndex));
        }

        $this->writer->setCurrentSheet($sheets[$sheetIndex]);

        return $this;
    }

    /**
     * {@inheritdoc}
     */
    public function setSheetTitle(string $title)
    {
        if ($this->isNotInitialised()) {
            $this->initialise();
        }

        $this->writer->getCurrentSheet()->setName($title);

        return $this;
    }

    /**
     * {@inheritdoc}
     */
    public function createNewSheet(): void
    {
        $this->writer->addNewSheetAndMakeItCurrent();
    }

    /**
     * @return Style
     */
    public function buildRowStyle(array $style)
    {
        $finalStyle = new StyleBuilder();

        // Spout doesn't offer support for changing column alignment.

        if (\array_key_exists('font', $style)) {
            if (\array_key_exists('color', $style['font'])) {
                $finalStyle->setFontColor(
                    Color::toARGB($style['font']['color']['rgb'])
                );
            }

            if (\array_key_exists('bold', $style['font']) && (true === $style['font']['bold'])) {
                $finalStyle->setFontBold();
            }
        }

        return $finalStyle->build();
    }

    /**
     * {@inheritdoc}
     */
    public function buildHeaderRow(
        array $columns,
        $style = null
    ): void {
        if ($this->isNotInitialised()) {
            $this->initialise();
        }

        if ($style instanceof Style) {
            $this->writer->addRowWithStyle($columns, $style);
        } else {
            $this->writer->addRow($columns);
        }

        return;
    }

    /**
     * {@inheritdoc}
     *
     * Spout does not require the $rowIndex parameter.
     */
    public function buildRow(
        array $row,
        $style = null,
        $rowIndex = 1
    ): void {
        if ($this->isNotInitialised()) {
            $this->initialise();
        }

        if ($style instanceof Style) {
            $this->writer->addRowWithStyle($row, $style);
        } else {
            $this->writer->addRow($row);
        }

        return;
    }

    /**
     * {@inheritdoc}
     */
    public function buildRows(
        array $rows,
        $style = null
    ): void {
        if ($this->isNotInitialised()) {
            $this->initialise();
        }

        if ($style instanceof Style) {
            $this->writer->addRowsWithStyle($rows, $style);
        } else {
            $this->writer->addRows($rows);
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
        // Spout doesn't support setting fixed column widths yet.
        return;
    }

    /**
     * {@inheritdoc}
     */
    public function autoSizeColumns(
        array $columns,
        $sheet = null
    ): void {
        // Spout doesn't support auto-sizing of columns yet.
        return;
    }

    /**
     * {@inheritdoc}
     */
    public function closeAndWrite(string $type = ''): void
    {
        $this->writer->close();
    }
}

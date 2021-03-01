<?php

declare(strict_types=1);

namespace Iter8\Builder;

use Iter8\Builder\Interfaces\BuilderInterface;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Fill;

/**
 * Helper for easily generating cached Excel or CSV files, ready to download, from an array of data.
 *
 * Class Builder
 */
class Builder
{
    public const REPORT_EXCEL = 0;
    public const REPORT_CSV = 1;

    public const ALIGNMENT_CENTER = 0;
    public const ALIGNMENT_LEFT = 1;
    public const ALIGNMENT_RIGHT = 2;

    /**
     * @var int
     */
    private $reportType = self::REPORT_EXCEL;

    /**
     * @var BuilderInterface
     */
    private $builder;

    /**
     * @var string
     */
    private $reportCacheDir;

    /**
     * @var array
     */
    private $headers = [];

    /**
     * @var array
     */
    private $data = [];

    /**
     * @var string|null
     */
    private $creator;

    /**
     * @var string|null
     */
    private $title;

    /**
     * @var array
     */
    private $sheetTitles = [];

    /**
     * @var string|null
     */
    private $description;

    /**
     * @var string|null
     */
    private $filename;

    /**
     * @var array
     */
    private $sheets = [];

    /**
     * @var array
     */
    private $columnWidths = [];

    /**
     * @var array
     */
    private $columnStyles = [];

    /**
     * Builder constructor.
     */
    public function __construct(
        BuilderInterface $builder,
        string $reportCacheDir
    ) {
        $this->setBuilder($builder);
        $this->setReportCacheDir($reportCacheDir);

        $this->prepareBuilder();
    }

    /**
     * Prepares the builders and creates the cache directory if it doesn't exist.
     */
    private function prepareBuilder(): void
    {
        if (!file_exists($this->getReportCacheDir())) {
            @mkdir($this->getReportCacheDir(), 0777, true);
        }

        $this->builder
             ->setCacheDir($this->getReportCacheDir())
             ->initialise();

        return;
    }

    /**
     * Get the temp filename for the Excel builder.
     */
    public function getTempName(): string
    {
        return $this->builder->getTempName();
    }

    /**
     * Generate the final report using whatever the set format is.
     *
     * @param bool $unlinkFlag
     *
     * @return void
     */
    public function generate($unlinkFlag = true)
    {
        // Determine which format we are using and call the appropriate method.
        if (self::REPORT_EXCEL === $this->getReportType()) {
            $this->generateExcel();

            // Output headers
            header('Pragma: public');
            header('Expires: 0');
            header('Cache-Control: must-revalidate, post-check=0, pre-check=0');
            header('Content-Type: application/force-download');
            header('Content-Type: application/octet-stream');
            header('Content-Type: application/download');
            header('Content-Disposition: attachment;filename='.$this->getFilename().'.xlsx');
            header('Content-Transfer-Encoding: binary ');

            readfile($this->builder->getTempName());
            unlink($this->builder->getTempName());

            exit;
        }

        throw new \UnexpectedValueException('Attempted to generate a report in an unsupported format.');
    }

    /**
     * Generates an Excel document.
     */
    public function generateExcel(): void
    {
        // Set Document Properties from Service values.
        $this->builder
             ->setCreator($this->getCreator())
             ->setLastModifiedBy($this->getCreator())
             ->setTitle($this->getTitle())
             ->setSubject($this->getTitle())
             ->setDescription($this->getDescription());

        if ($this->hasSheets()) {
            // Multiple sheets.
            $this->createSheets();
        } else {
            // Single sheet - these will be an array and a string.
            $headers = $this->getHeaders();
            $reportArray = $this->getData();
            $sheetTitle = $this->getSheetTitles();

            $this->builder->setActiveSheetIndex(0);

            $this->createSheet(
                $headers,
                $reportArray,
                $sheetTitle
            );
        }

        // Close the builder and write the file.
        $this->builder->closeAndWrite();

        return;
    }

    public function createSheets(): void
    {
        $sheets = $this->getSheets();
        $titles = $this->getSheetTitles();

        $totalSheets = \count($sheets);
        $sheetCount = 0;

        if (empty($sheets) || !\is_array($sheets)) {
            throw new \UnexpectedValueException('Expected an array of sheets data but got an empty value or non-array.');
        }

        // We have to set the initial active sheet.
        $this->builder->setActiveSheetIndex($sheetCount);

        foreach ($sheets as $sheet) {
            $this->createSheet(
                $sheet['headers'],
                $sheet['rows'],
                $titles[$sheetCount]
            );

            // Only create a new sheet if we actually have a data array for it.
            if ($sheetCount < ($totalSheets - 1)) {
                $this->builder->createNewSheet();

                // Increment the active sheet count and move to that sheet.
                ++$sheetCount;
                $this->builder->setActiveSheetIndex($sheetCount);
            }
        }

        // Finally switch back to the first sheet.
        $this->builder->setActiveSheetIndex(0);

        return;
    }

    public function createSheet(
        array $headers,
        array $rows,
        string $title
    ): void {
        // Check if we are setting any custom column widths.
        if ($this->hasColumnWidths()) {
            // We have a numeric index array, so create an array of letters that we can use to map to Excel columns.
            // e.g. 0 = A, 3 = D, etc.
            $columns = range('A', 'Z');

            $this->builder->applyColumnWidths($columns, $this->getColumnWidths());
        }

        // Style settings for agent headers
        $style = $this->builder->buildRowStyle([
            'alignment' => BuilderInterface::ALIGNMENT_CENTRE,
            'font' => [
                'color' => [
                    'rgb' => BuilderInterface::COLOUR_BLACK_RGB,
                ],
                'bold' => true,
            ],
        ]);

        // Build column headers.
        $this->builder->buildHeaderRow($headers, $style);

        // Build all the rows now.
        $this->builder->buildRows($rows);

        // If no column widths are specified, then simply auto-size all columns.
        if (!$this->hasColumnWidths()) {
            $this->builder->autoSizeColumns($rows[0]);
        }

        // Rename sheet.
        $this->builder->setSheetTitle($title);
    }

    /**
     * TODO: Implement to output in CSV format but with an .xls extension to open in excel.
     */
    public function generateCSV()
    {
    }

    public function getReportType(): int
    {
        return $this->reportType;
    }

    /**
     * @return $this
     */
    public function setReportType(int $reportType): self
    {
        $this->reportType = $reportType;

        return $this;
    }

    public function getCreator(): ?string
    {
        return $this->creator;
    }

    /**
     * @param string $creator
     *
     * @return $this
     */
    public function setCreator($creator): self
    {
        $this->creator = $creator;

        return $this;
    }

    public function getHeaders(): array
    {
        return $this->headers;
    }

    public function setHeaders(array $headers): self
    {
        $this->headers = $headers;

        return $this;
    }

    public function getData(): array
    {
        return $this->data;
    }

    public function setData(array $data): self
    {
        $this->data = $data;

        return $this;
    }

    public function getDescription(): ?string
    {
        return $this->description;
    }

    /**
     * @return $this
     */
    public function setDescription(string $description): self
    {
        $this->description = $description;

        return $this;
    }

    public function getFilename(): string
    {
        return $this->filename;
    }

    /**
     * @return $this
     */
    public function setFilename(string $filename): self
    {
        $this->filename = $filename;

        return $this;
    }

    public function getTitle(): ?string
    {
        return $this->title;
    }

    /**
     * Note: if the title is longer than 31 characters it'll be trimmed.
     * This is due to a limit in Excel.
     *
     * @return $this
     */
    public function setTitle(string $title): self
    {
        if (\strlen($title) > 31) {
            $title = mb_substr($title, 0, 31);
        }

        $this->title = $title;

        return $this;
    }

    public function getBuilder(): BuilderInterface
    {
        return $this->builder;
    }

    /**
     * @return $this
     */
    public function setBuilder(BuilderInterface $builder): self
    {
        $this->builder = $builder;

        return $this;
    }

    public function getReportCacheDir(): string
    {
        return $this->reportCacheDir;
    }

    /**
     * @return $this
     */
    public function setReportCacheDir(string $reportCacheDir): self
    {
        $this->reportCacheDir = $reportCacheDir;

        return $this;
    }

    public function hasColumnWidths(): bool
    {
        $columnWidths = $this->getColumnWidths();

        return !empty($columnWidths);
    }

    /**
     * @return mixed
     */
    public function getColumnWidths()
    {
        return $this->columnWidths;
    }

    /**
     * @return $this
     */
    public function setColumnWidths(array $columnWidths): self
    {
        $this->columnWidths = $columnWidths;

        return $this;
    }

    public function hasColumnStyles(): bool
    {
        $columnStyles = $this->getColumnStyles();

        return !empty($columnStyles);
    }

    /**
     * @param int $columnIndex
     */
    public function hasColumnStylesForColumn($columnIndex): bool
    {
        if (!$this->hasColumnStyles()) {
            return false;
        }

        $columnStyles = $this->getColumnStyles();

        return isset($columnStyles[$columnIndex]) ? true : false;
    }

    /**
     * @param int $columnIndex
     *
     * @return bool|array
     */
    public function getColumnStylesForColumn($columnIndex)
    {
        if (!$this->hasColumnStylesForColumn($columnIndex)) {
            return false;
        }

        $columnStyles = $this->getColumnStyles();

        return $columnStyles[$columnIndex] ?? false;
    }

    /**
     * TODO: Allow specific (col, row) styles
     * TODO: Allow text styles and colours to be applied.
     *
     * @param $columnIndex
     *
     * @return array
     */
    private function getPhpSpreadsheetColumnStylesForColumn($columnIndex)
    {
        $phpExcelStyleArray = [];

        $styleArray = $this->getColumnStylesForColumn($columnIndex);

        if (empty($styleArray)) {
            return $phpExcelStyleArray;
        }

        // Alignment
        if (isset($styleArray['alignment'])) {
            if (self::ALIGNMENT_CENTER === $styleArray['alignment']) {
                $phpExcelStyleArray['alignment']['horizontal'] = Alignment::HORIZONTAL_CENTER;
            } elseif (self::ALIGNMENT_LEFT === $styleArray['alignment']) {
                $phpExcelStyleArray['alignment']['horizontal'] = Alignment::HORIZONTAL_LEFT;
            } elseif (self::ALIGNMENT_RIGHT === $styleArray['alignment']) {
                $phpExcelStyleArray['alignment']['horizontal'] = Alignment::HORIZONTAL_RIGHT;
            }
        }

        if (isset($styleArray['fill'])) {
            $phpExcelStyleArray['fill'] = [
                'type' => Fill::FILL_SOLID,
                'color' => [
                    'rgb' => $styleArray['fill'],
                ],
            ];
        }

        /*
        $styleArray = array(
            'font' => array(
                'color' => array(
                    'rgb' => 'FFFFFF'
                ),
                'bold' => true
            ),
        );
        */

        return $phpExcelStyleArray;
    }

    /**
     * @return mixed
     */
    public function getColumnStyles()
    {
        return $this->columnStyles;
    }

    /**
     * @return $this
     */
    public function setColumnStyles(array $columnStyles): self
    {
        $this->columnStyles = $columnStyles;

        return $this;
    }

    public function hasSheets(): bool
    {
        return \count($this->sheets) > 0;
    }

    public function setSheets(array $sheets): self
    {
        $this->sheets = $sheets;

        return $this;
    }

    public function getSheets(): array
    {
        return $this->sheets;
    }

    /**
     * @param string|array $sheetTitles
     *
     * @return $this
     */
    public function setSheetTitles($sheetTitles)
    {
        $this->sheetTitles = $sheetTitles;

        return $this;
    }

    /**
     * @return string|array
     */
    public function getSheetTitles()
    {
        return $this->sheetTitles;
    }
}

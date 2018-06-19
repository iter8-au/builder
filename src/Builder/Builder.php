<?php

namespace Builder;

use Builder\Interfaces\BuilderInterface;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use UnexpectedValueException;

/**
 * Helper for easily generating cached Excel or CSV files, ready to download, from an array of data.
 *
 * Class Builder
 * @package Builder
 */
class Builder
{
    const REPORT_EXCEL = 0;
    const REPORT_CSV   = 1;

    const ALIGNMENT_CENTER  = 0;
    const ALIGNMENT_LEFT    = 1;
    const ALIGNMENT_RIGHT   = 2;

    /**
     * @var int
     */
    private $reportType;

    /**
     * @var BuilderInterface
     */
    private $builder;

    private $reportCacheDir;

    /**
     * @var array
     */
    private $data;

    private $creator;

    private $title;

    private $sheetTitles;

    private $description;

    private $filename;

    /**
     * @var array
     */
    private $sheets;

    /**
     * @var array
     */
    private $columnWidths;

    /**
     * @var array
     */
    private $columnStyles;

    /**
     * Builder constructor.
     *
     * @param  \Builder\Interfaces\BuilderInterface $builder
     * @param  string                               $reportCacheDir
     *
     * @return self
     */
    public function __construct(
        BuilderInterface $builder,
        $reportCacheDir
    ) {
        $this->setBuilder($builder);
        $this->setReportCacheDir($reportCacheDir);

        $this->prepareBuilder();

        // Default the report to Excel format.
        $this->setReportType(self::REPORT_EXCEL);
    }

    /**
     * Prepares the builders.
     *
     * @return void
     */
    private function prepareBuilder()
    {
        $this->builder
             ->setCacheDir($this->getReportCacheDir())
             ->initialise();
    }

    /**
     * Get the temp filename for the Excel builder.
     *
     * @return string
     */
    public function getTempName()
    {
        return $this->builder->getTempName();
    }

    /**
     * Generate the final report using whatever the set format is.
     *
     * @param  bool $unlinkFlag
     *
     * @return void
     *
     * @throws \UnexpectedValueException
     */
    public function generate($unlinkFlag = true) {
        // Determine which format we are using and call the appropriate method.
        if ($this->getReportType() === self::REPORT_EXCEL) {
            $this->generateExcel();

            // Output headers
            header('Pragma: public');
            header('Expires: 0');
            header('Cache-Control: must-revalidate, post-check=0, pre-check=0');
            header('Content-Type: application/force-download');
            header('Content-Type: application/octet-stream');
            header('Content-Type: application/download');
            header('Content-Disposition: attachment;filename=' . $this->getFilename() . '.xlsx');
            header('Content-Transfer-Encoding: binary ');

            readfile($this->builder->getTempName());
            unlink($this->builder->getTempName());

            exit;
        }

        throw new UnexpectedValueException('Attempted to generate a report in an unsupported format.');
    }

    /**
     * Generates an Excel document.
     *
     * @return void
     */
    public function generateExcel()
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
            $reportArray = $this->getData();
            $sheetTitle  = $this->getSheetTitles();

            $this->builder->setActiveSheetIndex(0);

            $this->createSheet(
                $reportArray,
                $sheetTitle
            );
        }

        // Close the builder and write the file.
        $this->builder->closeAndWrite();
    }

    /**
     * @return void
     *
     * @throws \UnexpectedValueException
     */
    public function createSheets()
    {
        $sheets = $this->getSheets();
        $titles = $this->getSheetTitles();

        $totalSheets = count($sheets);
        $sheetCount  = 0;

        if (empty($sheets) || !is_array($sheets)) {
            throw new UnexpectedValueException(
                'Expected an array of sheets data but got an empty value or non-array.'
            );
        }

        // We have to set the initial active sheet.
        $this->builder->setActiveSheetIndex($sheetCount);

        foreach ($sheets as $sheet) {
            $this->createSheet(
                $sheet,
                $titles[$sheetCount]
            );

            // Only create a new sheet if we actually have a data array for it.
            if ($sheetCount < ($totalSheets - 1)) {
                $this->builder->createNewSheet();

                // Increment the active sheet count and move to that sheet.
                $sheetCount++;
                $this->builder->setActiveSheetIndex($sheetCount);
            }
        }

        // Finally switch back to the first sheet.
        $this->builder->setActiveSheetIndex(0);
    }

    /**
     * @param  array  $data
     * @param  string $title
     *
     * @return void
     */
    public function createSheet(array $data, $title) {
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
            'font'      => [
                'color' => [
                    'rgb' => BuilderInterface::COLOUR_BLACK_RGB,
                ],
                'bold'  => true,
            ],
            //'fill'      => [
            //    'type'  => BuilderInterface::FILL_SOLID,
            //    'color' => [
            //        'rgb' => BuilderInterface::COLOUR_BLACK_RGB,
            //    ],
            //],
        ]);

        // Build column headers.
        $this->builder->buildHeaderRow($data[0], $style);

        // Build all the rows now.
        $this->builder->buildRows($data);

        // If no column widths are specified, then simply auto-size all columns.
        if (!$this->hasColumnWidths()) {
            $this->builder->autoSizeColumns($data[0]);
        }

        // Rename sheet.
        $this->builder->setSheetTitle($title);
    }


    /**
     * TODO: Implement to output in CSV format but with an .xls extension to open in excel
     */
    public function generateCSV()
    {

    }

    /**
     * @return int
     */
    public function getReportType()
    {
        return $this->reportType;
    }

    /**
     * @param  int $reportType
     *
     * @return $this
     */
    public function setReportType($reportType)
    {
        $this->reportType = $reportType;

        return $this;
    }

    /**
     * @return string
     */
    public function getCreator()
    {
        return $this->creator;
    }

    /**
     * @param  string $creator
     *
     * @return $this
     */
    public function setCreator($creator)
    {
        $this->creator = $creator;

        return $this;
    }

    /**
     * @return array
     */
    public function getData()
    {
        return $this->data;
    }

    /**
     * @param  array $data
     *
     * @return $this
     */
    public function setData(array $data)
    {
        $this->data = $data;

        return $this;
    }

    /**
     * @return string
     */
    public function getDescription()
    {
        return $this->description;
    }

    /**
     * @param  string $description
     *
     * @return $this
     */
    public function setDescription($description)
    {
        $this->description = $description;

        return $this;
    }

    /**
     * @return string
     */
    public function getFilename()
    {
        return $this->filename;
    }

    /**
     * @param  string $filename
     *
     * @return $this
     */
    public function setFilename($filename)
    {
        $this->filename = $filename;

        return $this;
    }

    /**
     * TODO: If > 31 characters, and PhpSpreadsheet, then sub_str? (31 characters in an Excel title limit)
     *
     * @return string
     */
    public function getTitle()
    {
        return $this->title;
    }

    /**
     * @param  string $title
     *
     * @return $this
     */
    public function setTitle($title)
    {
        $this->title = $title;

        return $this;
    }

    /**
     * @return \Builder\BuilderInterface
     */
    public function getBuilder()
    {
        return $this->builder;
    }

    /**
     * @param  \Builder\BuilderInterface $builder
     *
     * @return $this
     */
    public function setBuilder(BuilderInterface $builder)
    {
        $this->builder = $builder;

        return $this;
    }

    /**
     * @return string
     */
    public function getReportCacheDir()
    {
        return $this->reportCacheDir;
    }

    /**
     * @param  string$reportCacheDir
     *
     * @return $this
     */
    public function setReportCacheDir($reportCacheDir)
    {
        $this->reportCacheDir = $reportCacheDir;

        return $this;
    }

    /**
     * @return bool
     */
    public function hasColumnWidths() {
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
     * @param $columnWidths
     * @return $this
     */
    public function setColumnWidths(array $columnWidths)
    {
        $this->columnWidths = $columnWidths;

        return $this;
    }

    /**
     * @return bool
     */
    public function hasColumnStyles()
    {
        $columnStyles = $this->getColumnStyles();

        return !empty($columnStyles);
    }

    /**
     * @param  int $columnIndex
     *
     * @return bool
     */
    public function hasColumnStylesForColumn($columnIndex)
    {
        if (!$this->hasColumnStyles()) {
            return false;
        }

        $columnStyles = $this->getColumnStyles();

        return isset($columnStyles[$columnIndex]) ? true : false;
    }

    /**
     * @param  int $columnIndex
     *
     * @return bool|array
     */
    public function getColumnStylesForColumn($columnIndex)
    {
        if (!$this->hasColumnStylesForColumn($columnIndex)) {
            return false;
        }

        $columnStyles = $this->getColumnStyles();

        return isset($columnStyles[$columnIndex]) ? $columnStyles[$columnIndex] : false;
    }

    /**
     * TODO: Allow specific (col, row) styles
     * TODO: Allow text styles and colours to be applied
     *
     * @param $columnIndex
     * @return array
     */
    private function getPhpSpreadsheetColumnStylesForColumn($columnIndex) {
        $phpExcelStyleArray = [];

        $styleArray = $this->getColumnStylesForColumn($columnIndex);

        if (empty($styleArray)) {
            return $phpExcelStyleArray;
        }

        // Alignment
        if (isset($styleArray['alignment'])) {
            if ($styleArray['alignment'] === self::ALIGNMENT_CENTER) {
                $phpExcelStyleArray['alignment']['horizontal'] = Alignment::HORIZONTAL_CENTER;
            } else if ($styleArray['alignment'] === self::ALIGNMENT_LEFT) {
                $phpExcelStyleArray['alignment']['horizontal'] = Alignment::HORIZONTAL_LEFT;
            } else if ($styleArray['alignment'] === self::ALIGNMENT_RIGHT) {
                $phpExcelStyleArray['alignment']['horizontal'] = Alignment::HORIZONTAL_RIGHT;
            }
        }

        if (isset($styleArray['fill'])) {
            $phpExcelStyleArray['fill'] = [
                'type' => Fill::FILL_SOLID,
                'color' => [
                    'rgb' => $styleArray['fill']
                ]
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
     * @param  array $columnStyles
     *
     * @return $this
     */
    public function setColumnStyles(array $columnStyles)
    {
        $this->columnStyles = $columnStyles;

        return $this;
    }

    /**
     * @return bool
     */
    public function hasSheets()
    {
        $sheets = $this->sheets;

        return !empty($sheets);
    }

    /**
     * @param  array $sheets
     *
     * @return $this
     */
    public function setSheets(array $sheets)
    {
        $this->sheets = $sheets;

        return $this;
    }

    /**
     * @return array
     */
    public function getSheets()
    {
        return $this->sheets;
    }

    /**
     * @param  string|array $sheetTitles
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

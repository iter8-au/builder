<?php
namespace Builder;

/**
 * Helper for easily generating cached Excel or CSV files, ready to download, from an array of data.
 *
 * Class ReportService
 * @package SD\Vendo\Service
 */
class Builder
{
    const REPORT_EXCEL = 0;
    const REPORT_CSV = 1;

    const ALIGNMENT_CENTER  = 0;
    const ALIGNMENT_LEFT    = 1;
    const ALIGNMENT_RIGHT   = 2;

    /**
     * @var int
     */
    private $reportType;

    /**
     * @var \PHPExcel
     */
    private $phpexcel;

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

    public function __construct(
        \PHPExcel $phpexcel,
        $reportCacheDir
    )
    {
        $this->setPHPExcel($phpexcel);
        $this->setReportCacheDir($reportCacheDir);

        // Default the report to Excel format (using PHPExcel)
        $this->setReportType(self::REPORT_EXCEL);
    }

    /**
     * Generate the final report using whatever the set format is
     */
    public function generate(
        $unlinkFlag = true
    )
    {
        // Determine which format we are using and call the apppriate method
        if ($this->getReportType() === self::REPORT_EXCEL) {
            $this->generateExcel();
        } else {
            throw new \UnexpectedValueException("Attempted to generate a report in an unsupported format.");
        }
    }

    public function generateExcel()
    {
        $objPHPExcel = $this->getPhpexcel();
        $reportArray = $this->getData();

        // Set properties from Service values
        $objPHPExcel->getProperties()->setCreator($this->getCreator());
        $objPHPExcel->getProperties()->setLastModifiedBy($this->getCreator());
        $objPHPExcel->getProperties()->setTitle($this->getTitle());
        $objPHPExcel->getProperties()->setSubject($this->getTitle());
        $objPHPExcel->getProperties()->setDescription($this->getDescription());

        if ($this->hasSheets()) {
            // Multiple sheets
            $this->createSheets(
                $objPHPExcel
            );
        } else {
            // Single sheet - these will be an array and a string
            $reportArray = $this->getData();
            $sheetTitle = $this->getSheetTitles();

            $objPHPExcel->setActiveSheetIndex(0);

            $this->createSheet(
                $objPHPExcel,
                $reportArray,
                $sheetTitle
            );
        }

        // Output headers
        header("Pragma: public");
        header("Expires: 0");
        header("Cache-Control: must-revalidate, post-check=0, pre-check=0");
        header("Content-Type: application/force-download");
        header("Content-Type: application/octet-stream");
        header("Content-Type: application/download");;
        header("Content-Disposition: attachment;filename=" . $this->getFilename() . ".xlsx");
        header("Content-Transfer-Encoding: binary ");

        $objWriter = \PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
        ob_end_clean();

        // Write output then exit() to prevent page content being appended to .xlsx file
        $filePath = $this->getReportCacheDir() . "\\" . rand(0, getrandmax()) . rand(0, getrandmax()) . ".tmp";
        $objWriter->save($filePath);
        readfile($filePath);
        unlink($filePath);
        exit();
    }

    /**
     * @param $phpExcel
     */
    public function createSheets(&$phpExcel)
    {
        $sheets = $this->getSheets();

        $titles = $this->getSheetTitles();

        $totalSheets = count($sheets);
        $sheetCount = 0;

        if (empty($sheets) || !is_array($sheets)) {
            throw new \UnexpectedValueException(
                'Expected an array of sheets data but got an empty value or non-array.'
            );
        }

        // We have to set the initial active sheet
        $phpExcel->setActiveSheetIndex($sheetCount);

        foreach ($sheets as $sheet) {
            $this->createSheet(
                $phpExcel,
                $sheet,
                $titles[$sheetCount]
            );

            // Only create a new sheet if we actually have a data array for it
            if ($sheetCount < ($totalSheets - 1)) {
                $phpExcel->createSheet();

                // Increment the active sheet count and move to that sheet
                $sheetCount++;
                $phpExcel->setActiveSheetIndex($sheetCount);
            }
        }

        // Finally switch back to the first sheet
        $phpExcel->setActiveSheetIndex(0);

        return;
    }

    public function createSheet(
        &$phpExcel,
        $data,
        $title
    ) {
        // Style settings for agent headers
        // http://stackoverflow.com/questions/12918586/phpexcel-specific-cell-formatting-from-style-object
        $styleArray = array(
            'alignment' => array(
                'horizontal' => \PHPExcel_Style_Alignment::HORIZONTAL_CENTER,
            ),
            'font' => array(
                'color' => array(
                    'rgb' => 'FFFFFF'
                ),
                'bold' => true
            ),
            'fill' => array(
                'type' => \PHPExcel_Style_Fill::FILL_SOLID,
                'color' => array('rgb' => '000000')
            )
        );

        // Check if we are setting any custom column values
        if ($this->hasColumnWidths()) {
            // We have a numeric index array, so create an array of letters that we can use to map to Excel columns.
            // e.g. 0 = A, 3 = D, etc.
            $columns = range('A', 'Z');

            // Loop through all of our column values -  we only set values for columns that we actually have
            foreach ($this->getColumnWidths() as $columnKey => $columnWidth) {
                $phpExcel->getActiveSheet()->getColumnDimension($columns[$columnKey])->setWidth($columnWidth);
            }

        }

        // The row needs to start at 1 at the beginning of execution
        // the top left corner of the sheet is actually position (col = 0, row = 1)
        // http://stackoverflow.com/questions/2584954/phpexcel-how-to-set-cell-value-dynamically
        $row = 1;

        // Build column headers
        $col = 0;
        foreach (array_keys($data[0]) as $key) {
            // Set the header value
            $phpExcel->getActiveSheet()->setCellValueByColumnAndRow($col, $row, $key);

            // Apply header style to column headers
            $phpExcel->getActiveSheet()->getStyleByColumnAndRow($col, $row)->applyFromArray($styleArray);

            $col++;
        }

        $row++;

        // Output each record
        foreach ($data as $record) {

            $col = 0;
            foreach ($record as $column) {
                $phpExcel->getActiveSheet()->setCellValueByColumnAndRow($col, $row, $column);

                if ($this->hasColumnStylesForColumn($col)) {
                    // Apply style to column
                    $phpExcel->getActiveSheet()->getStyleByColumnAndRow($col, $row)->applyFromArray(
                        $this->getPHPExcelColumnStylesForColumn($col)
                    );
                }

                $col++;
            }

            $row++;
        }

        // Rename sheet
        $phpExcel->getActiveSheet()->setTitle($title);

        return;
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
     * @param int $reportType
     * @return $this
     */
    public function setReportType($reportType)
    {
        $this->reportType = $reportType;

        return $this;
    }

    /**
     * @return mixed
     */
    public function getCreator()
    {
        return $this->creator;
    }

    /**
     * @param mixed $creator
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
     * @param $data
     * @return $this
     */
    public function setData($data)
    {
        $this->data = $data;

        return $this;
    }

    /**
     * @return mixed
     */
    public function getDescription()
    {
        return $this->description;
    }

    /**
     * @param mixed $description
     * @return $this
     */
    public function setDescription($description)
    {
        $this->description = $description;

        return $this;
    }

    /**
     * @return mixed
     */
    public function getFilename()
    {
        return $this->filename;
    }

    /**
     * @param mixed $filename
     * @return $this
     */
    public function setFilename($filename)
    {
        $this->filename = $filename;

        return $this;
    }

    /**
     * TODO: If > 31 characters, and PHPExcel, then sub_str? (31 characters in an Excel title limit)
     *
     * @return mixed
     */
    public function getTitle()
    {
        return $this->title;
    }

    /**
     * @param mixed $title
     * @return $this
     */
    public function setTitle($title)
    {
        $this->title = $title;

        return $this;
    }

    /**
     * @return \PHPExcel
     */
    public function getPhpexcel()
    {
        return $this->phpexcel;
    }

    /**
     * @param \PHPExcel $phpexcel
     * @return $this
     */
    public function setPhpexcel(\PHPExcel $phpexcel)
    {
        $this->phpexcel = $phpexcel;

        return $this;
    }

    /**
     * @return mixed
     */
    public function getReportCacheDir()
    {
        return $this->reportCacheDir;
    }

    /**
     * @param $reportCacheDir
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
     * @param $columnIndex
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
     * @param $columnIndex
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
    private function getPHPExcelColumnStylesForColumn($columnIndex) {
        $phpExcelStyleArray = [];

        $styleArray = $this->getColumnStylesForColumn($columnIndex);

        if (empty($styleArray)) {
            return $phpExcelStyleArray;
        }

        // Alignment
        if (isset($styleArray['alignment'])) {
            if ($styleArray['alignment'] === self::ALIGNMENT_CENTER) {
                $phpExcelStyleArray['alignment']['horizontal'] = \PHPExcel_Style_Alignment::HORIZONTAL_CENTER;
            } else if ($styleArray['alignment'] === self::ALIGNMENT_LEFT) {
                $phpExcelStyleArray['alignment']['horizontal'] = \PHPExcel_Style_Alignment::HORIZONTAL_LEFT;
            } else if ($styleArray['alignment'] === self::ALIGNMENT_RIGHT) {
                $phpExcelStyleArray['alignment']['horizontal'] = \PHPExcel_Style_Alignment::HORIZONTAL_RIGHT;
            }
        }

        if (isset($styleArray['fill'])) {
            $phpExcelStyleArray['fill'] = [
                'type' => \PHPExcel_Style_Fill::FILL_SOLID,
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
     * @param array $columnStyles
     * @return $this
     */
    public function setColumnStyles(array $columnStyles)
    {
        $this->columnStyles = $columnStyles;

        return $this;
    }


    public function hasSheets()
    {
        $sheets = $this->sheets;

        return !empty($sheets);
    }

    public function setSheets(array $sheets)
    {
        $this->sheets = $sheets;
    }

    public function getSheets()
    {
        return $this->sheets;
    }

    public function setSheetTitles($sheetTitles)
    {
        $this->sheetTitles = $sheetTitles;
    }

    public function getSheetTitles()
    {
        return $this->sheetTitles;
    }
}

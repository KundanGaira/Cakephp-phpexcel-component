<?php
/*
For cakephp version 2.*
*/
class PhpExcelComponent extends Component {

    private $phpExcelName = 'PHPExcel';
    private $objPHPExcel = null, $objWorksheet = null;
    private $inputFileType = 'Excel2007';
    private $defaults = array('extension' => '.xlsx', 'excelName' => 'ExcelSheet', 'sheet1Name' => 'Sheet1');
    private $alphabets = null;

    /*
     * Create workbook,and return 
     */

    public function createExcel() {
        $loadStatus = App::import('Vendor', 'PHPExcel'); // Load PHPExcel from vender location
        if (!$loadStatus) {
            $msg = 'Unable to load ' . $this->phpExcelName . '.';
            $this->_requestError($msg);
        }
        $this->objPHPExcel = new PHPExcel(); // Make excel object, globally accessable

        $this->objWorksheet = $this->objPHPExcel->getActiveSheet(); // Make current sheet object, globally accessable
        $this->objWorksheet->setTitle($this->defaults['sheet1Name']);
        return $this->objPHPExcel;
    }

    public function openExcel($excelFileName) {
        $loadStatus = App::import('Vendor', 'PHPExcel'); // Load PHPExcel from vender location
        if (!$loadStatus) {
            $msg = 'Unable to load ' . $this->phpExcelName . '.';
            $this->_requestError($msg);
        }
        $objReader = PHPExcel_IOFactory::createReader($this->inputFileType); // Create reader
        $objReader->setIncludeCharts(TRUE);  // If charts are avilable use them
        
        if(!file_exists($excelFileName)){
            $this->_requestError('Unable to locate '.$excelFileName.' !');
        }
        $this->objPHPExcel = $objReader->load($excelFileName); // Make excel object, globally accessable
        $this->objWorksheet = $this->objPHPExcel->getActiveSheet(); // Make current sheet object, globally accessable

        return $this->objPHPExcel;
    }

    public function selectSheetByName($sheetName) {
        $this->objWorksheet = $this->objPHPExcel->getSheetByName($sheetName);
        return $this->objWorksheet;
    }

    /*
     * To create a new sheet at a particular index.
     * If sheets previous to that index not exists ,In this case it will create them as blank
     */

    public function additonalSheet($sheetLocIndex = null, $sheetName = null) {
        // index 1 based
        $objPHPExcel = $this->objPHPExcel;
        $alreadySheetsCount = $objPHPExcel->getSheetCount();  // Total number of existing sheets
        // If  total sheets are less then passed sheet index($sheetNo), Create rest sheets as empty 
        if ($sheetLocIndex > $alreadySheetsCount) {
            for ($i = $alreadySheetsCount; $i < $sheetLocIndex; $i++) {
                $objPHPExcel->createSheet($i);
                $objPHPExcel->setActiveSheetIndex($i);
                $objPHPExcel->getActiveSheet()->setTitle('Sheet' . ($i + 1));
            }
            $this->objWorksheet = $objPHPExcel->getActiveSheet(); // Make object accessable globally
        }
        if ($sheetName) {
            $this->objWorksheet->setTitle($sheetName);
        }
        $this->objWorksheet = $this->objPHPExcel->getActiveSheet(); // Make object accessable globally
    }

    public function downloadFile($fileName = null) {
        $objWriter = PHPExcel_IOFactory::createWriter($this->objPHPExcel, $this->inputFileType);  // writer object

        $objWriter->setIncludeCharts(TRUE); // Include charts if any
        $this->objPHPExcel->setActiveSheetIndex(0);  // Make first sheet as active.

        if ($fileName == null) {
            $fileName = $this->defaults['excelName']; // Default name 
        }
        $fileName = $fileName . $this->defaults['extension'];  // Full name

        header('Content-type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header('Content-Disposition: attachment; filename=' . $fileName);
        // Write file to the browser
        $objWriter->save('php://output');
        die;
    }

    public function _makeChartFromSheetData($dataStartCol, $dataStartRow, $dataEndCol, $dataEndRow, $chartHeight, $chartWidth) {
        $worksheetName = $this->objWorksheet->getTitle();

        $seriesStartRow = $dataStartRow + 1;
        $seriesTotalCols = $dataStartCol + ($dataEndCol - 1);
        $dataSeriesLabels = $xAxisTickValues = $dataSeriesValues = array();
        $totalRows = $dataEndRow - 1;

        // Basic structure  of sheet data
        /* $objWorksheet->fromArray(
          array(
          array('', 2010, 2011, 2012),
          array('Q1', 12, 15, 21),
          array('Q2', 56, 73, 86),
          array('Q3', 52, 61, 69),
          array('Q4', 30, 32, 0),
          )
          ); */

        // Labels for dataseries.  Direction ===>
        for ($i = $dataStartCol; $i <= $seriesTotalCols; $i++) {
            array_push($dataSeriesLabels, new PHPExcel_Chart_DataSeriesValues('String', $worksheetName . '!$' . $this->alphabets[$dataStartCol] . '$' . $i, NULL, 1));
        }

        for ($j = $dataStartRow; $j <= $totalRows; $j++) {
            array_push($xAxisTickValues, new PHPExcel_Chart_DataSeriesValues('String', $worksheetName . '!$' . $this->alphabets[$j] . '$' . $dataStartCol, NULL, 1));
        }

        for ($i = $dataStart; $i <= $dataEnd; $i++) {
            array_push($dataSeriesValues, new PHPExcel_Chart_DataSeriesValues('Number', $worksheetName . '!$' . $this->alphabets[$col + 1] . '$' . $i . ':$' . $this->alphabets[$col + 3] . '$' . $i, NULL, 3));
        }

        $series = new PHPExcel_Chart_DataSeries(
                PHPExcel_Chart_DataSeries::TYPE_BARCHART, // plotType
                PHPExcel_Chart_DataSeries::GROUPING_STACKED, // plotGrouping
                range(0, count($dataSeriesValues) - 1), // plotOrder
                $dataSeriesLabels, // plotLabel
                $xAxisTickValues, // plotCategory
                $dataSeriesValues        // plotValues
        );
        // new PHPExcel_Chart_DataSeries($plotType, $plotGrouping, $plotOrder, $plotLabel, $plotCategory, $plotValues, $smoothLine, $plotStyle);
        $series->setPlotDirection(PHPExcel_Chart_DataSeries::DIRECTION_COL);
        //	Set the series in the plot area
        $plotArea = new PHPExcel_Chart_PlotArea(NULL, array($series));
        //	Set the chart legend
        if ($type == 3) {
            $legend = null;
        } else {
            $legend = new PHPExcel_Chart_Legend(PHPExcel_Chart_Legend::POSITION_RIGHT, NULL, false);
        }
        //$legend->
        //$legend->setOverlay(true);
        $chrtTitle = ' ';
        $title = new PHPExcel_Chart_Title($chrtTitle);
        $yAxisLabel = new PHPExcel_Chart_Title('');
        $chart = new PHPExcel_Chart(
                null, // name
                $title, // title
                $legend, // legend
                $plotArea, // plotArea
                true, // plotVisibleOnly
                0, // displayBlanksAs
                NULL, // xAxisLabel
                $yAxisLabel  // yAxisLabel
        );
        //	Set the position where the chart should appear in the worksheet
        $chart->setTopLeftPosition($this->alphabets[$col] . $row);
        $chart->setBottomRightPosition($this->alphabets[$col + $width] . ($row + $chartHeight));
        //
        //	Add the chart to the worksheet

        $this->objWorksheet->addChart($chart);
    }

    /*
     * ************** Style related functions*****************************
     */
    /*     * *****Fill colour of a particular cell.Colour code in hex value without "#" *********** */

    function fillCellColour($cellDim, $color) {
        $cellBckStyle = array(
            'fill' => array(
                'type' => PHPExcel_Style_Fill::FILL_SOLID,
                'color' => array('rgb' => $color)
        ));
        $this->objWorksheet->getStyle($cellDim)->applyFromArray($cellBckStyle); // fill colour
    }

    /*     * *****Draw line.For example botton line for range "A3:A20" *********** */

    // Allowed type: "bottom","right","left","top"
    function drawLine($lineRange, $type, $colour = 'D9D9D9') {
        $border_style = array('borders' => array($type =>
                array('style' =>
                    PHPExcel_Style_Border::BORDER_THIN,
                    'color' => array('rgb' => $colour)
        )));
        $this->objWorksheet->getStyle($lineRange)->applyFromArray($border_style);
    }

    function addComment($cellDim, $comment) {
        $this->objWorksheet->getComment($cellDim)->getText()->createTextRun($comment); // Add comment
    }

    function writeCellValue($cell, $value) {
        $this->objWorksheet->getCell($cell)->setValue($value);
    }

    function getCellValue($cell) {
        return $this->objWorksheet->getCell($cell)->getCalculatedValue();
    }

    function mapAlphabets($colIndex) {
        if (empty($this->alphabets)) {
            $this->alphabets(1);
        }
        if (!empty($this->alphabets[$colIndex])) {
            return $this->alphabets[$colIndex];
        } else {
            $this->_requestError('Invalid column Index passed !');
        }
    }

    function getTotalRows($column = 0) {
        $returnVal = 1;
        if (!empty($this->objWorksheet->getHighestRow($column))) {
            $returnVal = $this->objWorksheet->getHighestRow($column);
        }
        return $returnVal;
    }

    public function alphabets($level) {
        //  Alphabets Array
        $this->alphabets = $alphabets = range('A', 'Z'); // Array containing latters from A to Z
        for ($i = 0; $i < $level; $i++) {
            foreach ($alphabets as $alpha) {
                array_push($this->alphabets, $alphabets[$i] . $alpha);
            }
        }
    }

    function getExcelObj() {
        return $this->objPHPExcel;
    }

    function getWorksheetObj() {
        return $this->objWorksheet;
    }

    public function _requestError($msg) {
        echo json_encode(array('status' => 'error', 'msg' => $msg));
        die;
    }

}

<?php

/**
 * ExportXLSX controller.
 *
 * This file will perform all function related with export of xls sheet along with chart
 *
 * @copyright     EQUIST
 * @package       EQUIST.Controller
 * @since         26 Oct 2015
 */
App::uses('AppController', 'Controller');

class ExportController extends AppController {

    public $chartMinLength = 4;
    public $chartMinwidth = 3;
    public $csvDelimiter = ',';
    public $params_missing_error = 'Parameters missing !';
    public $lang_data = null;
    public $country_data = null;
    public $alphabets = [];
    public $lastUsedRow = 10;

    //public $uses = array('');

    /*
      function to execute before any controller action
     */
    public function beforeFilter() {
        parent::beforeFilter();
    }

    public function index() {
        // Initail setting
        App::import('Vendor', 'PHPExcel'); // Load PHPExcel
        $this->autoRender = FALSE; // No view

        $objPHPExcel = new PHPExcel();
        $objWorksheet = $objPHPExcel->getActiveSheet();

        //  Alphabets Array
        $this->alphabets = $alphabetsTmp = range('A', 'Z'); // Array containing latters from A to Z
        foreach ($alphabetsTmp as $alp) {
            array_push($this->alphabets, 'A' . $alp); // Array containing latters from A to Z and AA,AB,AC and so on..
        }

        $selLanguage = Configure::read('Config.language'); // default language
        //  Get language file
        $lang_data_file = file_get_contents(DATA_FILE_PATH . '/lang_data/master_tbls_' . $selLanguage . '.json');
        $this->lang_data = json_decode($lang_data_file, true);

        /*  Get country file
         *  First get selected coutry name from cookies
         * 
         */
        if (!empty($_COOKIE[SEL_COUNTRY_COOKIE_NAME])) {
            $selCountry = $_COOKIE[SEL_COUNTRY_COOKIE_NAME];
        } else {
            $selCountry = DEF_AREA_ID;
        }
        // Country specific data file for subnational list

        $country_file = file_get_contents(DATA_FILE_PATH . '/country_data/' . $selCountry . '/country_master_tbl_' . $selLanguage . '.json');
        $this->country_data = json_decode($country_file, true);

//        echo'<pre>';
//        print_r($this->data);
//        die;
        if (empty($this->data['csv'])) {
            echo json_encode(array('msg' => $this->params_missing_error));
            die;
        }

        $csvData = $this->data['csv'];

        $dataArry = $this->_formatCsv($csvData);
        if (empty($dataArry)) {
            echo json_encode(array('msg' => 'Invalid data !'));
            die;
        }
        // Other post parametrs
        $otherParams = [];
        if (!empty($this->data['other'])) {
            $otherParams = $this->data['other'];
        }
        foreach ($dataArry as $indx => $chrtData) {
            $this->_createChartTable($chrtData, json_decode($otherParams[$indx], true), $objWorksheet);
        }

        $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
        $objWriter->setIncludeCharts(TRUE);

        // We'll be outputting an excel file
        header('Content-type: application/vnd.ms-excel');
        header('Content-Disposition: attachment; filename=' . 'Equist.xlsx');
        // Write file to the browser
        $objWriter->save('php://output');
        die;
    }

    /*
     * Create chartObj and series data for Excel file 
     * 
     */

    public function _createChartTable($chartData, $otherParams, $objWorksheet) {
        /*         * ********************* DATA FORMAT FOR HIGHCHART***************************
          array(
          array('', 2010, 2011, 2012), Step1
          array('Q1', 12, 15, 21),   Step2
          array('Q2', 56, 73, 86),   Step3
          array('Q3', 52, 61, 69),    ---
          array('Q4', 30, 32, 0),     ---
          );
         * ********************* DATA FORMAT*************************** */


        // Chart grouping type
        $chartGroup = '';
        if (!empty($otherParams['cGroup'])) {
            $chartGroup = $otherParams['cGroup'];
        }
        if (!in_array($chartGroup, array(PHPExcel_Chart_DataSeries::GROUPING_CLUSTERED, PHPExcel_Chart_DataSeries::GROUPING_STACKED))) {
            $chartGroup = PHPExcel_Chart_DataSeries::GROUPING_CLUSTERED;  // Default value=clustered;
        }
        // Start preparing data for csv/chart as per Highchart format.
        $workshettArry = $horizonLabl = $horizontalData = $verticalData = $xAxisTickValues = $dataSeriesLabels = $dataSeriesValues = [];

        $horizontalData = explode($this->csvDelimiter, $chartData['category']);
        array_shift($horizontalData); // remove unwanted data.
        $horzntlLen = count($horizontalData);

        $verticalData = $chartData['data'];
        $verticalLength = count($verticalData);

        //  Create data for horizontal axis
        for ($i = 0; $i <= $horzntlLen; $i++) { //length+1 first cell need to be empty
            if ($i == 0) {
                $horizonLabl[$i] = '';
            } else {
                $tmptxt = trim(str_replace('"', '', $horizontalData[$i - 1]));
                $tmptxt = $this->_translate($tmptxt);
                array_push($horizonLabl, $tmptxt);  // Step1
                array_push($dataSeriesLabels, new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$' . $this->alphabets[$i] . '$' . ($this->lastUsedRow + 1), NULL, 1)); //Map horizontal labels
            }
        }
        $workshettArry[0] = $horizonLabl; // Place at 0 index always
        //              Vertical labels and cell data
        for ($i = 0; $i < $verticalLength; $i++) {
            $row = explode($this->csvDelimiter, $verticalData[$i]);  // seprate each cell data from csv string
            $tmparr = [];
            for ($j = 0; $j <= $horzntlLen; $j++) {

                //  Prepare label
                if ($j == 0) {
                    $tmparr[$j] = $this->_translate($row[$j]); // Step 2 , First cell                     
                } else {
                    // Fix for scenario ,When first cell is blank
                    //var_dump(empty($row[$j]);die;
                    if (($j == 1) && ($row[$j] == ' ')) {
                        unset($row[$j]);
                        $row = array_values($row); // re index
                    }
                    $tmparr[$j] = $row[$j];  // Step2, rest cells
                }
            }
            array_push($workshettArry, $tmparr);
            array_push($dataSeriesValues, new PHPExcel_Chart_DataSeriesValues('Number', 'Worksheet!$' . $this->alphabets[$i + 1] . '$' . ($this->lastUsedRow + 2) . ':$' . $this->alphabets[$i + 1] . '$' . ($verticalLength + 1), NULL, $verticalLength)); // Map cell data,
        }


        $xAxisTickValues = array(
            new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$' . $this->alphabets[0] . '$' . ($this->lastUsedRow + 2) . ':$' . $this->alphabets[0] . '$' . ($verticalLength + 1), NULL, $verticalLength), //	Map vertical labels
        );
        $objWorksheet->fromArray($workshettArry);
        //  Build the dataseries

        $series = new PHPExcel_Chart_DataSeries(
                PHPExcel_Chart_DataSeries::TYPE_BARCHART, // plotType
                $chartGroup, // plotGrouping
                range(0, ($horzntlLen - 1)), // plotOrder
                $dataSeriesLabels, // plotLabel
                $xAxisTickValues, // plotCategory
                $dataSeriesValues        // plotValues
        );

        $series->setPlotDirection(PHPExcel_Chart_DataSeries::DIRECTION_COL);
        $plotArea = new PHPExcel_Chart_PlotArea(NULL, array($series));
        $legend = new PHPExcel_Chart_Legend(PHPExcel_Chart_Legend::POSITION_BOTTOM, NULL, false);
        $title = new PHPExcel_Chart_Title('Chart');
        $yAxisLabel = new PHPExcel_Chart_Title('');
        $chart = new PHPExcel_Chart(
                'chart', // name
                $title, // title
                $legend, // legend
                $plotArea, // plotArea
                true, // plotVisibleOnly
                0, // displayBlanksAs
                NULL, // xAxisLabel
                $yAxisLabel  // yAxisLabel
        );

        // Chart size and placement calculation. 
        if ($horzntlLen < $this->chartMinwidth) {
            $x2 = ($horzntlLen * 2) + ($this->chartMinwidth * 2);
        } else {
            $x2 = ($horzntlLen * 4);
            if ($x2 > 17) {
                $x2 = 17;
            }
        }
        if ($verticalLength < $this->chartMinLenght) {
            $y2 = ($this->chartMinLenght * 2) + 15;
        } else {
            $y2 = ($verticalLength * 2) + 15;
        }
        $y2 = $y2 + $this->lastUsedRow;
        $x1 = 0;
        $y1 = ($verticalLength + 3) + $this->lastUsedRow;

        $chart->setTopLeftPosition($this->alphabets[$x1] . $y1);
        $chart->setBottomRightPosition($this->alphabets[$x2] . $y2);
        $objWorksheet->addChart($chart);


        $this->lastUsedRow = $this->lastUsedRow + ($verticalLength + 1); // Set index for next chart location 
    }

    /*
     * Seprate category Data from rest data.
     * Return formatted array.Empty array in case of invalid data
     * 
     */

    public function _formatCsv($csv) {
        $returnArry = [];
        foreach ($csv as $c) {
            $tmp1 = str_getcsv($c, "\n");
            $pushArray = [];
            if (count($tmp1) > 1) {

                $pushArray['category'] = $tmp1[0];
                array_shift($tmp1);
                $pushArray['data'] = $tmp1; // remove first element.Already assigned to previous line
                array_push($returnArry, $pushArray);
            }
        }
        return $returnArry;
    }

    /*
     * Function to translate indicators 
     * @input:plain string
     * @output: translated string 
     * 
     */

    public function _translate($str) {
        $subGroupData = $this->lang_data['sgrpValueList'];
        $subNationalData = $this->country_data['groupList']['list'][GEOGRAPHY_TP]['selOpts'];
        $epiCauseData = $this->lang_data[EPIC_TP];
        $intvnsData = $this->lang_data[INTRV_TP];
        $bottelNeckData = $this->lang_data[BTLNK_TP];
        $startegyData = $this->lang_data[STRTG_TP];
        $shortName = 'snm';
        $longName = 'lnm';
        // Copy from layout/view
        $LANG_STATIC_TXT_OBJ['total_deaths_averatble'] = 'Total preventable deaths';
        $LANG_STATIC_TXT_OBJ['amenable_equity_gap'] = 'Excess deaths';
        $LANG_STATIC_TXT_OBJ['scenario_deaths_avertable'] = 'Deaths averted in scenario';

        // Check in subgroup list
        if (!empty($subGroupData[$str])) {
            return $subGroupData[$str];
        }
        // Check in sub-National list
        foreach ($subNationalData as $lvl) {
            if (!empty($lvl[$str])) {
                return $lvl[$str];
            }
        }
        // Check in EpiCauses list
        foreach ($epiCauseData as $epi) {
            if (!empty($epi['details'][$str])) {
                return $epi['details'][$str][$longName];
            }
        }
        // Check in EpiCauses list
        foreach ($epiCauseData as $epi) {
            if (!empty($epi['details'][$str])) {
                return $epi['details'][$str][$longName];
            }
        }
        // Check in Interventions list
        if (!empty($intvnsData[$str])) {
            return $intvnsData[$str][$shortName];
        }
        // Check in Bottleneck list
        if (!empty($bottelNeckData[$str])) {
            return $bottelNeckData[$str][$longName];
        }
        // Check in STRATEGY list
        if (!empty($startegyData[$str])) {
            return $startegyData[$str][$shortName];
        }
        // Special case
        if (!empty($LANG_STATIC_TXT_OBJ[$str])) {
            return $LANG_STATIC_TXT_OBJ[$str];
        }

        // Check if it is with in language file with "TXT" suffix
        if (__($str . '_TXT') !== $str . '_TXT') {
            return __($str . '_TXT');
        }

        // Use default case
        return __($str);
    }

    function test() {
        App::import('Vendor', 'PHPExcel'); // Load PHPExcel
        $objPHPExcel = new PHPExcel();
        $objWorksheet = $objPHPExcel->getActiveSheet();
        $objWorksheet->fromArray(
                array(
                    array('', 2010, 2011, 2012),
                    array('Q1', 12, 15, 21),
                    array('Q2', 56, 73, 86),
                    array('Q3', 52, 61, 69),
                    array('Q4', 30, 32, 0),
                )
        );
//	Set the Labels for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
        $dataSeriesLabels1 = array(
            new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$B$1', NULL, 1), //	2010
            new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$C$1', NULL, 1), //	2011
            new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$D$1', NULL, 1), //	2012
        );
//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
        $xAxisTickValues1 = array(
            new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$A$2:$A$5', NULL, 4), //	Q1 to Q4
        );
//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
        $dataSeriesValues1 = array(
            new PHPExcel_Chart_DataSeriesValues('Number', 'Worksheet!$B$2:$B$5', NULL, 4),
            new PHPExcel_Chart_DataSeriesValues('Number', 'Worksheet!$C$2:$C$5', NULL, 4),
            new PHPExcel_Chart_DataSeriesValues('Number', 'Worksheet!$D$2:$D$5', NULL, 4),
        );
//	Build the dataseries
        $series1 = new PHPExcel_Chart_DataSeries(
                PHPExcel_Chart_DataSeries::TYPE_AREACHART, // plotType
                PHPExcel_Chart_DataSeries::GROUPING_PERCENT_STACKED, // plotGrouping
                range(0, count($dataSeriesValues1) - 1), // plotOrder
                $dataSeriesLabels1, // plotLabel
                $xAxisTickValues1, // plotCategory
                $dataSeriesValues1          // plotValues
        );
//	Set the series in the plot area
        $plotArea1 = new PHPExcel_Chart_PlotArea(NULL, array($series1));
//	Set the chart legend
        $legend1 = new PHPExcel_Chart_Legend(PHPExcel_Chart_Legend::POSITION_TOPRIGHT, NULL, false);
        $title1 = new PHPExcel_Chart_Title('Test %age-Stacked Area Chart');
        $yAxisLabel1 = new PHPExcel_Chart_Title('Value ($k)');
//	Create the chart
        $chart1 = new PHPExcel_Chart(
                'chart1', // name
                $title1, // title
                $legend1, // legend
                $plotArea1, // plotArea
                true, // plotVisibleOnly
                0, // displayBlanksAs
                NULL, // xAxisLabel
                $yAxisLabel1 // yAxisLabel
        );
//	Set the position where the chart should appear in the worksheet
        $chart1->setTopLeftPosition('A7');
        $chart1->setBottomRightPosition('H20');
//	Add the chart to the worksheet
        $objWorksheet->addChart($chart1);
//	Set the Labels for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
        $dataSeriesLabels2 = array(
            new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$B$1', NULL, 1), //	2010
            new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$C$1', NULL, 1), //	2011
            new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$D$1', NULL, 1), //	2012
        );
//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
        $xAxisTickValues2 = array(
            new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$A$2:$A$5', NULL, 4), //	Q1 to Q4
        );
//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
        $dataSeriesValues2 = array(
            new PHPExcel_Chart_DataSeriesValues('Number', 'Worksheet!$B$2:$B$5', NULL, 4),
            new PHPExcel_Chart_DataSeriesValues('Number', 'Worksheet!$C$2:$C$5', NULL, 4),
            new PHPExcel_Chart_DataSeriesValues('Number', 'Worksheet!$D$2:$D$5', NULL, 4),
        );
//	Build the dataseries
        $series2 = new PHPExcel_Chart_DataSeries(
                PHPExcel_Chart_DataSeries::TYPE_BARCHART, // plotType
                PHPExcel_Chart_DataSeries::GROUPING_STANDARD, // plotGrouping
                range(0, count($dataSeriesValues2) - 1), // plotOrder
                $dataSeriesLabels2, // plotLabel
                $xAxisTickValues2, // plotCategory
                $dataSeriesValues2        // plotValues
        );
//	Set additional dataseries parameters
//		Make it a vertical column rather than a horizontal bar graph
        $series2->setPlotDirection(PHPExcel_Chart_DataSeries::DIRECTION_COL);
//	Set the series in the plot area
        $plotArea2 = new PHPExcel_Chart_PlotArea(NULL, array($series2));
//	Set the chart legend
        $legend2 = new PHPExcel_Chart_Legend(PHPExcel_Chart_Legend::POSITION_RIGHT, NULL, false);
        $title2 = new PHPExcel_Chart_Title('Test Column Chart');
        $yAxisLabel2 = new PHPExcel_Chart_Title('Value ($k)');
//	Create the chart
        $chart2 = new PHPExcel_Chart(
                'chart2', // name
                $title2, // title
                $legend2, // legend
                $plotArea2, // plotArea
                true, // plotVisibleOnly
                0, // displayBlanksAs
                NULL, // xAxisLabel
                $yAxisLabel2 // yAxisLabel
        );
//	Set the position where the chart should appear in the worksheet
        $chart2->setTopLeftPosition('I7');
        $chart2->setBottomRightPosition('P20');
//	Add the chart to the worksheet
        $objWorksheet->addChart($chart2);

// Save Excel 2007 file
        $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
        $objWriter->setIncludeCharts(TRUE);
        // We'll be outputting an excel file
        header('Content-type: application/vnd.ms-excel');
        header('Content-Disposition: attachment; filename=' . 'Equist.xlsx');
        // Write file to the browser
        $objWriter->save('php://output');
        die;
    }

    function single() {
        App::import('Vendor', 'PHPExcel'); // Load PHPExcel
        $objPHPExcel = new PHPExcel();
        $objWorksheet = $objPHPExcel->getActiveSheet();
        $objWorksheet->fromArray(
                array(
                    array('tet', 'test', 'tes','test'),
                    array('', 2010, 2011, 2012),
                    array('Q1', 12, 15, 21),
                    array('Q2', 56, 73, 86),
                    array('Q3', 52, 61, 69),
                    array('Q4', 30, 32, 0),
                )
        );
//	Set the Labels for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
        $dataSeriesLabels = array(
            new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$B$2', NULL, 1), //	2010
            new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$C$2', NULL, 1), //	2011
            new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$D$2', NULL, 1), //	2012
        );
//	Set the X-Axis Labels
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
        $xAxisTickValues = array(
            new PHPExcel_Chart_DataSeriesValues('String', 'Worksheet!$A$3:$A$6', NULL, 4), //	Q1 to Q4
        );
//	Set the Data values for each data series we want to plot
//		Datatype
//		Cell reference for data
//		Format Code
//		Number of datapoints in series
//		Data values
//		Data Marker
        $dataSeriesValues = array(
            new PHPExcel_Chart_DataSeriesValues('Number', 'Worksheet!$B$3:$B$6', NULL, 4),
            new PHPExcel_Chart_DataSeriesValues('Number', 'Worksheet!$C$3:$C$6', NULL, 4),
            new PHPExcel_Chart_DataSeriesValues('Number', 'Worksheet!$D$3:$D$6', NULL, 4),
        );
//	Build the dataseries
        $series = new PHPExcel_Chart_DataSeries(
                PHPExcel_Chart_DataSeries::TYPE_BARCHART, // plotType
                PHPExcel_Chart_DataSeries::GROUPING_CLUSTERED, // plotGrouping
                range(0, count($dataSeriesValues) - 1), // plotOrder
                $dataSeriesLabels, // plotLabel
                $xAxisTickValues, // plotCategory
                $dataSeriesValues        // plotValues
        );
//	Set additional dataseries parameters
//		Make it a horizontal bar rather than a vertical column graph
        $series->setPlotDirection(PHPExcel_Chart_DataSeries::DIRECTION_COL);
//	Set the series in the plot area
        $plotArea = new PHPExcel_Chart_PlotArea(NULL, array($series));
//	Set the chart legend
        $legend = new PHPExcel_Chart_Legend(PHPExcel_Chart_Legend::POSITION_RIGHT, NULL, false);
        $title = new PHPExcel_Chart_Title('Test Bar Chart');
        $yAxisLabel = new PHPExcel_Chart_Title('Value ($k)');
//	Create the chart
        $chart = new PHPExcel_Chart(
                'chart1', // name
                $title, // title
                $legend, // legend
                $plotArea, // plotArea
                true, // plotVisibleOnly
                0, // displayBlanksAs
                NULL, // xAxisLabel
                $yAxisLabel  // yAxisLabel
        );
//	Set the position where the chart should appear in the worksheet
        $chart->setTopLeftPosition('A7');
        $chart->setBottomRightPosition('H20');
//	Add the chart to the worksheet
        $objWorksheet->addChart($chart);
// Save Excel 2007 file
        $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
        $objWriter->setIncludeCharts(TRUE);
        header('Content-type: application/vnd.ms-excel');
        header('Content-Disposition: attachment; filename=' . 'Equist.xlsx');
        // Write file to the browser
        $objWriter->save('php://output');
        die;
    }

}

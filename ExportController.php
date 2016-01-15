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

    public $components = array('CommonMethod', 'Auth');
    public $objPHPExcel = null, $objWorksheet = null, $lang_data = null, $scr_desc_data = null, $default_lang = 'en';
    public $rel_data = null, $epi_nn = null, $epi_stunt, $epi_u5mr = null;
    public $sheetCol = 1; // IF you want to leave some columns blank  at extreme left
    public $userInfo = [], $alphabets = [],$selPackages=[];
    public $summaryTemplate = 'Scenario_Summary.xlsx';
    public $inputFileType = 'Excel2007';
    public $headingInfo = array('scnrName' => [0, 20], 'countryName' => [0, 16], 'trgetPpl' => [0, 16], 'trgetPpl2' => [6, 11],
        'epicause' => [0, 16], 'intrv' => [0, 16], 'btlnk' => [0, 16], 'btlnk2' => [0, 14], 'csp1' => [0, 16], 'csp2' => [0, 12], 'csp3' => [3, 12], 'csp4' => [13, 12],
        'impctDth' => [0, 16], 'epicChrt1' => [-1, 14], 'epicChrt11' => [3, 12], 'epicChrt12' => [16, 12], 'epicChrt13' => [29, 12], 'epicChrt21' => [-1, 14], 'epicChrt22' => [22, 14],
        'cost1' => [0, 16], 'cost21' => [0, 14], 'cost22' => [17, 14], 'cost23' => [31, 14], 'ee' => [0, 14]
    ); // [col,font] col is 0 index based 
    public $colIndexArray = array('btlnkSrvHead' => 17, 'causStrgyPck' => 29, 'csp_cause' => 3, 'csp_strtgy' => 13, 'epic2BaseHead' => 2);
    public $rowLimit = array('deprivedGrp' => 10);
    public $otherInfo = array('colLimit' => 38, 'depGrpNtlGap' => 4, 'epicGap' => 8, 'intrvGap' => 11, 'btlnkSvrGap' => 3, 'chartGrp1' => 22, 'chartGrp2' => 14
        , 'epicSheetSeriesGap' => 5); //colLimit=max columns will be used in sheet
    public $chartDim = array('epic11' => [0, 12], 'epic12' => [13, 12], 'epic13' => [26, 11], 'epic21' => [0, 20], 'epic22' => [21, 16],
        'cost1' => [0, 12], 'cost2' => [13, 12], 'cost3' => [26, 11]
    );
    public $secWidths = array(COMMUNITY_TP => 11, SCHEDULE_TP => 12, CLINICAL_TP => 12);
    public $colorArray = array('h1' => '00B0F0', 'h2' => 'D8D8D8', 'h3' => 'F2F2F2', 'SDM_1' => 'C5D9F1', 'SDM_2' => 'F1AC51', 'SDM_3' => 'C4D79B');
    public $skippedBTNKGids = array(BTLNK_HUMAN_RSRC_CVRG_GID, BTLNK_INITIAL_UTLZ_CVRG_GID);
    public $sdmColorCode = array(COMMUNITY_TP => COMMUNITY_CLR, SCHEDULE_TP => SCHEDULE_CLR, CLINICAL_TP => CLINICAL_CLR);
    public $sheetRow = 1;

    /*
      function to execute before any controller action
     */

    public function beforeFilter() {
        parent::beforeFilter();
        $this->autoRender = FALSE; // No view
    }

    public function makeSummary() {
        //      Accept post request only
        $this->autoRender = false;
        if ($this->request->is('post')) {
            $this->_validateRequest(array('scr', 'area_id','btnk'));  //      Post data validation
            $this->_initSetting();   //     Initial and common setting
            $this->_scrDescData($this->userInfo['userId'], $this->userInfo['areaId'], $this->userInfo['scnrId']);
            $this->_loadWorkbook();
            $this->_sheet1(1);
            $this->_sheet2(2);
            $this->_sheet3(3);

            $fileNameParms = array($this->userInfo['areaName'], $this->userInfo['scnrDataset'], $this->userInfo['scnrName']);
            $fileName = $this->_getFileName($fileNameParms);
            $fileName = $fileName . '.xlsx';
            
            $this->_downloadFile($fileName);
        }
    }
    
    public function makeEpic() {
        //      Accept post request only
        $this->autoRender = false;
        if ($this->request->is('post')) {
            $this->_validateRequest(array('scr', 'area_id','btnk'));  //      Post data validation
            $this->_initSetting();   //     Initial and common setting
            $this->_scrDescData($this->userInfo['userId'], $this->userInfo['areaId'], $this->userInfo['scnrId']);
            $this->_loadWorkbook();
            $this->_sheet2(1);
            
            $fileNameParms = array($this->userInfo['areaName'], $this->userInfo['scnrDataset'], $this->userInfo['scnrName'],__('Epic_priorities'));
            $fileName = $this->_getFileName($fileNameParms);
            $fileName = $fileName . '.xlsx';
            
            $this->_downloadFile($fileName);
        }
    }
    public function makeIntrv() {
        //      Accept post request only
        $this->autoRender = false;
        if ($this->request->is('post')) {
            $this->_validateRequest(array('scr', 'area_id','btnk'));  //      Post data validation
            $this->_initSetting();   //     Initial and common setting
            $this->_scrDescData($this->userInfo['userId'], $this->userInfo['areaId'], $this->userInfo['scnrId']);
            $this->_loadWorkbook();
            $this->_sheet3(1);
            
            $fileNameParms = array($this->userInfo['areaName'], $this->userInfo['scnrDataset'], $this->userInfo['scnrName'],__('Intervention'));
            $fileName = $this->_getFileName($fileNameParms);
            $fileName = $fileName . '.xlsx';
            
            $this->_downloadFile($fileName);
        }
    }

    public function _getNameFromIndGid($indId) {
        $returnName = '';
        // return name of indicator of given Id
        if ($this->lang_data == null) {
            $this->_langData();
        }
        // Check package name
        if (!empty($this->lang_data[SRV_DLV_MD_TP][$indId]['lnm'])) { // Check  for SDM
            $returnName = $this->lang_data[SRV_DLV_MD_TP][$indId]['lnm'];
        } else if (!empty($this->lang_data[PKG_TP][$indId]['lnm'])) {  // Check package name
            $returnName = $this->lang_data[PKG_TP][$indId]['lnm'];
        } else if (!empty($this->lang_data[INTRV_TP][$indId]['lnm'])) {   // Check intervention name
            $returnName = $this->lang_data[INTRV_TP][$indId]['lnm'];
        } else if (!empty($this->lang_data[BTLNK_TP][$indId]['lnm'])) {   // Check bottleneck name
            $returnName = $this->lang_data[BTLNK_TP][$indId]['lnm'];
        } else if (!empty($this->lang_data[CUS_TP][$indId]['lnm'])) {   // Check cause name
            $returnName = $this->lang_data[CUS_TP][$indId]['lnm'];
        } else if (!empty($this->lang_data[STRTG_TP][$indId]['lnm'])) {   // Check startegy name
            $returnName = $this->lang_data[STRTG_TP][$indId]['lnm'];
        }

        return $returnName;
    }

    public function _scrDescData($userId, $areaCode, $scrId) {
        //  Get scenario file
        $scrPath = $this->CommonMethod->getUserCountryScenarioFolderPath($userId, $areaCode, $scrId);

        $scrFile = $scrPath . 'decision_tree_data.json';
        if (!file_exists($scrFile)) {
            $this->requestError('Invalid scenario !');
        }
        $scrFileData = file_get_contents($scrFile);
        $this->scr_desc_data = json_decode($scrFileData, true);
    }

    public function _scrChartsData($userId, $areaCode, $scrId) {
        //  Get scenario file
        $scrPath = $this->CommonMethod->getUserCountryScenarioFolderPath($userId, $areaCode, $scrId);
        $prefix = $scrPath . 'chart_data_';
        $EPI_NN_file = $prefix . EPIC_GRP_NN . '.json';
        $EPI_STUNTING_file = $prefix . EPIC_GRP_STUNT . '.json';
        $EPI_U5MR_file = $prefix . EPIC_GRP_U5MR . '.json';

        if (file_exists($EPI_NN_file)) {
            $fileData = file_get_contents($EPI_NN_file);
            $this->epi_nn = json_decode($fileData, true);
        }

        if (file_exists($EPI_STUNTING_file)) {
            $fileData = file_get_contents($EPI_STUNTING_file);
            $this->epi_stunt = json_decode($fileData, true);
        }
        if (file_exists($EPI_U5MR_file)) {
            $fileData = file_get_contents($EPI_U5MR_file);
            $this->epi_u5mr = json_decode($fileData, true);
        }
    }

    public function _getFileName($params) {
        $fileName = 'EQUIST';
        foreach ($params as $prm) {
            if ($prm != null) {
                $fileName.='_' . str_replace(' ', '_', $prm);
            }
        }
        if (strlen($fileName) > 250) {
            $fileName = substr($fileName, 0, 249); // filename limit
        }
        return $fileName;
    }

    public function _getICNameSource($ICId, $type, $btnckId = null) {
        $returnArray = array();
        $indctName = '';
        $indctSrcName = '';

        if ($type == EPIC_TP) {
            if (!empty($this->rel_data['epicIndicatorRelList'][$ICId])) {
                $indctId = $this->rel_data['epicIndicatorRelList'][$ICId];  // Id of indicator
                // Get name form language json
                if (!empty($this->lang_data['indicatorList'][$indctId])) {
                    $indctName = $this->lang_data['indicatorList'][$indctId];
                } else {
                    $indctName = __($indctId);
                }
                $selLanguage = Configure::read('Config.language'); // default language
                $countryMasterFile = $this->CommonMethod->getSelLangCountryMasterFile($this->userInfo['areaId'], $selLanguage);
                $countryMasterFileData = file_get_contents($countryMasterFile);
                $countryMasterFileData = json_decode($countryMasterFileData, true);

                if (!empty($countryMasterFileData['INDCT_DATA_SRC_LIST'][$this->userInfo['scnrDataset']][$indctId])) {
                    $dataSrcArry = $countryMasterFileData['INDCT_DATA_SRC_LIST'][$this->userInfo['scnrDataset']][$indctId];
                    $selDprvGrp = $this->userInfo['depGroup'];
                    $DATA_SRC_INDEX_ARR = json_decode(DATA_SRC_INDEX_ARR, true);
                    if ($selDprvGrp == GEOGRAPHY_TP) {
                        $indctSrc = $dataSrcArry[$DATA_SRC_INDEX_ARR['SUBNATIONAL1']];
                    } else if ($selDprvGrp == QNUINTILE_TP) {
                        $indctSrc = $dataSrcArry[$DATA_SRC_INDEX_ARR[QNUINTILE_TP]];
                    } else {
                        $indctSrc = $dataSrcArry[$DATA_SRC_INDEX_ARR['OTHER_GROUP']];
                    }
                    if (!empty($countryMasterFileData['DATA_SRC_LIST'][$indctSrc])) {
                        $indctSrcName = $countryMasterFileData['DATA_SRC_LIST'][$indctSrc];
                    }
                }
            }
        } else if ($type == INTRV_TP) {
            if (!empty($this->rel_data['visDataIndctRel'][COVERAGE_INDCT][INTRV_TP][$btnckId][$ICId])) {
                $indctId = null;
                foreach ($this->rel_data['visDataIndctRel'][COVERAGE_INDCT][INTRV_TP][$btnckId][$ICId] as $indId => $v) {
                    $indctId = $indId;
                }

                if (!empty($this->lang_data['indicatorList'][$indctId])) {
                    $indctName = $this->lang_data['indicatorList'][$indctId]; // Name
                } else {
                    $indctName = __($indctId);
                }

                // Prepare source name
                // First check scenario file 
                $userDefined = null;
                if (!empty($this->scr_desc_data[BTLNK_TP][INTRV_TP][$ICId][$btnckId]['src'])) { // Intrv level
                    $userDefined = $this->scr_desc_data[BTLNK_TP][INTRV_TP][$ICId][$btnckId]['src'];
                } else {
                    if (!empty($this->rel_data['intrvPckgRel'][$ICId])) { // Pckg level
                        $pckg = $this->rel_data['intrvPckgRel'][$ICId];
                        if (!empty($this->scr_desc_data[BTLNK_TP][PKG_TP][$pckg][$btnckId]['src'])) {
                            $userDefined = $this->scr_desc_data[BTLNK_TP][PKG_TP][$pckg][$btnckId]['src'];
                        } else if (!empty($this->rel_data['pckgSdmRel'][$pckg])) {
                            $sdm = $this->rel_data['pckgSdmRel'][$pckg];
                            if (!empty($this->scr_desc_data[BTLNK_TP][SRV_DLV_MD_TP][$sdm][$btnckId]['src'])) {
                                $userDefined = $this->scr_desc_data[BTLNK_TP][SRV_DLV_MD_TP][$sdm][$btnckId]['src'];
                            }
                        }
                    }
                }

                if ($userDefined == null) {
                    $selLanguage = Configure::read('Config.language'); // default language
                    $countryMasterFile = $this->CommonMethod->getSelLangCountryMasterFile($this->userInfo['areaId'], $selLanguage);
                    $countryMasterFileData = file_get_contents($countryMasterFile);
                    $countryMasterFileData = json_decode($countryMasterFileData, true);
                    if (!empty($countryMasterFileData['INDCT_DATA_SRC_LIST'][$this->userInfo['scnrDataset']][$indctId])) {
                        $dataSrcArry = $countryMasterFileData['INDCT_DATA_SRC_LIST'][$this->userInfo['scnrDataset']][$indctId];
                        $selDprvGrp = $this->userInfo['depGroup'];
                        $DATA_SRC_INDEX_ARR = json_decode(DATA_SRC_INDEX_ARR, true);
                        if ($selDprvGrp == GEOGRAPHY_TP) {
                            $indctSrc = $dataSrcArry[$DATA_SRC_INDEX_ARR['SUBNATIONAL1']];
                        } else if ($selDprvGrp == QNUINTILE_TP) {
                            $indctSrc = $dataSrcArry[$DATA_SRC_INDEX_ARR[QNUINTILE_TP]];
                        } else {
                            $indctSrc = $dataSrcArry[$DATA_SRC_INDEX_ARR['OTHER_GROUP']];
                        }
                        if (!empty($countryMasterFileData['DATA_SRC_LIST'][$indctSrc])) {
                            $indctSrcName = $countryMasterFileData['DATA_SRC_LIST'][$indctSrc];
                        }
                    }
                } else {
                    $indctSrcName = $userDefined;
                }
            }
        }
        $returnArray = array(
            'name' => $indctName,
            'src' => $indctSrcName
        );
        return $returnArray;
    }

    public function _epicInfoMastr() {
        
    }

    public function _loadWorkbook() {
        App::import('Vendor', 'PHPExcel'); // Load PHPExcel

        /* EDIT */
        //  $objReader = PHPExcel_IOFactory::createReader($this->inputFileType);
        // $objReader->setIncludeCharts(TRUE);
        // $templateFile = TEMPLATE_FILES_PATH . $this->summaryTemplate;
        // $this->objPHPExcel = $objReader->load($templateFile);
        /* READ */
        $this->objPHPExcel = new PHPExcel();
    }

    // Validate request (POST) data 
    public function _validateRequest($paramsArray) {
        $requestData = $this->request->data;
        foreach ($paramsArray as $param) {
            if (empty($requestData[$param])) {
                $this->requestError('parameter : ' . $param . ' missing !');
            }
        }
    }

    // Terminate request with message 
    public function requestError($msg) {
        echo json_encode(array('status' => 'error', 'msg' => $msg));
        die;
    }

    // Set user infomation and scenario data to a global variable 
    public function _initSetting() {
        //  Default values
        $this->userInfo['scnrId'] = $this->userInfo['scnrName'] = $this->userInfo['scnrDataset'] = $this->userInfo['depGroup'] = null;
        $this->userInfo['areaName'] = null;
        $this->userInfo['areaId'] = DEF_AREA_ID;  // Default as definded in config file;

        if (empty($this->Auth->user('usr_NId'))) {
            $this->requestError('Session expired .Please login..!');
        }
        $this->userInfo['userId'] = $this->Auth->user('usr_NId');  // user id 
//   Check country name from cookies and post data
        if (!empty($_COOKIE['selCountry']) && ($_COOKIE['selCountry'] == $this->request->data['area_id'])) {
            $this->userInfo['areaId'] = $_COOKIE['selCountry'];
        }

//  Scenario id
        $this->userInfo['scnrId'] = $this->request->data['scr']; // request data already validated
        //  Get scenario master file
        $scrMasterPath = $this->CommonMethod->getUserCountryScenarioMasterFile($this->userInfo['userId'], $this->userInfo['areaId']);
        if (file_exists($scrMasterPath)) {
            $scrFileDataEncode = file_get_contents($scrMasterPath);
            $scrFileData = json_decode($scrFileDataEncode, true);
            // Set scenario name 
            if (!empty($scrFileData[$this->userInfo['scnrId']]['name'])) {
                $this->userInfo['scnrName'] = $scrFileData[$this->userInfo['scnrId']]['name'];
            }
            // Set scenario dataset
            if (!empty($scrFileData[$this->userInfo['scnrId']]['dataset'])) {
                $this->userInfo['scnrDataset'] = $scrFileData[$this->userInfo['scnrId']]['dataset'];
            }
            // Set scenario "Focus deprived population"
            if (!empty($scrFileData[$this->userInfo['scnrId']]['focus_area']['selNode'])) {
                $this->userInfo['depGroup'] = $scrFileData[$this->userInfo['scnrId']]['focus_area']['selNode'];
            }

//            App::import('Vendor', 'PHPExcel'); // Load PHPExcel
//            $this->objPHPExcel = new PHPExcel();
//            $this->objWorksheet = $this->objPHPExcel->getActiveSheet();
            $this->_alphabets(1);
            $this->_langData();
            $this->_relData();
            /* ----------------------Country name-------------------------------- */
            if (!empty($this->lang_data['nationalAreaIDNameList'][$this->userInfo['areaId']])) {
                $this->userInfo['areaName'] = $this->lang_data['nationalAreaIDNameList'][$this->userInfo['areaId']];
            }
        }
    }

    // Alphabates array
    public function _alphabets($level) {
        //  Alphabets Array
        if ($level == 1) {
            $this->alphabets = $alphabetsTmp = range('A', 'Z'); // Array containing latters from A to Z
            foreach ($alphabetsTmp as $alp) {
                array_push($this->alphabets, 'A' . $alp); // Array containing latters from A to Z and AA,AB,AC and so on..
            }
        }
    }

    //  Get language file 
    public function _langData() {
        $selLanguage = Configure::read('Config.language'); // default language
        $this->default_lang = $selLanguage;
        //  Get language file
        $lanMstrFile = $this->CommonMethod->getSelLangMasterFile($selLanguage);
        $lang_data_file = file_get_contents($lanMstrFile);
        $this->lang_data = json_decode($lang_data_file, true);
    }

    // Master relationship file
    public function _relData() {
        //  Get master relationship file
        $mstrRelFile = $this->CommonMethod->getRelationshipFile();
        $relData = file_get_contents($mstrRelFile);
        $this->rel_data = json_decode($relData, true);
    }

    // Make excel reader for editing excel file
    public function _createSheet($sheetNo = null) {
        $sheets = $this->objPHPExcel->getSheetCount();
        // IF sheets are less then passed args ,Create rest sheets
        if ($sheetNo > $sheets) {
            for ($i = $sheets; $i < $sheetNo; $i++) {
                $this->objPHPExcel->createSheet($i);
            }
        }
        $this->objPHPExcel->setActiveSheetIndex($sheetNo - 1);
        $this->objWorksheet = $this->objPHPExcel->getActiveSheet();
        $this->objWorksheet->getDefaultColumnDimension()->setWidth('4.43');
        $this->objWorksheet->getColumnDimension('A')->setWidth('2.5');
        $this->objWorksheet->setShowGridlines(false);
    }

    // Output file as a download
    public function _downloadFile($fileName) {
        $objWriter = PHPExcel_IOFactory::createWriter($this->objPHPExcel, $this->inputFileType);
        $objWriter->setIncludeCharts(TRUE);
        $this->objPHPExcel->setActiveSheetIndex(0);
        header('Content-type: application/vnd.ms-excel');
        header('Content-Disposition: attachment; filename=' . $fileName);
        // Write file to the browser
        $objWriter->save('php://output');
        die;
    }

    public function _getNameFromIndId($indId) {
        $returnName = '';
        // return name of indicator of given Id
        if ($this->lang_data == null) {
            $this->_langData();
        }
        // Check package name
        if (!empty($this->lang_data[SRV_DLV_MD_TP][$indId]['lnm'])) { // Check  for SDM
            $returnName = $this->lang_data[SRV_DLV_MD_TP][$indId]['lnm'];
        } else if (!empty($this->lang_data[PKG_TP][$indId]['lnm'])) {  // Check package name
            $returnName = $this->lang_data[PKG_TP][$indId]['lnm'];
        }
        return $returnName;
    }

    public function _fillRowColour($color, $rowCount = 1, $startRow = null) {
        $startCol = 0;
        $endCol = $this->otherInfo['colLimit'];
        if ($startRow == null) {
            $startRow = $this->sheetRow;
        }
        $cellBckStyle = array(
            'fill' => array(
                'type' => PHPExcel_Style_Fill::FILL_SOLID,
                'color' => array('rgb' => $color)
        ));
        for ($i = 0; $i < $rowCount; $i++) {
            $this->objWorksheet->getStyle($this->alphabets[$startCol] . $startRow . ':' . $this->alphabets[$endCol] . $startRow)->applyFromArray($cellBckStyle); // fill colour
            $startRow++;
        }
    }

    public function _fillDataColour($startCol, $row, $color, $rowCount = 1) {
        $endCol = $this->otherInfo['colLimit'];

        $cellBckStyle = array(
            'fill' => array(
                'type' => PHPExcel_Style_Fill::FILL_SOLID,
                'color' => array('rgb' => $color)
        ));
        for ($i = 0; $i < $rowCount; $i++) {
            //$this->objWorksheet->mergeCellsByColumnAndRow($startCol, $row, $endCol);
            $this->objWorksheet->getStyle($this->alphabets[$startCol] . $row . ':' . $this->alphabets[$endCol] . $row)->applyFromArray($cellBckStyle); // fill colour
            $row++;
        }
    }

    public function _putHeadingData($type, $value) {
        if (!empty($this->headingInfo[$type])) {
            $headingObj = $this->headingInfo[$type];
            $cellCol = $this->sheetCol + $headingObj[0];
            $cellRow = $this->sheetRow;
            $fontSize = $headingObj[1];
            $headingStyle = ['font' => ['size' => $fontSize, 'bold' => true]];
            $cellDim = $this->alphabets[$cellCol] . $cellRow;
            $this->objWorksheet->getCell($cellDim)->setValue($value); // Put value
            $this->objWorksheet->getStyle($cellDim)->applyFromArray($headingStyle); // apply style

            $this->sheetRow = $this->sheetRow + 1; // next time start from next row
        }
    }

    public function _targetPopulationSelection() {
        $deprivedGroup = GEOGRAPHY_TP; //default
        $depGroupText = '';
        $depriveGrpStrtRow = $this->sheetRow;
        // get selected deprived group from scnr_desc tree
        if (!empty($this->scr_desc_data[GEOGRAPHY_TP]['selectedGroup'])) {
            $deprivedGroup = $this->scr_desc_data[GEOGRAPHY_TP]['selectedGroup'];
        }

        if ($deprivedGroup == GEOGRAPHY_TP) {
            $depGroupText = __('Geography'); // Group label
            // Get all subnational list 
            $countryMasterFIle = $this->CommonMethod->getSelLangCountryMasterFile($this->userInfo['areaId'], $this->default_lang);
            $countryMasrFileEncd = file_get_contents($countryMasterFIle);
            $countryMasrFile = json_decode($countryMasrFileEncd, true);
            if (!empty($countryMasrFile['groupList']['list'][GEOGRAPHY_TP]['selOpts']['lvl2'])) {
                $cellRow = $depriveGrpStrtRow; // data start row

                $subntlList = $countryMasrFile['groupList']['list'][GEOGRAPHY_TP]['selOpts']['lvl2'];
                $selSubntls = [];
                // create a list of selected subnational
                if (!empty($this->scr_desc_data[GEOGRAPHY_TP][GEOGRAPHY_TP]['lvl2'])) {
                    $selSubntls = $this->scr_desc_data[GEOGRAPHY_TP][GEOGRAPHY_TP]['lvl2'];
                }
                $loopCounter = 1; // Start from extream left
                $cellCol = $this->sheetCol;
                $gap = $this->sheetCol;

                /* Start looping for each subnational and make a "x" if that subnational is selected in scenario
                 */
                foreach ($subntlList as $areaId => $subnational) {
                    $cellCol = $gap;
                    // First cell
                    if (in_array($areaId, $selSubntls)) { // if this subnational is selected then only put value
                        $cellDim = $this->alphabets[$cellCol] . $cellRow;
                        $this->_writeCellValue($cellDim, $subnational); // put text
                        if ($loopCounter == $this->rowLimit['deprivedGrp']) { // This row is complete, Move to next row
                            $cellRow = $depriveGrpStrtRow; // reset data start row
                            $gap = $cellCol + $this->otherInfo['depGrpNtlGap']; // change columns for next set of data
                            $loopCounter = 1; // Reset counter
                        } else {
                            $cellRow++; // Next row
                            $loopCounter++;
                        }
                    }
                }
            }
        } else if ($deprivedGroup == QNUINTILE_TP) {
            $depGroupText = __('Quintile'); // Group label
            // Get all wealth quantile list 
            if (!empty($this->rel_data['sgrpTypeValueRel'][QNUINTILE_TP])) {
                $quantileList = $this->rel_data['sgrpTypeValueRel'][QNUINTILE_TP];
                $selQuntiles = [];
                $cellRow = $depriveGrpStrtRow; // data start row
                if (!empty($this->scr_desc_data[GEOGRAPHY_TP][QNUINTILE_TP])) {
                    $selQuntiles = $this->scr_desc_data[GEOGRAPHY_TP][QNUINTILE_TP];
                }
                /* Start looping for each quantile and make a "x" if that quantile is selected in scenario
                 */
                foreach ($quantileList as $quantile) {
                    $quantileTxt = '';
                    if (!empty($this->lang_data['sgrpValueList'][$quantile])) {
                        $quantileTxt = $this->lang_data['sgrpValueList'][$quantile];
                    }
                    if (in_array($quantile, $selQuntiles)) { // if this quantile is selected
                        $cellDim = $this->alphabets[$this->sheetCol] . $cellRow;
                        $this->_writeCellValue($cellDim, $quantileTxt); // checked mark
                    }
                }
            }
        } else if ($deprivedGroup == LOCATION_TP) {
            $depGroupText = ''; // Group label
            if (!empty($this->lang_data['sgrpTypeList'][LOCATION_TP])) {
                $depGroupText = $this->lang_data['sgrpTypeList'][LOCATION_TP];
            }
            if (!empty($this->rel_data['sgrpTypeValueRel'][LOCATION_TP])) {
                $selLocations = [];
                $cellRow = $depriveGrpStrtRow; // data start row
                $locationList = $this->rel_data['sgrpTypeValueRel'][LOCATION_TP];
                // remove total from list
                $chk = array_search('TOTAL', $locationList);
                if ($chk) {
                    unset($locationList[$chk]);
                }
                if (!empty($this->scr_desc_data[GEOGRAPHY_TP][LOCATION_TP])) {
                    $selLocations = $this->scr_desc_data[GEOGRAPHY_TP][LOCATION_TP];
                }
                /* Start looping for each location and make a "x" if that location is selected in scenario
                 */
                foreach ($locationList as $location) {
                    $locTxt = '';  // textual value
                    if (!empty($this->lang_data['sgrpValueList'][$location])) {
                        $locTxt = $this->lang_data['sgrpValueList'][$location];
                    }
                    if (in_array($location, $selLocations)) { // if this quantile is selected
                        $cellDim = $this->alphabets[$this->sheetCol] . $cellRow;
                        $this->_writeCellValue($cellDim, $locTxt); // put text
                    }
                }
            }
        }
        //  Edit heading
        $this->sheetRow = $this->sheetRow - 1;
        $this->_putHeadingData('trgetPpl2', $depGroupText);
        $this->_fillDataColour(0, $depriveGrpStrtRow, $this->colorArray['h3'], $this->rowLimit['deprivedGrp']);

        $this->sheetRow = $this->sheetRow + $this->rowLimit['deprivedGrp'];
        $this->_fillRowColour($this->colorArray['h1']); // Blank row
        $this->sheetRow = $this->sheetRow + 1;
    }

    public function _epicauseSelection() {

        if (!empty($this->lang_data[EPIC_TP])) {
            // Write heading
            $this->_fillRowColour($this->colorArray['h2']); // 
            $this->_putHeadingData('epicause', __('Epic_priorities'));


            $selEpicauses = [];
            // get list of selected epicauses from scnr desc json
            if (!empty($this->scr_desc_data[EPIC_TP])) {
                $selEpicausesTmp = $this->scr_desc_data[EPIC_TP];
                foreach ($selEpicausesTmp as $epid => $epInfo) {
                    if (!empty($epInfo['chkd']) && ($epInfo['chkd'] == '1')) {
                        $selEpicauses[$epid] = '';
                    }
                }
            }
            // Lopp for each epicause group
            $cellRow = $this->sheetRow;
            $cellCol = $this->sheetCol;
            $maxElemnts = 0;

            foreach ($this->lang_data[EPIC_TP] as $epicGrpDetails) {
                // get group name 
                $groupName = '';
                $writeCounter = 0;
                if (!empty($epicGrpDetails['name'])) {
                    $groupName = $epicGrpDetails['name'];
                }
                //Add group name to sheet
                $cellDim = $this->alphabets[$cellCol] . $cellRow;
                $this->_writeCellValue($cellDim, $groupName); // Group name
                $this->objWorksheet->getStyle($cellDim)->applyFromArray(['font' => ['size' => 12, 'bold' => true]]); // Make bold
                // Move to next row
                $cellRow++;
                $epicList = [];
                if (!empty($epicGrpDetails['details'])) {
                    $epicList = $epicGrpDetails['details']; // List of all epicauses in group
                }
                // Loop for each epicause in this group
                foreach ($epicList as $epicId => $epicInfo) {
                    // Name of epic cause
                    $epicName = '';
                    if (!empty($epicInfo['lnm'])) {
                        $epicName = $epicInfo['lnm'];
                    }
                    if (isset($selEpicauses[$epicId])) { // If this epic group is selected put "x" sign
                        $cellDim = $this->alphabets[$cellCol] . $cellRow;
                        $this->_writeCellValue($cellDim, $epicName); // put value
                        $cellRow++; // Next row 
                        $writeCounter++;
                    }
                }
                // Now goto next group, and make a space of $otherInfo['epicGap'].
                $cellCol = $cellCol + $this->otherInfo['epicGap'];
                $cellRow = $this->sheetRow;
                if ($writeCounter > $maxElemnts) {
                    $maxElemnts = $writeCounter;
                }
            }
            $maxElemnts++;
            $this->_fillDataColour(0, $this->sheetRow, $this->colorArray['h3'], $maxElemnts);
            $this->sheetRow = $this->sheetRow + $maxElemnts;
            $this->_fillRowColour($this->colorArray['h1']); // Blank row
            $this->sheetRow = $this->sheetRow + 1;
        }
    }

    public function _intrvSelection() {
        $maxElemnts = 0;
        // Write heading
        $this->_fillRowColour($this->colorArray['h2']); // 
        $this->_putHeadingData('intrv', __('Intervention'));
        // Loop  through master relation json to get list of interventions, w.r.t their relations
        if (!empty($this->rel_data['sdmPckgIntrvRel'])) {

            // Get list of selected interventations
            $selIntrvns = [];
            if (!empty($this->scr_desc_data[INTRV_TP]['selList'])) {
                $selIntrvns = $this->scr_desc_data[INTRV_TP]['selList'];
            }

            $rowData = $this->rel_data['sdmPckgIntrvRel'];
            // Loop for sdm
            $cellCol = $this->sheetCol;
            $sdmDim = [];
            foreach ($rowData as $sdmId => $sdmInfo) {
                $writeCounter = 0;
                $sdmDim[$sdmId] = $cellCol;
                // Name of SDM
                $cellRow = $this->sheetRow;
                $sdmName = $this->_getNameFromIndGid($sdmId);
                // Write name of SDM to sheet
                $cellDim = $this->alphabets[$cellCol] . $cellRow;
                $this->_writeCellValue($cellDim, $sdmName);
                $this->objWorksheet->getStyle($cellDim)->applyFromArray(['font' => ['size' => 14, 'bold' => true]]); // Make bold
                //  move to next row and put package heading and interventions
                $cellRow++;
                $writeCounter++;
                // Loop for package
                foreach ($sdmInfo as $pckId => $pckInfo) {
                    // Name of Package
                    $pckName = $this->_getNameFromIndGid($pckId);
                    // Write name of package to sheet
                    $cellDimP = $this->alphabets[$cellCol] . $cellRow;
                    // Now move to next row and put interventions
                    $cellRow++;
                    $writeCounter++;
                    // Loop for interventions
                    $notEmpty = false;
                    foreach ($pckInfo as $intrvId) {
                        $intrvName = $this->_getNameFromIndGid($intrvId);
                        if (in_array($intrvId, $selIntrvns)) { // If this intervention  is selected 
                            $notEmpty = true;
                            $this->selPackages[$pckId]='';
                            
                            $cellDim = $this->alphabets[$cellCol] . $cellRow;
                            $this->_writeCellValue($cellDim, $intrvName); // put value
                            $cellRow++;
                            $writeCounter++;
                        }
                    }

                    if ($notEmpty) {
                        $this->_writeCellValue($cellDimP, $pckName);
                        //Make package name bold
                        $this->objWorksheet->getStyle($cellDimP)->applyFromArray(['font' => ['size' => 12, 'bold' => true]]);
                        // Now move to next row and put interventions
                    }
                }
                $cellCol = $cellCol + $this->secWidths[$sdmId]; // Make proper gap for next SDM

                if ($writeCounter > $maxElemnts) {
                    $maxElemnts = $writeCounter;    // maximum possible height of this section
                }
            }
            $maxElemnts++;
            foreach ($sdmDim as $id => $col) {
                if ($id == COMMUNITY_TP) {
                    $col--;  // start filling colour, from 0 index to cover blank space at left;
                }
                $this->_fillDataColour($col, $this->sheetRow, $this->colorArray[$id], $maxElemnts);
            }
            $this->sheetRow = $this->sheetRow + $maxElemnts;
            $this->_fillRowColour($this->colorArray['h1']); // Blank row
            $this->sheetRow = $this->sheetRow + 1;
        }
    }

    public function _severityBtnkPack() {
        $skippedBTNKGids = $this->skippedBTNKGids;
        // Write heading
        $this->_fillRowColour($this->colorArray['h2']); // 
        $this->_putHeadingData('btlnk', __('BTNK_CAS_STRTGY'));

        $this->_fillRowColour($this->colorArray['h3']); // 
        $this->_putHeadingData('btlnk2', __('SVRT_BTNK'));

        $this->sheetRow = $this->sheetRow - 1;  // Move one row up
        // Create colour ramp
        $scaleStartCol = $this->sheetCol + 19;
        $cellDim = $this->alphabets[$scaleStartCol] . $this->sheetRow;
        $this->_writeCellValue($cellDim, __('NO_SERVTY')); // colour ramp left text
        // create colourfull cells
        $colorRampStartCol = $scaleStartCol + 3;
        $colorArray = array('00B050', 'AAE41B', 'FFEF00', 'FFCD00', 'FFAB00', 'FF8800', 'FF6600', 'FF4400', 'FF2200', 'FF0000');
        for ($i = 0; $i < count($colorArray); $i++) {
            $cellDim = $this->alphabets[$colorRampStartCol] . $this->sheetRow;
            $this->_fillCellColour($cellDim, $colorArray[$i]);
            $colorRampStartCol++;
        }

        $cellDim = $this->alphabets[$colorRampStartCol + 1] . $this->sheetRow;
        $this->_writeCellValue($cellDim, __('SRV_BTLNCK')); // colour ramp right text
        $this->sheetRow = $this->sheetRow + 1;

        $this->_fillRowColour($this->colorArray['h3'], 3); // First three rows
        $this->sheetRow = $this->sheetRow + 1;
        // create column heading for bottneck
        $cellRow = $this->sheetRow;
        $btlnkDim = [];
        if (!empty($this->lang_data[BTLNK_TP])) {
            $cellCol = $this->sheetCol + $this->colIndexArray['btlnkSrvHead'];
            $gap = $this->otherInfo['btlnkSvrGap'];
            foreach ($this->lang_data[BTLNK_TP] as $gid => $btlnck) {
                if (in_array($gid, $skippedBTNKGids)) {
                    continue; // Goto next iteration
                }
                // Put bottleneck Name 
                $btlnckName = $this->_getNameFromIndGid($gid);
                $cellDim = $this->alphabets[$cellCol] . ($cellRow);
                $this->_writeCellValue($cellDim, $btlnckName);

                $btlnckHeadStyle = array(
                    'alignment' => ['horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER], // Central aligned 
                    'font' => ['size' => 11, 'bold' => true] // bold
                );

                $btlnkDim[$gid] = $cellCol;
                $this->objWorksheet->getStyle($cellDim)->applyFromArray($btlnckHeadStyle);

                $this->objWorksheet->mergeCellsByColumnAndRow($cellCol, $cellRow, ($cellCol + ($gap - 1)), $cellRow); // Merge 3 cell
                $cellCol = $cellCol + $gap;
            }
        }
        $lineRange = $this->alphabets[$this->sheetCol] . $this->sheetRow . ':' . $this->alphabets[$this->otherInfo['colLimit'] - 3] . $this->sheetRow;
        $this->_drawLine($lineRange);
        $this->sheetRow = $this->sheetRow + 1; // Gap
        // Create packages list, in a single column
        if (!empty($this->lang_data[PKG_TP]) && (!empty($this->scr_desc_data[BTLNK_TP][PKG_TP]))) {  // Data for bottleneck svr is avialable in scr json
            $packageData = $this->lang_data[PKG_TP];
            ksort($packageData); // Arrange in order
            $counter = 0;
            $cellCol = $this->sheetCol;
            $cellRow = $this->sheetRow;
            foreach ($packageData as $pckgGid => $pckgDetails) {
                // Now loop through the bottlenecks for this package in scnr json to get "svr data"
                if (isset($this->selPackages[$pckgGid])&&!empty($this->scr_desc_data[BTLNK_TP][PKG_TP][$pckgGid])) {
                    // Get package Name
                    $pckgName = $this->_getNameFromIndGid($pckgGid);
                    $pckgName = '';
                    if (!empty($pckgDetails['lnm'])) {
                        $pckgName = $pckgDetails['lnm'];
                    }
                    $cellDim2 = $this->alphabets[$cellCol] . ($cellRow);
                    $this->_writeCellValue($cellDim2, $pckgName);

                    $lineRange = $this->alphabets[$this->sheetCol] . $cellRow . ':' . $this->alphabets[$this->otherInfo['colLimit'] - 3] . $cellRow;
                    $this->_drawLine($lineRange);
                    $counter++;
                    $this->_fillRowColour($this->colorArray['h3'], 1, $cellRow); // 

                    foreach ($this->scr_desc_data[BTLNK_TP][PKG_TP][$pckgGid] as $btlnckGid => $svrData) {
                        if (isset($btlnkDim[$btlnckGid])) {
                            $cellCol2 = $btlnkDim[$btlnckGid] + 1; // move colour cell to right, for styling
                            $svrValue = 0;
                            if (!empty($svrData['svr'])) {
                                $svrValue = $svrData['svr'];
                            }
                            $colour = $this->getSvrtyColour($svrValue);
                            $cellBckStyle = array(
                                'fill' => array(
                                    'type' => PHPExcel_Style_Fill::FILL_SOLID,
                                    'color' => array('rgb' => $colour)
                                ),
                                'font' => array('color' => array('rgb' => $colour))
                            );


                            $cellDim = $this->alphabets[$cellCol2] . ($cellRow);
                            $this->objWorksheet->getStyle($cellDim)->applyFromArray($cellBckStyle); // fill colour
                            $this->_writeCellValue($cellDim, $svrValue);
                        }
                    }
                    $cellRow++;
                }
            }
            $counter++;
            // $this->_fillDataColour(0, $this->sheetRow, $this->colorArray['h3'], $counter);
            $this->sheetRow = $this->sheetRow + $counter;
            $this->_fillDataColour(0, $this->sheetRow - 1, $this->colorArray['h3']);
        }
    }

    public function _causeStrtgyPckg() {
        $logicalRelArray = $selPackages = [];
        $skippedBTNKGids = $this->skippedBTNKGids;

        // Write heading
        $this->_fillRowColour($this->colorArray['h2'], 2);
        $this->_putHeadingData('csp1', __('CAS_STRT_PCK'));
        // Create list of all selected packages
        $cellCol = $this->sheetCol + $this->colIndexArray['causStrgyPck'];
        $cellRow = $this->sheetRow;
        if (!empty($this->scr_desc_data[INTRV_TP]['selList'])) {
            $this->objWorksheet->getRowDimension($cellRow)->setRowHeight(100);
            foreach ($this->scr_desc_data[INTRV_TP]['selList'] as $intrvGid) {
                if (!empty($this->rel_data['intrvPckgRel'][$intrvGid])) {
                    $packGid = $this->rel_data['intrvPckgRel'][$intrvGid];
                    if (!isset($selPackages[$packGid])) { // Packge uniquness  
                        // Put package name
                        $pckHeadStyle = array(
                            'alignment' => ['horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER], // Central aligned 
                            'font' => ['size' => 11, 'bold' => true] // bold
                        );
                        $packgName = $this->_getNameFromIndGid($packGid);
                        $cellDim = $this->alphabets[$cellCol] . ($cellRow);
                        $this->_writeCellValue($cellDim, $packgName);
                        $this->objWorksheet->getStyle($cellDim)->getAlignment()->setTextRotation(90);
                        $this->objWorksheet->getStyle($cellDim)->applyFromArray($pckHeadStyle);
                        $selPackages[$packGid] = $cellCol;
                        $cellCol++;
                    }
                }
            }
        }
        //'csp_cause' => 3, 'csp_strtgy' =>
        $this->_putHeadingData('csp2', __('BTLNK_SINGLR'));
        $this->sheetRow = $this->sheetRow - 1;
        $this->_putHeadingData('csp3', __('CAUSE'));
        $this->sheetRow = $this->sheetRow - 1;
        $this->_putHeadingData('csp4', __('STRTG_SNGLR'));
        if (!empty($this->lang_data[BTLNK_TP])) {
            foreach ($this->lang_data[BTLNK_TP] as $btlnkGid => $bltnckInfo) {
                if (!in_array($btlnkGid, $skippedBTNKGids)) { // Only allowed bottleneck
                    if (!empty($this->rel_data['btlnkCauseStrtgList'][$btlnkGid])) {
                        foreach ($this->rel_data['btlnkCauseStrtgList'][$btlnkGid] as $causeGid => $strtgyInfo) {
                            foreach ($strtgyInfo as $strtgyGid) {

                                foreach ($selPackages as $pckGid => $col) {
                                    $isSelected = false;
                                    if (!empty($this->scr_desc_data[STRTG_TP][PKG_TP][$pckGid][$btlnkGid][$causeGid][$strtgyGid]['chkd'])) { // first check at package level
                                        if ($this->scr_desc_data[STRTG_TP][PKG_TP][$pckGid][$btlnkGid][$causeGid][$strtgyGid]['chkd'] == '1') {
                                            $isSelected = true;
                                        }
                                    } else {// Get data realeted to  parent SDM
                                        if (!empty($this->rel_data['pckgSdmRel'][$pckGid])) {
                                            $parentSdmGid = $this->rel_data['pckgSdmRel'][$pckGid];
                                            if (!empty($this->scr_desc_data[STRTG_TP][SRV_DLV_MD_TP][$parentSdmGid][$btlnkGid][$causeGid][$strtgyGid]['chkd'])) {
                                                if ($this->scr_desc_data[STRTG_TP][SRV_DLV_MD_TP][$parentSdmGid][$btlnkGid][$causeGid][$strtgyGid]['chkd'] == '1') {
                                                    $isSelected = true;
                                                }
                                            }
                                        }
                                    }
                                    if ($isSelected) {
                                        $logicalRelArray[$btlnkGid][$causeGid][$strtgyGid][$pckGid] = $col;
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }


        $causeDataStartRow = $this->sheetRow;
        foreach ($logicalRelArray as $btlnckGid => $causeInfo) {
            $counter = 0;
            $btnckName = $this->_getNameFromIndGid($btlnckGid);
            // Put bottleneck Heading
            $cellRow = $this->sheetRow;
            $cellCol = $this->sheetCol;
            $cellDim = $this->alphabets[$cellCol] . ($cellRow);
            $this->_writeCellValue($cellDim, $btnckName);

            // Put causes and startegy for this bottlenck

            foreach ($causeInfo as $causeGid => $strtgyInfo) {
                // Cause name
                $cellCol = $this->sheetCol + $this->colIndexArray['csp_cause'];
                $causeName = $this->_getNameFromIndGid($causeGid);
                $cellDim = $this->alphabets[$cellCol] . ($cellRow);
                $this->_writeCellValue($cellDim, $causeName);
                foreach ($strtgyInfo as $strtgGId => $pckInfo) {
                    // Startegy name
                    $cellCol = $this->sheetCol + $this->colIndexArray['csp_strtgy'];
                    $startgyName = $this->_getNameFromIndGid($strtgGId);
                    $cellDim = $this->alphabets[$cellCol] . ($cellRow);
                    $this->_writeCellValue($cellDim, $startgyName);
                    foreach ($pckInfo as $pckGId => $locIndex) {
                        $cellColTmp = $locIndex;
                        $cellStyle = array(
                            'alignment' => ['horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER] // Central aligned 
                        );
                        $cellDim = $this->alphabets[$cellColTmp] . ($cellRow);
                        $this->_writeCellValue($cellDim, 'x');
                        $this->objWorksheet->getStyle($cellDim)->applyFromArray($cellStyle);
                    }
                }
                $counter++;
                $cellRow++;
            }
            $this->sheetRow = $this->sheetRow + ($counter - 1); // Incremented inside loop,So dcrease it.
            $this->_drawFullLine();
            $this->sheetRow = $this->sheetRow + 1;
        }
        $this->sheetRow = $this->sheetRow + 1;

        $this->_fillRowColour($this->colorArray['h3'], ($this->sheetRow - $causeDataStartRow), $causeDataStartRow);
    }

    public function _enablingEnv() {
        $deprvdGroup = $this->userInfo['depGroup'];
        $eeInfo = [];
        $sumEe = 0;
        $chartHeight = $this->otherInfo['chartGrp2'];
        // Write heading
        $this->_fillRowColour($this->colorArray['h1']);
        $this->sheetRow = $this->sheetRow + 1;
        $this->_fillRowColour($this->colorArray['h2']);
        $this->_putHeadingData('ee', __('EE_SCORE'));
        $this->sheetRow = $this->sheetRow + 1;


        if (!empty($this->scr_desc_data[EE_DATA][$deprvdGroup])) {
            foreach ($this->scr_desc_data[EE_DATA][$deprvdGroup] as $subAreaId => $eeData) {
                $ttl = 0;
                $countEE = count($eeData);
                foreach ($eeData as $eeGid => $eeVal) {
                    if (!isset($eeInfo[$eeGid])) {
                        $eeInfo[$eeGid] = $eeVal;
                    } else {
                        $eeInfo[$eeGid] = $eeInfo[$eeGid] + $eeVal;
                    }
                    $ttl = $ttl + $eeVal;
                }
                $sumEeTmp = $ttl / $countEE;
                $sumEe = $sumEe + $sumEeTmp;
            }
        }
        // Average 
        $totalsubGrp = count($this->scr_desc_data[EE_DATA][$deprvdGroup]);

        $eeScore = ($sumEe / $totalsubGrp) * 100;

        $cellCol = $this->sheetCol;
        $cellRow = $this->sheetRow;

        $cellDim = $this->alphabets[$cellCol] . $cellRow;
        $this->_writeCellValue($cellDim, '');
        $cellDim = $this->alphabets[$cellCol + 1] . $cellRow;

        $cellRow++;

        // Start putting data
        $cellDim = $this->alphabets[$cellCol] . $cellRow;
        $this->_writeCellValue($cellDim, __('EE_SCOR'));
        $cellDim = $this->alphabets[$cellCol + 1] . $cellRow;
        $this->_writeCellValue($cellDim, $eeScore); // Group value
        $cellRow++;
        foreach ($eeInfo as $ee => $val) {
            $avgVal = ($val / $totalsubGrp) * 100;

            $eeName = '';
            if (!empty($this->lang_data['indicatorList'][$ee])) {
                $eeName = $this->lang_data['indicatorList'][$ee];
            }
            $cellDim = $this->alphabets[$cellCol] . $cellRow;
            $this->_writeCellValue($cellDim, $eeName); // Group name
            $cellDim = $this->alphabets[$cellCol + 1] . $cellRow;
            $this->_writeCellValue($cellDim, $avgVal); //  value
            $cellRow++;
        }
        $this->_makeChart2($cellCol, ($this->sheetRow), (count($eeInfo) + 1), $this->otherInfo['colLimit'] - 1, $chartHeight);
        $this->sheetRow = $this->sheetRow + ($chartHeight + 1);
    }

    public function _epicChart() {
        /* Create chart for epiccauses, packages and bottlenecks
         */
        // Basic structure  DON'T delete
        /* $objWorksheet->fromArray(
          array(
          array('', 2010, 2011, 2012),
          array('Q1', 12, 15, 21),
          array('Q2', 56, 73, 86),
          array('Q3', 52, 61, 69),
          array('Q4', 30, 32, 0),
          )
          ); */
        $chartTypes = array(__('SCNR_INDCT_OPERATIONAL_TXT'), __('EXS_DTHS'), __('SCNR_INDCT_DEATHS_TXT'));
        // Write heading
        $this->_fillRowColour($this->colorArray['h1']);
        $this->sheetRow = $this->sheetRow + 1;

        $this->_fillRowColour($this->colorArray['h2']);
        $this->_putHeadingData('impctDth', __('IMPCT_DTH_AVRT'));

        $this->_fillRowColour($this->colorArray['h3']);
        $this->objWorksheet->mergeCellsByColumnAndRow(0, $this->sheetRow, $this->otherInfo['colLimit'], $this->sheetRow); // Merge cell
        $this->objWorksheet->getStyle($this->alphabets[0] . $this->sheetRow)->applyFromArray(['alignment' => ['horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER]]);
        $this->_putHeadingData('epicChrt1', __('EPIC_FULL_TXT'));

        $this->_fillRowColour($this->colorArray['h3']);
        $this->_putHeadingData('epicChrt11', __('UNDER_FIVE_MORT'));
        $this->sheetRow = $this->sheetRow - 1;
        $this->_putHeadingData('epicChrt12', __('NEO_MORT'));
        $this->sheetRow = $this->sheetRow - 1;
        $this->_putHeadingData('epicChrt13', __('STUNTING'));
        $this->sheetRow = $this->sheetRow + 1;


        $this->_scrChartsData($this->userInfo['userId'], $this->userInfo['areaId'], $this->userInfo['scnrId']);

        //--------------------- under 5 mortality-epicause
        $chartHeight = $this->otherInfo['chartGrp1'];
        $chartCombinedU5mr = [];
        $typeArray = [];

        if (!empty($this->epi_u5mr)) {
            $dataObj = $this->epi_u5mr;
            if (!empty($dataObj[SCNR_INDCT_AEG][EPIC_TP][UNIT_NUMBER_GID]['seriesData'])) {
                $typeArray[SCNR_INDCT_AEG] = $dataObj[SCNR_INDCT_AEG][EPIC_TP][UNIT_NUMBER_GID]['seriesData'];
            }
            if (!empty($dataObj[SCNR_INDCT_OPERATIONAL][EPIC_TP][UNIT_NUMBER_GID]['seriesData'])) {
                $typeArray[SCNR_INDCT_OPERATIONAL] = $dataObj[SCNR_INDCT_OPERATIONAL][EPIC_TP][UNIT_NUMBER_GID]['seriesData'];
            }
            if (!empty($dataObj[SCNR_INDCT_DEATHS][EPIC_TP][UNIT_NUMBER_GID]['seriesData'])) {
                $typeArray[SCNR_INDCT_DEATHS] = $dataObj[SCNR_INDCT_DEATHS][EPIC_TP][UNIT_NUMBER_GID]['seriesData'];
            }

            foreach ($typeArray as $type => $tData) {
                foreach ($tData as $seriesData) {
                    if (isset($seriesData['data'][0]['y']) && (!empty($seriesData['name']))) {
                        $epicGid = $seriesData['name'];
                        $epicvalue = $seriesData['data'][0]['y'];
                        $chartCombinedU5mr[$epicGid][$type] = $epicvalue;
                    }
                }
            }

            $cellCol = $this->sheetCol + $this->chartDim['epic11'][0];

            // Initial row
            $this->_writeCellValue($this->alphabets[$cellCol] . $this->sheetRow, ''); //  
            $this->_writeCellValue($this->alphabets[$cellCol + 1] . $this->sheetRow, $chartTypes[0]); //  operational
            $this->_writeCellValue($this->alphabets[$cellCol + 2] . $this->sheetRow, $chartTypes[1]); //  equity
            $this->_writeCellValue($this->alphabets[$cellCol + 3] . $this->sheetRow, $chartTypes[2]); //  death avertable
            //----------
            $cellRow = $this->sheetRow + 1;
            foreach ($chartCombinedU5mr as $epicGid => $colSeries) {
                $epicName = $epicGid;
                if (!empty($this->lang_data['indicatorList'][$epicGid])) {
                    $epicName = $this->lang_data['indicatorList'][$epicGid];
                }else{
                    $epicName =__($epicGid);
                }
                $opFrntVal = $eqFrntrVal = $dthAvrtVal = 0;
                if (!empty($colSeries[SCNR_INDCT_OPERATIONAL]) && ($colSeries[SCNR_INDCT_OPERATIONAL] > 0)) {
                    $opFrntVal = round($colSeries[SCNR_INDCT_OPERATIONAL], 3);
                }
                if (!empty($colSeries[SCNR_INDCT_AEG]) && ($colSeries[SCNR_INDCT_AEG] > 0)) {
                    $eqFrntrVal = round($colSeries[SCNR_INDCT_AEG], 3);
                }
                if (!empty($colSeries[SCNR_INDCT_DEATHS]) && ($colSeries[SCNR_INDCT_DEATHS] > 0)) {
                    $dthAvrtVal = round($colSeries[SCNR_INDCT_DEATHS], 3);
                }
                $this->_writeCellValue($this->alphabets[$cellCol] . $cellRow, $epicName); //  name
                $this->_writeCellValue($this->alphabets[$cellCol + 1] . $cellRow, $opFrntVal); //  operational
                $this->_writeCellValue($this->alphabets[$cellCol + 2] . $cellRow, $eqFrntrVal); //  equity
                $this->_writeCellValue($this->alphabets[$cellCol + 3] . $cellRow, $dthAvrtVal); //  death avertable

                $cellRow++; //next row
            }

            $this->_makeChart($cellCol, ($this->sheetRow), count($chartCombinedU5mr), $this->chartDim['epic11'][1], $chartHeight, 1);
        }
        //--------------------- Neonatal mortality -epicause
        $chartCombinedNeonatal = [];
        $typeArray = [];
        if (!empty($this->epi_nn)) {
            $dataObj = $this->epi_nn;
            if (!empty($dataObj[SCNR_INDCT_AEG][EPIC_TP][UNIT_NUMBER_GID]['seriesData'])) {
                $typeArray[SCNR_INDCT_AEG] = $dataObj[SCNR_INDCT_AEG][EPIC_TP][UNIT_NUMBER_GID]['seriesData'];
            }
            if (!empty($dataObj[SCNR_INDCT_OPERATIONAL][EPIC_TP][UNIT_NUMBER_GID]['seriesData'])) {
                $typeArray[SCNR_INDCT_OPERATIONAL] = $dataObj[SCNR_INDCT_OPERATIONAL][EPIC_TP][UNIT_NUMBER_GID]['seriesData'];
            }
            if (!empty($dataObj[SCNR_INDCT_DEATHS][EPIC_TP][UNIT_NUMBER_GID]['seriesData'])) {
                $typeArray[SCNR_INDCT_DEATHS] = $dataObj[SCNR_INDCT_DEATHS][EPIC_TP][UNIT_NUMBER_GID]['seriesData'];
            }

            foreach ($typeArray as $type => $tData) {
                foreach ($tData as $seriesData) {
                    if (isset($seriesData['data'][0]['y']) && (!empty($seriesData['name']))) {
                        $epicGid = $seriesData['name'];
                        $epicvalue = $seriesData['data'][0]['y'];
                        $chartCombinedNeonatal[$epicGid][$type] = $epicvalue;
                    }
                }
            }

            $cellCol = $this->sheetCol + $this->chartDim['epic12'][0];

            // Initial row
            $this->_writeCellValue($this->alphabets[$cellCol] . $this->sheetRow, ''); //  
            $this->_writeCellValue($this->alphabets[$cellCol + 1] . $this->sheetRow, $chartTypes[0]); //  operational
            $this->_writeCellValue($this->alphabets[$cellCol + 2] . $this->sheetRow, $chartTypes[1]); //  equity
            $this->_writeCellValue($this->alphabets[$cellCol + 3] . $this->sheetRow, $chartTypes[2]); //  death avertable
            $cellRow = $this->sheetRow + 1;
            foreach ($chartCombinedNeonatal as $epicGid => $colSeries) {
                $epicName = $epicGid;
                if (!empty($this->lang_data['indicatorList'][$epicGid])) {
                    $epicName = $this->lang_data['indicatorList'][$epicGid];
                }else{
                    $epicName =__($epicGid);
                }
                $opFrntVal = $eqFrntrVal = $dthAvrtVal = 0;
                if (!empty($colSeries[SCNR_INDCT_OPERATIONAL]) && ($colSeries[SCNR_INDCT_OPERATIONAL] > 0)) {
                    $opFrntVal = round($colSeries[SCNR_INDCT_OPERATIONAL], 3);
                }
                if (!empty($colSeries[SCNR_INDCT_AEG]) && ($colSeries[SCNR_INDCT_AEG] > 0)) {
                    $eqFrntrVal = round($colSeries[SCNR_INDCT_AEG], 3);
                }
                if (!empty($colSeries[SCNR_INDCT_DEATHS]) && ($colSeries[SCNR_INDCT_DEATHS] > 0)) {
                    $dthAvrtVal = round($colSeries[SCNR_INDCT_DEATHS], 3);
                }
                $this->_writeCellValue($this->alphabets[$cellCol] . $cellRow, $epicName); //  name
                $this->_writeCellValue($this->alphabets[$cellCol + 1] . $cellRow, $opFrntVal); //  operational
                $this->_writeCellValue($this->alphabets[$cellCol + 2] . $cellRow, $eqFrntrVal); //  equity
                $this->_writeCellValue($this->alphabets[$cellCol + 3] . $cellRow, $dthAvrtVal); //  death avertable

                $cellRow++; //next row
            }
            $this->_makeChart($cellCol, ($this->sheetRow), count($chartCombinedNeonatal), $this->chartDim['epic12'][1], $chartHeight, 1);
        }
        //--------------------- Stunting -epicause
        $chartCombinedStunting = [];
        $typeArray = [];
        if (!empty($this->epi_stunt)) {
            $dataObj = $this->epi_stunt;
            if (!empty($dataObj[SCNR_INDCT_AEG][EPIC_TP][UNIT_NUMBER_GID]['seriesData'])) {
                $typeArray[SCNR_INDCT_AEG] = $dataObj[SCNR_INDCT_AEG][EPIC_TP][UNIT_NUMBER_GID]['seriesData'];
            }
            if (!empty($dataObj[SCNR_INDCT_OPERATIONAL][EPIC_TP][UNIT_NUMBER_GID]['seriesData'])) {
                $typeArray[SCNR_INDCT_OPERATIONAL] = $dataObj[SCNR_INDCT_OPERATIONAL][EPIC_TP][UNIT_NUMBER_GID]['seriesData'];
            }
            if (!empty($dataObj[SCNR_INDCT_DEATHS][EPIC_TP][UNIT_NUMBER_GID]['seriesData'])) {
                $typeArray[SCNR_INDCT_DEATHS] = $dataObj[SCNR_INDCT_DEATHS][EPIC_TP][UNIT_NUMBER_GID]['seriesData'];
            }
            foreach ($typeArray as $type => $tData) {
                foreach ($tData as $seriesData) {
                    if (isset($seriesData['data'][0]['y']) && (!empty($seriesData['name']))) {
                        $epicGid = $seriesData['name'];
                        $epicvalue = $seriesData['data'][0]['y'];
                        $chartCombinedStunting[$epicGid][$type] = $epicvalue;
                    }
                }
            }

            $cellCol = $this->sheetCol + $this->chartDim['epic13'][0];

            // Initial row
            $this->_writeCellValue($this->alphabets[$cellCol] . $this->sheetRow, ''); //  
            $this->_writeCellValue($this->alphabets[$cellCol + 1] . $this->sheetRow, $chartTypes[0]); //  operational
            $this->_writeCellValue($this->alphabets[$cellCol + 2] . $this->sheetRow, $chartTypes[1]); //  equity
            $this->_writeCellValue($this->alphabets[$cellCol + 3] . $this->sheetRow, $chartTypes[2]); //  death avertable
            $cellRow = $this->sheetRow + 1;
            foreach ($chartCombinedStunting as $epicGid => $colSeries) {
                $epicName = $epicGid;
                if (!empty($this->lang_data['indicatorList'][$epicGid])) {
                    $epicName = $this->lang_data['indicatorList'][$epicGid];
                }else{
                    $epicName =__($epicGid);
                }
                $opFrntVal = $eqFrntrVal = $dthAvrtVal = 0;
                if (!empty($colSeries[SCNR_INDCT_OPERATIONAL]) && ($colSeries[SCNR_INDCT_DEATHS] > 0)) {
                    $opFrntVal = round($colSeries[SCNR_INDCT_OPERATIONAL], 3);
                }
                if (!empty($colSeries[SCNR_INDCT_AEG]) && ($colSeries[SCNR_INDCT_AEG] > 0)) {
                    $eqFrntrVal = round($colSeries[SCNR_INDCT_AEG], 3);
                }
                if (!empty($colSeries[SCNR_INDCT_DEATHS]) && ($colSeries[SCNR_INDCT_DEATHS] > 0)) {
                    $dthAvrtVal = round($colSeries[SCNR_INDCT_DEATHS], 3);
                }
                $this->_writeCellValue($this->alphabets[$cellCol] . $cellRow, $epicName); //  name
                $this->_writeCellValue($this->alphabets[$cellCol + 1] . $cellRow, $opFrntVal); //  operational
                $this->_writeCellValue($this->alphabets[$cellCol + 2] . $cellRow, $eqFrntrVal); //  equity
                $this->_writeCellValue($this->alphabets[$cellCol + 3] . $cellRow, $dthAvrtVal); //  death avertable

                $cellRow++; //next row
            }
            $this->_makeChart($cellCol, ($this->sheetRow), count($chartCombinedStunting), $this->chartDim['epic13'][1], $chartHeight, 1);
        }

        //Start of second group of charts
        $this->sheetRow = $this->sheetRow + $chartHeight + 1; // Sepration between two groups of chart
        $this->_fillRowColour($this->colorArray['h3']);
        $col1 = ($this->sheetCol + $this->headingInfo['epicChrt22'][0] - 1);
        $this->objWorksheet->mergeCellsByColumnAndRow(0, $this->sheetRow, $col1, ($this->sheetRow + 1)); // Merge cell
        $this->objWorksheet->mergeCellsByColumnAndRow($col1 + 1, $this->sheetRow, $this->otherInfo['colLimit'], ($this->sheetRow + 1)); // Merge cell
        $this->objWorksheet->getStyleByColumnAndRow(0, $this->sheetRow, $this->otherInfo['colLimit'])->applyFromArray(['alignment' => ['horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER]]);
        $this->_putHeadingData('epicChrt21', __('PCK_SINGLR'));
        $this->sheetRow = $this->sheetRow - 1;
        $this->_putHeadingData('epicChrt22', __('BTLNK_SINGLR'));
        $this->sheetRow = $this->sheetRow + 2;
        //--------------------- U5MR -Packages
        $chartCombinedU5mrPackges = [];
        $typeArray = [];
        
        // U5mr-Packages
        if (!empty($this->epi_u5mr)) {
            $dataObj = $this->epi_u5mr;
            if (!empty($dataObj[SCNR_INDCT_AEG][PKG_TP][UNIT_NUMBER_GID]['seriesData'])) {
                $typeArray[SCNR_INDCT_AEG] = $dataObj[SCNR_INDCT_AEG][PKG_TP][UNIT_NUMBER_GID]['seriesData'];
            }
            if (!empty($dataObj[SCNR_INDCT_OPERATIONAL][PKG_TP][UNIT_NUMBER_GID]['seriesData'])) {
                $typeArray[SCNR_INDCT_OPERATIONAL] = $dataObj[SCNR_INDCT_OPERATIONAL][PKG_TP][UNIT_NUMBER_GID]['seriesData'];
            }
            if (!empty($dataObj[SCNR_INDCT_DEATHS][PKG_TP][UNIT_NUMBER_GID]['seriesData'])) {
                $typeArray[SCNR_INDCT_DEATHS] = $dataObj[SCNR_INDCT_DEATHS][PKG_TP][UNIT_NUMBER_GID]['seriesData'];
            }

            foreach ($typeArray as $type => $tData) {
                foreach ($tData as $seriesData) {
                    if (isset($seriesData['data'][0]['y']) && (!empty($seriesData['name']))) {
                        $epicGid = $seriesData['name'];
                        $epicvalue = $seriesData['data'][0]['y'];
                        $chartCombinedU5mrPackges[$epicGid][$type] = $epicvalue;
                    }
                }
            }

            krsort($chartCombinedU5mrPackges);
            $cellCol = $this->sheetCol + $this->chartDim['epic21'][0];
            // Initial row
            $this->_writeCellValue($this->alphabets[$cellCol] . $this->sheetRow, ''); //  
            $this->_writeCellValue($this->alphabets[$cellCol + 1] . $this->sheetRow, $chartTypes[0]); //  operational
            $this->_writeCellValue($this->alphabets[$cellCol + 2] . $this->sheetRow, $chartTypes[1]); //  equity
            $this->_writeCellValue($this->alphabets[$cellCol + 3] . $this->sheetRow, $chartTypes[2]); //  death avertable
            $cellRow = $this->sheetRow + 1;
            foreach ($chartCombinedU5mrPackges as $epicGid => $colSeries) {
                $epicName = $epicGid;
                if (!empty($this->lang_data[PKG_TP][$epicGid]['snm'])) {
                    $epicName = $this->lang_data[PKG_TP][$epicGid]['snm'];
                }else{
                    $epicName =__($epicGid);
                }

                $opFrntVal = $eqFrntrVal = $dthAvrtVal = 0;
                if (!empty($colSeries[SCNR_INDCT_OPERATIONAL]) && ($colSeries[SCNR_INDCT_OPERATIONAL] > 0)) {
                    $opFrntVal = round($colSeries[SCNR_INDCT_OPERATIONAL], 3);
                }
                if (!empty($colSeries[SCNR_INDCT_AEG]) && ($colSeries[SCNR_INDCT_AEG] > 0)) {
                    $eqFrntrVal = round($colSeries[SCNR_INDCT_AEG], 3);
                }
                if (!empty($colSeries[SCNR_INDCT_DEATHS]) && ($colSeries[SCNR_INDCT_DEATHS] > 0)) {
                    $dthAvrtVal = round($colSeries[SCNR_INDCT_DEATHS], 3);
                }
                $this->_writeCellValue($this->alphabets[$cellCol] . $cellRow, $epicName); //  name
                $this->_writeCellValue($this->alphabets[$cellCol + 1] . $cellRow, $opFrntVal); //  operational
                $this->_writeCellValue($this->alphabets[$cellCol + 2] . $cellRow, $eqFrntrVal); //  equity
                $this->_writeCellValue($this->alphabets[$cellCol + 3] . $cellRow, $dthAvrtVal); //  death avertable

                $cellRow++; //next row
            }
            $this->_makeChart($cellCol, ($this->sheetRow), count($chartCombinedU5mrPackges), $this->chartDim['epic21'][1], $chartHeight, 2);
        }
        //--------------------- U5MR -Bottlenecks
        $chartCombinedU5mrBtlnck = [];
        $typeArray = [];
        if (!empty($this->epi_u5mr)) {
            $dataObj = $this->epi_u5mr;
            if (!empty($dataObj[SCNR_INDCT_AEG][BTLNK_TP][UNIT_NUMBER_GID]['seriesData'])) {
                $typeArray[SCNR_INDCT_AEG] = $dataObj[SCNR_INDCT_AEG][BTLNK_TP][UNIT_NUMBER_GID]['seriesData'];
            }
            if (!empty($dataObj[SCNR_INDCT_OPERATIONAL][BTLNK_TP][UNIT_NUMBER_GID]['seriesData'])) {
                $typeArray[SCNR_INDCT_OPERATIONAL] = $dataObj[SCNR_INDCT_OPERATIONAL][BTLNK_TP][UNIT_NUMBER_GID]['seriesData'];
            }
            if (!empty($dataObj[SCNR_INDCT_DEATHS][BTLNK_TP][UNIT_NUMBER_GID]['seriesData'])) {
                $typeArray[SCNR_INDCT_DEATHS] = $dataObj[SCNR_INDCT_DEATHS][BTLNK_TP][UNIT_NUMBER_GID]['seriesData'];
            }

            foreach ($typeArray as $type => $tData) {
                foreach ($tData as $seriesData) {
                    if (isset($seriesData['data'][0]['y']) && (!empty($seriesData['name']))) {
                        $epicGid = $seriesData['name'];
                        $epicvalue = $seriesData['data'][0]['y'];
                        $chartCombinedU5mrBtlnck[$epicGid][$type] = $epicvalue;
                    }
                }
            }

            //krsort($chartCombinedU5mrBtlnck);
            $cellCol = $this->sheetCol + $this->chartDim['epic22'][0];
            // Initial row
            $this->_writeCellValue($this->alphabets[$cellCol] . $this->sheetRow, ''); //  
            $this->_writeCellValue($this->alphabets[$cellCol + 1] . $this->sheetRow, $chartTypes[0]); //  operational
            $this->_writeCellValue($this->alphabets[$cellCol + 2] . $this->sheetRow, $chartTypes[1]); //  equity
            $this->_writeCellValue($this->alphabets[$cellCol + 3] . $this->sheetRow, $chartTypes[2]); //  death avertable
            $cellRow = $this->sheetRow + 1;
            foreach ($chartCombinedU5mrBtlnck as $epicGid => $colSeries) {
                $epicName = $epicGid;
                if (!empty($this->lang_data[BTLNK_TP][$epicGid]['snm'])) {
                    $epicName = $this->lang_data[BTLNK_TP][$epicGid]['snm'];
                }else{
                    $epicName =__($epicGid);
                }

                $opFrntVal = $eqFrntrVal = $dthAvrtVal = 0;
                if (!empty($colSeries[SCNR_INDCT_OPERATIONAL]) && ($colSeries[SCNR_INDCT_OPERATIONAL] > 0)) {
                    $opFrntVal = round($colSeries[SCNR_INDCT_OPERATIONAL], 3);
                }
                if (!empty($colSeries[SCNR_INDCT_AEG]) && ($colSeries[SCNR_INDCT_AEG] > 0)) {
                    $eqFrntrVal = round($colSeries[SCNR_INDCT_AEG], 3);
                }
                if (!empty($colSeries[SCNR_INDCT_DEATHS]) && ($colSeries[SCNR_INDCT_DEATHS] > 0)) {
                    $dthAvrtVal = round($colSeries[SCNR_INDCT_DEATHS], 3);
                }
                $this->_writeCellValue($this->alphabets[$cellCol] . $cellRow, $epicName); //  name
                //            /echo $cellCol.' '.$cellRow.' '.$epicName;die;
                $this->_writeCellValue($this->alphabets[$cellCol + 1] . $cellRow, $opFrntVal); //  operational
                $this->_writeCellValue($this->alphabets[$cellCol + 2] . $cellRow, $eqFrntrVal); //  equity
                $this->_writeCellValue($this->alphabets[$cellCol + 3] . $cellRow, $dthAvrtVal); //  death avertable

                $cellRow++; //next row
            }
            $this->_makeChart($cellCol, ($this->sheetRow), count($chartCombinedU5mrBtlnck), $this->chartDim['epic22'][1], $chartHeight, 2);
        }
        $this->sheetRow = $this->sheetRow + ($chartHeight + 1);
    }

    public function _costChart() {
        // Write heading
        $this->_fillRowColour($this->colorArray['h1']);
        $this->sheetRow = $this->sheetRow + 1;

        $this->_fillRowColour($this->colorArray['h2']);
        $this->_putHeadingData('cost1', __('IMPT_COST'));

        $this->_fillRowColour($this->colorArray['h3']);
        $this->_putHeadingData('cost21', __('STRTGY'));
        $this->sheetRow = $this->sheetRow - 1;

        $this->_fillRowColour($this->colorArray['h3']);
        $this->_putHeadingData('cost22', __('PER_MILLION_LIFE_SAVE'));
        $this->sheetRow = $this->sheetRow - 1;

        $this->_fillRowColour($this->colorArray['h3']);
        $this->_putHeadingData('cost23', __('ECO_GDP_CAPITA'));
        $this->sheetRow = $this->sheetRow + 1;
        //  Get scenario costing file
        $scrPath = $this->CommonMethod->getUserCountryScenarioFolderPath($this->userInfo['userId'], $this->userInfo['areaId'], $this->userInfo['scnrId']);
        $file = $scrPath . 'costing_output_data.json';
        if (file_exists($file)) {
            $fileData = file_get_contents($file);
            $costingData = json_decode($fileData, true);


            // Cost by startegy
            $chartHeight = $this->otherInfo['chartGrp1'];
            if (!empty($costingData[STRTG_TP]['seriesData'])) {
                $cellCol = $this->sheetCol + $this->chartDim['cost1'][0];
                $cellRow = $this->sheetRow;

                //Initial data
                $this->_writeCellValue($this->alphabets[$cellCol] . $cellRow, '');
                $this->_writeCellValue($this->alphabets[$cellCol + 1] . $cellRow, '');
                $cellRow++;

                $costByStrgy = $costingData[STRTG_TP]['seriesData'];
                $strtCount = 0;
                foreach ($costByStrgy as $Gid => $colVal) {
                    $indName = '';
                    // Two type of indicators .1>STRATEGY 2>NDC

                    if (!empty($this->lang_data[STRTG_TP][$Gid]['snm'])) {
                        $indName = $this->lang_data[STRTG_TP][$Gid]['snm'];
                    } else if (!empty($this->lang_data['NDC'][$Gid]['lnm'])) {
                        $indName = $this->lang_data['NDC'][$Gid]['lnm'];
                    }
                    if (!empty($colVal)) { // Put only if value exists
                        $costVal = round($colVal, 3);
                        $this->_writeCellValue($this->alphabets[$cellCol] . $cellRow, $indName); //  name
                        $this->_writeCellValue($this->alphabets[$cellCol + 1] . $cellRow, $costVal); // value
                        $cellRow++;
                        $strtCount++;
                    }
                }
                $this->_makeChart($cellCol, ($this->sheetRow), $strtCount, $this->chartDim['cost1'][1], $chartHeight, 2);
            }

            // Cost-Lives saved per 1 million US$
            if (!empty($costingData['totalData']['seriesData'])) {
                $cellCol = $this->sheetCol + $this->chartDim['cost2'][0];
                $cellRow = $this->sheetRow;
                $colVal = 0;

                //Initial data
                $this->_writeCellValue($this->alphabets[$cellCol] . $cellRow, '');
                $this->_writeCellValue($this->alphabets[$cellCol + 1] . $cellRow, '');
                $cellRow++;

                $costPerLive = $costingData['totalData']['seriesData'];
                foreach ($costPerLive as $Gid => $colVal) {
                    $indName = __('PER_MILLION_LIFE_SAVE');
                    $costVal = round($colVal, 3);
                    $this->_writeCellValue($this->alphabets[$cellCol] . $cellRow, $indName); //  name
                    $this->_writeCellValue($this->alphabets[$cellCol + 1] . $cellRow, $costVal); // value
                    $cellRow++; //next row
                }
                $this->_makeChart3($cellCol, ($this->sheetRow), count($costPerLive), $this->chartDim['cost2'][1], $chartHeight, 3);
            }

            // Cost-Per Capita
            if (!empty($costingData['ECO_GDP_CAPITA']['seriesData'])) {
                $cellCol = $this->sheetCol + $this->chartDim['cost3'][0];
                $cellRow = $this->sheetRow;

                //Initial data
                $this->_writeCellValue($this->alphabets[$cellCol] . $cellRow, '');
                $this->_writeCellValue($this->alphabets[$cellCol + 1] . $cellRow, '');
                $cellRow++;

                $colVal = 0;
                $costPerCapita = $costingData['ECO_GDP_CAPITA']['seriesData'];
                foreach ($costPerCapita as $Gid => $colVal) {
                    $indName = __('ECO_GDP_CAPITA');
                    $costVal = round($colVal, 3);
                    $this->_writeCellValue($this->alphabets[$cellCol] . $cellRow, $indName); //  name
                    $this->_writeCellValue($this->alphabets[$cellCol + 1] . $cellRow, $costVal); // value
                    $cellRow++; //next row
                }
                $this->_makeChart3($cellCol, ($this->sheetRow), count($costPerLive), $this->chartDim['cost3'][1], $chartHeight, 3);
            }

            $this->sheetRow = $this->sheetRow + ($chartHeight + 1);
        } else {
            $this->sheetRow = $this->sheetRow + 1;
        }
    }

    public function _makeChart($col, $row, $dataCount, $width, $chartHeight, $type) {
        $worksheetName = $this->objWorksheet->getTitle();
        $dataStart = $row + 1;
        $dataEnd = $dataStart + ($dataCount - 1);
        $dataSeriesLabels = $dataSeriesValues = array();

        for ($i = $dataStart; $i <= $dataEnd; $i++) {
            array_push($dataSeriesLabels, new PHPExcel_Chart_DataSeriesValues('String', $worksheetName . '!$' . $this->alphabets[$col] . '$' . $i, NULL, 1));
        }
        $xAxisTickValues = array(
            new PHPExcel_Chart_DataSeriesValues('String', $worksheetName . '!$' . $this->alphabets[$col + 1] . '$' . $row . ':$' . $this->alphabets[$col + 3] . '$' . $row, NULL, 3), //	e.g. Epicauses
        );

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

    public function _makeChartSheetEpic($col, $row, $dataCount, $width, $chartHeight, $type) {

        $worksheetName = $this->objWorksheet->getTitle();
        $dataStart = $row + 1;
        $dataEnd = $dataStart + ($dataCount - 1);
        $dataSeriesLabels = $dataSeriesValues = $xAxisTickValues = array();

        // Vertical lables
        for ($i = $dataStart; $i <= $dataEnd; $i++) {
            array_push($dataSeriesLabels, new PHPExcel_Chart_DataSeriesValues('String', $worksheetName . '!$' . $this->alphabets[$col] . '$' . $i, NULL, 1));
        }

        // Horizontal lables
        $cellCol = $this->colIndexArray['epic2BaseHead'];
        $xAxisTickValues = array(
            new PHPExcel_Chart_DataSeriesValues('String', $worksheetName . '!$' . $this->alphabets[$cellCol] . '$' . $row . ':$' . $this->alphabets[$cellCol + 2] . '$' . $row, NULL, 3), //	e.g. Epicauses
        );


        // Data values 
        $cellCol = $this->colIndexArray['epic2BaseHead'];
        for ($i = $dataStart; $i <= $dataEnd; $i++) {
            array_push($dataSeriesValues, new PHPExcel_Chart_DataSeriesValues('Number', $worksheetName . '!$' . $this->alphabets[$cellCol] . '$' . $i . ':$' . $this->alphabets[$cellCol + 2] . '$' . $i, NULL, 3));
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
        $chartStart = $dataEnd + 1;
        $colStartChart = $col - 1;
        $chart->setTopLeftPosition($this->alphabets[$colStartChart] . $chartStart); // Start after data in sheet
        $chart->setBottomRightPosition($this->alphabets[$colStartChart + $width] . ($chartStart + $chartHeight));
        $this->objWorksheet->addChart($chart);
    }

    public function _makeChart2($col, $row, $dataCount, $width, $chartHeight) {
        $worksheetName = $this->objWorksheet->getTitle();
        $dataStart = $row + 1;
        $dataEnd = $dataStart + ($dataCount - 1);
        $dataSeriesLabels = array(
            new PHPExcel_Chart_DataSeriesValues('String', $worksheetName . '!$' . $this->alphabets[$col + 1] . '$' . $row, NULL, 1), //	Deprived Baseline
            new PHPExcel_Chart_DataSeriesValues('String', $worksheetName . '!$' . $this->alphabets[$col + 2] . '$' . $row, NULL, 1), //	Deprived Target
            new PHPExcel_Chart_DataSeriesValues('String', $worksheetName . '!$' . $this->alphabets[$col + 3] . '$' . $row, NULL, 1), //	Least Deprived Baseline
        );
        //	Set the X-Axis Labels
        $xAxisTickValues = array(
            new PHPExcel_Chart_DataSeriesValues('String', $worksheetName . '!$' . $this->alphabets[$col] . '$' . $dataStart . ':$' . $this->alphabets[$col] . '$' . $dataEnd, NULL, $dataCount), //	e.g. Epicauses
        );
        //	Set the Data values for each data series we want to plot
        $dataSeriesValues = array(
            new PHPExcel_Chart_DataSeriesValues('Number', $worksheetName . '!$' . $this->alphabets[$col + 1] . '$' . $dataStart . ':$' . $this->alphabets[$col + 1] . '$' . $dataEnd, NULL, $dataCount),
            new PHPExcel_Chart_DataSeriesValues('Number', $worksheetName . '!$' . $this->alphabets[$col + 2] . '$' . $dataStart . ':$' . $this->alphabets[$col + 2] . '$' . $dataEnd, NULL, $dataCount),
            new PHPExcel_Chart_DataSeriesValues('Number', $worksheetName . '!$' . $this->alphabets[$col + 3] . '$' . $dataStart . ':$' . $this->alphabets[$col + 3] . '$' . $dataEnd, NULL, $dataCount),
        );

        $series = new PHPExcel_Chart_DataSeries(
                PHPExcel_Chart_DataSeries::TYPE_BARCHART, // plotType
                PHPExcel_Chart_DataSeries::GROUPING_STACKED, // plotGrouping
                range(0, count($dataSeriesValues) - 1), // plotOrder
                $dataSeriesLabels, // plotLabel
                $xAxisTickValues, // plotCategory
                $dataSeriesValues        // plotValues
        );
        // new PHPExcel_Chart_DataSeries($plotType, $plotGrouping, $plotOrder, $plotLabel, $plotCategory, $plotValues, $smoothLine, $plotStyle);
        $series->setPlotDirection(PHPExcel_Chart_DataSeries::DIRECTION_BAR);
        //	Set the series in the plot area
        $plotArea = new PHPExcel_Chart_PlotArea(NULL, array($series));
        //	Set the chart legend
        // $legend = new PHPExcel_Chart_Legend(PHPExcel_Chart_Legend::POSITION_BOTTOM, NULL, false);
        //$legend->setOverlay(true);
        $chrtTitle = ' ';
        $title = new PHPExcel_Chart_Title($chrtTitle);
        $yAxisLabel = new PHPExcel_Chart_Title('');
        $chart = new PHPExcel_Chart(
                null, // name
                $title, // title
                null, // legend
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

    public function _makeChart3($col, $row, $dataCount, $width, $chartHeight) {
        $worksheetName = $this->objWorksheet->getTitle();
        $dataStart = $row + 1;
        $dataEnd = $dataStart + ($dataCount - 1);
        $dataSeriesLabels = array(
            new PHPExcel_Chart_DataSeriesValues('String', $worksheetName . '!$' . $this->alphabets[$col + 1] . '$' . $row, NULL, 1), //	Deprived Baseline
            new PHPExcel_Chart_DataSeriesValues('String', $worksheetName . '!$' . $this->alphabets[$col + 2] . '$' . $row, NULL, 1), //	Deprived Target
            new PHPExcel_Chart_DataSeriesValues('String', $worksheetName . '!$' . $this->alphabets[$col + 3] . '$' . $row, NULL, 1), //	Least Deprived Baseline
        );
        //	Set the X-Axis Labels
        $xAxisTickValues = array(
            new PHPExcel_Chart_DataSeriesValues('String', $worksheetName . '!$' . $this->alphabets[$col] . '$' . $dataStart . ':$' . $this->alphabets[$col] . '$' . $dataEnd, NULL, $dataCount), //	e.g. Epicauses
        );
        //	Set the Data values for each data series we want to plot
        $dataSeriesValues = array(
            new PHPExcel_Chart_DataSeriesValues('Number', $worksheetName . '!$' . $this->alphabets[$col + 1] . '$' . $dataStart . ':$' . $this->alphabets[$col + 1] . '$' . $dataEnd, NULL, $dataCount),
            new PHPExcel_Chart_DataSeriesValues('Number', $worksheetName . '!$' . $this->alphabets[$col + 2] . '$' . $dataStart . ':$' . $this->alphabets[$col + 2] . '$' . $dataEnd, NULL, $dataCount),
            new PHPExcel_Chart_DataSeriesValues('Number', $worksheetName . '!$' . $this->alphabets[$col + 3] . '$' . $dataStart . ':$' . $this->alphabets[$col + 3] . '$' . $dataEnd, NULL, $dataCount),
        );

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
        $legend = new PHPExcel_Chart_Legend(PHPExcel_Chart_Legend::POSITION_BOTTOM, NULL, false);
        //$legend->setOverlay(true);
        $chrtTitle = ' ';
        $title = new PHPExcel_Chart_Title($chrtTitle);
        $yAxisLabel = new PHPExcel_Chart_Title('');
        $chart = new PHPExcel_Chart(
                null, // name
                $title, // title
                null, // legend
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

    public function _makeChart4($col, $row, $dataCount, $width, $chartHeight) {
        $worksheetName = $this->objWorksheet->getTitle();
        $dataStart = $row + 1;
        $dataEnd = $dataStart + ($dataCount - 1);
        $dataSeriesLabels = array(
            new PHPExcel_Chart_DataSeriesValues('String', $worksheetName . '!$' . $this->alphabets[$col + 1] . '$' . $row, NULL, 1), //	Deprived Baseline
            new PHPExcel_Chart_DataSeriesValues('String', $worksheetName . '!$' . $this->alphabets[$col + 2] . '$' . $row, NULL, 1), //	Deprived Target
            new PHPExcel_Chart_DataSeriesValues('String', $worksheetName . '!$' . $this->alphabets[$col + 3] . '$' . $row, NULL, 1), //	Least Deprived Baseline
        );
        //	Set the X-Axis Labels
        $xAxisTickValues = array(
            new PHPExcel_Chart_DataSeriesValues('String', $worksheetName . '!$' . $this->alphabets[$col] . '$' . $dataStart . ':$' . $this->alphabets[$col] . '$' . $dataEnd, NULL, $dataCount), //	e.g. Epicauses
        );
        //	Set the Data values for each data series we want to plot
        $dataSeriesValues = array(
            new PHPExcel_Chart_DataSeriesValues('Number', $worksheetName . '!$' . $this->alphabets[$col + 1] . '$' . $dataStart . ':$' . $this->alphabets[$col + 1] . '$' . $dataEnd, NULL, $dataCount),
            new PHPExcel_Chart_DataSeriesValues('Number', $worksheetName . '!$' . $this->alphabets[$col + 2] . '$' . $dataStart . ':$' . $this->alphabets[$col + 2] . '$' . $dataEnd, NULL, $dataCount),
            new PHPExcel_Chart_DataSeriesValues('Number', $worksheetName . '!$' . $this->alphabets[$col + 3] . '$' . $dataStart . ':$' . $this->alphabets[$col + 3] . '$' . $dataEnd, NULL, $dataCount),
        );

        $series = new PHPExcel_Chart_DataSeries(
                PHPExcel_Chart_DataSeries::TYPE_BARCHART, // plotType
                PHPExcel_Chart_DataSeries::GROUPING_CLUSTERED, // plotGrouping
                range(0, count($dataSeriesValues) - 1), // plotOrder
                $dataSeriesLabels, // plotLabel
                $xAxisTickValues, // plotCategory
                $dataSeriesValues        // plotValues
        );
        // new PHPExcel_Chart_DataSeries($plotType, $plotGrouping, $plotOrder, $plotLabel, $plotCategory, $plotValues, $smoothLine, $plotStyle);
        $series->setPlotDirection(PHPExcel_Chart_DataSeries::DIRECTION_BAR);
        //	Set the series in the plot area
        $plotArea = new PHPExcel_Chart_PlotArea(NULL, array($series));
        //	Set the chart legend
        $legend = new PHPExcel_Chart_Legend(PHPExcel_Chart_Legend::POSITION_BOTTOM, NULL, false);
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
        $chart->setTopLeftPosition($this->alphabets[0] . $row);
        $chart->setBottomRightPosition($this->alphabets[$width] . ($row + $chartHeight));
        //
        //	Add the chart to the worksheet
        $this->objWorksheet->addChart($chart);
    }

    function _sheet1($sheetNumber) {
        $this->_createSheet($sheetNumber);
        $this->objWorksheet->setTitle('Summary');

        //Make scnario name heading
        $this->_fillRowColour($this->colorArray['h1'], 2); // First two rows 
        $this->_putHeadingData('scnrName', $this->userInfo['scnrName']);
        // country name
        $countryVal = $this->userInfo['areaName'] . ' - ' . $this->userInfo['scnrDataset'];
        $this->_putHeadingData('countryName', $countryVal);

        //target population heading
        $this->_fillRowColour($this->colorArray['h2']);
        $this->_putHeadingData('trgetPpl', __('TRGETD_POP'));

        $this->_targetPopulationSelection();
        $this->_epicauseSelection();
        $this->_intrvSelection();
        $this->_severityBtnkPack();
        $this->_causeStrtgyPckg();
        $this->_epicChart();
        $this->_costChart();
        $this->_enablingEnv();
        $this->_fillRowColour($this->colorArray['h1']);
        $this->_resetGlobalVars();
    }

    function _sheet2($sheetNumber) {

        $this->_createSheet($sheetNumber);
        $this->objWorksheet->setTitle('Epicause');
        //Make scnario name heading
        $this->_fillRowColour($this->colorArray['h1'], 4); // First two rows 
        $this->_putHeadingData('scnrName', $this->userInfo['scnrName']);
        // country name
        $countryVal = $this->userInfo['areaName'] . ' - ' . $this->userInfo['scnrDataset'];
        $this->_putHeadingData('countryName', $countryVal);
        // Group heading
        $this->_putHeadingData('epicause', __('Epic_priorities'));


        $this->sheetRow = $this->sheetRow + 1; // blank row

        if (!empty($this->lang_data[EPIC_TP])) {    //  EPICAUSE 
            /* --------Start looping for each group  --------- */
            foreach ($this->lang_data[EPIC_TP] as $grpData) {  //       EPI_NN:{name} 
                $cellCol = $this->sheetCol;
                $cellRow = $this->sheetRow;
                $grpNameMain = '';
                if (!empty($grpData['name'])) {
                    $grpNameMain = $grpData['name'];
                }

                $this->_fillRowColour($this->colorArray['h2'], 3);
                $cellDim = $this->alphabets[$cellCol] . $cellRow;
                $this->_writeCellValue($cellDim, $grpNameMain);
                $this->objWorksheet->getStyle($cellDim)->applyFromArray(['font' => ['size' => 12, 'bold' => true]]); // Make bold  

                $this->sheetRow = $this->sheetRow + 1; // Next row 
                $this->_seriesTypeHeading();

                // Put values for different epicauses
                $cellCol = $this->sheetCol;
                $cellRow = $this->sheetRow;
                $loopCounter = 0;
                if (!empty($grpData['details'])) {
                    foreach ($grpData['details'] as $inct => $indInfo) {
                        if (!empty($indInfo['lnm'])) {
                            $indName = $indInfo['lnm'];
                            $cellDim = $this->alphabets[$cellCol] . $cellRow;

                            $indctSource = $indctActualName = '';
                            $icNStmp = $this->_getICNameSource($inct, EPIC_TP);
                            if (!empty($icNStmp['name'])) {
                                $indctActualName = $icNStmp['name'];
                            }
                            if (!empty($icNStmp['src'])) {
                                $indctSource = $icNStmp['src'];
                            }

                            $this->_writeCellValue($cellDim, $indName); // Name
                            $this->_addComment($cellDim, $indctActualName); // Actual name as comment

                            if (!empty($this->scr_desc_data[EPIC_TP][$inct]['series'])) {
                                $series = $this->scr_desc_data[EPIC_TP][$inct]['series'];
                                $this->_seriesTypeVal($cellRow, $series, $indctSource);  // Put  value
                            }
                            $cellRow++;
                            $loopCounter++;
                        }
                    }
                }
                //echo $loopCounter;die;
                $this->_fillRowColour($this->colorArray['h3'], $loopCounter); // Background colour of inserted rows
// Make hidden sheet
                $this->sheetRow = $cellRow + 1;
                $this->_seriesTypeHeadingHidden();
                // Put values for different epicauses
                $cellCol = $this->sheetCol;
                $cellRow = $this->sheetRow;
                $loopCounter = 0;
                if (!empty($grpData['details'])) {
                    foreach ($grpData['details'] as $inct => $indInfo) {
                        if (!empty($indInfo['lnm'])) {
                            $indName = $indInfo['lnm'];
                            $cellDim = $this->alphabets[$cellCol] . $cellRow;

                            $indctSource = $indctActualName = '';
                            $icNStmp = $this->_getICNameSource($inct, EPIC_TP);
                            if (!empty($icNStmp['name'])) {
                                $indctActualName = $icNStmp['name'];
                            }
                            if (!empty($icNStmp['src'])) {
                                $indctSource = $icNStmp['src'];
                            }
                            $this->_writeCellValue($cellDim, $indName); // Name


                            if (!empty($this->scr_desc_data[EPIC_TP][$inct]['series'])) {
                                $series = $this->scr_desc_data[EPIC_TP][$inct]['series'];
                                $this->_seriesTypeValHidden($cellRow, $series, $indctSource);  // Put  value
                            }
                            $cellRow++;
                            $loopCounter++;
                        }
                    }
                }

                // Create chart from above data
                $chartHeight = 20;
                $this->_makeChart4($cellCol, ($this->sheetRow - 1), $loopCounter, $this->otherInfo['colLimit'], $chartHeight, 1);
                $this->sheetRow = $this->sheetRow + $chartHeight;
                $this->_fillRowColour($this->colorArray['h1']); // Coloured row as seprator
                $this->sheetRow = $this->sheetRow + 1;
            }
        }

        $this->_resetGlobalVars();
    }

    function _sheet3($sheetNumber) {

        $this->_createSheet($sheetNumber);
        $this->objWorksheet->setTitle('Intervention');
        //Make scnario name heading
        $this->_fillRowColour($this->colorArray['h1'], 4); // First two rows 
        $this->_putHeadingData('scnrName', $this->userInfo['scnrName']);
        // country name
        $countryVal = $this->userInfo['areaName'] . ' - ' . $this->userInfo['scnrDataset'];
        $this->_putHeadingData('countryName', $countryVal);
        // Group heading
        $this->_putHeadingData('epicause', __('Intervention'));


        $this->sheetRow = $this->sheetRow + 1; // blank row
        $selBtlnck = $this->request->data['btnk'];
        $btlnckName = '';
        if (!empty($this->lang_data[BTLNK_TP][$selBtlnck]['snm'])) {
            $btlnckName = $this->lang_data[BTLNK_TP][$selBtlnck]['snm'];
        }

        /* --------Start looping for each SDM Package and interventation --------- */

        $loopCounter = 0;
        $type = INTRV_TP;
        if (!empty($this->rel_data['sdmPckgIntrvRel'])) {
            $cellCol = $this->sheetCol;
            foreach ($this->rel_data['sdmPckgIntrvRel'] as $sdm => $pkg) {  // packages inside a SDM
                $sdmName = $this->_getNameFromIndId($sdm);
                $this->_fillRowColour($this->colorArray['h2']);

                $cellDim = $this->alphabets[$cellCol] . $this->sheetRow;
                $this->_writeCellValue($cellDim, $sdmName); // Group name
                $this->objWorksheet->getStyle($cellDim)->applyFromArray(['font' => ['size' => 16, 'bold' => true]]); // Make bold
                //$loopCounter++;
                $this->sheetRow = $this->sheetRow + 1; // Next row
                foreach ($pkg as $pkg => $intrvn) { // Interventions inside a package
                    $packageName = $this->_getNameFromIndId($pkg);
                    $this->_fillRowColour($this->colorArray['h2']);

                    $cellDim = $this->alphabets[$cellCol] . $this->sheetRow;
                    $this->_writeCellValue($cellDim, $packageName);
                    $this->objWorksheet->getStyle($cellDim)->applyFromArray(['font' => ['size' => 12, 'bold' => true]]); // Make bold

                    $this->sheetRow = $this->sheetRow + 1; // Next row

                    $color = $this->sdmColorCode[$sdm];
                    $color = str_replace('#', '', $color);
                    $totalIntrvn = count($intrvn);
                    $this->_fillDataColour(0, $this->sheetRow, $color, $totalIntrvn + 2); // Coloured of selected sdm
                    $this->_seriesTypeHeading();

                    $cellRow = $this->sheetRow;
                    for ($i = 0; $i < $totalIntrvn; $i++) {
                        $intrvnId = $intrvn[$i];
                        $interventionName = '';
                        if (!empty($this->lang_data[$type][$intrvnId]['lnm']))
                            $interventionName = $this->lang_data[$type][$intrvnId]['lnm'];

                        $indctSource = $indctActualName = '';
                        $icNStmp = $this->_getICNameSource($intrvnId, INTRV_TP, $selBtlnck);
                        if (!empty($icNStmp['name'])) {
                            $indctActualName = $icNStmp['name'];
                        }
                        if (!empty($icNStmp['src'])) {
                            $indctSource = $icNStmp['src'];
                        }
                        $dpBaselineValue = $dpTargetValue = $ldpBaselineValue = 0;
                        if (!empty($this->scr_desc_data[$type]['chartData'][$selBtlnck][$intrvnId])) {
                            $series = $this->scr_desc_data[$type]['chartData'][$selBtlnck][$intrvnId];
                            if (!empty($series[DP_BASELINE_CVRG_INDEX]) && (!empty($series[DP_BASELINE_CVRG_INDEX][DP_INDEX]))&&($series[DP_BASELINE_CVRG_INDEX][DP_INDEX]>0)) {
                                $dpBaselineValue = floatval($series[DP_BASELINE_CVRG_INDEX][DP_INDEX]);
                            }
                            if (!empty($series[DP_TARGET_CVRG_INDEX]) && (!empty($series[DP_TARGET_CVRG_INDEX][DP_INDEX]))&&($series[DP_TARGET_CVRG_INDEX][DP_INDEX]>0)) {
                                $dpTargetValue = floatval($series[DP_TARGET_CVRG_INDEX][DP_INDEX]);
                            }
                            if (!empty($series[LDP_BASELINE_CVRG_INDEX]) && (!empty($series[LDP_BASELINE_CVRG_INDEX][DP_INDEX]))&&($series[LDP_BASELINE_CVRG_INDEX][DP_INDEX]>0)) {
                                $ldpBaselineValue = floatval($series[LDP_BASELINE_CVRG_INDEX][DP_INDEX]);
                            }
                        }

                        $cellDim = $this->alphabets[$cellCol] . $cellRow;
                        $this->_writeCellValue($cellDim, $interventionName);
                        $this->_addComment($cellDim, $indctActualName); // Actual name as comment

                        $cellCol = $this->sheetCol + 10; //colIndexArray['epic2BaseHead'];
                        $cellDim = $this->alphabets[$cellCol] . $cellRow;
                        $this->_writeCellValue($cellDim, $dpBaselineValue);

                        $cellCol = $cellCol + 5;
                        $cellDim = $this->alphabets[$cellCol] . $cellRow;
                        $this->_writeCellValue($cellDim, $dpTargetValue);

                        $cellCol = $cellCol + 5;
                        $cellDim = $this->alphabets[$cellCol] . $cellRow;
                        $this->_writeCellValue($cellDim, $ldpBaselineValue);

                        $cellCol = $cellCol + 5;
                        $cellDim = $this->alphabets[$cellCol] . $cellRow;
                        $this->_writeCellValue($cellDim, $indctSource);

                        $cellRow++; // Next row
                        $cellCol = $this->sheetCol;
                    }

                    $cellRow++;
                    $this->sheetRow = $cellRow;
                    $this->_seriesTypeHeadingHidden();
                    $cellRow = $this->sheetRow;
                    for ($i = 0; $i < $totalIntrvn; $i++) {
                        $intrvnId = $intrvn[$i];
                        $interventionName = '';
                        if (!empty($this->lang_data[$type][$intrvnId]['lnm']))
                            $interventionName = $this->lang_data[$type][$intrvnId]['lnm'];

                        $indctSource = $indctActualName = '';
                        $icNStmp = $this->_getICNameSource($intrvnId, INTRV_TP, $selBtlnck);
                        if (!empty($icNStmp['name'])) {
                            $indctActualName = $icNStmp['name'];
                        }
                        if (!empty($icNStmp['src'])) {
                            $indctSource = $icNStmp['src'];
                        }
                        $dpBaselineValue = $dpTargetValue = $ldpBaselineValue = 0;
                        if (!empty($this->scr_desc_data[$type]['chartData'][$selBtlnck][$intrvnId])) {
                            $series = $this->scr_desc_data[$type]['chartData'][$selBtlnck][$intrvnId];
                            if (!empty($series[DP_BASELINE_CVRG_INDEX]) && (!empty($series[DP_BASELINE_CVRG_INDEX][DP_INDEX]))&&($series[DP_BASELINE_CVRG_INDEX][DP_INDEX]>0)) {
                                $dpBaselineValue = floatval($series[DP_BASELINE_CVRG_INDEX][DP_INDEX]);
                            }
                            if (!empty($series[DP_TARGET_CVRG_INDEX]) && (!empty($series[DP_TARGET_CVRG_INDEX][DP_INDEX]))&&($series[DP_TARGET_CVRG_INDEX][DP_INDEX]>0)) {
                                $dpTargetValue = floatval($series[DP_TARGET_CVRG_INDEX][DP_INDEX]);
                            }
                            if (!empty($series[LDP_BASELINE_CVRG_INDEX]) && (!empty($series[LDP_BASELINE_CVRG_INDEX][DP_INDEX]))&&($series[LDP_BASELINE_CVRG_INDEX][DP_INDEX]>0)) {
                                $ldpBaselineValue = floatval($series[LDP_BASELINE_CVRG_INDEX][DP_INDEX]);
                            }
                        }

                        $cellDim = $this->alphabets[$cellCol] . $cellRow;
                        $this->_writeCellValue($cellDim, $interventionName);

                        $cellCol = $this->colIndexArray['epic2BaseHead'];
                        $cellDim = $this->alphabets[$cellCol] . $cellRow;
                        $this->_writeCellValue($cellDim, $dpBaselineValue);

                        $cellCol++;
                        $cellDim = $this->alphabets[$cellCol] . $cellRow;
                        $this->_writeCellValue($cellDim, $dpTargetValue);

                        $cellCol++;
                        $cellDim = $this->alphabets[$cellCol] . $cellRow;
                        $this->_writeCellValue($cellDim, $ldpBaselineValue);
                        $cellRow++; // Next row
                        $cellCol = $this->sheetCol;
                    }


                    $chartHeight = 20;
                    $this->_makeChart4($cellCol, ($this->sheetRow - 1), $totalIntrvn, $this->otherInfo['colLimit'] + 1, $chartHeight, 1);
                    $this->sheetRow = $this->sheetRow + $chartHeight;
                    $this->_fillRowColour($this->colorArray['h1']); // Coloured row as seprator
                    $this->sheetRow = $this->sheetRow + 1;
                }
            }
        }

//            if (!empty($this->rel_data['sdmPckgIntrvRel'])) {
//                foreach ($this->rel_data['sdmPckgIntrvRel'] as $sdm => $pkg) {
//                    $sdmName = $this->_getNameFromIndId($sdm);
//                    // Put SDM name at to for each group.Also store index to make it bold
//                    array_push($workSheetDataArry, array($sdmName));
//                    $sdmInfo[$sdmName]['index'] = $indexCounter;
//                    $indexCounter++; // Pick next row
//                    foreach ($pkg as $pkg => $intrvn) {
//                        // Set heading for each set of interventions
//                        $packageName = $this->_getNameFromIndId($pkg);
//                        $groupDimensions[$packageName]['index'] = $indexCounter; // Store the location of heading
//                        array_push($workSheetDataArry, array($packageName));  // Insert package name at first row for each set of interventions
//                        array_push($groupHeadingIndex, $indexCounter);
//                        $indexCounter++;  // No more data should be inserted here. Goto next row
//                        array_push($workSheetDataArry, $initialRow);
//                        $indexCounter++; // Pick next row
//                        $totalIntrvn = count($intrvn);
//                        $interventionName = '';
//                        $loopCount = 0;
//                        for ($i = 0; $i < $totalIntrvn; $i++) {
//                            $intrvnId = $intrvn[$i];
//                            if (!empty($this->lang_data[$type][$intrvnId]['lnm']))
//                                $interventionName = $this->lang_data[$type][$intrvnId]['lnm'];
//
//                            //   Set max width of first row
//                            if (strlen($interventionName) > $firstColWidth) {
//                                $firstColWidth = strlen($interventionName);
//                            }
//
//                            $indctSource = $indctActualName = '';
//                            $icNStmp = $this->_getICNameSource($intrvnId, INTRV_TP, $selBtlnck);
//                            if (!empty($icNStmp['name'])) {
//                                $indctActualName = $icNStmp['name'];
//                            }
//                            if (!empty($icNStmp['src'])) {
//                                $indctSource = $icNStmp['src'];
//                            }
//                            $commentArray[$indexCounter] = $indctActualName;
//                            // Now get values from desc tree json
//                            $dpBaselineValue = $dpTargetValue = $ldpBaselineValue = 0;
//                            if (!empty($this->scr_desc_data[$type]['chartData'][$selBtlnck][$intrvnId])) {
//                                $series = $this->scr_desc_data[$type]['chartData'][$selBtlnck][$intrvnId];
//                                if (!empty($series[DP_BASELINE_CVRG_INDEX]) && (!empty($series[DP_BASELINE_CVRG_INDEX][DP_INDEX]))) {
//                                    $dpBaselineValue = floatval($series[DP_BASELINE_CVRG_INDEX][DP_INDEX]);
//                                }
//                                if (!empty($series[DP_TARGET_CVRG_INDEX]) && (!empty($series[DP_TARGET_CVRG_INDEX][DP_INDEX]))) {
//                                    $dpTargetValue = floatval($series[DP_TARGET_CVRG_INDEX][DP_INDEX]);
//                                }
//                                if (!empty($series[LDP_BASELINE_CVRG_INDEX]) && (!empty($series[LDP_BASELINE_CVRG_INDEX][DP_INDEX]))) {
//                                    $ldpBaselineValue = floatval($series[LDP_BASELINE_CVRG_INDEX][DP_INDEX]);
//                                }
//                            }
//                            /* -------- Push data for current intervention--------- */
//                            array_push($workSheetDataArry, array($interventionName, $dpBaselineValue, $dpTargetValue, $ldpBaselineValue, $indctSource));
//                            $indexCounter++;
//                            $loopCount++;
//                        }
//                        $groupDimensions[$packageName]['count'] = $loopCount;
//
//                        /* --------Add seprators --------- 
//                         * Take count for chart length also
//                         *                          
//                         */
//                        $sepratorsLength = $loopCount + $groupSeprationFactor;
//                        for ($ic = 0; $ic < $sepratorsLength; $ic++) {
//                            array_push($workSheetDataArry, $blankRow);
//                            $indexCounter++;
//                        }
//                    }
//                }
//            }
        $this->_resetGlobalVars();
    }

    function _seriesTypeHeading() {
        $rowHeadings = array(__('BSLNE_4_TRGT_POP'), __('PRJT_ELN_TRGT_POP'), __('LST_DPRVD_POP'), __('Source'));
        $gap=5;
        $cellRow = $this->sheetRow;
        $cellCol = $this->sheetCol + 10; //$this->colIndexArray['epic2BaseHead'];
        
        
        
        foreach ($rowHeadings as $key => $heading) {
            $this->objWorksheet->mergeCellsByColumnAndRow($cellCol, $cellRow, $cellCol+($gap-1), $cellRow+1);
            $cellDim = $this->alphabets[$cellCol] . $cellRow;
            $this->_writeCellValue($cellDim, $heading);
            $this->objWorksheet->getStyle($cellDim)->applyFromArray(['font' => ['bold' => true]]); // Make bold  
            $this->objWorksheet->getStyle($cellDim)->getAlignment()->setWrapText(true);
            $cellCol = $cellCol + $gap; // $this->otherInfo['epicSheetSeriesGap'];
        }
        $this->sheetRow = $this->sheetRow + 2; // Next row ,after heading
    }

    function _seriesTypeHeadingHidden() {
        $rowHeadings = array(__('BSLNE_4_TRGT_POP'), __('PRJT_ELN_TRGT_POP'), __('LST_DPRVD_POP'));
        $cellRow = $this->sheetRow;
        $cellCol = $this->colIndexArray['epic2BaseHead'];
        foreach ($rowHeadings as $key => $heading) {
            $cellDim = $this->alphabets[$cellCol] . $cellRow;
            $this->_writeCellValue($cellDim, $heading);
            $this->objWorksheet->getStyle($cellDim)->applyFromArray(['font' => ['bold' => true]]); // Make bold  
            $cellCol = $cellCol + 1; // $this->otherInfo['epicSheetSeriesGap'];
        }
        $this->sheetRow = $this->sheetRow + 1; // Next row ,after heading
    }

    function _seriesTypeVal($cellRow, $series, $source) {

        $cellCol = $this->sheetCol + 10;
        //Baseline for targeted population
        $dpBaselineValue = $dpTargetValue = $ldpBaselineValue = 0;
        if (!empty($series[DP_BASELINE_CVRG_INDEX]) && (!empty($series[DP_BASELINE_CVRG_INDEX][DP_INDEX]))&&($series[DP_BASELINE_CVRG_INDEX][DP_INDEX]>0)) {
            $dpBaselineValue = floatval($series[DP_BASELINE_CVRG_INDEX][DP_INDEX]);
        }
        $cellDim = $this->alphabets[$cellCol] . $cellRow;
        $this->_writeCellValue($cellDim, $dpBaselineValue);

        $cellCol = $cellCol + 5; // $this->otherInfo['epicSheetSeriesGap'];
        //Projected endline for target population
        if (!empty($series[DP_TARGET_CVRG_INDEX]) && (!empty($series[DP_TARGET_CVRG_INDEX][DP_INDEX]))&&($series[DP_TARGET_CVRG_INDEX][DP_INDEX]>0)) {
            $dpTargetValue = floatval($series[DP_TARGET_CVRG_INDEX][DP_INDEX]);
        }
        $cellDim = $this->alphabets[$cellCol] . $cellRow;
        $this->_writeCellValue($cellDim, $dpTargetValue);

        $cellCol = $cellCol + 5; // $this->otherInfo['epicSheetSeriesGap'];
        //Least deprived (richest) population
        if (!empty($series[LDP_BASELINE_CVRG_INDEX]) && (!empty($series[LDP_BASELINE_CVRG_INDEX][DP_INDEX]))&&($series[LDP_BASELINE_CVRG_INDEX][DP_INDEX]>0)) {
            $ldpBaselineValue = floatval($series[LDP_BASELINE_CVRG_INDEX][DP_INDEX]);
        }
        $cellDim = $this->alphabets[$cellCol] . $cellRow;
        $this->_writeCellValue($cellDim, $ldpBaselineValue);

        // Source
        $cellCol = $cellCol + 5; //$this->otherInfo['epicSheetSeriesGap'];
        $cellDim = $this->alphabets[$cellCol] . $cellRow;
        $this->_writeCellValue($cellDim, $source);
    }

    function _seriesTypeValHidden($cellRow, $series, $source) {

        $cellCol = $this->colIndexArray['epic2BaseHead'];
        //Baseline for targeted population
        $dpBaselineValue = $dpTargetValue = $ldpBaselineValue = 0;
        if (!empty($series[DP_BASELINE_CVRG_INDEX]) && (!empty($series[DP_BASELINE_CVRG_INDEX][DP_INDEX]))&&($series[DP_BASELINE_CVRG_INDEX][DP_INDEX]>0)) {
            $dpBaselineValue = floatval($series[DP_BASELINE_CVRG_INDEX][DP_INDEX]);
        }
        $cellDim = $this->alphabets[$cellCol] . $cellRow;
        $this->_writeCellValue($cellDim, $dpBaselineValue);

        $cellCol = $cellCol + 1; // $this->otherInfo['epicSheetSeriesGap'];
        //Projected endline for target population
        if (!empty($series[DP_TARGET_CVRG_INDEX]) && (!empty($series[DP_TARGET_CVRG_INDEX][DP_INDEX]))&&($series[DP_TARGET_CVRG_INDEX][DP_INDEX]>0)) {
            $dpTargetValue = floatval($series[DP_TARGET_CVRG_INDEX][DP_INDEX]);
        }
        $cellDim = $this->alphabets[$cellCol] . $cellRow;
        $this->_writeCellValue($cellDim, $dpTargetValue);

        $cellCol = $cellCol + 1; // $this->otherInfo['epicSheetSeriesGap'];
        //Least deprived (richest) population
        if (!empty($series[LDP_BASELINE_CVRG_INDEX]) && (!empty($series[LDP_BASELINE_CVRG_INDEX][DP_INDEX]))&&($series[LDP_BASELINE_CVRG_INDEX][DP_INDEX]>0)) {
            $ldpBaselineValue = floatval($series[LDP_BASELINE_CVRG_INDEX][DP_INDEX]);
        }
        $cellDim = $this->alphabets[$cellCol] . $cellRow;
        $this->_writeCellValue($cellDim, $ldpBaselineValue);
    }

    function _writeCellValue($cell, $value) {
        $this->objWorksheet->getCell($cell)->setValue($value);
    }

    function _addComment($cellDim, $comment) {
        $this->objWorksheet->getComment($cellDim)->getText()->createTextRun($comment); // Add comment
    }

    function _drawFullLine() {
        $border_style = array('borders' => array('bottom' =>
                array('style' =>
                    PHPExcel_Style_Border::BORDER_THIN,
                    'color' => array('rgb' => 'D9D9D9')
        )));
        $lineRange = $this->alphabets[$this->sheetCol] . $this->sheetRow . ':' . $this->alphabets[$this->otherInfo['colLimit']] . $this->sheetRow;
        $this->objWorksheet->getStyle($lineRange)->applyFromArray($border_style);
    }

    function _drawLine($lineRange) {
        $border_style = array('borders' => array('bottom' =>
                array('style' =>
                    PHPExcel_Style_Border::BORDER_THIN,
                    'color' => array('rgb' => 'D9D9D9')
        )));
        // $lineRange = $this->alphabets[$this->sheetCol] . $this->sheetRow . ':' . $this->alphabets[$this->otherInfo['colLimit']] . $this->sheetRow;
        $this->objWorksheet->getStyle($lineRange)->applyFromArray($border_style);
    }

    function _fillCellColour($cellDim, $color) {
        $cellBckStyle = array(
            'fill' => array(
                'type' => PHPExcel_Style_Fill::FILL_SOLID,
                'color' => array('rgb' => $color)
        ));
        $this->objWorksheet->getStyle($cellDim)->applyFromArray($cellBckStyle); // fill colour
    }

    function _resetGlobalVars() {
        $this->sheetRow = 1;
        $this->sheetCol = 1;
    }

    function getSvrtyColour($dataVal) {
        $dataVal = ceil($dataVal);
        $red = 255;
        $green = 255;
        $blue = 255;
        /*
         * green  (0,128,0)
         * yellow=(255,255,0)
         * red=(255,0,0)
         * 
         */
        if ($dataVal == 0) {  //green
            $red = 0;
            $green = 128;
            $blue = 0;
        } elseif ($dataVal == 25) {  //yellow
            $blue = 0;
        } elseif ($dataVal == 100) { //red
            $green = 0;
            $blue = 0;
        } elseif (in_array($dataVal, range(1, 24))) {  // 1 to 24
            $blue = 0;

            $Min = 0;
            $Max = 25;
            $maxMinDiff = $Max - $Min;
            $redDiff = 255 - 0;

            $greenDiff = 255 - 128;

            /*             * min+[(val-min)/(max-min)]*diff            */
            $red = ceil(0 + $Min + (($dataVal - $Min) / $maxMinDiff) * $redDiff);
            $green = ceil(128 + $Min + (($dataVal - $Min) / $maxMinDiff) * $greenDiff);
        } elseif (in_array($dataVal, range(26, 99))) { // 26 to 99
            /*        max-[(val-min)/(max-min)]*diff            */
            $red = 255;
            $blue = 0;

            $Min = 25;
            $Max = 100;
            $maxMinDiff = $Max - $Min;

            $greenDiff = 255 - 0;
            $green = floor(255 - (($dataVal - $Min) / $maxMinDiff) * $greenDiff);
        }
        $colour = $this->RGBToHex($red, $green, $blue);
        return $colour;

        /* $red = 255;
          $green = 255;
          $blue = 255;
          if ($val == 0) {  //green
          $red = 0;
          $green = 128;
          $blue = 0;
          } elseif ($val == 25) {  //yellow
          $blue = 0;
          } elseif ($val == 100) { //red
          $green = 0;
          $blue = 0;
          } elseif (in_array($val, range(1, 24))) {  // 1 to 24
          $blue = 0;
          $factor1 = 10.54; //=>253/24
          $red = 0 + floor($val * $factor1);

          $factor2 = 5.21; //=>125/24
          $green = 128 + floor($val * $factor2);
          } elseif (in_array($val, range(26, 99))) { // 26 to 99
          $red = 255;
          $blue = 0;
          $factor = 11; //=>253/23
          $green = 255 - floor(($val - 25) * $factor);
          }
          $colour =$this->RGBToHex($red,$green,$blue);
          return $colour; */
    }

    function RGBToHex($r, $g, $b) {
        //String padding bug found and the solution put forth by Pete Williams (http://snipplr.com/users/PeteW)
        $hex = str_pad(dechex($r), 2, "0", STR_PAD_LEFT);
        $hex.= str_pad(dechex($g), 2, "0", STR_PAD_LEFT);
        $hex.= str_pad(dechex($b), 2, "0", STR_PAD_LEFT);

        return $hex;
    }

}

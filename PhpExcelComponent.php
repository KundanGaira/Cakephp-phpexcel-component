<?php

class PhpExcelComponent extends Component {

    private $phpExcelName = 'PHPExcel';
    private $objPHPExcel = null, $objWorksheet = null;
    private $inputFileType = 'Excel2007';
    private $defaults = array('extension' => '.xlsx', 'excelName' => 'ExcelSheet','sheet1Name'=>'Sheet1');

   
     /*
     * Create workbook,and return 
     */

    public function createExcel() {
        $loadStatus = App::import('Vendor', 'PHPExcel'); // Load PHPExcel from vender location
        if (!$loadStatus) {
            $msg = 'Unable to load ' . $this->phpExcelName . '.';
            $this->_requestError($msg);
        }
        $this->objPHPExcel = new PHPExcel(); // Make object accessable globally
        $this->objPHPExcel->getActiveSheet()->setTitle($this->defaults['sheet1Name']);
        return $this->objPHPExcel;
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

        header('Content-type: application/vnd.ms-excel');
        header('Content-Disposition: attachment; filename=' . $fileName);
        // Write file to the browser
        $objWriter->save('php://output');
        die;
    }

    function getExcelObj() {
        return $this->objPHPExcel;
    }

    function getWorksheetObj() {
        return $this->objWorksheet;
    }

    private function _requestError($msg) {
        echo json_encode(array('status' => 'error', 'msg' => $msg));
        die;
    }

}

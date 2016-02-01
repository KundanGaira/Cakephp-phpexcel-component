# Cakephp Component for Phpexcel
  Cakephp component for using **Phpexcel** with cakephp application. Commonly used features of Phpexcel are  simplified in simple   methods.
  
### Dependencies
 * https://github.com/PHPOffice/PHPExcel
 * Cakephp versions 2.*
 * Your Patience :-)

### How to use
 * Download, and place https://github.com/PHPOffice/PHPExcel inside you vendor **(app/Vendor)**.
 * Place our Component inside Component folder **(app/Controller/Component)**.
 * Load the component in any Controller in order to use it.
    
  ```public $components = array('PhpExcel');``` 

### Conventions
    *Cell References should be alphanumeric value. e.g. "A23", "B23".
    *Colors  should be in hex code (without hash symbol) or a text like "red"
  
### Example usage

    $PhpExcel=$this->PhpExcel;
    $PhpExcel->createExcel();
    $PhpExcel->downloadFile();
   

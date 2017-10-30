# Cakephp "Phpexcel" Component 
  This component is for popular PHP based MVC framework "Cakephp",for using features of **Phpexcel**. Some common features of Phpexcel are  simplified in a form of simple methods.You can enhance your cakephp application with features like excel export, excel chart etc.
  
### Dependencies
 * Core Phpexcel Class.(Placed inside PhpExcel folder But you can also download from https://github.com/PHPOffice/PHPExcel). 
 * Your Patience :-)

### How to use
 For "Cakephp 3.*"
 
 * Put PhpExcel folder inside your vendor folder (/vendor).
 * Put PhpExcelComponent.php file (from cakephp3 folder) inside /src/Controller/Component.
 * Load the component in any Controller in order to use it.

     ```$this->loadComponent('PhpExcel');``` 
 
For "Cakephp 2.*"

  *Put PhpExcel folder inside your vendor folder (/app/Vendor).
  *Put PhpExcelComponent.php file (from cakephp2 folder) inside /app/Controller/Component.  
  * Load the component in any Controller in order to use it.

    ```public $components = array('PhpExcel');``` 

### Conventions
    *Cell References should be alphanumeric value. e.g. "A2", "B2". For first cell of excel ,It is "A1".
    *Colors  should be in hex code (without hash symbol) or a text like "red"
  
### Example usage

    $PhpExcel=$this->PhpExcel;
    $PhpExcel->createExcel();
    $PhpExcel->downloadFile();
   

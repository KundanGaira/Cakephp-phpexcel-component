# Cakephp "Phpexcel" Component 
  This component is for popular MVC based PHP framework "Cakephp",for using features of **Phpexcel**. Some common features of Phpexcel are  simplified in a form of simple methods.You can enhance your cakephp application with features like excel export, excel chart etc.
  
### Dependencies
 * Core Phpexcel Classes.(check more about it at https://github.com/PHPOffice/PHPExcel). 
 * Your Patience :-)

### How to use
 **For "Cakephp 3.*"**
 
 * First, Put PhpExcelComponent.php file (from cakephp3 folder) inside /src/Controller/Component.
 * Now second step is to, include directory containing Phpexcel classes.This can be done in two ways.
   
    1.Put "phpoffice" folder directly inside your vendor folder (/vendor). OR

    2.Using composer

       Add these lines in your composer.json

            "repositories": [
            {
                "type": "vcs",
                "url": "https://github.com/PHPOffice/PHPExcel.git"
            }
            ],
        
        
            "require": {
                "phpoffice/phpexcel": "1.8.0"
            },
            

    Then update composer using command

        ```composer update``` 

 * Now in order to use it, load the component in any Controller using following code.

     ```$this->loadComponent('PhpExcel');``` 
 
**For "Cakephp 2.*"**

  *Put PhpExcel folder inside your vendor folder (/app/Vendor).
  *Put PhpExcelComponent.php file (from cakephp2 folder) inside /app/Controller/Component.  
  * Load the component in any Controller in order to use it.

    ```public $components = array('PhpExcel');``` 

### Coding Conventions
    *Cell References should be alphanumeric value. e.g. "A2", "B2".Like for first cell of excel sheet,It is "A1".
    *Colors  should be in hex code (without hash symbol) or a text like "red"
  
### Example usage

    $PhpExcel=$this->PhpExcel;
    $PhpExcel->createExcel();
    $PhpExcel->downloadFile();
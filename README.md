# Cakephp "Phpexcel" Component 
  This component is for popular MVC based PHP framework "Cakephp",for using features of **Phpexcel**. Some common features of Phpexcel are  simplified in a form of simple methods.You can enhance your cakephp application with features like excel export, excel chart etc.
  
### Dependencies
 * Core Phpexcel Classes.(check more about it at https://github.com/PHPOffice/PHPExcel). 
 * Your Patience :-)
 
#### Update #1
Recently "PHPOffice/PHPExcel" was changed to read only. And in php **v >=7.4** this library is throwing error due to deprecated offset declaration [Array and string offset access syntax with curly braces is deprecated].

In our **phpoffice/phpexcel** folder we have "made changes to remove errors". You still can download phpoffice/phpexcel library from official website and can make following changes.

    For any offset ,replace {} with [] . 

  * phpoffice/phpexcel/Classes/PHPExcel/Shared/String.php  line # 529,530,536,537
  * phpoffice/phpexcel/Classes/PHPExcel/Calculation.php  line# 2186,2294,2296,2372,2374,2383,2632,2761,2763,2764,3039,3039,3042,3043,3459,3459,3501,3505,3558,3559
  * phpoffice/phpexcel/Classes/PHPExcel/Worksheet/AutoFilter.php  line# 729
  * phpoffice/phpexcel/Classes/PHPExcel/ReferenceHelper.php  line# 892,894
  * phpoffice/phpexcel/Classes/PHPExcel/Cell.php  line# 772,773,776,777,779,780,784
  * phpoffice/phpexcel/Classes/PHPExcel/Calculation/Functions.php  line# 311,313,534
  
phpoffice/phpexcel/Classes/PHPExcel/Calculation/Functions.php  Please remove unnecessary break; line#581 
    

### How to use
 **For "Cakephp 3.*"**
 
 * First, Put PhpExcelComponent.php file (from our cakephp3 folder) inside /src/Controller/Component.
 * Now second step is, to include directory containing Phpexcel classes.This can be done in two ways.
   
    1.Put "phpoffice" folder directly inside your vendor folder (/vendor). **OR**

    2.Using composer [Update:Officical phpoffice will not work for php7.4 and higher. **We recommend to use our version of phpoffice** OR make changes suggested above ]

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

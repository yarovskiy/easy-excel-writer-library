# Easy Excel Writer Library ![Library Image](http://www.fontstuff.com/excel/images/icon_excel_2010.jpg)

[![Average time to resolve an issue](http://isitmaintained.com/badge/resolution/svetel/easy-excel-writer-library.svg)](http://isitmaintained.com/project/svetel/easy-excel-writer-library "Average time to resolve an issue")
[![Percentage of issues still open](http://isitmaintained.com/badge/open/svetel/easy-excel-writer-library.svg)](http://isitmaintained.com/project/svetel/easy-excel-writer-library "Percentage of issues still open")

##Installation

Install via [Composer](https://getcomposer.org/) 
(```svetel/easy-excel-writer``` on [Packagist](https://packagist.org/))

Max columns = 1024

Library which provides easy way to create Excel files

Features:

- [x] Add rows like singledimensional array
- [x] Add cell values to position according with column value
- [ ] Sort columns alphabetical
- [ ] Sort columns using own rules

###Examples:

1. Add rows like singledimension array, where each array element it's cell value. 
Filling cells starts with column with index "A". 

    ```php
    <?php 
    #index.php
    require_once '__DIR__.'/vendor/autoload.php';
    $easyExcelWriter = new \Svetel\Library\EasyExcelWriter;
    $rows = [
      ['first row first cell value','first row first second value'],
      ['second row first cell value','second row first second value']
    ];
    foreach ($rows as $row){
      $easyExcelWriter->addRow($row);
    }
    $easyExcelWriter->setFileName('testFile.xlsx');//full path and filename of excel file
    $easyExcelWriter->save();
    ?>
    ```

2. Add row like singledimension associative array, where is key - column name

    ```php
    <?php 
    #index.php
    require_once '__DIR__.'/vendor/autoload.php';
    $excelColumns = new \Svetel\Library\Resources\ColumnNames;
            $excelColumns->addColumn('firstColumn Name')
                    ->addColumn('secondColumn Name')
                    ->addColumn('thirdColumn Name');//adding column names
    $easyExcelWriter = new \Svetel\Library\EasyExcelWriter;
    $easyExcelWriter->setColumnNames($excelColumns);
    $rows = [
     [ 
         'firstColumn Name' =>'firstColumn Value'
                , 'thirdColumn Name' => 'thirdColumn Value'
                , 'secondColumn Name' => 'secondColumn Value'
         ]
    ];
    foreach ($rows as $row){
      $easyExcelWriter->addRow($row);
    }
    $easyExcelWriter->setFileName('testFile.xlsx');//full path and filename of excel file
    $easyExcelWriter->save();
    ?>
    ```

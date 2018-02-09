# Easy Excel Writer Library  ![Library Image](http://www.fontstuff.com/excel/images/icon_excel_2010.jpg)
[![Average time to resolve an issue](http://isitmaintained.com/badge/resolution/svetel/easy-excel-writer-library.svg)](http://isitmaintained.com/project/svetel/easy-excel-writer-library "Average time to resolve an issue")
[![Percentage of issues still open](http://isitmaintained.com/badge/open/svetel/easy-excel-writer-library.svg)](http://isitmaintained.com/project/svetel/easy-excel-writer-library "Percentage of issues still open")


Максимальное число столбцов = 1024

Максимально простой способ создания Excel файлов

Что умеет делать:

- [x] Добавлять строки как одномерный массив
- [x] Добавлять значения ячеек в столбец в соотвествии со значением столбца колонки
- [ ] Сортировать значения колонок по алфавиту
- [ ] Использовать собственную сортировку значений колонок

###Примеры:

1. Добавление строки как одномерный массив, где каждый элемент 
массива - значение ячейки. Заполнение значений ячеек начинается со столбца с 
индексом "A". 

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
    $easyExcelWriter->setFileName('testFile.xlsx');//Полный путь и имя файла
    $easyExcelWriter->save();
    ?>
    ```

2. Добавление строки как одномерного ассоциативного массива, где ключ массива - 
название столбца ячейки

    ```php
    <?php 
    #index.php
    require_once '__DIR__.'/vendor/autoload.php';
    $excelColumns = new \Svetel\Library\Resources\ColumnNames;
            $excelColumns->addColumn('firstColumn Name')
                    ->addColumn('secondColumn Name')
                    ->addColumn('thirdColumn Name');//добавление названий колонок
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
    $easyExcelWriter->setFileName('testFile.xlsx');//Полный путь и имя файла
    $easyExcelWriter->save();
    ?>
    ```

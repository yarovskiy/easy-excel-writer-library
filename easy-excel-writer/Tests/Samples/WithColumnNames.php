<?php

use Svetel\Library\EasyExcelWriter;

$excelColumns = new \Svetel\Library\Resources\ColumnNames;
$excelColumns->addColumn('firstColumn Name')
        ->addColumn('secondColumn Name')
        ->addColumn('thirdColumn Name');
$easyExcelWriter = new EasyExcelWriter;
$easyExcelWriter->setColumnNames($excelColumns);
$rows = [
    [
        'firstColumn Name' => 'firstColumn Value'
        , 'thirdColumn Name' => 'thirdColumn Value'
        , 'secondColumn Name' => 'secondColumn Value'
    ]
];
foreach ($rows as $row) {
    $easyExcelWriter->addRow($row);
}
$easyExcelWriter->setFileName(
        __DIR__ . DIRECTORY_SEPARATOR
        . basename(__FILE__, ".php") . '.xlsx'); //full path and filename of excel file
$easyExcelWriter->save();

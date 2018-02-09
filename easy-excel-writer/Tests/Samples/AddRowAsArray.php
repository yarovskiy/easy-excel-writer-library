<?php


use Svetel\Library\EasyExcelWriter;

$easyExcelWriter = new EasyExcelWriter;
$rows = [
    ['first row first cell value', 'first row first second value'],
    ['second row first cell value', 'second row first second value']
];
foreach ($rows as $row) {
    $easyExcelWriter->addRow($row);
}

$easyExcelWriter->setFileName(
        __DIR__ . DIRECTORY_SEPARATOR
        . basename(__FILE__, ".php")
        . '.xlsx'); //full path and filename of excel file
$easyExcelWriter->save();

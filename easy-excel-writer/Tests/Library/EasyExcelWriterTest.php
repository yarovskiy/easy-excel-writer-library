<?php

namespace Svetel\Tests\Library;

use Svetel\Library\EasyExcelWriter;

class EasyExcelWriterTest extends \PHPUnit_Framework_TestCase
{

    /**
     * @var EasyExcelWriter
     */
    protected $object;

    /**
     * Sets up the fixture, for example, opens a network connection.
     * This method is called before a test is executed.
     */
    protected function setUp()
    {
        $this->object = new EasyExcelWriter;
    }

    /**
     * Tears down the fixture, for example, closes a network connection.
     * This method is called after a test is executed.
     */
    protected function tearDown()
    {
        $this->object = NULL;
    }

    /**
     * @covers Svetel\Library\EasyExcelWriter::addRow
     * @covers Svetel\Library\EasyExcelWriter::save
     */
    public function testAddRow()
    {
        $rowsCount = 9;
        $row = range(1, 1024);
        while ($rowsCount--) {
            $this->object->addRow($row);
        }
        $fileName = __DIR__ . '/../Samples/' . __FUNCTION__ . '.xlsx';
        $this->object->setFileName($fileName);
        $this->object->save();
    }

    /**
     * @covers Svetel\Library\Resources\ExcelColumns::addColumn
     * @covers Svetel\Library\EasyExcelWriter::addRow
     * @covers Svetel\Library\EasyExcelWriter::save
     */
    public function testCreateFileWithColumns()
    {
        $rowsCount = 9;
        $excelColumns = new \Svetel\Library\Resources\ColumnNames;
        $excelColumns->addColumn('firstColumn Name')
                ->addColumn('secondColumn Name')
                ->addColumn('thirdColumn Name');
        $this->object->setColumnNames($excelColumns);
        $row = [
            'firstColumn Name' => 1
            , 'secondColumn Name' => 2
            , 'thirdColumn Name' => 3
        ];
        while ($rowsCount--) {
            $this->object->addRow($row);
        }
        $fileName = __DIR__ . '/../Samples/' . __FUNCTION__ . '.xlsx';
        $this->object->setFileName($fileName);
        $this->object->save();
    }

    /**
     * @covers Svetel\Library\EasyExcelWriter::getFileName
     * @covers Svetel\Library\EasyExcelWriter::setFileName
     */
    public function testSetFileName()
    {
        $fileName = __DIR__ . '/../Samples/' . __FUNCTION__ . '.xlsx';

        $this->object->setFileName($fileName);
        $this->assertEquals($fileName, $this->object->getFileName());
    }

    /**
     * @covers Svetel\Library\EasyExcelWriter::generateContentTypesXML
     */
    public function testGenerateContentTypesXML()
    {
        $xml = $this->object->generateContentTypesXML();
    }

}

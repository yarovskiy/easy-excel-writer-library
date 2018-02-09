<?php

namespace Svetel\Tests\Library\Collections;

use Svetel\Library\Collections\RowCollection;
use Svetel\Library\Collections\CellCollection;

class RowCollectionTest extends \PHPUnit_Framework_TestCase
{

    /**
     * @var RowCollection
     */
    protected $object;

    /**
     * Sets up the fixture, for example, opens a network connection.
     * This method is called before a test is executed.
     */
    protected function setUp()
    {
        $this->object = new RowCollection;
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
     * @covers Svetel\Library\Collections\RowCollection::addRow
     * @covers Svetel\Library\Collections\RowCollection::getCollection
     * @covers Svetel\Library\Collections\RowCollection::count
     */
    public function testAddRow()
    {
        $cellCollection = new CellCollection;
        $cellList = [
            'firstValue'
            , 'secondValue'
        ];
        foreach ($cellList as $cellKey => $cellValue) {
            $cellCollection->addCell($cellKey, $cellValue);
        }
        $this->object->addRow($cellCollection);
        $this->assertTrue(is_array($this->object->getCollection()));
        $this->assertEquals(1, $this->object->count());
    }

}

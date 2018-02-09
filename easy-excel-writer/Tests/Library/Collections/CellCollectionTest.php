<?php

namespace Svetel\Tests\Library\Collections;

use Svetel\Library\Collections\CellCollection;

class CellCollectionTest extends \PHPUnit_Framework_TestCase
{

    /**
     * @var CellCollection
     */
    protected $object;

    /**
     * Sets up the fixture, for example, opens a network connection.
     * This method is called before a test is executed.
     */
    protected function setUp()
    {
        $this->object = new CellCollection;
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
     * @covers Svetel\Library\Collections\CellCollection::addCell
     * @covers Svetel\Library\Collections\CellCollection::getCollection
     * @covers Svetel\Library\Collections\CellCollection::count
     * @covers Svetel\Library\Collections\CellCollection::getColumnIndex
     */
    public function testAddCell()
    {
        $cellList = [
            'firstValue'
            , 'secondValue'
            , 'thirdValue'
        ];
        $cellLength = sizeof($cellList);
        while ($cellLength--) {
            foreach ($cellList as $cellKey => $cellValue) {
                $this->object->addCell($cellKey, $cellValue);
            }
        }
        $this->assertTrue(is_array($this->object->getCollection()));
        $this->assertEquals(
                sizeof($cellList)
                , sizeof($this->object->getCollection())
        );
        $this->assertEquals(
                sizeof($cellList)
                , $this->object->count()
        );
        $this->assertEquals(
                array_search('thirdValue', $cellList)
                , $this->object->getColumnIndex('thirdValue')
        );
        $this->assertEquals(
                10
                , array_search(
                        '10 Value'
                        , $this->object->addCell(10, '10 Value')
                                ->getCollection()
                )
        );
    }

}

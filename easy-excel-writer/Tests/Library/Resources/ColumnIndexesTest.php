<?php

namespace Svetel\Library\Resources;

class ColumnIndexesTest extends \PHPUnit_Framework_TestCase
{

    /**
     * @var ColumnIndexes
     */
    protected $object;

    /**
     * Sets up the fixture, for example, opens a network connection.
     * This method is called before a test is executed.
     */
    protected function setUp()
    {
        $this->object = ColumnIndexes::getInstance();
    }

    /**
     * @covers Svetel\Library\Resources\ColumnIndexes::getInstance
     */
    public function testGetInstance()
    {
        $this->assertTrue(ColumnIndexes::getInstance() instanceof ColumnIndexes);
    }

    /**
     * @covers Svetel\Library\Resources\ColumnIndexes::getColumnStringIndexFromNumber
     */
    public function testGetColumnStringIndexFromNumber()
    {
        $abcAssertions = [
            'A' => 0
            , 'AA' => 26
            , 'BA' => 52
            , 'CA' => 78
            , 'DA' => 104
            , 'ZA' => 676
            , 'ZZ' => 701
            , 'ZZ' => 701
            , 'AAA' => 702
        ];
        foreach ($abcAssertions as $expected => $value) {
            $this->assertEquals(
                    $expected
                    , $this->object->getColumnStringIndexFromNumber($value)
            );
        }
    }

}

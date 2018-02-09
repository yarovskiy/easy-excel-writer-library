<?php

/*
 * LICENSE: This source file is subject to version 3.01 of the PHP license
 * that is available through the world-wide-web at the following URI:
 * http://www.php.net/license/3_01.txt.  If you did not receive a copy of
 * the PHP License and are unable to obtain it through the web, please
 * send a note to license@php.net so we can mail you a copy immediately.
 */

namespace Svetel\Library\Collections;

/**
 * Cell list in Excel row
 *
 * @author Vladimir Oleynikov <olejnikov.vladimir@gmail.com>
 */
class CellCollection
{

    /**
     * @var array
     */
    private $collection;
    /**
     * @param integer $columnIndex
     * @param string $columnValue
     * @return \Svetel\Library\Collections\CellCollection
     */
    public function addCell($columnIndex, $columnValue)
    {
        $this->collection[$columnIndex] = $columnValue;
        ksort($this->collection);
        return $this;
    }
    /**
     * @return array
     */
    public function getCollection()
    {
        return $this->collection;
    }

    /**
     * @param string $columnValue
     * @return integer
     */
    public function getColumnIndex($columnValue)
    {
        return array_search($columnValue, $this->collection);
    }

    /**
     * @return integer
     */
    public function count()
    {
        return count($this->collection);
    }

}

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
 * Description of RowCollection
 *
 * @author Vladimir Oleynikov <olejnikov.vladimir@gmail.com>
 */
class RowCollection
{

    /**
     * @var CellCollection[]
     */
    private $collection;

    /**
     * @param \Svetel\Library\Collections\CellCollection $cellCollection
     * @return \Svetel\Library\Collections\RowCollection
     */
    public function addRow(CellCollection $cellCollection)
    {
        $this->collection[] = $cellCollection;

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
     * @return integer
     */
    public function count()
    {
        return sizeof($this->collection);
    }

}

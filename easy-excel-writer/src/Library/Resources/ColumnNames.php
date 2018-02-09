<?php

/*
 * LICENSE: This source file is subject to version 3.01 of the PHP license
 * that is available through the world-wide-web at the following URI:
 * http://www.php.net/license/3_01.txt.  If you did not receive a copy of
 * the PHP License and are unable to obtain it through the web, please
 * send a note to license@php.net so we can mail you a copy immediately.
 */

namespace Svetel\Library\Resources;

/**
 * Description of ExcelColumns
 *
 * @author Vladimir Oleynikov <olejnikov.vladimir@gmail.com>
 */
class ColumnNames
{

    /**
     * @var array
     */
    private $columns;

    /**
     * @var bool
     */
    private $sort = FALSE;

    public function __construct(array $columns = [])
    {
        $this->columns = $columns;
    }

    /**
     * Add column to stack of exitsts columns
     * Create value of last column in stack incremented by 1
     * If Array is empty create column with value 0
     * @param string $columnName
     * @return \Svetel\Library\Resources\ColumnNames
     */
    public function addColumn($columnName)
    {
        if (empty($this->columns)) {
            $this->columns[$columnName] = 0;
        } elseif (!isset($this->columns[$columnName])) {
            $this->columns[$columnName] = end($this->columns) + 1;
        }
        return $this;
    }
    /**
     * @return \Svetel\Library\Resources\ColumnNames
     */
    private function abcSortColumnNames()
    {
        $counter = 1;
        ksort($this->columns);
        foreach ($this->columns as $columnKey => $columnValue) {
            $this->columns[$columnKey] = $counter;
            $counter++;
        }
        return $this;
    }

    /**
     * @return array Names of columns
     */
    public function getColumns()
    {
        foreach ($this->columns as $key => $value) {
            $columns[$value] = $key;
        }
        return $columns;
    }
    /**
     * @param string $columnKey
     * @return boolean|string
     */
    public function getColumnValue($columnKey)
    {
        if (!isset($this->columns[$columnKey])) {
            return FALSE;
        }
        return $this->columns[$columnKey];
    }
    /**
     * @return bool
     */
    public function getSort()
    {
        return $this->sort;
    }

    /**
     * 
     * @param bool $sort
     * @return \Svetel\Library\Resources\ColumnNames
     */
    public function setSort($sort)
    {
        $this->sort = $sort;
        return $this;
    }

    /**
     * @return integer
     */
    public function count()
    {
        return sizeof($this->columns);
    }

}

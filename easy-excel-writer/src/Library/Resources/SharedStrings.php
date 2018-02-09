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
 * Description of SharedStrings
 *
 * @author Vladimir Oleynikov <olejnikov.vladimir@gmail.com>
 */
class SharedStrings
{

    private $stringCollection;

    /**
     * 
     * @param string $param
     * @return int  Key of XML string alias
     */
    public function addCollection($param)
    {
        $this->stringCollection[] = $param;

        return ($this->count() - 1);
    }

    public function count()
    {
        return count($this->stringCollection);
    }

    public function getCollection()
    {
        return $this->stringCollection;
    }

}

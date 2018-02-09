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
 * ColumnIndexes
 *
 * @author Vladimir Oleynikov <olejnikov.vladimir@gmail.com>
 */
class ColumnIndexes
{

    private static $instance;



    public static function getInstance()
    {
        if (!self::$instance instanceof self) {
            self::$instance = new self;
        }
        return self::$instance;
    }

    private function __construct()
    {
        
    }
    private function __clone()
    {
        
    }

    /**
     * @param integer $columnNumber maximum 1024 columns
     * @return string
     */
    public function getColumnStringIndexFromNumber($columnNumber)
    {
        if (26 > $columnNumber) {
            return chr(65 + $columnNumber);
        } elseif (702 > $columnNumber) {//ZZ
            return chr(64 + ($columnNumber / 26))
                    . chr(65 + $columnNumber % 26);
        }
        if (1024 < $columnNumber) {
            throw new \InvalidArgumentException('Maximum 1024 columns');
        }
        return chr(64 + (($columnNumber - 26) / 676))
                . chr(65 + ((($columnNumber - 26) % 676) / 26))
                . chr(65 + $columnNumber % 26);
    }

}

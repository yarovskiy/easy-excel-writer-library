<?php

/*
 * LICENSE: This source file is subject to version 3.01 of the PHP license
 * that is available through the world-wide-web at the following URI:
 * http://www.php.net/license/3_01.txt.  If you did not receive a copy of
 * the PHP License and are unable to obtain it through the web, please
 * send a note to license@php.net so we can mail you a copy immediately.
 */

namespace Svetel\Library;

/**
 * Description of EasyExcelWriter
 *
 * @author Vladimir Oleynikov <olejnikov.vladimir@gmail.com>
 */
use Svetel\Library\Resources\ColumnNames;
use Svetel\Library\Resources\ColumnIndexes;
use Svetel\Library\Resources\SharedStrings;
use Svetel\Library\Collections\CellCollection;
use Svetel\Library\Collections\RowCollection;
use XMLWriter;

class EasyExcelWriter
{

    /**
     * @var CellCollection
     */
    private $columnNames;

    /**
     * @var string
     */
    private $fileName;

    /**
     * @var ColumnIndexes
     */
    private $columnIndexes;

    /**
     * @var RowCollection
     */
    private $rowCollection;

    /**
     * @var SharedStrings
     */
    private $sharedStringCollection;

    /**
     * @param ColumnNames $excelColumns
     */
    public function __construct()
    {
        $this->columnIndexes = ColumnIndexes::getInstance();
        $this->rowCollection = new RowCollection;
    }

    /**
     * @param ColumnNames $excelColumns
     * @return CellCollection
     */
    public function setColumnNames(ColumnNames $excelColumns)
    {
        $columnNameListRow = new CellCollection;
        foreach ($excelColumns->getColumns() as $keyColumn => $valueColumn) {
            $columnNameListRow->addCell($keyColumn, $valueColumn);
        }
        $this->columnNames = $columnNameListRow;
        $this->rowCollection->addRow($columnNameListRow);
        return $this;
    }

    /**
     * @param array $row
     * @return \Svetel\Library\EasyExcelWriter
     */
    public function addRow(array $row)
    {
        $cellCollection = new CellCollection;
        if ($this->columnNames) {
            foreach ($row as $columnName => $columnValue) {
                if (is_array($columnValue) && !empty($columnValue)) {
                    foreach ($columnValue as $keyArray => $valueArray) {
                        $cellCollection->addCell(
                                $this->columnNames
                                        ->getColumnIndex($keyArray), $valueArray
                        );
                    }
                } elseif (!empty($columnValue)) {
                    $cellCollection->addCell(
                            $this->columnNames->getColumnIndex($columnName)
                            , $columnValue
                    );
                }
            }
        } else {
            foreach ($row as $columnName => $columnValue) {
                if (is_array($columnValue) && !empty($columnValue)) {
                    foreach ($columnValue as $keyArray => $valueArray) {
                        $cellCollection->addCell(
                                $keyArray, $valueArray
                        );
                    }
                } elseif (!empty($columnValue)) {
                    $cellCollection->addCell(
                            $columnName
                            , $columnValue
                    );
                }
            }
        }
        $this->rowCollection->addRow($cellCollection);

        return $this;
    }

    /**
     * @return string
     */
    public function getFileName()
    {
        return $this->fileName;
    }

    /**
     * 
     * @param string $fileName
     * @return \Svetel\Library\EasyExcelWriter
     */
    public function setFileName($fileName)
    {
        $this->fileName = $fileName;
        return $this;
    }

    /**
     * @return XMLWriter
     */
    private function generateXMLHeaderRow()
    {
        $xml = new XMLWriter;
        $xml->openMemory();
        $xml->setIndent(TRUE);
        $xml->startDocument('1.0', 'UTF-8', 'yes');
        $xml->startElement('sst'); //sst
        $xml->writeAttribute('xmlns'
                , 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
        ); //sst
        $xml->writeAttribute(
                'uniqueCount'
                , $this->sharedStringCollection->count()
        ); //sst
        foreach ($this->sharedStringCollection->getCollection() as $keyColumn => $valueColumn) {
            $valueColumn = preg_replace_callback('/.*\/\*.+\*\/(.+)/i', function($matches) {
                return trim($matches[1]);
            }, $valueColumn);
            $valueColumn = iconv('Windows-1252', 'ASCII//utf-8//IGNORE', $valueColumn);
            $xml->startElement('si'); //sst > si
            $xml->startElement('t'); //sst > si > t
            $xml->writeRaw(htmlentities($valueColumn)); //sst > si > t
            $xml->endElement(); //sst > si > t
            $xml->endElement(); //sst > si
        }
        $xml->endElement(); //sst
        return $xml;
    }

    private function generateXMLRows()
    {
        $xml = new XMLWriter;
        $xml->openMemory();
        $xml->setIndent(TRUE);
        $xml->startDocument('1.0', 'UTF-8', 'yes');
        $xml->startElement('worksheet'); //worksheet
        $xml->writeAttribute('xml:space', 'preserve'); //worksheet
        $xml->writeAttribute(
                'xmlns'
                , 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
        ); //worksheet
        $xml->writeAttribute(
                'xmlns:r'
                , 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
        ); //worksheet
        $xml->startElement('sheetPr'); //worksheet > sheetPr
        $xml->startElement('outlinePr'); //worksheet > sheetPr > outlinePr
        $xml->writeAttribute('summaryBelow', 1); //worksheet > sheetPr > outlinePr
        $xml->writeAttribute('summaryRight', 1); //worksheet > sheetPr > outlinePr
        $xml->endElement(); //worksheet > sheetPr
        $xml->endElement(); //worksheet
        $xml->startElement('dimension'); //worksheet > dimension
        /**
         * @todo <dimension ref="A1:K51"/>
         */
        $xml->writeAttribute('ref', 'A1:K' . $this->rowCollection->count()); //worksheet > dimension
        $xml->endElement(); //worksheet > dimension

        $xml->startElement('sheetViews'); //worksheet > sheetViews
        $xml->startElement('sheetView'); //worksheet > sheetViews > sheetView
        $xml->writeAttribute('tabSelected', 1); //worksheet > sheetViews > sheetView
        $xml->writeAttribute('workbookViewId', 0); //worksheet > sheetViews > sheetView
        $xml->writeAttribute('showGridLines', 'true'); //worksheet > sheetViews > sheetView
        $xml->writeAttribute('showRowColHeaders', 1); //worksheet > sheetViews > sheetView
        $xml->endElement(); //worksheet > sheetViews > sheetView
        $xml->endElement(); //worksheet > sheetViews
        $xml->startElement('sheetFormatPr'); //worksheet > sheetFormatPr
        $xml->writeAttribute('defaultRowHeight', '12.75'); //worksheet > sheetFormatPr
        $xml->writeAttribute('outlineLevelRow', '0'); //worksheet > sheetFormatPr
        $xml->writeAttribute('outlineLevelCol', '0'); //worksheet > sheetFormatPr
        $xml->endElement(); //worksheet > sheetFormatPr
        $xml->startElement('sheetData'); //worksheet > sheetData
        /**
         * @todo implement rows
         */
        $this->sharedStringCollection = new SharedStrings;
        /* @var $collection CellCollection */
        foreach ($this->rowCollection->getCollection() as $keyCollection => $collection) {
            $xml->startElement('row'); //worksheet > sheetData > row
            $xml->writeAttribute('r', ($keyCollection + 1)); //worksheet > sheetData > row
            $xml->writeAttribute('spans', '1:' . $collection->count()); //worksheet > sheetData > row
            foreach ($collection->getCollection() as $columnIntIndex => $columnValue) {
                if (!$columnValue) {
                    continue;
                }
                $xml->startElement('c'); //worksheet > sheetData > row > c
                $xml->writeAttribute(
                        'r'
                        , $this->columnIndexes
                                ->getColumnStringIndexFromNumber($columnIntIndex)
                        . ($keyCollection + 1)
                ); //worksheet > sheetData > row > c
                $xml->writeAttribute('t', 's'); //worksheet > sheetData > row > c
                $xml->startElement('v'); //worksheet > sheetData > row > c > v
                $xml->writeRaw(
                        $this->sharedStringCollection
                                ->addCollection(
                                        iconv(
                                                'windows-1251'
                                                , 'utf-8//IGNORE'
                                                , $columnValue
                                        )
                                )
                ); //worksheet > sheetData > row > c > v
                $xml->endElement(); //worksheet > sheetData > row > c > v
                $xml->endElement(); //worksheet > sheetData > row > c
            }
            $xml->endElement(); //worksheet > sheetData > row
        }
        $xml->endElement(); //worksheet > sheetData
        $xml->startElement('sheetProtection'); //worksheet > sheetProtection
        $xml->writeAttribute('sheet', "false"); //worksheet > sheetProtection
        $xml->writeAttribute('objects', "false"); //worksheet > sheetProtection
        $xml->writeAttribute('scenarios', "false"); //worksheet > sheetProtection
        $xml->writeAttribute('formatCells', "false"); //worksheet > sheetProtection
        $xml->writeAttribute('formatColumns', "false"); //worksheet > sheetProtection
        $xml->writeAttribute('formatRows', "false"); //worksheet > sheetProtection
        $xml->writeAttribute('insertColumns', "false"); //worksheet > sheetProtection
        $xml->writeAttribute('insertRows', "false"); //worksheet > sheetProtection
        $xml->writeAttribute('insertHyperlinks', "false"); //worksheet > sheetProtection
        $xml->writeAttribute('deleteColumns', "false"); //worksheet > sheetProtection
        $xml->writeAttribute('deleteRows', "false"); //worksheet > sheetProtection
        $xml->writeAttribute('selectLockedCells', "false"); //worksheet > sheetProtection
        $xml->writeAttribute('sort', "false"); //worksheet > sheetProtection
        $xml->writeAttribute('autoFilter', "false"); //worksheet > sheetProtection
        $xml->writeAttribute('pivotTables', "false"); //worksheet > sheetProtection
        $xml->writeAttribute('selectUnlockedCells', "false"); //worksheet > sheetProtection
        $xml->endElement(); //worksheet > sheetProtection
        $xml->startElement('printOptions'); //worksheet > printOptions
        $xml->writeAttribute('gridLines', 'false'); //worksheet > printOptions
        $xml->writeAttribute('gridLinesSet', 'true'); //worksheet > printOptions
        $xml->endElement(); //worksheet > printOptions
        $xml->startElement('pageMargins'); //worksheet > pageMargins
        $xml->writeAttribute('left', '0.7'); //worksheet > pageMargins
        $xml->writeAttribute('right', '0.7'); //worksheet > pageMargins
        $xml->writeAttribute('top', '0.75'); //worksheet > pageMargins
        $xml->writeAttribute('bottom', '0.75'); //worksheet > pageMargins
        $xml->writeAttribute('header', '0.3'); //worksheet > pageMargins
        $xml->writeAttribute('footer', '0.3'); //worksheet > pageMargins
        $xml->endElement(); //worksheet > pageMargins
        $xml->startElement('pageSetup'); //worksheet > pageSetup
        $xml->writeAttribute('paperSize', '1'); //worksheet > pageSetup
        $xml->writeAttribute('orientation', 'default'); //worksheet > pageSetup
        $xml->writeAttribute('scale', '100'); //worksheet > pageSetup
        $xml->writeAttribute('fitToHeight', '1'); //worksheet > pageSetup
        $xml->writeAttribute('fitToWidth', '1'); //worksheet > pageSetup
        $xml->endElement(); //worksheet > pageSetup
        $xml->startElement('headerFooter'); //worksheet > headerFooter
        $xml->writeAttribute('differentOddEven', 'false'); //worksheet > headerFooter
        $xml->writeAttribute('differentFirst', 'false'); //worksheet > headerFooter
        $xml->writeAttribute('scaleWithDoc', 'true'); //worksheet > headerFooter
        $xml->writeAttribute('alignWithMargins', 'true'); //worksheet > headerFooter
        $xml->startElement('oddHeader'); //worksheet > headerFooter > oddHeader
        $xml->writeRaw(''); //worksheet > headerFooter > oddHeader
        $xml->endElement(); //worksheet > headerFooter > oddHeader
        $xml->startElement('oddFooter'); //worksheet > headerFooter > oddFooter
        $xml->writeRaw(''); //worksheet > headerFooter > oddFooter
        $xml->endElement(); //worksheet > headerFooter > oddFooter
        $xml->startElement('evenHeader'); //worksheet > headerFooter > evenHeader
        $xml->writeRaw(''); //worksheet > headerFooter > evenHeader
        $xml->endElement(); //worksheet > headerFooter > evenHeader
        $xml->startElement('evenFooter'); //worksheet > headerFooter > evenFooter
        $xml->writeRaw(''); //worksheet > headerFooter > evenFooter
        $xml->endElement(); //worksheet > headerFooter > evenFooter
        $xml->startElement('firstHeader'); //worksheet > headerFooter > firstHeader
        $xml->writeRaw(''); //worksheet > headerFooter > firstHeader
        $xml->endElement(); //worksheet > headerFooter > firstHeader
        $xml->startElement('firstFooter'); //worksheet > headerFooter > firstFooter
        $xml->writeRaw(''); //worksheet > headerFooter > firstFooter
        $xml->endElement(); //worksheet > headerFooter > firstFooter
        $xml->endElement(); //worksheet > headerFooter
        $xml->endElement(); //worksheet 
        return $xml;
    }

    private function generateRelationshipsXML()
    {
        $xml = new XMLWriter;
        $xml->openMemory();
        $xml->setIndent(TRUE);
        $xml->startDocument('1.0', 'UTF-8', 'yes');
        $xml->startElement('Relationships'); //Relationships
        $xml->writeAttribute(
                'xmlns'
                , 'http://schemas.openxmlformats.org/package/2006/relationships'
        ); //Relationships

        $xml->startElement('Relationship'); //Relationships > Relationship
        $xml->writeAttribute('Id', 'rId3'); //Relationships > Relationship
        $xml->writeAttribute(
                'Type'
                , 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties'
        ); //Relationships > Relationship
        $xml->writeAttribute('Target', 'docProps/app.xml'); //Relationships > Relationship
        $xml->endElement(); //Relationships > Relationship
        $xml->startElement('Relationship'); //Relationships > Relationship
        $xml->writeAttribute('Id', 'rId2'); //Relationships > Relationship
        $xml->writeAttribute(
                'Type'
                , 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties'
        ); //Relationships > Relationship
        $xml->writeAttribute('Target', 'docProps/core.xml'); //Relationships > Relationship
        $xml->endElement(); //Relationships > Relationship
        $xml->startElement('Relationship'); //Relationships > Relationship
        $xml->writeAttribute('Id', 'rId1'); //Relationships > Relationship
        $xml->writeAttribute(
                'Type'
                , 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument'
        ); //Relationships > Relationship
        $xml->writeAttribute('Target', 'xl/workbook.xml'); //Relationships > Relationship
        $xml->endElement(); //Relationships > Relationship
        $xml->endElement(); //Relationships
        return $xml;
    }

    public function generateContentTypesXML()
    {
        $xml = new XMLWriter;
        $xml->openMemory();
        $xml->setIndent(TRUE);
        $xml->startDocument('1.0', 'UTF-8', 'yes');
        $xml->startElement('Types'); //Types
        $xml->writeAttribute(
                'xmlns'
                , 'http://schemas.openxmlformats.org/package/2006/content-types'
        ); //Types
        $xml->startElement('Default'); //Types > Default
        $xml->writeAttribute('Extension', 'bin'); //Types > Default
        $xml->writeAttribute(
                'ContentType', 'application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings'
        ); //Types > Default
        $xml->endElement(); //Types > Default
        $xml->startElement('Default'); //Types > Default
        $xml->writeAttribute('Extension', 'rels'); //Types > Default
        $xml->writeAttribute(
                'ContentType'
                , 'application/vnd.openxmlformats-package.relationships+xml'
        ); //Types > Default
        $xml->endElement(); //Types > Default
        $xml->startElement('Default'); //Types > Default
        $xml->writeAttribute('Extension', 'xml'); //Types > Default
        $xml->writeAttribute('ContentType', 'application/xml'); //Types > Default
        $xml->endElement(); //Types > Default
        $xml->startElement('Override'); //Types > Override
        $xml->writeAttribute('PartName', '/xl/workbook.xml'); //Types > Override
        $xml->writeAttribute(
                'ContentType'
                , 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml'
        ); //Types > Override
        $xml->endElement(); //Types > Override
        $xml->startElement('Override'); //Types > Override
        $xml->writeAttribute('PartName', '/xl/worksheets/sheet1.xml'); //Types > Override
        $xml->writeAttribute(
                'ContentType'
                , 'application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml'
        ); //Types > Override
        $xml->endElement(); //Types > Override
        $xml->startElement('Override'); //Types > Override
        $xml->writeAttribute('PartName', '/xl/theme/theme1.xml'); //Types > Override
        $xml->writeAttribute(
                'ContentType'
                , 'application/vnd.openxmlformats-officedocument.theme+xml'
        ); //Types > Override
        $xml->endElement(); //Types > Override
        $xml->startElement('Override'); //Types > Override
        $xml->writeAttribute('PartName', '/xl/styles.xml'); //Types > Override
        $xml->writeAttribute(
                'ContentType'
                , 'application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml'
        ); //Types > Override
        $xml->endElement(); //Types > Override
        $xml->startElement('Override'); //Types > Override
        $xml->writeAttribute('PartName', '/xl/sharedStrings.xml'); //Types > Override
        $xml->writeAttribute(
                'ContentType'
                , 'application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml'
        ); //Types > Override
        $xml->endElement(); //Types > Override


        $xml->startElement('Override'); //Types > Override
        $xml->writeAttribute('PartName', '/docProps/app.xml'); //Types > Override
        $xml->writeAttribute(
                'ContentType'
                , 'application/vnd.openxmlformats-officedocument.extended-properties+xml'
        ); //Types > Override
        $xml->endElement(); //Types > Override
        $xml->startElement('Override'); //Types > Override
        $xml->writeAttribute('PartName', '/docProps/core.xml'); //Types > Override
        $xml->writeAttribute(
                'ContentType'
                , 'application/vnd.openxmlformats-package.core-properties+xml'
        ); //Types > Override
        $xml->endElement(); //Types > Override
        $xml->endElement(); //Types
        return $xml;
    }

    private function generateRefWorkbookXML()
    {
        $xml = new XMLWriter;
        $xml->openMemory();
        $xml->setIndent(TRUE);
        $xml->startDocument('1.0', 'UTF-8', 'yes');
        $xml->startElement('Relationships'); //Relationships
        $xml->writeAttribute(
                'xmlns'
                , 'http://schemas.openxmlformats.org/package/2006/relationships'
        ); //Relationships
        $xml->startElement('Relationship'); //Relationships > Relationship
        $xml->writeAttribute('Id', 'rId1'); //Relationships > Relationship
        $xml->writeAttribute(
                'Type'
                , 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles'
        ); //Relationships > Relationship
        $xml->writeAttribute('Target', 'styles.xml'); //Relationships > Relationship
        $xml->endElement(); //Relationships > Relationship
        $xml->startElement('Relationship'); //Relationships > Relationship
        $xml->writeAttribute('Id', 'rId2'); //Relationships > Relationship
        $xml->writeAttribute(
                'Type'
                , 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme'
        ); //Relationships > Relationship
        $xml->writeAttribute('Target', 'theme/theme1.xml'); //Relationships > Relationship
        $xml->endElement(); //Relationships > Relationship
        $xml->startElement('Relationship'); //Relationships > Relationship
        $xml->writeAttribute('Id', 'rId3'); //Relationships > Relationship
        $xml->writeAttribute(
                'Type'
                , 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings'
        ); //Relationships > Relationship
        $xml->writeAttribute('Target', 'sharedStrings.xml'); //Relationships > Relationship
        $xml->endElement(); //Relationships > Relationship
        $xml->startElement('Relationship'); //Relationships > Relationship
        $xml->writeAttribute('Id', 'rId4'); //Relationships > Relationship
        $xml->writeAttribute(
                'Type'
                , 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet'
        ); //Relationships > Relationship
        $xml->writeAttribute('Target', 'worksheets/sheet1.xml'); //Relationships > Relationship
        $xml->endElement(); //Relationships > Relationship
        $xml->endElement(); //Relationships
        return $xml;
    }

    private function generateCoreXML()
    {
        $xml = new XMLWriter;
        $xml->openMemory();
        $xml->setIndent(TRUE);
        $xml->startDocument('1.0', 'UTF-8', 'yes');
        $xml->startElement('cp:coreProperties'); //cp:coreProperties
        $xml->writeAttribute(
                'xmlns:cp'
                , 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties'
        ); //cp:coreProperties
        $xml->writeAttribute(
                'xmlns:dc'
                , 'http://purl.org/dc/elements/1.1/'
        ); //cp:coreProperties
        $xml->writeAttribute(
                'xmlns:dcterms'
                , 'http://purl.org/dc/terms/'
        ); //cp:coreProperties
        $xml->writeAttribute(
                'xmlns:dcmitype'
                , 'http://purl.org/dc/dcmitype/'
        ); //cp:coreProperties
        $xml->writeAttribute(
                'xmlns:xsi'
                , 'http://www.w3.org/2001/XMLSchema-instance'
        ); //cp:coreProperties
        $xml->startElement('dc:creator'); //cp:coreProperties > dc:creator
        $xml->writeRaw('Svetel'); //cp:coreProperties > dc:creator
        $xml->endElement(); //cp:coreProperties > dc:creator
        $xml->startElement('cp:lastModifiedBy'); //cp:coreProperties > cp:lastModifiedBy
        $xml->writeRaw('Svetel'); //cp:coreProperties > cp:lastModifiedBy
        $xml->endElement(); //cp:coreProperties > cp:lastModifiedBy
        $xml->startElement('dcterms:created'); //cp:coreProperties > dcterms:created
        $xml->writeAttribute('xsi:type', 'dcterms:W3CDTF'); //cp:coreProperties > dcterms:created
        $xml->writeRaw((new \DateTime)->format('c')); //cp:coreProperties > dcterms:created
        $xml->endElement(); //cp:coreProperties > dcterms:created
        $xml->startElement('dcterms:modified'); //cp:coreProperties > dcterms:modified
        $xml->writeAttribute('xsi:type', 'dcterms:W3CDTF'); //cp:coreProperties > dcterms:modified
        $xml->writeRaw((new \DateTime)->format('c')); //cp:coreProperties > dcterms:modified
        $xml->endElement(); //cp:coreProperties > dcterms:modified
        $xml->endElement(); //cp:coreProperties
        return $xml;
    }

    private function generateWorkbookXML()
    {
        $xml = new XMLWriter;
        $xml->openMemory();
        $xml->setIndent(TRUE);
        $xml->startDocument('1.0', 'UTF-8', 'yes');
        $xml->startElement('workbook'); //workbook
        $xml->writeAttribute('xml:space', 'preserve'); //workbook
        $xml->writeAttribute(
                'xmlns'
                , 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
        ); //workbook
        $xml->writeAttribute(
                'xmlns:r'
                , 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
        ); //workbook
        $xml->startElement('fileVersion'); //workbook > fileVersion
        $xml->writeAttribute('appName', 'xl'); //workbook > fileVersion
        $xml->writeAttribute('lastEdited', '4'); //workbook > fileVersion
        $xml->writeAttribute('lowestEdited', '4'); //workbook > fileVersion
        $xml->writeAttribute('rupBuild', '4505'); //workbook > fileVersion
        $xml->endElement(); //workbook > fileVersion
        $xml->startElement('workbookPr'); //workbook > workbookPr
        $xml->writeAttribute('codeName', 'ThisWorkbook'); //workbook > workbookPr
        $xml->endElement(); //workbook > workbookPr
        $xml->startElement('bookViews'); //workbook > bookViews
        $xml->startElement('workbookView'); //workbook > bookViews > workbookView
        $xml->writeAttribute('activeTab', '0'); //workbook > bookViews > workbookView
        $xml->writeAttribute('autoFilterDateGrouping', '1'); //workbook > bookViews > workbookView
        $xml->writeAttribute('firstSheet', '0'); //workbook > bookViews > workbookView
        $xml->writeAttribute('minimized', '0'); //workbook > bookViews > workbookView
        $xml->writeAttribute('showHorizontalScroll', '1'); //workbook > bookViews > workbookView
        $xml->writeAttribute('showSheetTabs', '1'); //workbook > bookViews > workbookView
        $xml->writeAttribute('showVerticalScroll', '1'); //workbook > bookViews > workbookView
        $xml->writeAttribute('tabRatio', '600'); //workbook > bookViews > workbookView
        $xml->writeAttribute('visibility', 'visible'); //workbook > bookViews > workbookView
        $xml->endElement(); //workbook > bookViews > workbookView
        $xml->endElement(); //workbook > bookViews
        $xml->startElement('sheets'); //workbook > sheets
        $xml->startElement('sheet'); //workbook > sheets > sheet
        $xml->writeAttribute('name', 'Worksheet'); //workbook > sheets > sheet
        $xml->writeAttribute('sheetId', '1'); //workbook > sheets > sheet
        $xml->writeAttribute('r:id', 'rId4'); //workbook > sheets > sheet
        $xml->endElement(); //workbook > sheets > sheet
        $xml->endElement(); //workbook > sheets
        $xml->startElement('definedNames'); //workbook > definedNames
        $xml->endElement(); //workbook > definedNames
        $xml->startElement('calcPr'); //workbook > calcPr
        $xml->writeAttribute('calcId', '124519'); //workbook > calcPr
        $xml->writeAttribute('calcMode', 'auto'); //workbook > calcPr
        $xml->writeAttribute('fullCalcOnLoad', '1'); //workbook > calcPr
        $xml->endElement(); //workbook > calcPr
        $xml->endElement(); //workbook

        return $xml;
    }

    private function generateThemeXML()
    {
        $xml = new XMLWriter;
        $xml->openMemory();
        $xml->setIndent(TRUE);
        $xml->startDocument('1.0', 'UTF-8', 'yes');
        $xml->startElement('a:theme'); //a:theme
        $xml->writeAttribute(
                'xmlns:a'
                , 'http://schemas.openxmlformats.org/drawingml/2006/main'
        ); //a:theme
        $xml->writeAttribute('name', 'Office Theme'); //a:theme
        $xml->startElement('a:themeElements'); //a:theme > a:themeElements
        $xml->startElement('a:clrScheme'); //a:theme > a:themeElements > a:clrScheme
        $xml->writeAttribute('name', 'Office'); //a:theme > a:themeElements > a:clrScheme
        $xml->startElement('a:dk1'); //a:theme > a:themeElements > a:clrScheme > a:dk1
        $xml->startElement('a:sysClr'); //a:theme > a:themeElements > a:clrScheme > a:dk1 > a:sysClr
        $xml->writeAttribute('val', 'windowText'); //a:theme > a:themeElements > a:clrScheme > a:dk1 > a:sysClr
        $xml->writeAttribute('lastClr', '000000'); //a:theme > a:themeElements > a:clrScheme > a:dk1 > a:sysClr
        $xml->endElement(); //a:theme > a:themeElements > a:clrScheme > a:dk1 > a:sysClr
        $xml->endElement(); //a:theme > a:themeElements > a:clrScheme > a:dk1 
        $xml->startElement('a:lt1'); //a:theme > a:themeElements > a:clrScheme > a:lt1
        $xml->startElement('a:sysClr'); //a:theme > a:themeElements > a:clrScheme > a:lt1 > a:sysClr
        $xml->writeAttribute('val', 'window'); //a:theme > a:themeElements > a:clrScheme > a:lt1 > a:sysClr
        $xml->writeAttribute('lastClr', 'FFFFFF'); //a:theme > a:themeElements > a:clrScheme > a:lt1 > a:sysClr
        $xml->endElement(); //a:theme > a:themeElements > a:clrScheme > a:lt1 > a:sysClr
        $xml->endElement(); //a:theme > a:themeElements > a:clrScheme > a:lt1 
        $xml->startElement('a:dk2'); //a:theme > a:themeElements > a:clrScheme > a:dk2
        $xml->startElement('a:srgbClr'); //a:theme > a:themeElements > a:clrScheme > a:dk2 > a:srgbClr
        $xml->writeAttribute('val', '1F497D'); //a:theme > a:themeElements > a:clrScheme > a:dk2 > a:srgbClr
        $xml->endElement(); //a:theme > a:themeElements > a:clrScheme > a:dk2 > a:srgbClr
        $xml->endElement(); //a:theme > a:themeElements > a:clrScheme > a:dk2 
        $xml->startElement('a:lt2'); //a:theme > a:themeElements > a:clrScheme > a:lt2
        $xml->startElement('a:srgbClr'); //a:theme > a:themeElements > a:clrScheme > a:lt2 > a:srgbClr
        $xml->writeAttribute('val', 'EEECE1'); //a:theme > a:themeElements > a:clrScheme > a:lt2 > a:srgbClr
        $xml->endElement(); //a:theme > a:themeElements > a:clrScheme > a:lt2 > a:srgbClr
        $xml->endElement(); //a:theme > a:themeElements > a:clrScheme > a:lt2
        $xml->startElement('a:accent1'); //a:theme > a:themeElements > a:clrScheme > a:accent1
        $xml->startElement('a:srgbClr'); //a:theme > a:themeElements > a:clrScheme > a:accent1 > a:srgbClr
        $xml->writeAttribute('val', '4F81BD'); //a:theme > a:themeElements > a:clrScheme > a:accent1 > a:srgbClr
        $xml->endElement(); //a:theme > a:themeElements > a:clrScheme > a:accent1 > a:srgbClr
        $xml->endElement(); //a:theme > a:themeElements > a:clrScheme > a:accent1
        $xml->startElement('a:accent2'); //a:theme > a:themeElements > a:clrScheme > a:accent2
        $xml->startElement('a:srgbClr'); //a:theme > a:themeElements > a:clrScheme > a:accent2 > a:srgbClr
        $xml->writeAttribute('val', 'C0504D'); //a:theme > a:themeElements > a:clrScheme > a:accent2 > a:srgbClr
        $xml->endElement(); //a:theme > a:themeElements > a:clrScheme > a:accent2 > a:srgbClr
        $xml->endElement(); //a:theme > a:themeElements > a:clrScheme > a:accent2
        $xml->startElement('a:accent3'); //a:theme > a:themeElements > a:clrScheme > a:accent3
        $xml->startElement('a:srgbClr'); //a:theme > a:themeElements > a:clrScheme > a:accent3 > a:srgbClr
        $xml->writeAttribute('val', '9BBB59'); //a:theme > a:themeElements > a:clrScheme > a:accent3 > a:srgbClr
        $xml->endElement(); //a:theme > a:themeElements > a:clrScheme > a:accent3 > a:srgbClr
        $xml->endElement(); //a:theme > a:themeElements > a:clrScheme > a:accent3
        $xml->startElement('a:accent4'); //a:theme > a:themeElements > a:clrScheme > a:accent4
        $xml->startElement('a:srgbClr'); //a:theme > a:themeElements > a:clrScheme > a:accent4 > a:srgbClr
        $xml->writeAttribute('val', '8064A2'); //a:theme > a:themeElements > a:clrScheme > a:accent4 > a:srgbClr
        $xml->endElement(); //a:theme > a:themeElements > a:clrScheme > a:accent4 > a:srgbClr
        $xml->endElement(); //a:theme > a:themeElements > a:clrScheme > a:accent4
        $xml->startElement('a:accent5'); //a:theme > a:themeElements > a:clrScheme > a:accent5
        $xml->startElement('a:srgbClr'); //a:theme > a:themeElements > a:clrScheme > a:accent5 > a:srgbClr
        $xml->writeAttribute('val', '4BACC6'); //a:theme > a:themeElements > a:clrScheme > a:accent5 > a:srgbClr
        $xml->endElement(); //a:theme > a:themeElements > a:clrScheme > a:accent5 > a:srgbClr
        $xml->endElement(); //a:theme > a:themeElements > a:clrScheme > a:accent5
        $xml->startElement('a:accent6'); //a:theme > a:themeElements > a:clrScheme > a:accent6
        $xml->startElement('a:srgbClr'); //a:theme > a:themeElements > a:clrScheme > a:accent6 > a:srgbClr
        $xml->writeAttribute('val', 'F79646'); //a:theme > a:themeElements > a:clrScheme > a:accent6 > a:srgbClr
        $xml->endElement(); //a:theme > a:themeElements > a:clrScheme > a:accent6 > a:srgbClr
        $xml->endElement(); //a:theme > a:themeElements > a:clrScheme > a:accent6
        $xml->startElement('a:hlink'); //a:theme > a:themeElements > a:clrScheme > a:hlink
        $xml->startElement('a:srgbClr'); //a:theme > a:themeElements > a:clrScheme > a:hlink > a:srgbClr
        $xml->writeAttribute('val', '0000FF'); //a:theme > a:themeElements > a:clrScheme > a:hlink > a:srgbClr
        $xml->endElement(); //a:theme > a:themeElements > a:clrScheme > a:hlink > a:srgbClr
        $xml->endElement(); //a:theme > a:themeElements > a:clrScheme > a:hlink
        $xml->startElement('a:folHlink'); //a:theme > a:themeElements > a:clrScheme > a:folHlink
        $xml->startElement('a:srgbClr'); //a:theme > a:themeElements > a:clrScheme > a:folHlink > a:srgbClr
        $xml->writeAttribute('val', '800080'); //a:theme > a:themeElements > a:clrScheme > a:folHlink > a:srgbClr
        $xml->endElement(); //a:theme > a:themeElements > a:clrScheme > a:folHlink > a:srgbClr
        $xml->endElement(); //a:theme > a:themeElements > a:clrScheme > a:folHlink
        $xml->endElement(); //a:theme > a:themeElements > a:clrScheme
        $xml->startElement('a:fontScheme'); //a:theme > a:themeElements > a:fontScheme
        $xml->writeAttribute('name', 'Office'); //a:theme > a:themeElements > a:fontScheme
        $xml->startElement('a:majorFont'); //a:theme > a:themeElements > a:fontScheme > a:majorFont
        $xml->startElement('a:latin'); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:latin 
        $xml->writeAttribute('typeface', 'Cambria'); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:latin 
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:latin 
        $xml->startElement('a:ea'); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:ea
        $xml->writeAttribute('typeface', ''); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:ea 
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:ea
        $xml->startElement('a:cs'); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:cs
        $xml->writeAttribute('typeface', ''); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:cs
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:cs
        $xml->startElement('a:font'); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->writeAttribute('script', 'Jpan'); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->writeAttribute('typeface', '?? ?????'); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->startElement('a:font'); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->writeAttribute('script', 'Hang'); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->writeAttribute('typeface', '?? ??'); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->startElement('a:font'); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->writeAttribute('script', 'Hans'); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->writeAttribute('typeface', '??'); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->startElement('a:font'); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->writeAttribute('script', 'Hant'); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->writeAttribute('typeface', '????'); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->startElement('a:font'); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->writeAttribute('script', 'Arab'); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->writeAttribute('typeface', 'Times New Roman'); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->startElement('a:font'); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->writeAttribute('script', 'Hebr'); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->writeAttribute('typeface', 'Times New Roman'); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->startElement('a:font'); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->writeAttribute('script', 'Thai'); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->writeAttribute('typeface', 'Tahoma'); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->startElement('a:font'); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->writeAttribute('script', 'Ethi'); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->writeAttribute('typeface', 'Nyala'); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->startElement('a:font'); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->writeAttribute('script', 'Beng'); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->writeAttribute('typeface', 'Vrinda'); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->startElement('a:font'); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->writeAttribute('script', 'Gujr'); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->writeAttribute('typeface', 'Shruti'); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->startElement('a:font'); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->writeAttribute('script', 'Khmr'); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->writeAttribute('typeface', 'MoolBoran'); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->startElement('a:font'); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->writeAttribute('script', 'Knda'); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->writeAttribute('typeface', 'Tunga'); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->startElement('a:font'); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->writeAttribute('script', 'Guru'); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->writeAttribute('typeface', 'Raavi'); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->startElement('a:font'); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->writeAttribute('script', 'Cans'); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->writeAttribute('typeface', 'Euphemia'); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->startElement('a:font'); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->writeAttribute('script', 'Cher'); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->writeAttribute('typeface', 'Plantagenet Cherokee'); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->startElement('a:font'); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->writeAttribute('script', 'Yiii'); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->writeAttribute('typeface', 'Microsoft Yi Baiti'); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->startElement('a:font'); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->writeAttribute('script', 'Tibt'); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->writeAttribute('typeface', 'Microsoft Himalaya'); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->startElement('a:font'); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->writeAttribute('script', 'Thaa'); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->writeAttribute('typeface', 'MV Boli'); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->startElement('a:font'); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->writeAttribute('script', 'Deva'); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->writeAttribute('typeface', 'Mangal'); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->startElement('a:font'); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->writeAttribute('script', 'Telu'); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->writeAttribute('typeface', 'Gautami'); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->startElement('a:font'); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->writeAttribute('script', 'Taml'); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->writeAttribute('typeface', 'Latha'); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->startElement('a:font'); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->writeAttribute('script', 'Syrc'); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->writeAttribute('typeface', 'Estrangelo Edessa'); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->startElement('a:font'); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->writeAttribute('script', 'Orya'); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->writeAttribute('typeface', 'Kalinga'); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->startElement('a:font'); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->writeAttribute('script', 'Mlym'); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->writeAttribute('typeface', 'Kartika'); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->startElement('a:font'); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->writeAttribute('script', 'Laoo'); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->writeAttribute('typeface', 'DokChampa'); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->startElement('a:font'); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->writeAttribute('script', 'Sinh'); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->writeAttribute('typeface', 'Iskoola Pota'); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->startElement('a:font'); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->writeAttribute('script', 'Mong'); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->writeAttribute('typeface', 'Mongolian Baiti'); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->startElement('a:font'); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->writeAttribute('script', 'Viet'); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->writeAttribute('typeface', 'Times New Roman'); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->startElement('a:font'); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->writeAttribute('script', 'Uigh'); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->writeAttribute('typeface', 'Microsoft Uighur'); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:majorFont > a:font
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:majorFont
        $xml->startElement('a:minorFont'); //a:theme > a:themeElements > a:fontScheme > a:minorFont
        $xml->startElement('a:latin'); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:latin 
        $xml->writeAttribute('typeface', 'Calibri'); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:latin 
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:latin 
        $xml->startElement('a:ea'); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:ea 
        $xml->writeAttribute('typeface', ''); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:ea 
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:ea 
        $xml->startElement('a:cs'); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:cs 
        $xml->writeAttribute('typeface', ''); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:cs 
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:cs 
        $xml->startElement('a:font'); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->writeAttribute('script', 'Jpan'); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->writeAttribute('typeface', '?? ?????'); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->startElement('a:font'); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->writeAttribute('script', 'Hang'); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->writeAttribute('typeface', '?? ??'); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->startElement('a:font'); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->writeAttribute('script', 'Hans'); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->writeAttribute('typeface', '??'); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->startElement('a:font'); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->writeAttribute('script', 'Hant'); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->writeAttribute('typeface', '????'); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->startElement('a:font'); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->writeAttribute('script', 'Arab'); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->writeAttribute('typeface', 'Arial'); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->startElement('a:font'); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->writeAttribute('script', 'Hebr'); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->writeAttribute('typeface', 'Arial'); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->startElement('a:font'); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->writeAttribute('script', 'Thai'); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->writeAttribute('typeface', 'Tahoma'); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->startElement('a:font'); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->writeAttribute('script', 'Ethi'); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->writeAttribute('typeface', 'Nyala'); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->startElement('a:font'); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->writeAttribute('script', 'Beng'); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->writeAttribute('typeface', 'Vrinda'); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->startElement('a:font'); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->writeAttribute('script', 'Gujr'); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->writeAttribute('typeface', 'Shruti'); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->startElement('a:font'); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->writeAttribute('script', 'Khmr'); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->writeAttribute('typeface', 'DaunPenh'); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->startElement('a:font'); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->writeAttribute('script', 'Knda'); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->writeAttribute('typeface', 'Tunga'); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->startElement('a:font'); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->writeAttribute('script', 'Guru'); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->writeAttribute('typeface', 'Raavi'); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->startElement('a:font'); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->writeAttribute('script', 'Cans'); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->writeAttribute('typeface', 'Euphemia'); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->startElement('a:font'); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->writeAttribute('script', 'Cher'); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->writeAttribute('typeface', 'Plantagenet Cherokee'); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->startElement('a:font'); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->writeAttribute('script', 'Yiii'); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->writeAttribute('typeface', 'Microsoft Yi Baiti'); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->startElement('a:font'); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->writeAttribute('script', 'Tibt'); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->writeAttribute('typeface', 'Microsoft Himalaya'); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->startElement('a:font'); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->writeAttribute('script', 'Thaa'); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->writeAttribute('typeface', 'MV Boli'); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->startElement('a:font'); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->writeAttribute('script', 'Deva'); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->writeAttribute('typeface', 'Mangal'); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->startElement('a:font'); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->writeAttribute('script', 'Telu'); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->writeAttribute('typeface', 'Gautami'); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->startElement('a:font'); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->writeAttribute('script', 'Taml'); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->writeAttribute('typeface', 'Latha'); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->startElement('a:font'); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->writeAttribute('script', 'Syrc'); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->writeAttribute('typeface', 'Estrangelo Edessa'); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->startElement('a:font'); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->writeAttribute('script', 'Orya'); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->writeAttribute('typeface', 'Kalinga'); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->startElement('a:font'); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->writeAttribute('script', 'Mlym'); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->writeAttribute('typeface', 'Kartika'); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->startElement('a:font'); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->writeAttribute('script', 'Laoo'); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->writeAttribute('typeface', 'DokChampa'); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->startElement('a:font'); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->writeAttribute('script', 'Sinh'); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->writeAttribute('typeface', 'Iskoola Pota'); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->startElement('a:font'); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->writeAttribute('script', 'Mong'); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->writeAttribute('typeface', 'Mongolian Baiti'); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->startElement('a:font'); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->writeAttribute('script', 'Viet'); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->writeAttribute('typeface', 'Arial'); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->startElement('a:font'); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->writeAttribute('script', 'Uigh'); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->writeAttribute('typeface', 'Microsoft Uighur'); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:minorFont > a:font 
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:minorFont
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme
        $xml->startElement('a:fmtScheme'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme
        $xml->writeAttribute('name', 'Office'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme
        $xml->startElement('a:fillStyleLst'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:fillStyleLst
        $xml->startElement('a:solidFill'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:fillStyleLst > a:solidFill
        $xml->startElement('a:schemeClr'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:fillStyleLst > a:solidFill > a:schemeClr
        $xml->writeAttribute('val', 'phClr'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:fillStyleLst > a:solidFill > a:schemeClr
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:fillStyleLst > a:solidFill > a:schemeClr
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:fillStyleLst > a:solidFill
        $xml->startElement('a:gradFill'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:fillStyleLst > a:gradFill
        $xml->writeAttribute('rotWithShape', '1'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:fillStyleLst > a:gradFill
        $xml->startElement('a:gsLst'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:fillStyleLst > a:gradFill > a:gsLst
        $xml->startElement('a:gs'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:fillStyleLst > a:gradFill > a:gsLst > a:gs
        $xml->writeAttribute('pos', '0'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:fillStyleLst > a:gradFill > a:gsLst > a:gs
        $xml->startElement('a:schemeClr'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:fillStyleLst > a:gradFill > a:gsLst > a:gs > a:schemeClr
        $xml->writeAttribute('val', 'phClr'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:fillStyleLst > a:gradFill > a:gsLst > a:gs > a:schemeClr
        $xml->startElement('a:tint'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:fillStyleLst > a:gradFill > a:gsLst > a:gs > a:schemeClr > a:tint
        $xml->writeAttribute('val', '50000'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:fillStyleLst > a:gradFill > a:gsLst > a:gs > a:schemeClr > a:tint
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:fillStyleLst > a:gradFill > a:gsLst > a:gs > a:schemeClr > a:tint
        $xml->startElement('a:satMod'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:fillStyleLst > a:gradFill > a:gsLst > a:gs > a:schemeClr > a:satMod
        $xml->writeAttribute('val', '300000'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:fillStyleLst > a:gradFill > a:gsLst > a:gs > a:schemeClr > a:satMod
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:fillStyleLst > a:gradFill > a:gsLst > a:gs > a:schemeClr > a:satMod
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:fillStyleLst > a:gradFill > a:gsLst > a:gs > a:schemeClr
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:fillStyleLst > a:gradFill > a:gsLst > a:gs
        $xml->startElement('a:gs'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:fillStyleLst > a:gradFill > a:gsLst > a:gs
        $xml->writeAttribute('pos', '35000'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:fillStyleLst > a:gradFill > a:gsLst > a:gs
        $xml->startElement('a:schemeClr'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:fillStyleLst > a:gradFill > a:gsLst > a:gs > a:schemeClr
        $xml->writeAttribute('val', 'phClr'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:fillStyleLst > a:gradFill > a:gsLst > a:gs > a:schemeClr
        $xml->startElement('a:tint'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:fillStyleLst > a:gradFill > a:gsLst > a:gs > a:schemeClr > a:tint
        $xml->writeAttribute('val', '37000'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:fillStyleLst > a:gradFill > a:gsLst > a:gs > a:schemeClr > a:tint
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:fillStyleLst > a:gradFill > a:gsLst > a:gs > a:schemeClr > a:tint
        $xml->startElement('a:satMod'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:fillStyleLst > a:gradFill > a:gsLst > a:gs > a:schemeClr > a:satMod
        $xml->writeAttribute('val', '300000'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:fillStyleLst > a:gradFill > a:gsLst > a:gs > a:schemeClr > a:satMod
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:fillStyleLst > a:gradFill > a:gsLst > a:gs > a:schemeClr > a:satMod
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:fillStyleLst > a:gradFill > a:gsLst > a:gs > a:schemeClr
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:fillStyleLst > a:gradFill > a:gsLst > a:gs
        $xml->startElement('a:gs'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:fillStyleLst > a:gradFill > a:gsLst > a:gs
        $xml->writeAttribute('pos', '100000'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:fillStyleLst > a:gradFill > a:gsLst > a:gs
        $xml->startElement('a:schemeClr'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:fillStyleLst > a:gradFill > a:gsLst > a:gs > a:schemeClr
        $xml->writeAttribute('val', 'phClr'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:fillStyleLst > a:gradFill > a:gsLst > a:gs > a:schemeClr
        $xml->startElement('a:tint'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:fillStyleLst > a:gradFill > a:gsLst > a:gs > a:schemeClr > a:tint
        $xml->writeAttribute('val', '15000'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:fillStyleLst > a:gradFill > a:gsLst > a:gs > a:schemeClr > a:tint
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:fillStyleLst > a:gradFill > a:gsLst > a:gs > a:schemeClr > a:tint
        $xml->startElement('a:satMod'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:fillStyleLst > a:gradFill > a:gsLst > a:gs > a:schemeClr > a:satMod
        $xml->writeAttribute('val', '350000'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:fillStyleLst > a:gradFill > a:gsLst > a:gs > a:schemeClr > a:satMod
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:fillStyleLst > a:gradFill > a:gsLst > a:gs > a:schemeClr > a:satMod
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:fillStyleLst > a:gradFill > a:gsLst > a:gs > a:schemeClr
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:fillStyleLst > a:gradFill > a:gsLst > a:gs
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:fillStyleLst > a:gradFill > a:gsLst
        $xml->startElement('a:lin'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:fillStyleLst > a:gradFill > a:lin
        $xml->writeAttribute('ang', '16200000'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:fillStyleLst > a:gradFill > a:lin
        $xml->writeAttribute('scaled', '1'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:fillStyleLst > a:gradFill > a:lin
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:fillStyleLst > a:gradFill > a:lin
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:fillStyleLst > a:gradFill
        $xml->startElement('a:gradFill'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:fillStyleLst > a:gradFill
        $xml->writeAttribute('rotWithShape', '1'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:fillStyleLst > a:gradFill
        $xml->startElement('a:gsLst'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:fillStyleLst > a:gradFill > a:gsLst
        $xml->startElement('a:gs'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:fillStyleLst > a:gradFill > a:gsLst > a:gs
        $xml->writeAttribute('pos', '0'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:fillStyleLst > a:gradFill > a:gsLst > a:gs
        $xml->startElement('a:schemeClr'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:fillStyleLst > a:gradFill > a:gsLst > a:gs > a:schemeClr
        $xml->writeAttribute('val', 'phClr'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:fillStyleLst > a:gradFill > a:gsLst > a:gs > a:schemeClr
        $xml->startElement('a:shade'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:fillStyleLst > a:gradFill > a:gsLst > a:gs > a:schemeClr > a:shade
        $xml->writeAttribute('val', '51000'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:fillStyleLst > a:gradFill > a:gsLst > a:gs > a:schemeClr > a:shade
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:fillStyleLst > a:gradFill > a:gsLst > a:gs > a:schemeClr > a:shade
        $xml->startElement('a:satMod'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:fillStyleLst > a:gradFill > a:gsLst > a:gs > a:schemeClr > a:satMod
        $xml->writeAttribute('val', '130000'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:fillStyleLst > a:gradFill > a:gsLst > a:gs > a:schemeClr > a:satMod
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:fillStyleLst > a:gradFill > a:gsLst > a:gs > a:schemeClr > a:satMod
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:fillStyleLst > a:gradFill > a:gsLst > a:gs > a:schemeClr
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:fillStyleLst > a:gradFill > a:gsLst > a:gs
        $xml->startElement('a:gs'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:fillStyleLst > a:gradFill > a:gsLst > a:gs
        $xml->writeAttribute('pos', '80000'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:fillStyleLst > a:gradFill > a:gsLst > a:gs
        $xml->startElement('a:schemeClr'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:fillStyleLst > a:gradFill > a:gsLst > a:gs > a:schemeClr
        $xml->writeAttribute('val', 'phClr'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:fillStyleLst > a:gradFill > a:gsLst > a:gs > a:schemeClr
        $xml->startElement('a:tint'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:fillStyleLst > a:gradFill > a:gsLst > a:gs > a:schemeClr > a:tint
        $xml->writeAttribute('val', '93000'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:fillStyleLst > a:gradFill > a:gsLst > a:gs > a:schemeClr > a:tint
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:fillStyleLst > a:gradFill > a:gsLst > a:gs > a:schemeClr > a:tint
        $xml->startElement('a:satMod'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:fillStyleLst > a:gradFill > a:gsLst > a:gs > a:schemeClr > a:satMod
        $xml->writeAttribute('val', '130000'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:fillStyleLst > a:gradFill > a:gsLst > a:gs > a:schemeClr > a:satMod
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:fillStyleLst > a:gradFill > a:gsLst > a:gs > a:schemeClr > a:satMod
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:fillStyleLst > a:gradFill > a:gsLst > a:gs > a:schemeClr
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:fillStyleLst > a:gradFill > a:gsLst > a:gs
        $xml->startElement('a:gs'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:fillStyleLst > a:gradFill > a:gsLst > a:gs
        $xml->writeAttribute('pos', '100000'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:fillStyleLst > a:gradFill > a:gsLst > a:gs
        $xml->startElement('a:schemeClr'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:fillStyleLst > a:gradFill > a:gsLst > a:gs > a:schemeClr
        $xml->writeAttribute('val', 'phClr'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:fillStyleLst > a:gradFill > a:gsLst > a:gs > a:schemeClr
        $xml->startElement('a:shade'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:fillStyleLst > a:gradFill > a:gsLst > a:gs > a:schemeClr > a:shade
        $xml->writeAttribute('val', '94000'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:fillStyleLst > a:gradFill > a:gsLst > a:gs > a:schemeClr > a:shade
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:fillStyleLst > a:gradFill > a:gsLst > a:gs > a:schemeClr > a:shade
        $xml->startElement('a:satMod'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:fillStyleLst > a:gradFill > a:gsLst > a:gs > a:schemeClr > a:satMod
        $xml->writeAttribute('val', '135000'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:fillStyleLst > a:gradFill > a:gsLst > a:gs > a:schemeClr > a:satMod
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:fillStyleLst > a:gradFill > a:gsLst > a:gs > a:schemeClr > a:satMod
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:fillStyleLst > a:gradFill > a:gsLst > a:gs > a:schemeClr
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:fillStyleLst > a:gradFill > a:gsLst > a:gs
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:fillStyleLst > a:gradFill > a:gsLst
        $xml->startElement('a:lin'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:fillStyleLst > a:gradFill > a:lin
        $xml->writeAttribute('ang', '16200000'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:fillStyleLst > a:gradFill > a:lin
        $xml->writeAttribute('scaled', '0'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:fillStyleLst > a:gradFill > a:lin
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:fillStyleLst > a:gradFill > a:lin
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:fillStyleLst > a:gradFill
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:fillStyleLst
        $xml->startElement('a:lnStyleLst'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:lnStyleLst
        $xml->startElement('a:ln'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:lnStyleLst > a:ln
        $xml->writeAttribute('w', '9525'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:lnStyleLst > a:ln
        $xml->writeAttribute('cap', 'flat'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:lnStyleLst > a:ln
        $xml->writeAttribute('cmpd', 'sng'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:lnStyleLst > a:ln
        $xml->writeAttribute('algn', 'ctr'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:lnStyleLst > a:ln
        $xml->startElement('a:solidFill'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:lnStyleLst > a:ln > a:solidFill
        $xml->startElement('a:schemeClr'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:lnStyleLst > a:ln > a:solidFill > a:schemeClr
        $xml->writeAttribute('val', 'phClr'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:lnStyleLst > a:ln > a:solidFill > a:schemeClr
        $xml->startElement('a:shade'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:lnStyleLst > a:ln > a:solidFill > a:schemeClr > a:shade
        $xml->writeAttribute('val', '95000'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:lnStyleLst > a:ln > a:solidFill > a:schemeClr > a:shade
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:lnStyleLst > a:ln > a:solidFill > a:schemeClr > a:shade
        $xml->startElement('a:satMod'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:lnStyleLst > a:ln > a:solidFill > a:schemeClr > a:satMod
        $xml->writeAttribute('val', '105000'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:lnStyleLst > a:ln > a:solidFill > a:schemeClr > a:satMod
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:lnStyleLst > a:ln > a:solidFill > a:schemeClr > a:satMod
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:lnStyleLst > a:ln > a:solidFill > a:schemeClr
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:lnStyleLst > a:ln > a:solidFill
        $xml->startElement('a:prstDash'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:lnStyleLst > a:ln > a:prstDash
        $xml->writeAttribute('val', 'solid'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:lnStyleLst > a:ln > a:prstDash
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:lnStyleLst > a:ln > a:prstDash
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:lnStyleLst > a:ln
        $xml->startElement('a:ln'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:lnStyleLst > a:ln
        $xml->writeAttribute('w', '25400'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:lnStyleLst > a:ln
        $xml->writeAttribute('cap', 'flat'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:lnStyleLst > a:ln
        $xml->writeAttribute('cmpd', 'sng'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:lnStyleLst > a:ln
        $xml->writeAttribute('algn', 'ctr'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:lnStyleLst > a:ln
        $xml->startElement('a:solidFill'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:lnStyleLst > a:ln > a:solidFill
        $xml->startElement('a:schemeClr'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:lnStyleLst > a:ln > a:solidFill > a:schemeClr
        $xml->writeAttribute('val', 'phClr'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:lnStyleLst > a:ln > a:solidFill > a:schemeClr
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:lnStyleLst > a:ln > a:solidFill > a:schemeClr
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:lnStyleLst > a:ln > a:solidFill
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:lnStyleLst > a:ln
        $xml->startElement('a:ln'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:lnStyleLst > a:ln
        $xml->writeAttribute('w', '38100'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:lnStyleLst > a:ln
        $xml->writeAttribute('cap', 'flat'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:lnStyleLst > a:ln
        $xml->writeAttribute('cmpd', 'sng'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:lnStyleLst > a:ln
        $xml->writeAttribute('algn', 'ctr'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:lnStyleLst > a:ln
        $xml->startElement('a:solidFill'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:lnStyleLst > a:ln > a:solidFill
        $xml->startElement('a:schemeClr'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:lnStyleLst > a:ln > a:solidFill > a:schemeClr
        $xml->writeAttribute('val', 'phClr'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:lnStyleLst > a:ln > a:solidFill > a:schemeClr
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:lnStyleLst > a:ln > a:solidFill > a:schemeClr
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:lnStyleLst > a:ln > a:solidFill
        $xml->startElement('a:prstDash'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:lnStyleLst > a:ln > a:prstDash
        $xml->writeAttribute('val', 'solid'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:lnStyleLst > a:ln > a:prstDash
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:lnStyleLst > a:ln > a:prstDash
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:lnStyleLst > a:ln
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:lnStyleLst
        $xml->startElement('a:effectStyleLst'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:effectStyleLst
        $xml->startElement('a:effectStyle'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:effectStyleLst > a:effectStyle
        $xml->startElement('a:effectLst'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:effectStyleLst > a:effectStyle > a:effectLst
        $xml->startElement('a:outerShdw'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:effectStyleLst > a:effectStyle > a:effectLst > a:outerShdw
        $xml->writeAttribute('blurRad', '40000'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:effectStyleLst > a:effectStyle > a:effectLst > a:outerShdw
        $xml->writeAttribute('dist', '20000'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:effectStyleLst > a:effectStyle > a:effectLst > a:outerShdw
        $xml->writeAttribute('dir', '5400000'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:effectStyleLst > a:effectStyle > a:effectLst > a:outerShdw
        $xml->writeAttribute('rotWithShape', '0'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:effectStyleLst > a:effectStyle > a:effectLst > a:outerShdw
        $xml->startElement('a:srgbClr'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:effectStyleLst > a:effectStyle > a:effectLst > a:outerShdw > a:srgbClr
        $xml->writeAttribute('val', '000000'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:effectStyleLst > a:effectStyle > a:effectLst > a:outerShdw > a:srgbClr
        $xml->startElement('a:alpha'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:effectStyleLst > a:effectStyle > a:effectLst > a:outerShdw > a:srgbClr > a:alpha
        $xml->writeAttribute('val', '38000'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:effectStyleLst > a:effectStyle > a:effectLst > a:outerShdw > a:srgbClr > a:alpha
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:effectStyleLst > a:effectStyle > a:effectLst > a:outerShdw > a:srgbClr > a:alpha
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:effectStyleLst > a:effectStyle > a:effectLst > a:outerShdw > a:srgbClr
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:effectStyleLst > a:effectStyle > a:effectLst > a:outerShdw
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:effectStyleLst > a:effectStyle > a:effectLst
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:effectStyleLst > a:effectStyle
        $xml->startElement('a:effectStyle'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:effectStyleLst > a:effectStyle
        $xml->startElement('a:effectLst'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:effectStyleLst > a:effectStyle > a:effectLst
        $xml->startElement('a:outerShdw'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:effectStyleLst > a:effectStyle > a:effectLst > a:outerShdw
        $xml->writeAttribute('blurRad', '40000'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:effectStyleLst > a:effectStyle > a:effectLst > a:outerShdw
        $xml->writeAttribute('dist', '23000'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:effectStyleLst > a:effectStyle > a:effectLst > a:outerShdw
        $xml->writeAttribute('dir', '5400000'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:effectStyleLst > a:effectStyle > a:effectLst > a:outerShdw
        $xml->writeAttribute('rotWithShape', '0'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:effectStyleLst > a:effectStyle > a:effectLst > a:outerShdw
        $xml->startElement('a:srgbClr'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:effectStyleLst > a:effectStyle > a:effectLst > a:outerShdw > a:srgbClr
        $xml->writeAttribute('val', '000000'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:effectStyleLst > a:effectStyle > a:effectLst > a:outerShdw > a:srgbClr
        $xml->startElement('a:alpha'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:effectStyleLst > a:effectStyle > a:effectLst > a:outerShdw > a:srgbClr > a:alpha
        $xml->writeAttribute('val', '38000'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:effectStyleLst > a:effectStyle > a:effectLst > a:outerShdw > a:srgbClr > a:alpha
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:effectStyleLst > a:effectStyle > a:effectLst > a:outerShdw > a:srgbClr > a:alpha
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:effectStyleLst > a:effectStyle > a:effectLst > a:outerShdw > a:srgbClr
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:effectStyleLst > a:effectStyle > a:effectLst > a:outerShdw
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:effectStyleLst > a:effectStyle > a:effectLst
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:effectStyleLst > a:effectStyle
        $xml->startElement('a:effectStyle'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:effectStyleLst > a:effectStyle
        $xml->startElement('a:effectLst'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:effectStyleLst > a:effectStyle > a:effectLst
        $xml->startElement('a:outerShdw'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:effectStyleLst > a:effectStyle > a:effectLst > a:outerShdw
        $xml->writeAttribute('blurRad', '40000'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:effectStyleLst > a:effectStyle > a:effectLst > a:outerShdw
        $xml->writeAttribute('dist', '23000'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:effectStyleLst > a:effectStyle > a:effectLst > a:outerShdw
        $xml->writeAttribute('dir', '5400000'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:effectStyleLst > a:effectStyle > a:effectLst > a:outerShdw
        $xml->writeAttribute('rotWithShape', '0'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:effectStyleLst > a:effectStyle > a:effectLst > a:outerShdw
        $xml->startElement('a:srgbClr'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:effectStyleLst > a:effectStyle > a:effectLst > a:outerShdw > a:srgbClr
        $xml->writeAttribute('val', '000000'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:effectStyleLst > a:effectStyle > a:effectLst > a:outerShdw > a:srgbClr
        $xml->startElement('a:alpha'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:effectStyleLst > a:effectStyle > a:effectLst > a:outerShdw > a:srgbClr > a:alpha
        $xml->writeAttribute('val', '35000'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:effectStyleLst > a:effectStyle > a:effectLst > a:outerShdw > a:srgbClr > a:alpha
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:effectStyleLst > a:effectStyle > a:effectLst > a:outerShdw > a:srgbClr > a:alpha
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:effectStyleLst > a:effectStyle > a:effectLst > a:outerShdw > a:srgbClr
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:effectStyleLst > a:effectStyle > a:effectLst > a:outerShdw
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:effectStyleLst > a:effectStyle > a:effectLst
        $xml->startElement('a:scene3d'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:effectStyleLst > a:effectStyle > a:scene3d
        $xml->startElement('a:camera'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:effectStyleLst > a:effectStyle > a:scene3d > a:camera
        $xml->writeAttribute('prst', 'orthographicFront'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:effectStyleLst > a:effectStyle > a:scene3d > a:camera
        $xml->startElement('a:rot'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:effectStyleLst > a:effectStyle > a:scene3d > a:camera > a:rot
        $xml->writeAttribute('lat', '0'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:effectStyleLst > a:effectStyle > a:scene3d > a:camera > a:rot
        $xml->writeAttribute('lon', '0'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:effectStyleLst > a:effectStyle > a:scene3d > a:camera > a:rot
        $xml->writeAttribute('rev', '0'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:effectStyleLst > a:effectStyle > a:scene3d > a:camera > a:rot
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:effectStyleLst > a:effectStyle > a:scene3d > a:camera > a:rot
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:effectStyleLst > a:effectStyle > a:scene3d > a:camera
        $xml->startElement('a:lightRig'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:effectStyleLst > a:effectStyle > a:scene3d > a:lightRig
        $xml->writeAttribute('rig', 'threePt'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:effectStyleLst > a:effectStyle > a:scene3d > a:lightRig
        $xml->writeAttribute('dir', 't'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:effectStyleLst > a:effectStyle > a:scene3d > a:lightRig
        $xml->startElement('a:rot'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:effectStyleLst > a:effectStyle > a:scene3d > a:lightRig > a:rot
        $xml->writeAttribute('lat', '0'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:effectStyleLst > a:effectStyle > a:scene3d > a:lightRig > a:rot
        $xml->writeAttribute('lon', '0'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:effectStyleLst > a:effectStyle > a:scene3d > a:lightRig > a:rot
        $xml->writeAttribute('rev', '1200000'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:effectStyleLst > a:effectStyle > a:scene3d > a:lightRig > a:rot
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:effectStyleLst > a:effectStyle > a:scene3d > a:lightRig > a:rot
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:effectStyleLst > a:effectStyle > a:scene3d > a:lightRig
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:effectStyleLst > a:effectStyle > a:scene3d
        $xml->startElement('a:sp3d'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:effectStyleLst > a:effectStyle > a:sp3d
        $xml->startElement('a:sp3d'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:effectStyleLst > a:effectStyle > a:sp3d > a:bevelT
        $xml->writeAttribute('w', '63500'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:effectStyleLst > a:effectStyle > a:sp3d > a:bevelT
        $xml->writeAttribute('h', '25400'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:effectStyleLst > a:effectStyle > a:sp3d > a:bevelT
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:effectStyleLst > a:effectStyle > a:sp3d > a:bevelT
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:effectStyleLst > a:effectStyle > a:sp3d
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:effectStyleLst > a:effectStyle
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:effectStyleLst
        $xml->startElement('a:bgFillStyleLst'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:bgFillStyleLst
        $xml->startElement('a:solidFill'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:bgFillStyleLst > a:solidFill
        $xml->startElement('a:schemeClr'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:bgFillStyleLst > a:solidFill > a:schemeClr
        $xml->writeAttribute('val', 'phClr'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:bgFillStyleLst > a:solidFill > a:schemeClr
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:bgFillStyleLst > a:solidFill > a:schemeClr
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:bgFillStyleLst > a:solidFill
        $xml->startElement('a:gradFill'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:bgFillStyleLst > a:gradFill
        $xml->startElement('a:gsLst'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:bgFillStyleLst > a:gradFill > a:gsLst
        $xml->startElement('a:gs'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:bgFillStyleLst > a:gradFill > a:gsLst > a:gs
        $xml->writeAttribute('pos', '0'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:bgFillStyleLst > a:gradFill > a:gsLst > a:gs
        $xml->startElement('a:schemeClr'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:bgFillStyleLst > a:gradFill > a:gsLst > a:gs > a:schemeClr
        $xml->writeAttribute('val', 'phClr'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:bgFillStyleLst > a:gradFill > a:gsLst > a:gs > a:schemeClr
        $xml->startElement('a:tint'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:bgFillStyleLst > a:gradFill > a:gsLst > a:gs > a:schemeClr > a:tint
        $xml->writeAttribute('val', '40000'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:bgFillStyleLst > a:gradFill > a:gsLst > a:gs > a:schemeClr > a:tint
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:bgFillStyleLst > a:gradFill > a:gsLst > a:gs > a:schemeClr > a:tint
        $xml->startElement('a:satMod'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:bgFillStyleLst > a:gradFill > a:gsLst > a:gs > a:schemeClr > a:satMod
        $xml->writeAttribute('val', '350000'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:bgFillStyleLst > a:gradFill > a:gsLst > a:gs > a:schemeClr > a:satMod
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:bgFillStyleLst > a:gradFill > a:gsLst > a:gs > a:schemeClr > a:satMod
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:bgFillStyleLst > a:gradFill > a:gsLst > a:gs > a:schemeClr
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:bgFillStyleLst > a:gradFill > a:gsLst > a:gs
        $xml->startElement('a:gs'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:bgFillStyleLst > a:gradFill > a:gsLst > a:gs
        $xml->writeAttribute('pos', '40000'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:bgFillStyleLst > a:gradFill > a:gsLst > a:gs
        $xml->startElement('a:schemeClr'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:bgFillStyleLst > a:gradFill > a:gsLst > a:gs > a:schemeClr
        $xml->writeAttribute('val', 'phClr'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:bgFillStyleLst > a:gradFill > a:gsLst > a:gs > a:schemeClr
        $xml->startElement('a:tint'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:bgFillStyleLst > a:gradFill > a:gsLst > a:gs > a:schemeClr > a:tint
        $xml->writeAttribute('val', '45000'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:bgFillStyleLst > a:gradFill > a:gsLst > a:gs > a:schemeClr > a:tint
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:bgFillStyleLst > a:gradFill > a:gsLst > a:gs > a:schemeClr > a:tint
        $xml->startElement('a:shade'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:bgFillStyleLst > a:gradFill > a:gsLst > a:gs > a:schemeClr > a:shade
        $xml->writeAttribute('val', '99000'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:bgFillStyleLst > a:gradFill > a:gsLst > a:gs > a:schemeClr > a:shade
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:bgFillStyleLst > a:gradFill > a:gsLst > a:gs > a:schemeClr > a:shade
        $xml->startElement('a:satMod'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:bgFillStyleLst > a:gradFill > a:gsLst > a:gs > a:schemeClr > a:satMod
        $xml->writeAttribute('val', '350000'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:bgFillStyleLst > a:gradFill > a:gsLst > a:gs > a:schemeClr > a:satMod
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:bgFillStyleLst > a:gradFill > a:gsLst > a:gs > a:schemeClr > a:satMod
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:bgFillStyleLst > a:gradFill > a:gsLst > a:gs > a:schemeClr
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:bgFillStyleLst > a:gradFill > a:gsLst > a:gs
        $xml->startElement('a:gs'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:bgFillStyleLst > a:gradFill > a:gsLst > a:gs
        $xml->writeAttribute('pos', '100000'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:bgFillStyleLst > a:gradFill > a:gsLst > a:gs
        $xml->startElement('a:schemeClr'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:bgFillStyleLst > a:gradFill > a:gsLst > a:gs > a:schemeClr
        $xml->writeAttribute('val', 'phClr'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:bgFillStyleLst > a:gradFill > a:gsLst > a:gs > a:schemeClr
        $xml->startElement('a:shade'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:bgFillStyleLst > a:gradFill > a:gsLst > a:gs > a:schemeClr > a:shade
        $xml->writeAttribute('val', '20000'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:bgFillStyleLst > a:gradFill > a:gsLst > a:gs > a:schemeClr > a:shade
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:bgFillStyleLst > a:gradFill > a:gsLst > a:gs > a:schemeClr > a:shade
        $xml->startElement('a:satMod'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:bgFillStyleLst > a:gradFill > a:gsLst > a:gs > a:schemeClr > a:satMod
        $xml->writeAttribute('val', '255000'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:bgFillStyleLst > a:gradFill > a:gsLst > a:gs > a:schemeClr > a:satMod
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:bgFillStyleLst > a:gradFill > a:gsLst > a:gs > a:schemeClr > a:satMod
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:bgFillStyleLst > a:gradFill > a:gsLst > a:gs > a:schemeClr
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:bgFillStyleLst > a:gradFill > a:gsLst > a:gs
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:bgFillStyleLst > a:gradFill > a:gsLst
        $xml->startElement('a:path'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:bgFillStyleLst > a:gradFill > a:path
        $xml->writeAttribute('path', 'circle'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:bgFillStyleLst > a:gradFill > a:path
        $xml->startElement('a:fillToRect'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:bgFillStyleLst > a:gradFill > a:path > a:fillToRect
        $xml->writeAttribute('l', '50000'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:bgFillStyleLst > a:gradFill >  a:path > a:fillToRect
        $xml->writeAttribute('t', '-80000'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:bgFillStyleLst > a:gradFill > a:path > a:fillToRect
        $xml->writeAttribute('r', '50000'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:bgFillStyleLst > a:gradFill >  a:path > a:fillToRect
        $xml->writeAttribute('b', '180000'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:bgFillStyleLst > a:gradFill > a:path > a:fillToRect
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:bgFillStyleLst > a:gradFill > a:path > a:fillToRect
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:bgFillStyleLst > a:gradFill > a:path
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:bgFillStyleLst > a:gradFill
        $xml->startElement('a:gradFill'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:bgFillStyleLst > a:gradFill
        $xml->writeAttribute('rotWithShape', '1'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:bgFillStyleLst > a:gradFill
        $xml->startElement('a:gsLst'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:bgFillStyleLst > a:gradFill > a:gsLst
        $xml->startElement('a:gs'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:bgFillStyleLst > a:gradFill > a:gsLst > a:gs
        $xml->writeAttribute('pos', '0'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:bgFillStyleLst > a:gradFill > a:gsLst > a:gs
        $xml->startElement('a:schemeClr'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:bgFillStyleLst > a:gradFill > a:gsLst > a:gs > a:schemeClr
        $xml->writeAttribute('val', 'phClr'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:bgFillStyleLst > a:gradFill > a:gsLst > a:gs > a:schemeClr
        $xml->startElement('a:tint'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:bgFillStyleLst > a:gradFill > a:gsLst > a:gs > a:schemeClr > a:tint
        $xml->writeAttribute('val', '80000'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:bgFillStyleLst > a:gradFill > a:gsLst > a:gs > a:schemeClr > a:tint
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:bgFillStyleLst > a:gradFill > a:gsLst > a:gs > a:schemeClr > a:tint
        $xml->startElement('a:satMod'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:bgFillStyleLst > a:gradFill > a:gsLst > a:gs > a:schemeClr > a:satMod
        $xml->writeAttribute('val', '300000'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:bgFillStyleLst > a:gradFill > a:gsLst > a:gs > a:schemeClr > a:satMod
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:bgFillStyleLst > a:gradFill > a:gsLst > a:gs > a:schemeClr > a:satMod
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:bgFillStyleLst > a:gradFill > a:gsLst > a:gs > a:schemeClr
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:bgFillStyleLst > a:gradFill > a:gsLst > a:gs
        $xml->startElement('a:gs'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:bgFillStyleLst > a:gradFill > a:gsLst > a:gs
        $xml->startElement('a:schemeClr'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:bgFillStyleLst > a:gradFill > a:gsLst > a:gs > a:schemeClr
        $xml->writeAttribute('val', 'phClr'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:bgFillStyleLst > a:gradFill > a:gsLst > a:gs > a:schemeClr
        $xml->startElement('a:shade'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:bgFillStyleLst > a:gradFill > a:gsLst > a:gs > a:schemeClr > a:shade
        $xml->writeAttribute('val', '30000'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:bgFillStyleLst > a:gradFill > a:gsLst > a:gs > a:schemeClr > a:shade
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:bgFillStyleLst > a:gradFill > a:gsLst > a:gs > a:schemeClr > a:shade
        $xml->startElement('a:satMod'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:bgFillStyleLst > a:gradFill > a:gsLst > a:gs > a:schemeClr > a:satMod
        $xml->writeAttribute('val', '200000'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:bgFillStyleLst > a:gradFill > a:gsLst > a:gs > a:schemeClr > a:satMod
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:bgFillStyleLst > a:gradFill > a:gsLst > a:gs > a:schemeClr > a:satMod
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:bgFillStyleLst > a:gradFill > a:gsLst > a:gs > a:schemeClr
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:bgFillStyleLst > a:gradFill > a:gsLst > a:gs
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:bgFillStyleLst > a:gradFill
        $xml->startElement('a:path'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:bgFillStyleLst > a:gradFill > a:path
        $xml->writeAttribute('path', 'circle'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:bgFillStyleLst > a:gradFill > a:path
        $xml->startElement('a:fillToRect'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:bgFillStyleLst > a:gradFill > a:path > a:fillToRect
        $xml->writeAttribute('l', '50000'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:bgFillStyleLst > a:gradFill > a:path > a:fillToRect
        $xml->writeAttribute('t', '50000'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:bgFillStyleLst > a:gradFill > a:path > a:fillToRect
        $xml->writeAttribute('r', '50000'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:bgFillStyleLst > a:gradFill > a:path > a:fillToRect
        $xml->writeAttribute('b', '50000'); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:bgFillStyleLst > a:gradFill > a:path > a:fillToRect
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:bgFillStyleLst > a:gradFill > a:gsLst > a:path > a:fillToRect
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:bgFillStyleLst > a:gradFill > a:gsLst > a:path
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:bgFillStyleLst > a:gradFill
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme > a:bgFillStyleLst
        $xml->endElement(); //a:theme > a:themeElements > a:fontScheme > a:fmtScheme
        $xml->endElement(); //a:theme > a:themeElements
        $xml->startElement('a:objectDefaults'); //a:theme > a:objectDefaults
        $xml->endElement(); //a:theme > a:objectDefaults
        $xml->startElement('a:extraClrSchemeLst'); //a:theme > a:extraClrSchemeLst
        $xml->endElement(); //a:theme > a:extraClrSchemeLst
        $xml->endElement(); //a:theme
        return $xml;
    }

    private function generateStylesXML()
    {

        $xml = new XMLWriter;
        $xml->openMemory();
        $xml->setIndent(TRUE);
        $xml->startDocument('1.0', 'UTF-8', 'yes');
        $xml->startElement('styleSheet'); //styleSheet
        $xml->writeAttribute(
                'xmlns'
                , 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
        ); //styleSheet
        $xml->writeAttribute(
                'xmlns:mc'
                , 'http://schemas.openxmlformats.org/markup-compatibility/2006'
        ); //styleSheet
        $xml->writeAttribute(
                'mc:Ignorable'
                , 'x14ac'
        ); //styleSheet
        $xml->writeAttribute(
                'xmlns:x14ac'
                , 'http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac'
        ); //styleSheet

        $xml->startElement('fonts'); //styleSheet > fonts
        $xml->writeAttribute('count', 1); //styleSheet > fonts
        $xml->writeAttribute('x14ac:knownFonts', 1); //styleSheet > fonts
        $xml->startElement('font'); //styleSheet > fonts > font 
        $xml->startElement('name'); //styleSheet > fonts > font > name 
        $xml->writeAttribute('val', 'Calibri'); //styleSheet > fonts > font > name 
        $xml->endElement(); //styleSheet > fonts > font > name 
        $xml->startElement('sz'); //styleSheet > fonts > font > sz 
        $xml->writeAttribute('val', 11); //styleSheet > fonts > font > sz 
        $xml->endElement(); //styleSheet > fonts > font > sz 
        $xml->startElement('b'); //styleSheet > fonts > font > b 
        $xml->writeAttribute('val', 0); //styleSheet > fonts > font > b 
        $xml->endElement(); //styleSheet > fonts > font > b 

        $xml->endElement(); //styleSheet > fonts > font 
        $xml->endElement(); //styleSheet > fonts
        $xml->startElement('fills'); //styleSheet > fills 
        $xml->writeAttribute('count', '2'); //styleSheet > fills 
        $xml->startElement('fill'); //styleSheet > fills > fill
        $xml->startElement('patternFill'); //styleSheet > fills > fill > patternFill
        $xml->writeAttribute('patternType', 'none'); //styleSheet > fills > fill > patternFill
        $xml->startElement('fgColor'); //styleSheet > fills > fill > patternFill > fgColor
        $xml->writeAttribute('rgb', 'FFFFFFFF'); //styleSheet > fills > fill > patternFill > fgColor
        $xml->endElement(); //styleSheet > fills > fill > patternFill > fgColor
        $xml->startElement('bgColor'); //styleSheet > fills > fill > patternFill > bgColor
        $xml->writeAttribute('rgb', 'FF000000'); //styleSheet > fills > fill > patternFill > bgColor
        $xml->endElement(); //styleSheet > fills > fill > patternFill > bgColor
        $xml->endElement(); //styleSheet > fills > fill > patternFill 
        $xml->endElement(); //styleSheet > fills > fill 
        $xml->startElement('fill'); //styleSheet > fills > fill
        $xml->startElement('patternFill'); //styleSheet > fills > fill > patternFill
        $xml->writeAttribute('patternType', 'gray125'); //styleSheet > fills > fill > patternFill
        $xml->startElement('fgColor'); //styleSheet > fills > fill > patternFill > fgColor
        $xml->writeAttribute('rgb', 'FFFFFFFF'); //styleSheet > fills > fill > patternFill > fgColor
        $xml->endElement(); //styleSheet > fills > fill > patternFill > fgColor
        $xml->startElement('bgColor'); //styleSheet > fills > fill > patternFill > bgColor
        $xml->writeAttribute('rgb', 'FF000000'); //styleSheet > fills > fill > patternFill > bgColor
        $xml->endElement(); //styleSheet > fills > fill > patternFill > bgColor
        $xml->endElement(); //styleSheet > fills > fill > patternFill 
        $xml->endElement(); //styleSheet > fills > fill 
        $xml->endElement(); //styleSheet > fills 
        $xml->startElement('borders'); //styleSheet > borders
        $xml->writeAttribute('count', 1); //styleSheet > borders
        $xml->startElement('border'); //styleSheet > borders > border
        $xml->endElement(); //styleSheet > borders > border
        $xml->endElement(); //styleSheet > borders
        $xml->startElement('cellStyleXfs'); //styleSheet > cellStyleXfs
        $xml->writeAttribute('count', 1); //styleSheet > cellStyleXfs
        $xml->startElement('xf'); //styleSheet > cellStyleXfs > xf
        $xml->writeAttribute('numFmtId', 0); //styleSheet > cellStyleXfs > xf
        $xml->writeAttribute('fontId', 0); //styleSheet > cellStyleXfs > xf
        $xml->writeAttribute('fillId', 0); //styleSheet > cellStyleXfs > xf
        $xml->writeAttribute('borderId', 0); //styleSheet > cellStyleXfs > xf
        $xml->endElement(); //styleSheet > cellStyleXfs > xf
        $xml->endElement(); //styleSheet > cellStyleXfs
        $xml->startElement('cellXfs'); //styleSheet > cellXfs
        $xml->writeAttribute('count', 1); //styleSheet > cellXfs
        $xml->startElement('xf'); //styleSheet > cellXfs > xf
        $xml->writeAttribute('xfId', 0); //styleSheet > cellXfs > xf
        $xml->writeAttribute('fontId', 0); //styleSheet > cellXfs > xf
        $xml->writeAttribute('numFmtId', 0); //styleSheet > cellXfs > xf
        $xml->writeAttribute('fillId', 0); //styleSheet > cellXfs > xf
        $xml->writeAttribute('borderId', 0); //styleSheet > cellXfs > xf
        $xml->writeAttribute('applyFont', 0); //styleSheet > cellXfs > xf
        $xml->writeAttribute('applyNumberFormat', 0); //styleSheet > cellXfs > xf
        $xml->writeAttribute('applyFill', 0); //styleSheet > cellXfs > xf
        $xml->writeAttribute('applyBorder', 0); //styleSheet > cellXfs > xf
        $xml->writeAttribute('applyAlignment', 0); //styleSheet > cellXfs > xf
        $xml->startElement('alignment'); //styleSheet > cellXfs > xf > alignment
        $xml->writeAttribute('horizontal', 'general'); //styleSheet > cellXfs > xf > alignment
        $xml->writeAttribute('vertical', 'bottom'); //styleSheet > cellXfs > xf > alignment
        $xml->writeAttribute('textRotation', '0'); //styleSheet > cellXfs > xf > alignment
        $xml->writeAttribute('wrapText', 'false'); //styleSheet > cellXfs > xf > alignment
        $xml->writeAttribute('shrinkToFit', 'false'); //styleSheet > cellXfs > xf > alignment
        $xml->endElement(); //styleSheet > cellXfs > xf > alignment
        $xml->endElement(); //styleSheet > cellXfs > xf
        $xml->endElement(); //styleSheet > cellXfs
        $xml->startElement('cellStyles'); //styleSheet > cellStyles
        $xml->writeAttribute('count', '1'); //styleSheet > cellStyles
        $xml->startElement('cellStyle'); //styleSheet > cellStyles > cellStyle
        $xml->writeAttribute('name', 'Normal'); //styleSheet > cellStyles > cellStyle
        $xml->writeAttribute('xfId', '0'); //styleSheet > cellStyles > cellStyle
        $xml->writeAttribute('builtinId', '0'); //styleSheet > cellStyles > cellStyle
        $xml->endElement(); //styleSheet > cellStyles > cellStyle
        $xml->endElement(); //styleSheet > cellStyles
        $xml->startElement('dxfs'); //styleSheet > dxfs
        $xml->writeAttribute('count', '0'); //styleSheet > dxfs
        $xml->endElement(); //styleSheet > dxfs
        $xml->startElement('tableStyles'); //styleSheet > tableStyles
        $xml->writeAttribute('defaultTableStyle', 'TableStyleMedium9'); //styleSheet > tableStyles
        $xml->writeAttribute('defaultPivotStyle', 'PivotTableStyle1'); //styleSheet > tableStyles
        $xml->endElement(); //styleSheet > tableStyles
        $xml->endElement(); //styleSheet
        return $xml;
    }

    private function generateDocPropsAppXML()
    {
        $xml = new XMLWriter;
        $xml->openMemory();
        $xml->setIndent(TRUE);
        $xml->startDocument('1.0', 'UTF-8', 'yes');
        $xml->startElement('Properties'); //Properties
        $xml->writeAttribute(
                'xmlns'
                , 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties'
        ); //Properties
        $xml->writeAttribute(
                'xmlns:vt'
                , 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes'
        ); //Properties
        $xml->startElement('Application'); //Properties > Application
        $xml->writeRaw('Microsoft Excel'); //Properties > Application
        $xml->endElement(); //Properties > Application
        $xml->startElement('DocSecurity'); //Properties > DocSecurity
        $xml->writeRaw('0'); //Properties > DocSecurity
        $xml->endElement(); //Properties > DocSecurity
        $xml->startElement('ScaleCrop'); //Properties > ScaleCrop
        $xml->writeRaw('false'); //Properties > ScaleCrop
        $xml->endElement(); //Properties > ScaleCrop
        $xml->startElement('HeadingPairs'); //Properties > HeadingPairs
        $xml->startElement('vt:vector'); //Properties > HeadingPairs > vt:vector
        $xml->writeAttribute('size', '2'); //Properties > HeadingPairs > vt:vector
        $xml->writeAttribute('baseType', 'variant'); //Properties > HeadingPairs > vt:vector
        $xml->startElement('vt:variant'); //Properties > HeadingPairs > vt:vector > vt:variant
        $xml->startElement('vt:lpstr'); //Properties > HeadingPairs > vt:vector > vt:variant > vt:lpstr
        $xml->writeRaw('Worksheets'); //Properties > HeadingPairs > vt:vector > vt:variant > vt:lpstr
        $xml->endElement(); //Properties > HeadingPairs > vt:vector > vt:variant > vt:lpstr
        $xml->endElement(); //Properties > HeadingPairs > vt:vector > vt:variant
        $xml->startElement('vt:variant'); //Properties > HeadingPairs > vt:vector > vt:variant
        $xml->startElement('vt:i4'); //Properties > HeadingPairs > vt:vector > vt:variant > vt:i4
        $xml->writeRaw('1'); //Properties > HeadingPairs > vt:vector > vt:variant > vt:i4
        $xml->endElement(); //Properties > HeadingPairs > vt:vector > vt:variant > vt:i4
        $xml->endElement(); //Properties > HeadingPairs > vt:vector > vt:variant
        $xml->endElement(); //Properties > HeadingPairs > vt:vector
        $xml->endElement(); //Properties > HeadingPairs
        $xml->startElement('TitlesOfParts'); //Properties > HeadingPairs > TitlesOfParts
        $xml->startElement('vt:vector'); //Properties > HeadingPairs > TitlesOfParts > vt:vector
        $xml->writeAttribute('size', '1'); //Properties > HeadingPairs > TitlesOfParts > vt:vector
        $xml->writeAttribute('baseType', 'lpstr'); //Properties > HeadingPairs > TitlesOfParts > vt:vector
        $xml->startElement('vt:lpstr'); //Properties > HeadingPairs > TitlesOfParts > vt:vector > vt:lpstr
        $xml->writeRaw('Worksheet'); //Properties > HeadingPairs > TitlesOfParts > vt:vector > vt:lpstr
        $xml->endElement(); //Properties > HeadingPairs > TitlesOfParts > vt:vector > vt:lpstr
        $xml->endElement(); //Properties > HeadingPairs > TitlesOfParts > vt:vector
        $xml->endElement(); //Properties > HeadingPairs > TitlesOfParts
        $xml->startElement('Company'); //Properties > HeadingPairs > Company 
        $xml->endElement(); //Properties > HeadingPairs > Company
        $xml->startElement('LinksUpToDate'); //Properties > HeadingPairs > LinksUpToDate
        $xml->writeRaw('false'); //Properties > HeadingPairs > LinksUpToDate
        $xml->endElement(); //Properties > HeadingPairs > LinksUpToDate
        $xml->startElement('SharedDoc'); //Properties > HeadingPairs > SharedDoc
        $xml->writeRaw('false'); //Properties > HeadingPairs > SharedDoc
        $xml->endElement(); //Properties > HeadingPairs > SharedDoc
        $xml->startElement('HyperlinksChanged'); //Properties > HeadingPairs > HyperlinksChanged
        $xml->writeRaw('false'); //Properties > HeadingPairs > HyperlinksChanged
        $xml->endElement(); //Properties > HeadingPairs > HyperlinksChanged
        $xml->startElement('AppVersion'); //Properties > HeadingPairs > AppVersion
        $xml->writeRaw('14.0300'); //Properties > HeadingPairs > AppVersion
        $xml->endElement(); //Properties > HeadingPairs > AppVersion
        $xml->endElement(); //Properties
        return $xml;
    }

    public function save()
    {
        if (!$this->fileName) {
            throw new \RuntimeException('Not isset file to save');
        }
        $zip = new \ZipArchive;
        $res = $zip->open($this->fileName, \ZipArchive::CREATE);
        if (TRUE === $res) {
            $zip->addFromString(
                    "docProps/core.xml"
                    , $this->generateCoreXML()->outputMemory()
            /*
              '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
              <cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:creator>Unknown Creator</dc:creator><cp:lastModifiedBy>Unknown Creator</cp:lastModifiedBy><dcterms:created xsi:type="dcterms:W3CDTF">2016-08-11T15:58:11+03:00</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">2016-08-11T15:58:11+03:00</dcterms:modified><dc:title>Untitled Spreadsheet</dc:title><dc:description></dc:description><dc:subject></dc:subject><cp:keywords></cp:keywords><cp:category></cp:category></cp:coreProperties>'
             * 
             */);
            $zip->addFromString(
                    "docProps/custom.xml", '');
            $zip->addFromString(
                    "xl/theme/theme1.xml"
                    , $this->generateThemeXML()->outputMemory()
                    /* '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                      <a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme"><a:themeElements><a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="1F497D"/></a:dk2><a:lt2><a:srgbClr val="EEECE1"/></a:lt2><a:accent1><a:srgbClr val="4F81BD"/></a:accent1><a:accent2><a:srgbClr val="C0504D"/></a:accent2><a:accent3><a:srgbClr val="9BBB59"/></a:accent3><a:accent4><a:srgbClr val="8064A2"/></a:accent4><a:accent5><a:srgbClr val="4BACC6"/></a:accent5><a:accent6><a:srgbClr val="F79646"/></a:accent6><a:hlink><a:srgbClr val="0000FF"/></a:hlink><a:folHlink><a:srgbClr val="800080"/></a:folHlink></a:clrScheme><a:fontScheme name="Office"><a:majorFont><a:latin typeface="Cambria"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="?? ?????"/><a:font script="Hang" typeface="?? ??"/><a:font script="Hans" typeface="??"/><a:font script="Hant" typeface="????"/><a:font script="Arab" typeface="Times New Roman"/><a:font script="Hebr" typeface="Times New Roman"/><a:font script="Thai" typeface="Tahoma"/><a:font script="Ethi" typeface="Nyala"/><a:font script="Beng" typeface="Vrinda"/><a:font script="Gujr" typeface="Shruti"/><a:font script="Khmr" typeface="MoolBoran"/><a:font script="Knda" typeface="Tunga"/><a:font script="Guru" typeface="Raavi"/><a:font script="Cans" typeface="Euphemia"/><a:font script="Cher" typeface="Plantagenet Cherokee"/><a:font script="Yiii" typeface="Microsoft Yi Baiti"/><a:font script="Tibt" typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface="MV Boli"/><a:font script="Deva" typeface="Mangal"/><a:font script="Telu" typeface="Gautami"/><a:font script="Taml" typeface="Latha"/><a:font script="Syrc" typeface="Estrangelo Edessa"/><a:font script="Orya" typeface="Kalinga"/><a:font script="Mlym" typeface="Kartika"/><a:font script="Laoo" typeface="DokChampa"/><a:font script="Sinh" typeface="Iskoola Pota"/><a:font script="Mong" typeface="Mongolian Baiti"/><a:font script="Viet" typeface="Times New Roman"/><a:font script="Uigh" typeface="Microsoft Uighur"/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="?? ?????"/><a:font script="Hang" typeface="?? ??"/><a:font script="Hans" typeface="??"/><a:font script="Hant" typeface="????"/><a:font script="Arab" typeface="Arial"/><a:font script="Hebr" typeface="Arial"/><a:font script="Thai" typeface="Tahoma"/><a:font script="Ethi" typeface="Nyala"/><a:font script="Beng" typeface="Vrinda"/><a:font script="Gujr" typeface="Shruti"/><a:font script="Khmr" typeface="DaunPenh"/><a:font script="Knda" typeface="Tunga"/><a:font script="Guru" typeface="Raavi"/><a:font script="Cans" typeface="Euphemia"/><a:font script="Cher" typeface="Plantagenet Cherokee"/><a:font script="Yiii" typeface="Microsoft Yi Baiti"/><a:font script="Tibt" typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface="MV Boli"/><a:font script="Deva" typeface="Mangal"/><a:font script="Telu" typeface="Gautami"/><a:font script="Taml" typeface="Latha"/><a:font script="Syrc" typeface="Estrangelo Edessa"/><a:font script="Orya" typeface="Kalinga"/><a:font script="Mlym" typeface="Kartika"/><a:font script="Laoo" typeface="DokChampa"/><a:font script="Sinh" typeface="Iskoola Pota"/><a:font script="Mong" typeface="Mongolian Baiti"/><a:font script="Viet" typeface="Arial"/><a:font script="Uigh" typeface="Microsoft Uighur"/></a:minorFont></a:fontScheme><a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="50000"/><a:satMod val="300000"/></a:schemeClr></a:gs><a:gs pos="35000"><a:schemeClr val="phClr"><a:tint val="37000"/><a:satMod val="300000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:tint val="15000"/><a:satMod val="350000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="16200000" scaled="1"/></a:gradFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:shade val="51000"/><a:satMod val="130000"/></a:schemeClr></a:gs><a:gs pos="80000"><a:schemeClr val="phClr"><a:shade val="93000"/><a:satMod val="130000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="94000"/><a:satMod val="135000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="16200000" scaled="0"/></a:gradFill></a:fillStyleLst><a:lnStyleLst><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"><a:shade val="95000"/><a:satMod val="105000"/></a:schemeClr></a:solidFill><a:prstDash val="solid"/></a:ln><a:ln w="25400" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln><a:ln w="38100" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst><a:outerShdw blurRad="40000" dist="20000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="38000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="35000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="40000" dist="23000" dir="5400000" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="35000"/></a:srgbClr></a:outerShdw></a:effectLst><a:scene3d><a:camera prst="orthographicFront"><a:rot lat="0" lon="0" rev="0"/></a:camera><a:lightRig rig="threePt" dir="t"><a:rot lat="0" lon="0" rev="1200000"/></a:lightRig></a:scene3d><a:sp3d><a:bevelT w="63500" h="25400"/></a:sp3d></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="40000"/><a:satMod val="350000"/></a:schemeClr></a:gs><a:gs pos="40000"><a:schemeClr val="phClr"><a:tint val="45000"/><a:shade val="99000"/><a:satMod val="350000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="20000"/><a:satMod val="255000"/></a:schemeClr></a:gs></a:gsLst><a:path path="circle"><a:fillToRect l="50000" t="-80000" r="50000" b="180000"/></a:path></a:gradFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="80000"/><a:satMod val="300000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="30000"/><a:satMod val="200000"/></a:schemeClr></a:gs></a:gsLst><a:path path="circle"><a:fillToRect l="50000" t="50000" r="50000" b="50000"/></a:path></a:gradFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>'
                     */
            );
            $zip->addFromString(
                    "xl/styles.xml", $this->generateStylesXML()->outputMemory()
            );
            $zip->addFromString(
                    "xl/workbook.xml"
                    , $this->generateWorkbookXML()->outputMemory()
                    /* '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                      <workbook xml:space="preserve" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><fileVersion appName="xl" lastEdited="4" lowestEdited="4" rupBuild="4505"/><workbookPr codeName="ThisWorkbook"/><bookViews><workbookView activeTab="0" autoFilterDateGrouping="1" firstSheet="0" minimized="0" showHorizontalScroll="1" showSheetTabs="1" showVerticalScroll="1" tabRatio="600" visibility="visible"/></bookViews><sheets><sheet name="Worksheet" sheetId="1" r:id="rId4"/></sheets><definedNames/><calcPr calcId="124519" calcMode="auto" fullCalcOnLoad="1"/></workbook>'
                     */
            );
            $zip->addFromString(
                    "xl/worksheets/sheet1.xml"
                    , $this->generateXMLRows()
                            ->outputMemory()
            );
            $zip->addFromString(
                    "xl/sharedStrings.xml"
                    , $this->generateXMLHeaderRow()
                            ->outputMemory()
            );

            $zip->addFromString(
                    '[Content_Types].xml'
                    , $this->generateContentTypesXML()
                            ->outputMemory()
            );
            $zip->addFromString(
                    '_rels/.rels'
                    , $this->generateRelationshipsXML()
                            ->outputMemory()
            );
            $zip->addFromString(
                    'xl/_rels/workbook.xml.rels'
                    , $this->generateRefWorkbookXML()->outputMemory()
            );
            $zip->addFromString(
                    'docProps/app.xml'
                    , $this->generateDocPropsAppXML()->outputMemory()
            );

            $zip->close();
            return TRUE;
        }
        return FALSE;
    }

}

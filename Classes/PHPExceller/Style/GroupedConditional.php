<?php
namespace PHPExceller\Style;

use PHPExceller\Style\PHPExceller_Style_Supervisor;
use PHPExceller\PHPExceller_IComparable;
use PHPExceller\Style\PHPExceller_Style_CFVOType;
use PHPExceller\PHPExceller_Cell;

class PHPExceller_Style_GroupedConditional extends PHPExceller_Style_Supervisor
{
    /**
    * cellReference definition : (eg:A1:C5)
    *
    * @var string
    */
    protected $cellReference;

    public function __construct()
    {
        // set these cfvo values by default because without them no IconSet is shown
        $this->cfvos = array(PHPExceller_Style_CFVOType::fromString('min'),PHPExceller_Style_CFVOType::fromString('max'));
        // uniq id
        $this->id = md5(uniqid('',true));
    }

    public function removeCellReference($position)
    {
        if ((is_array($this->cellReference)) && (in_array($position, $this->cellReference)))
        {
            $index = array_search($position, $this->cellReference);
            if ($index !== FALSE ) { unset($this->cellReference[$index]); }
        }
        elseif ($this->cellReference == $position)
        {
            $this->cellReference = null;
        }
    }

    public function updateCellReference($oldreference, $newreference)
    {
        if ((is_array($this->cellReference)) && (in_array($oldreference, $this->cellReference)))
        {
            $index = array_search($oldreference, $this->cellReference);
            if($index !== FALSE) 
            { 
                $this->cellReference[$index] = $newreference; 
            }
        }
        elseif ($this->cellReference == $oldreference)
        {
            $this->cellReference = $newreference;
        }
    }
    
        /*
     * set the group of cells that this colorscale setting applies to
     *
     * <code>
     * $worksheetstyles = $objPHPExceller->getActiveSheet()->getConditionalStyles();
     * if ($worksheetstyles[0]->getConditionType() == PHPExceller_Style_Conditional::CONDITION_COLORSCALE) 
     * { 
     *     $worksheetstyles[0]->setCellReference('A1:A5') ; 
     * }
     * </code>
     *
     * @return void
    */
    public function setCellReference($cellReference = null) 
    {
        $this->cellReference = array();
        list($startColumn, $endColumn, $startRow, $endRow) = $this->extractRowColumns($cellReference);
        for($row = $startRow;$row <= $endRow;$row++)
        {
            for($column = $startColumn; $column <= $endColumn; $column++)
            {
                $position = PHPExceller_Cell::stringFromColumnIndex($column-1).$row;
                array_push($this->cellReference, $position);
            }
        }
    }

    protected function _extractRowColumns($cellReference)
    {
        $parts = explode(':',$cellReference);
        // start
        list($startColumChar, $startRow) = PHPExceller_Cell::coordinateFromString($parts[0]);
        $startColumn = PHPExceller_Cell::columnIndexFromString($startColumChar);
        if (count($parts) > 1)
        {
            // range
            list($endColumChar, $endRow) = PHPExceller_Cell::coordinateFromString($parts[1]);
            $endColumn = PHPExceller_Cell::columnIndexFromString($endColumChar);
        }
        else
        {
            // single reference
            $endColumn = $startColumn;
            $endRow = $startRow;
        }
        return array($startColumn, $endColumn, $startRow, $endRow);
    }

    /*
     * get the group of cells that this IconSet setting applies to
     * NOTE : THE RANGE IS DETERMINED ON THE MIX/MAX ROW/COLUMNS AND PRESUMES THE CELLS IN BETWEEN
     *       ARE ALSO SUBJECT TO THIS CONDITIONAL FORMATTING
     * <code>
     * $worksheetstyles = $objPHPExceller->getActiveSheet()->getConditionalStyles();
     * if ($worksheetstyles[0]->getConditionType() == PHPExceller_Style_Conditional::CONDITION_IconSet) { $color = $worksheetstyles[0]->getCellReference() ; }
     * </code>
     *
     * @return string
    */
    public function getCellReference() 
    {
        $minColumn = null; $maxColumn = null; $minRow = null; $maxRow = null;
                $cellRefsByRow = [];
        foreach($this->cellReference as $cellRef)
        {
            list($columnChar, $row) = PHPExceller_Cell::coordinateFromString($cellRef);
            $column = PHPExceller_Cell::columnIndexFromString($columnChar);
            $minColumn = (($minColumn == null)||($column < $minColumn))?$column:$minColumn;
            $maxColumn = (($maxColumn == null)||($column > $maxColumn))?$column:$maxColumn;
            $minRow = (($minRow == null)||($row < $minRow))?$row:$minRow;
            $maxRow = (($maxRow == null)||($row > $maxRow))?$row:$maxRow;
                        if(!isset($cellRefsByRow[$row])) { $cellRefsByRow[$row] = []; }
                        $cellRefsByRow[$row][] = $column;
        }

        if (($minRow == $maxRow) && ($minColumn == $maxColumn))
        {
            // single reference
            $result = PHPExceller_Cell::stringFromColumnIndex($minColumn-1).$minRow;
        }
        else //if (count($this->cellReference) == (($maxColumn - $minColumn + 1) * ($maxRow - $minRow + 1)))
        {
            // range as 1 single block
            $result = PHPExceller_Cell::stringFromColumnIndex($minColumn-1).$minRow.':'.PHPExceller_Cell::stringFromColumnIndex($maxColumn-1).$maxRow;
        }
/*                else
                {
                        // various cells and blocks -> optimize blocks
                        $result = $this->getOptimizedReferences($cellRefsByRow);
                }
*/
        return $result;
    }
/*
    protected function getOptimizedReferences($cellRefsByRow)
    {
    }
*/
    /*
     * Get the hashcode for this object
     *
     * <code>
     * $worksheetstyles = $objPHPExceller->getActiveSheet()->getConditionalStyles();
     * $hashcode = $worksheetstyles[0]->getHashCode();.
     * </code>
     *
     * @return string containing md5 hashcode of the object
    */
    public function getHashCode()
    {
        return $this->id;
    }

        /*
         * Get a unique classID for this object
         *
         * @return string CLASSID v3 string : e.g. {1546058F-5A25-4334-85AE-E68F2A44BBAF}
         *
         */
         public function getClassID()
         {
            $hash = $this->getHashCode();
            return strtoupper(sprintf('{%08s-%04s-%04x-%04x-%12s}',
                                      substr($hash, 0, 8),
                                      substr($hash, 8, 4),
                                      (hexdec(substr($hash, 12, 4)) & 0x0fff) | 0x3000,
                                      (hexdec(substr($hash, 16, 4)) & 0x3fff) | 0x8000,
                                       substr($hash, 20, 12)));
        }
}

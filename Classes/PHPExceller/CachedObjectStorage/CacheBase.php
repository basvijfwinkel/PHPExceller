<?php
namespace PHPExceller\CachedObjectStorage;

use PHPExceller\Cell;
use PHPExceller\Worksheet;

/**
 * Based on PHPExcel_CachedObjectStorage_CacheBase
 */
abstract class CacheBase
{
    /**
     * Parent worksheet
     *
     * @var PHPExceller\Worksheet
     */
    protected $parent;

    /**
     * The currently active Cell
     *
     * @var PHPExceller\Cell
     */
    protected $currentObject = null;

    /**
     * Coordinate address of the currently active Cell
     *
     * @var string
     */
    protected $currentObjectID = null;

    /**
     * Flag indicating whether the currently active Cell requires saving
     *
     * @var boolean
     */
    protected $currentCellIsDirty = true;

    /**
     * An array of cells or cell pointers for the worksheet cells held in this cache,
     *        and indexed by their coordinate address within the worksheet
     *
     * @var array of mixed
     */
    protected $cellCache = array();

    /**
     * Initialise this new cell collection
     *
     * @param    PHPExceller\Worksheet    $parent        The worksheet for this cell collection
     */
    public function __construct(Worksheet $parent)
    {
        //    Set our parent worksheet.
        //    This is maintained within the cache controller to facilitate re-attaching it to PHPExceller\Cell objects when
        //        they are woken from a serialized state
        $this->parent = $parent;
    }

    /**
     * Return the parent worksheet for this cell collection
     *
     * @return    PHPExceller\Worksheet
     */
    public function getParent()
    {
        return $this->parent;
    }

    /**
     * Is a value set in the current PHPExceller\CachedObjectStorage\ICache for an indexed cell?
     *
     * @param    string        $pCoord        Coordinate address of the cell to check
     * @return    boolean
     */
    public function isDataSet($pCoord)
    {
        if ($pCoord === $this->currentObjectID)
        {
            return true;
        }
        //    Check if the requested entry exists in the cache
        return isset($this->cellCache[$pCoord]);
    }

    /**
     * Move a cell object from one address to another
     *
     * @param    string        $fromAddress    Current address of the cell to move
     * @param    string        $toAddress        Destination address of the cell to move
     * @return    boolean
     */
    public function moveCell($fromAddress, $toAddress)
    {
        if ($fromAddress === $this->currentObjectID)
        {
            $this->currentObjectID = $toAddress;
        }
        $this->currentCellIsDirty = true;
        if (isset($this->cellCache[$fromAddress]))
        {
            $this->cellCache[$toAddress] = &$this->cellCache[$fromAddress];
            unset($this->cellCache[$fromAddress]);
        }

        return true;
    }

    /**
     * Add or Update a cell in cache
     *
     * @param    PHPExceller\Cell    $cell        Cell to update
     * @return    PHPExceller\Cell
     * @throws    PHPExceller\Exception
     */
    public function updateCacheData(Cell $cell)
    {
        return $this->addCacheData($cell->getCoordinate(), $cell);
    }

    /**
     * Delete a cell in cache identified by coordinate address
     *
     * @param    string            $pCoord        Coordinate address of the cell to delete
     * @throws    PHPExceller\Exception
     */
    public function deleteCacheData($pCoord)
    {
        if ($pCoord === $this->currentObjectID && !is_null($this->currentObject))
        {
            $this->currentObject->detach();
            $this->currentObjectID = $this->currentObject = null;
        }

        if (is_object($this->cellCache[$pCoord]))
        {
            $this->cellCache[$pCoord]->detach();
            unset($this->cellCache[$pCoord]);
        }
        $this->currentCellIsDirty = false;
    }

    /**
     * Get a list of all cell addresses currently held in cache
     *
     * @return    string[]
     */
    public function getCellList()
    {
        return array_keys($this->cellCache);
    }

    /**
     * Sort the list of all cell addresses currently held in cache by row and column
     *
     * @return    string[]
     */
    public function getSortedCellList()
    {
        $sortKeys = array();
        foreach ($this->getCellList() as $coord)
        {
            sscanf($coord, '%[A-Z]%d', $column, $row);
            $sortKeys[sprintf('%09d%3s', $row, $column)] = $coord;
        }
        ksort($sortKeys);

        return array_values($sortKeys);
    }

    /**
     * Get highest worksheet column and highest row that have cell records
     *
     * @return array Highest column name and highest row number
     */
    public function getHighestRowAndColumn()
    {
        // Lookup highest column and highest row
        $col = array('A' => '1A');
        $row = array(1);
        foreach ($this->getCellList() as $coord)
        {
            sscanf($coord, '%[A-Z]%d', $c, $r);
            $row[$r] = $r;
            $col[$c] = strlen($c).$c;
        }
        if (!empty($row))
        {
            // Determine highest column and row
            $highestRow = max($row);
            $highestColumn = substr(max($col), 1);
        }

        return array(
                      'row'    => $highestRow,
                      'column' => $highestColumn
                    );
    }

    /**
     * Return the cell address of the currently active cell object
     *
     * @return    string
     */
    public function getCurrentAddress()
    {
        return $this->currentObjectID;
    }

    /**
     * Return the column address of the currently active cell object
     *
     * @return    string
     */
    public function getCurrentColumn()
    {
        sscanf($this->currentObjectID, '%[A-Z]%d', $column, $row);
        return $column;
    }

    /**
     * Return the row address of the currently active cell object
     *
     * @return    integer
     */
    public function getCurrentRow()
    {
        sscanf($this->currentObjectID, '%[A-Z]%d', $column, $row);
        return (integer) $row;
    }

    /**
     * Get highest worksheet column
     *
     * @param   string     $row        Return the highest column for the specified row,
     *                                     or the highest column of any row if no row number is passed
     * @return  string     Highest column name
     */
    public function getHighestColumn($row = null)
    {
        if ($row == null)
        {
            $colRow = $this->getHighestRowAndColumn();
            return $colRow['column'];
        }

        $columnList = array(1);
        foreach ($this->getCellList() as $coord)
        {
            sscanf($coord, '%[A-Z]%d', $c, $r);
            if ($r != $row)
            {
                continue;
            }
            $columnList[] = Cell::columnIndexFromString($c);
        }
        return Cell::stringFromColumnIndex(max($columnList) - 1);
    }

    /**
     * Get highest worksheet row
     *
     * @param   string     $column     Return the highest row for the specified column,
     *                                     or the highest row of any column if no column letter is passed
     * @return  int        Highest row number
     */
    public function getHighestRow($column = null)
    {
        if ($column == null)
        {
            $colRow = $this->getHighestRowAndColumn();
            return $colRow['row'];
        }

        $rowList = array(0);
        foreach ($this->getCellList() as $coord)
        {
            sscanf($coord, '%[A-Z]%d', $c, $r);
            if ($c != $column)
            {
                continue;
            }
            $rowList[] = $r;
        }

        return max($rowList);
    }

    /**
     * Generate a unique ID for cache referencing
     *
     * @return string Unique Reference
     */
    protected function getUniqueID()
    {
        if (function_exists('posix_getpid'))
        {
            $baseUnique = posix_getpid();
        }
        else
        {
            $baseUnique = mt_rand();
        }
        return uniqid($baseUnique, true);
    }

    /**
     * Clone the cell collection
     *
     * @param    PHPExceller\Worksheet    $parent        The new worksheet
     * @return    void
     */
    public function copyCellCollection(Worksheet $parent)
    {
        $this->currentCellIsDirty;
        $this->storeData();

        $this->parent = $parent;
        if (($this->currentObject !== null) && (is_object($this->currentObject))) 
        {
            $this->currentObject->attach($this);
        }
    }

    /**
     * Remove a row, deleting all cells in that row
     *
     * @param string    $row    Row number to remove
     * @return void
     */
    public function removeRow($row)
    {
        foreach ($this->getCellList() as $coord)
        {
            sscanf($coord, '%[A-Z]%d', $c, $r);
            if ($r == $row)
            {
                $this->deleteCacheData($coord);
            }
        }
    }

    /**
     * Remove a column, deleting all cells in that column
     *
     * @param string    $column    Column ID to remove
     * @return void
     */
    public function removeColumn($column)
    {
        foreach ($this->getCellList() as $coord)
        {
            sscanf($coord, '%[A-Z]%d', $c, $r);
            if ($c == $column)
            {
                $this->deleteCacheData($coord);
            }
        }
    }

    /**
     * Identify whether the caching method is currently available
     * Some methods are dependent on the availability of certain extensions being enabled in the PHP build
     *
     * @return    boolean
     */
    public static function cacheMethodIsAvailable()
    {
        return true;
    }
}

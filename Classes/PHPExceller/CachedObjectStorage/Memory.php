<?php
namespace PHPExceller\CachedObjectStorage;

use PHPExceller\CachedObjectStorage\CacheBase;
use PHPExceller\CachedObjectStorage\ICache;
use PHPExceller\PHPExceller_Cell;
use PHPExceller\PHPExceller_Worksheet;

/**
 * Based on PHPExcel_CachedObjectStorage_Memory
 */
class Memory extends CacheBase implements ICache
{
    /**
     * Dummy method callable from CacheBase, but unused by Memory cache
     *
     * @return    void
     */
    protected function storeData()
    {
    }

    /**
     * Add or Update a cell in cache identified by coordinate address
     *
     * @param    string            $pCoord        Coordinate address of the cell to update
     * @param    PHPExceller\Cell    $cell        Cell to update
     * @return    PHPExceller\Cell
     * @throws    PHPExceller\Exception
     */
    public function addCacheData($pCoord, Cell $cell)
    {
        $this->cellCache[$pCoord] = $cell;

        //    Set current entry to the new/updated entry
        $this->currentObjectID = $pCoord;

        return $cell;
    }


    /**
     * Get cell at a specific coordinate
     *
     * @param     string             $pCoord        Coordinate of the cell
     * @throws     PHPExceller\Exception
     * @return     PHPExceller\Cell     Cell that was found, or null if not found
     */
    public function getCacheData($pCoord)
    {
        //    Check if the entry that has been requested actually exists
        if (!isset($this->cellCache[$pCoord]))
        {
            $this->currentObjectID = null;
            //    Return null if requested entry doesn't exist in cache
            return null;
        }

        //    Set current entry to the requested entry
        $this->currentObjectID = $pCoord;

        //    Return requested entry
        return $this->cellCache[$pCoord];
    }


    /**
     * Clone the cell collection
     *
     * @param    PHPExceller\Worksheet    $parent        The new worksheet
     */
    public function copyCellCollection(Worksheet $parent)
    {
        parent::copyCellCollection($parent);

        $newCollection = array();
        foreach ($this->cellCache as $k => &$cell)
        {
            $newCollection[$k] = clone $cell;
            $newCollection[$k]->attach($this);
        }

        $this->cellCache = $newCollection;
    }

    /**
     * Clear the cell collection and disconnect from our parent
     *
     */
    public function unsetWorksheetCells()
    {
        // Because cells are all stored as intact objects in memory, we need to detach each one from the parent
        foreach ($this->cellCache as $k => &$cell)
        {
            $cell->detach();
            $this->cellCache[$k] = null;
        }
        unset($cell);

        $this->cellCache = array();

        //    detach ourself from the worksheet, so that it can then delete this object successfully
        $this->parent = null;
    }
}

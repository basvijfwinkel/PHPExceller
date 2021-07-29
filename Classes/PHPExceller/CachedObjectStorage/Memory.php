<?php
namespace PHPExceller\CachedObjectStorage;

use PHPExceller\CachedObjectStorage\CacheBase;
use PHPExceller\CachedObjectStorage\ICache;
use PHPExceller\PHPExceller_Cell;
use PHPExceller\PHPExceller_Worksheet;

/**
 * PHPExceller_CachedObjectStorage_Memory
 *
 * Copyright (c) 2021 PHPExceller
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.
 *
 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
 * Lesser General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public
 * License along with this library; if not, write to the Free Software
 * Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA
 *
 * @category   PHPExceller
 * @package    PHPExceller_CachedObjectStorage
 * @copyright  Copyright (c) 2021
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt    LGPL
 * @version    ##VERSION##, ##DATE##
 */
class Memory extends PHPExceller_CachedObjectStorage_CacheBase implements PHPExceller_CachedObjectStorage_ICache
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
     * @param    PHPExceller_Cell    $cell        Cell to update
     * @return    PHPExceller_Cell
     * @throws    PHPExceller_Exception
     */
    public function addCacheData($pCoord, PHPExceller_Cell $cell)
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
     * @throws     PHPExceller_Exception
     * @return     PHPExceller_Cell     Cell that was found, or null if not found
     */
    public function getCacheData($pCoord)
    {
        //    Check if the entry that has been requested actually exists
        if (!isset($this->cellCache[$pCoord])) {
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
     * @param    PHPExceller_Worksheet    $parent        The new worksheet
     */
    public function copyCellCollection(PHPExceller_Worksheet $parent)
    {
        parent::copyCellCollection($parent);

        $newCollection = array();
        foreach ($this->cellCache as $k => &$cell) {
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
        foreach ($this->cellCache as $k => &$cell) {
            $cell->detach();
            $this->cellCache[$k] = null;
        }
        unset($cell);

        $this->cellCache = array();

        //    detach ourself from the worksheet, so that it can then delete this object successfully
        $this->parent = null;
    }
}

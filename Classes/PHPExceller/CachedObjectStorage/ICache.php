<?php
namespace PHPExceller\CachedObjectStorage;

use PHPExceller\Cell;
use PHPExceller\Worksheet;

/**
 * Based on PHPExcel_CachedObjectStorage_ICache
 */
interface ICache
{
    /**
     * Add or Update a cell in cache identified by coordinate address
     *
     * @param    string            $pCoord        Coordinate address of the cell to update
     * @param    PHPExceller\Cell    $cell        Cell to update
     * @return   PHPExceller\Cell
     * @throws   PHPExceller\Exception
     */
    public function addCacheData($pCoord, Cell $cell);

    /**
     * Add or Update a cell in cache
     *
     * @param    PHPExceller\Cell    $cell        Cell to update
     * @return    PHPExceller\Cell
     * @throws    PHPExceller\Exception
     */
    public function updateCacheData(PHPExceller\Cell $cell);

    /**
     * Fetch a cell from cache identified by coordinate address
     *
     * @param    string            $pCoord        Coordinate address of the cell to retrieve
     * @return PHPExceller\Cell     Cell that was found, or null if not found
     * @throws    PHPExceller\Exception
     */
    public function getCacheData($pCoord);

    /**
     * Delete a cell in cache identified by coordinate address
     *
     * @param    string            $pCoord        Coordinate address of the cell to delete
     * @throws    PHPExceller\Exception
     */
    public function deleteCacheData($pCoord);

    /**
     * Is a value set in the current PHPExceller\CachedObjectStorage\ICache for an indexed cell?
     *
     * @param    string        $pCoord        Coordinate address of the cell to check
     * @return    boolean
     */
    public function isDataSet($pCoord);

    /**
     * Get a list of all cell addresses currently held in cache
     *
     * @return    string[]
     */
    public function getCellList();

    /**
     * Get the list of all cell addresses currently held in cache sorted by column and row
     *
     * @return    string[]
     */
    public function getSortedCellList();

    /**
     * Clone the cell collection
     *
     * @param    PHPExceller\Worksheet    $parent        The new worksheet
     * @return    void
     */
    public function copyCellCollection(Worksheet $parent);

    /**
     * Identify whether the caching method is currently available
     * Some methods are dependent on the availability of certain extensions being enabled in the PHP build
     *
     * @return    boolean
     */
    public static function cacheMethodIsAvailable();
}

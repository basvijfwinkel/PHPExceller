<?php
namespace PHPExceller\Cell;

use PHPExceller\Cell;

/**
 * Based on PHPExcel_Cell_IValueBinder
 */

interface IValueBinder
{
    /**
     * Bind value to a cell
     *
     * @param  Cell $cell    Cell to bind value to
     * @param  mixed $value           Value to bind in cell
     * @return boolean
     */
    public function bindValue(Cell $cell, $value = null);
}

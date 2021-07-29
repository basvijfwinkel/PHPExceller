<?php
namespace PHPExceller\Cell;

use PHPExcel\Cell\IValueBinder;
use PHPExceller\Cell\DataType;
use PHPExceller\Shared\String;

/**
 * PHPExceller_Cell_DefaultValueBinder
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
 * @package    PHPExceller_Cell
 * @copyright  Copyright (c) 2021 PHPExceller
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt    LGPL
 * @version    ##VERSION##, ##DATE##
 */
class DefaultValueBinder implements IValueBinder
{
    /**
     * Bind value to a cell
     *
     * @param  PHPExceller_Cell  $cell   Cell to bind value to
     * @param  mixed          $value  Value to bind in cell
     * @return boolean
     */
    public function bindValue(PHPExceller_Cell $cell, $value = null)
    {
        // sanitize UTF-8 strings
        if (is_string($value)) {
            $value = PHPExceller_Shared_String::SanitizeUTF8($value);
        } elseif (is_object($value)) {
            // Handle any objects that might be injected
            if ($value instanceof DateTime) {
                $value = $value->format('Y-m-d H:i:s');
            } elseif (!($value instanceof PHPExceller_RichText)) {
                $value = (string) $value;
            }
        }

        // Set value explicit
        $cell->setValueExplicit($value, self::dataTypeForValue($value));

        // Done!
        return true;
    }

    /**
     * DataType for value
     *
     * @param   mixed  $pValue
     * @return  string
     */
    public static function dataTypeForValue($pValue = null)
    {
        // Match the value against a few data types
        if ($pValue === null) {
            return PHPExceller_Cell_DataType::TYPE_NULL;
        } elseif ($pValue === '') {
            return PHPExceller_Cell_DataType::TYPE_STRING;
        } elseif ($pValue instanceof PHPExceller_RichText) {
            return PHPExceller_Cell_DataType::TYPE_INLINE;
        } elseif ($pValue{0} === '=' && strlen($pValue) > 1) {
            return PHPExceller_Cell_DataType::TYPE_FORMULA;
        } elseif (is_bool($pValue)) {
            return PHPExceller_Cell_DataType::TYPE_BOOL;
        } elseif (is_float($pValue) || is_int($pValue)) {
            return PHPExceller_Cell_DataType::TYPE_NUMERIC;
        } elseif (preg_match('/^[\+\-]?([0-9]+\\.?[0-9]*|[0-9]*\\.?[0-9]+)([Ee][\-\+]?[0-2]?\d{1,3})?$/', $pValue)) {
            $tValue = ltrim($pValue, '+-');
            if (is_string($pValue) && $tValue{0} === '0' && strlen($tValue) > 1 && $tValue{1} !== '.') {
                return PHPExceller_Cell_DataType::TYPE_STRING;
            } elseif ((strpos($pValue, '.') === false) && ($pValue > PHP_INT_MAX)) {
                return PHPExceller_Cell_DataType::TYPE_STRING;
            }
            return PHPExceller_Cell_DataType::TYPE_NUMERIC;
        } elseif (is_string($pValue) && array_key_exists($pValue, PHPExceller_Cell_DataType::getErrorCodes())) {
            return PHPExceller_Cell_DataType::TYPE_ERROR;
        }

        return PHPExceller_Cell_DataType::TYPE_STRING;
    }
}

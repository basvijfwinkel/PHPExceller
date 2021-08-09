<?php
namespace PHPExceller\Cell;

use PHPExcel\Cell\IValueBinder;
use PHPExceller\Cell\DataType;
use PHPExceller\Shared\String;
use PHPExceller\RichText;

/**
 * Based on DefaultValueBinder
 */

class DefaultValueBinder implements IValueBinder
{
    /**
     * Bind value to a cell
     *
     * @param  Cell  $cell   Cell to bind value to
     * @param  mixed          $value  Value to bind in cell
     * @return boolean
     */
    public function bindValue(Cell $cell, $value = null)
    {
        // sanitize UTF-8 strings
        if (is_string($value)) {
            $value = String::SanitizeUTF8($value);
        } elseif (is_object($value)) {
            // Handle any objects that might be injected
            if ($value instanceof DateTime) {
                $value = $value->format('Y-m-d H:i:s');
            } elseif (!($value instanceof RichText)) {
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
            return Cell_DataType::TYPE_NULL;
        } elseif ($pValue === '') {
            return Cell_DataType::TYPE_STRING;
        } elseif ($pValue instanceof RichText) {
            return Cell_DataType::TYPE_INLINE;
        } elseif ($pValue{0} === '=' && strlen($pValue) > 1) {
            return Cell_DataType::TYPE_FORMULA;
        } elseif (is_bool($pValue)) {
            return Cell_DataType::TYPE_BOOL;
        } elseif (is_float($pValue) || is_int($pValue)) {
            return Cell_DataType::TYPE_NUMERIC;
        } elseif (preg_match('/^[\+\-]?([0-9]+\\.?[0-9]*|[0-9]*\\.?[0-9]+)([Ee][\-\+]?[0-2]?\d{1,3})?$/', $pValue)) {
            $tValue = ltrim($pValue, '+-');
            if (is_string($pValue) && $tValue{0} === '0' && strlen($tValue) > 1 && $tValue{1} !== '.') {
                return Cell_DataType::TYPE_STRING;
            } elseif ((strpos($pValue, '.') === false) && ($pValue > PHP_INT_MAX)) {
                return Cell_DataType::TYPE_STRING;
            }
            return Cell_DataType::TYPE_NUMERIC;
        } elseif (is_string($pValue) && array_key_exists($pValue, Cell_DataType::getErrorCodes())) {
            return Cell_DataType::TYPE_ERROR;
        }
        return Cell_DataType::TYPE_STRING;
    }
}

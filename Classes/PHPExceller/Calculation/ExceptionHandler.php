<?php
namespace PHPExceller\Calculation;

use PHPExceller\Calculation\Calculation_Exception;

/**
 * Based on PHPExcel_Calculation_ExceptionHandler
 */
class ExceptionHandler
{
    /**
     * Register errorhandler
     */
    public function __construct()
    {
        set_error_handler(array('Calculation_Exception', 'errorHandlerCallback'), E_ALL);
    }

    /**
     * Unregister errorhandler
     */
    public function __destruct()
    {
        restore_error_handler();
    }
}

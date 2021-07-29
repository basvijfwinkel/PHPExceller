<?php
namespace PHPExceller\Calculation;

use PHPExceller\Calculation\Exception;

/**
 * PHPExceller_Calculation_ExceptionHandler
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
 * @package    PHPExceller_Calculation
 * @copyright  Copyright (c) 2021 PHPExceller
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt    LGPL
 * @version    ##VERSION##, ##DATE##
 */
class ExceptionHandler
{
    /**
     * Register errorhandler
     */
    public function __construct()
    {
        set_error_handler(array('PHPExceller_Calculation_Exception', 'errorHandlerCallback'), E_ALL);
    }

    /**
     * Unregister errorhandler
     */
    public function __destruct()
    {
        restore_error_handler();
    }
}

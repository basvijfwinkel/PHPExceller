<?php
namespace PHPExceller\Style\DataBar;

use PHPExceller\Exception;
use PHPExceller\Style\DataBarAxisPosition;

/**
 * PHPExceller
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
 * @package    PHPExceller_Style
 * @copyright Copyright (c) 2021 PHPExceller
 * @license http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt    LGPL
 * @version ##VERSION##, ##DATE##
 */

/**
 * PHPExceller_Style_DataBar_DataBarDirection
 *
 * @category   PHPExceller
 * @package    PHPExceller_Style_DataBar
 * @copyright  Copyright (c) 2021 PHPExceller
 * @author    Bas Vijfwinkel
 */
class DataBarDirection
{
    const CONTEXT = 'context';
    const LEFTTORIGHT = 'leftToRight';
    const RIGHTTOLEFT = 'rightToLeft';

    protected $direction;

    /*
    * constructor : PHPExceller_Style_DataBar_DataBarDirection object should be created with 'fromString' method
    *
    * @params    string     $direction    valid databar axis direction
    * 
    */
    protected function __construct($direction)
    {
        $this->direction = $direction;
    }

    /**
     * check if the databar axis direction is correct and return a PHPExceller_Style_DataBar_DataBarDirection object representing this direction
     * in case an unknown value is passed, an exception will be thrown
     *
     * <code>
     * $databardirection = PHPExceller_Style_DataBar_DataBarDirection::fromString($stringValue)
     * </code>
     *
     * @param    string    $type    string containing databar direction information
     * @throws    PHPExceller_Exception
     * @return string    databar direction as a string value
     */
    public static function fromString($type)
    {
        if (is_string($type))
        {
            switch($type)
            {
                case 'context':
                                    return new PHPExceller_Style_DataBar_DataBarDirection(PHPExceller_Style_DataBar_DataBarDirection::CONTEXT);
                                    break;
                case 'leftToRight':
                                    return new PHPExceller_Style_DataBar_DataBarDirection(PHPExceller_Style_DataBar_DataBarDirection::LEFTTORIGHT);
                                    break;
                case 'rightToLeft':
                                    return new PHPExceller_Style_DataBar_DataBarDirection(PHPExceller_Style_DataBar_DataBarDirection::RIGHTTOLEFT);
                                    break;
            }
        }
        // unknown type
        throw new PHPExceller_Exception("Invalid DataBarDirection passed.:".$type);
    }

    /*
    * Return the databar axis direction as a string
    *
    * @return    string    databar axis direction 
    */
    public function toString()
    {
        return "$this->direction";
    }
}
?>
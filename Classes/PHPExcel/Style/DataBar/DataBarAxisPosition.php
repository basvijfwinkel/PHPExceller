<?PHP
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
 * PHPExceller_Style_DataBar_DataBarAxisPosition
 *
 * @category   PHPExceller
 * @package    PHPExceller_Style_DataBar
 * @copyright  Copyright (c) 2021 PHPExceller
 * @author    Bas Vijfwinkel
 */
class PHPExceller_Style_DataBar_DataBarAxisPosition
{
    const AUTOMATIC= 'automatic';
    const MIDDLE = 'middle';
    const NONE = 'none';

    protected $position;

    /*
    * constructor : PHPExceller_Style_DataBar_DataBarAxisPosition object should be created with 'fromString' method
    *
    * @params    string     $position    valid databar axis position
    * 
    */
    protected function __construct($position)
    {
        $this->position = $position;
    }
    
    /**
     * check if the databar axis position is correct and return a PHPExceller_Style_DataBar_DataBarAxisPosition that represents this position
     * in case an unknown value is passed, an exception will be thrown
     *
     * <code>
     * $databaraxisposition = PHPExceller_Style_DataBar_DataBarAxisPosition::fromString($typevalue)
     * </code>
     *
     * @param    string    $type    string containing databar axis position information
     * @throws    PHPExceller_Exception
     * @return PHPExceller_Style_DataBar_DataBarAxisPosition    databar axis position
     */
    public static function fromString($type)
    {
        if (is_string($type))
        {
            switch($type)
            {
                case 'automatic':
                                    return new PHPExceller_Style_DataBar_DataBarAxisPosition(PHPExceller_Style_DataBar_DataBarAxisPosition::AUTOMATIC);
                                    break;
                case 'middle':
                                    return new PHPExceller_Style_DataBar_DataBarAxisPosition(PHPExceller_Style_DataBar_DataBarAxisPosition::MIDDLE);
                                    break;
                case 'none':
                                    return new PHPExceller_Style_DataBar_DataBarAxisPosition(PHPExceller_Style_DataBar_DataBarAxisPosition::NONE);
                                    break;
            }
        }
        // unknown type
        throw new PHPExceller_Exception("Invalid DataBarAxisPosition string passed.");
    }
    
    /*
    * Return the databar axis position as a string
    *
    * @return    string    databar axis position 
    */
    public function toString()
    {
        return "$this->position";
    }
}
?>
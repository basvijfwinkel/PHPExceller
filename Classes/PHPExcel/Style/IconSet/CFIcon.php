<?php
namespace PHPExceller\Style;

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
 * @copyright Copyright (c) 2021 PHPExceller (http://www.codeplex.com/PHPExceller)
 * @license http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt    LGPL
 * @version ##VERSION##, ##DATE##
 */


/**
 * PHPExceller_Style_IconSet_CFIcon
 *
 * @category   PHPExceller
 * @package    PHPExceller_Style_IconSet
 * @copyright  Copyright (c) 2021 PHPExceller
 * @author    Bas Vijfwinkel
 */
class PHPExceller_Style_IconSet_CFIcon
{
    /*
    *    @var iconSet
    */
    protected $iconSet;

    /*
    *    @var iconId
    */
    protected $iconId;

    /*
    * constructor : PHPExceller_Style_IconSet_CFIcon object should be created with 'fromArray' or fromXML method
    *
    * @params    string     $type    type
    * 
    */
    protected function __construct($iconSet, $iconId)
    {
        $this->iconSet = $iconSet;
        $this->iconId = $iconId;
    }

    /*
    * set the object based upon an array
    *
    * @param    array    array('iconSet' =>,'iconId' => )
    * @throws    PHPExceller_Exception    in case iconSet or iconId information is missing
    */
    public static function fromArray($data)
    {
        if ((isset($data['iconId'])) && (isset($data['iconSet'])))
        {
            return new PHPExceller_Style_IconSet_CFIcon($data['iconSet'], $data['iconId']);
        }
        else
        {
            // missing info
            throw new PHPExceller_Exception("cfIcon : missing iconId or iconSet data");
        }
    }

    /*
     * set the information based upon an XML object
     *
     * @param    XML        cfIcon XML structure
     * @throws    PHPExceller_Exception    in case iconSet or iconId information is missing
     */

    public static function fromXML($data)
    {
        $attributes = $data->attributes();
        if ((isset($attributes->iconSet)) && (isset($attributes->iconId)))
        {
            return new PHPExceller_Style_IconSet_CFIcon($attributes->iconSet, $attributes->iconId);
        }
        else
        {
            // missing info
            throw new PHPExceller_Exception("cfIcon : missing iconId or iconSet data");
        }
    }

    /*
     * Return the cfIcon type information as an array
     *
     * @return    array    cfIcon value
     */
    public function toArray()
    {
        $result =  array('name' => 'cfIcon','attributes' => array(array('name' => 'iconSet', 
                                                                     'attributes' => $this->iconSet),
                                                               array('name' => 'iconId', 
                                                                     'attributes' => $this->iconId)));
        return $result;
    }

}
?>
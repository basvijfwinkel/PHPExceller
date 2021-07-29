<?php
namespace PHPExceller\Style;

use PHPExceller\Style\PHPExceller_Style_CFVOType;
use PHPExceller\PHPExceller_Exception;

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
 * PHPExceller_Style_CFVOType
 *
 * @category   PHPExceller
 * @package    PHPExceller_Style_DataBar
 * @copyright  Copyright (c) 2021 PHPExceller
 * @author    Bas Vijfwinkel
 * <p>Note : I haven't found enough examples to test and implement the full specification of this object</p>
 */
class PHPExceller_Style_CFVOType
{
    const NUM        = 'num';
    const PERCENT    = 'percent';
    const MAX        = 'max';
    const MIN        = 'min';
    const FORMULA    = 'formula';
    const PERCENTILE = 'percentile';
    const AUTOMIN    = 'autoMin';
    const AUTOMAX    = 'autoMax';
    /*
    *    @var type
    */
    protected $_type;

    /*
    * @var child value (percentile uses this field to define it's value)
    */
    protected $_childValue;


    /*
    * constructor : PHPExceller_Style_CFVOType object should be created with 'fromString' method
    *
    * @params    string     $type    type
    * 
    */
    protected function __construct($type, $data)
    {
        // set the type
        $this->_type = $type;
        // add additional data
        $this->addAdditionalData($data);
    }

    protected function addAdditionalData($data)
    {
        if (($this->_type == PHPExceller_Style_CFVOType::PERCENTILE) ||
            ($this->_type == PHPExceller_Style_CFVOType::PERCENT) ||
            ($this->_type == PHPExceller_Style_CFVOType::NUM))
        {
            // there should be an <xm:f>0000</xm:f> child object or a val object if type is PERCENTILE
            if (isset($data['xm:f']))
            {
                $this->_childValue = (string)$data['xm:f'];
            }
            elseif (isset($data['val']))
            {
                $this->_childValue = (string)$data['val'];
            }
        }
    }

    /**
     * check if a valid cfvo type is passed and return an object representing this type
     *
     * <code>
     * $cfvotype = PHPExceller_Style_CFVOType::fromString($stringValue, $data)
     * </code>
     *
     * @param    string    $type    string cfvo value information
     * @param    array    $data    additional data for percentile
     * @throws    PHPExceller_Exception
     * @return string    CFVOTYpe as a string value
     */
    public static function fromString($type, $data=null)
    {
        if (is_string($type))
        {
            switch($type)
            {
                case 'num':
                                    return new PHPExceller_Style_CFVOType(PHPExceller_Style_CFVOType::NUM, $data);
                                    break;
                case 'max':
                                    return new PHPExceller_Style_CFVOType(PHPExceller_Style_CFVOType::MAX, $data);
                                    break;
                case 'percent':
                                    return new PHPExceller_Style_CFVOType(PHPExceller_Style_CFVOType::PERCENT, $data);
                                    break;
                case 'min':
                                    return new PHPExceller_Style_CFVOType(PHPExceller_Style_CFVOType::MIN, $data);
                                    break;
                case 'formula':
                                    return new PHPExceller_Style_CFVOType(PHPExceller_Style_CFVOType::FORMULA, $data);
                                    break;
                case 'percentile':
                                    return new PHPExceller_Style_CFVOType(PHPExceller_Style_CFVOType::PERCENTILE, $data);
                                    break;
                case 'autoMin':
                                    return new PHPExceller_Style_CFVOType(PHPExceller_Style_CFVOType::AUTOMIN, $data);
                                    break;
                case 'autoMax':
                                    return new PHPExceller_Style_CFVOType(PHPExceller_Style_CFVOType::AUTOMAX, $data);
                                    break;
            }
        }
        // unknown type
        throw new PHPExceller_Exception("Invalid CFVOType value passed.:".$type);
    }

    public static function fromXML($data)
    {
        $attributes = $data->attributes();
        $type = (string)$data['type'];
        return PHPExceller_Style_CFVOType::fromString($type, $attributes);
    }

    /*
    * create a PHPExceller_Style_CFVOType object from an array of settings
    *
    * <code>
    *  $settings = array('type' => 'min');
    *  $cvfotype = PHPExceller_Style_CFVOType::fromArray($settings);
    * </code>
    *
    * TODO : find more examples how this CFVOType works (besides 'min' and 'max')
    */
    public static function fromArray($arr)
    {
        if (is_array($arr))
        {
            if (array_key_exists('type', $arr)) 
            { 
                return PHPExceller_Style_CFVOType::fromString($arr['type'],$arr); 
            }
        }
        // unknown type
        throw new PHPExceller_Exception("Invalid CFVOType value passed.");
    }

    /*
     * Return the cfvo type information as an array
     *
     * @return    array    cfvo type value
     */
    public function toArray($forExtLst = false, $name = 'cfvo')
    {
        if (($this->_type == PHPExceller_Style_CFVOType::PERCENTILE) ||
            ($this->_type == PHPExceller_Style_CFVOType::PERCENT) ||
            ($this->_type == PHPExceller_Style_CFVOType::NUM))
        {
            if ($forExtLst)
            {
                //<x14:cfvo type="percentile"><xm:f>90</xm:f></x14:cfvo>
                $result =  array('name' => $name,'attributes' => array(array('name' => 'type', 
                                                                             'attributes' => $this->_type),
                                                                       array('name' => 'f', 
                                                                             'namespace' => 'xm',
                                                                             'attributes' => array('value' => $this->_childValue))));
            }
            else
            {
                //<cfvo type="percentile" val="10"/>
                $result =  array('name' => $name,'attributes' => array(array('name' => 'type', 
                                                                             'attributes' => $this->_type),
                                                                       array('name' => 'val', 
                                                                             'attributes' => $this->_childValue)));
            }
            return $result;
        }
        else
        {
            return array('name' => $name,'attributes' => array(array('name'=>'type', 
                                                                     'attributes' => $this->_type)));
        }
    }

}
?>
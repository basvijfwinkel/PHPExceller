<?php
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
 * @copyright  Copyright (c) 2021 PHPExceller
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt    LGPL
 * @version    ##VERSION##, ##DATE##
 */

/**
 * PHPExceller_Style_IconSet
 *
 * @category   PHPExceller
 * @package    PHPExceller_Style
 * @copyright  Copyright (c) 2021 PHPExceller
 * @author    Bas Vijfwinkel
 */
class PHPExceller_Style_IconSet extends PHPExceller_Style_GroupedConditional implements PHPExceller_IComparable
{
    const ICONSET_NAME_3ARROWS         = '3Arrows';
    const ICONSET_NAME_3ARROWSGRAY     = '3ArrowsGray';
    const ICONSET_NAME_3Flags          = '3Flags';
    const ICONSET_NAME_3TrafficLights  = '3TrafficLights1';
    const ICONSET_NAME_3TrafficLights2 = '3TrafficLights2';
    const ICONSET_NAME_3Signs          = '3Signs';
    const ICONSET_NAME_3Symbols        = '3Symbols';
    const ICONSET_NAME_3Symbols2       = '3Symbols2';
    const ICONSET_NAME_4Arrows         = '4Arrows';
    const ICONSET_NAME_4ArrowsGray     = '4ArrowsGray';
    const ICONSET_NAME_4RedToBlack     = '4RedToBlack';
    const ICONSET_NAME_4Rating         = '4Rating';
    const ICONSET_NAME_4TrafficLights  = '4TrafficLights';
    const ICONSET_NAME_5Arrows         = '5Arrows';
    const ICONSET_NAME_5ArrowsGray     = '5ArrowsGray';
    const ICONSET_NAME_5Rating         = '5Rating';
    const ICONSET_NAME_5Quarters       = '5Quarters';
    const ICONSET_NAME_3Stars          = '3Stars';
    const ICONSET_NAME_5Boxes          = '5Boxes';
    const ICONSET_NAME_NoIcons         = 'NoIcons';
    /**
    * CFVO Array
    *
    * @var array
    */
    protected $cfvos;

    /**
    * icon set name that needs to be used
    * Known icon set names :
    *            3Arrows,3ArrowsGray,3Flags,3TrafficLights1,3TrafficLights2,3Signs,3Symbols,3Symbols2,
    *            4Arrows,4ArrowsGray,4RedToBlack,4Rating,4TrafficLights,5Arrows,5ArrowsGray,5Rating,
    *            5Quarters,3Stars,5Boxes,NoIcons
    *
    * @var string
    */
    protected $iconSetName;

    /**
    * use icons in the reverse direction (1=true; 0=false)
    *
    * @var integer
    */
    protected $reverse;

    /**
    * use custom iconsets for each range (1=true; 0=false)
    *
    * @var integer
    */
    protected $custom;

    /**
    * show cell value (1=true; 0=false)
    *
    * @var integer
    */
    protected $showValue;

    /**
    * uniq id for this object
    *
    * @var string
    */
    protected $id;

    /**
     * cfIcon objects (if object is a custom object)
     *
     * @var    array    array of cfIcons
     */
     protected $cfIcons;

    /**
     * Create a new PHPExceller_Style_IconSet
     *
     */
    public function __construct()
    {
        parent::__construct();
    }

    /**
     * Apply styles from array
     *
     * <code>
     * $pValue = new PHPExceller_Style_IconSet();
     *    $pValue->applyFromArray(
     *        array(
     *            'cellReference' => 'A1:A5',
     *            'iconSet' => "3Arrows",
     *            'reverse' => 1,
     *            'showValue' => 0,
     *          'custom' => 1,
     *            'cfIcons' => array(array('iconSet'=>'3Triangles','iconId=>0),array('iconSet'=>'3TrafficLights','iconId=>1),array('iconSet'=>'3Arrows','iconId=>2)),
     *          'cfvos' => array(array('type'=>'min'),array('type'=>'max'))
     *            ), true);
     * $objPHPExceller->getActiveSheet()->setConditionalStyles('A3',$pValue)
     * </code>
     *
     * @param    array    $pStyles    Array containing style information
     * @return    PHPExceller_Style_IconSet
     * Note : cfvos and color parameter will override all existing cfvo and color values; If you want to preserve them, add them manually with addCfvo or addColor
     */
    public function applyFromArray($pStyles = null, $checkInput=true, $isExtLstData= false) {
        if (is_array($pStyles))
        {
            if (array_key_exists('cellReference', $pStyles))
            {
                $this->setCellReference($pStyles['cellReference']);
            }
            if (array_key_exists('iconSet', $pStyles))
            {
                $this->setIconSetName($pStyles['iconSet']);
            }
            if (array_key_exists('reverse', $pStyles))
            {
                $this->setReverse($pStyles['reverse']);
            }
            if (array_key_exists('showValue', $pStyles))
            {
                $this->setShowValue($pStyles['showValue']);
            }
            if (array_key_exists('custom', $pStyles))
            {
                $this->setCustom($pStyles['custom']);
            }
            if (array_key_exists('cfvos', $pStyles))
            {
                $resultcfvos = array();
                foreach ($pStyles['cfvos'] as $cfvotype)
                {
                    $resultcfvos[] = PHPExceller_Style_CFVOType::fromArray($cfvotype);
                }
                $this->addCfvos($resultcfvos); // add cfvo to ext_cfvo to preserve both
            }
            if (array_key_exists('cfIcons', $pStyles))
            {
                $resultcfIcons = array();
                foreach ($pStyles['cfIcons'] as $cfIcon)
                {
                    $resultcfIcons[] = PHPExceller_Style_IconSet_CFIcon::fromArray($cfIcon);
                }
                $this->addCfIcons($resultcfIcons);
            }
        }
        else
        {
            throw new PHPExceller_Exception("IconSet : invalid input applyFromArray :".var_export($pStyles,true));
        }
        return $this;
    }


    /*
    * Create an array of elements containing all the information of the input xml object
    *
    * @param    string    $ref    cell reference (e.g A1:A5)
    * @param    SimpleXML $cfRule    cfRule xml structure containing the IconSet section
    * @param    --        not used
    * @return    Array    array containing all the information of the input xml object
    * @throws PHPExceller_Exception
    */
    public function applyFromXML($ref, $cfRule, $extst=null)
    {
        // check that at least the cfvo property exists
        if (isset($cfRule->iconSet->cfvo))
        {
            // add default properties
            $this->setCellReference($ref);

            $attr = $cfRule->iconSet->attributes();
            // iconSet
            if (isset($attr->iconSet))
            {
                $this->setIconSetName((string)$attr->iconSet);
            }
            // reverse
            if (isset($attr->reverse))
            {
                $this->setReverse((int)$attr->reverse);
            }
            // custom
            if (isset($attr->custom))
            {
                $this->setCustom((int)$attr->custom);
            }
            // showValue
            if (isset($attr->showValue))
            {
                $this->setShowValue((int)$attr->showValue);
            }
            // cfvo
            if (!is_null($cfRule->iconSet->cfvo))
            {
                $this->cfvos = array(); // clear our the list before adding new ones
                foreach($cfRule->iconSet->cfvo as $cfvo)
                {
                    $type = (string)$cfvo->attributes()->type;

                    $cfvoattr = $cfvo->attributes();
                    if (isset($cfvoattr['val']))
                    {
                        $data = $cfvoattr;
                    }
                    else
                    {
                        // fix this mess : namespace messing up the generic handling : xm:f not recognized
                        $c = $cfvo->children('xm',true);
                        $childValue = (string)$c->f;
                        $data = array('xm:f' => $childValue);
                    }
                    $this->addCfvo(PHPExceller_Style_CFVOType::fromString($type,$data));
                }
            }
            // cfIcon (only if custom=1)
            if ((isset($attr->custom)) && ((int)$attr->custom == 1))
            {
                // cfIcons
                $this->cfIcons = array(); // clear our the list before adding new ones
                foreach($cfRule->iconSet->cfIcon as $cfIcon)
                {
                    $this->addCfIcon(PHPExceller_Style_IconSet_CFIcon::fromXML($cfIcon));
                }
            }



        }
        else
        {
            // missing property
            throw new PHPExceller_Exception("IconSet : missing iconSet name or cvfo setting");
        }

        return $this;
    }

    /*
    /*
    * Create an array of elements containing all the information of the object
    *  in order to pass it on to the xmlwriter
    *
    * @param    bool    $isExtLst    whether the array is meant for the default section or the extlst block
    * @return    Array    array containing all the information of this IconSet
    * NOTE : The order in which the elements are written into the array should not matter but for Microsoft Excel it apparently _DOES MATTER_
    */
    public function getElementsAsArray($isExtLst=false)
    {
        $result = array();

        // 1. cfvo's
        if ($this->cfvos)
        {
            foreach($this->cfvos as $cfvotype)
            {
                $result[] = $cfvotype->toArray($isExtLst);
            }
        }
        // 2. iconset name
        if (!is_null($this->iconSetName)  )
        {
            $result[] = array('name' => 'iconSet', 'attributes' => $this->iconSetName);
        }

        // 3. reverse setting
        if (!is_null($this->reverse)  )
        {
            $result[] = array('name' => 'reverse', 'attributes' => $this->reverse);
        }

        // 4. showValue setting
        if (!is_null($this->showValue)  )
        {
            $result[] = array('name' => 'showValue', 'attributes' => $this->showValue);
        }

        // 5. custom attribute
        if (!is_null($this->custom))
        {
            $result[] = array('name' => 'custom', 'attributes' => $this->custom);
            // if custom : also write out the cfIcons
            foreach($this->cfIcons as $cfIcon)
            {
                $result[] = $cfIcon->toArray();
            }
        }

        // return the resulting array
        return $result;
    }

    /*
     * get the name of the IconSet
     *
     * <code>
     * $worksheetstyles = $objPHPExceller->getActiveSheet()->getConditionalStyles();
     * if ($worksheetstyles[0]->getConditionType() == PHPExceller_Style_Conditional::CONDITION_IconSet) { $color = $worksheetstyles[0]->getIconSetName() ; }
     * </code>
     *
     * @return string
    */
    public function getIconSetName()
    {
        return $this->iconSetName;
    }

    /*
     * Set the name of the IconSet
     *
     * <code>
     * $worksheetstyles = $objPHPExceller->getActiveSheet()->getConditionalStyles();
     * if ($worksheetstyles[0]->getConditionType() == PHPExceller_Style_Conditional::CONDITION_IconSet)
     * {
     *      $worksheetstyles[0]->setIconSetName('3Arrows');
     * }
     * </code>
     *
     * @param    string    name of the IconSet
     * @returns PHPExceller_Style_IconSet
    */
    public function setIconSetName($name)
    {
           $this->iconSetName = $name;
           return $this;
    }

    /*
     * get direction of the icons (1=reverse, 0=default direction)
     *
     * <code>
     * $worksheetstyles = $objPHPExceller->getActiveSheet()->getConditionalStyles();
     * if ($worksheetstyles[0]->getConditionType() == PHPExceller_Style_Conditional::CONDITION_IconSet) { $color = $worksheetstyles[0]->getReverse() ; }
     * </code>
     *
     * @return integer
    */
    public function getReverse()
    {
        return $this->reverse;
    }

    /*
     * Set the direction of the icons (1=reverse direction)
     *
     * <code>
     * $worksheetstyles = $objPHPExceller->getActiveSheet()->getConditionalStyles();
     * if ($worksheetstyles[0]->getConditionType() == PHPExceller_Style_Conditional::CONDITION_IconSet)
     * {
     *      $worksheetstyles[0]->setIconSetName(1);
     * }
     * </code>
     *
     * @param    integer    1=reverse direction;0=default direction
     * @returns PHPExceller_Style_IconSet
    */
    public function setReverse($dir)
    {
           $this->reverse = $dir;
           return $this;
    }

    /*
     * show values in the cell (1=show (default), 0=hide values)
     *
     * <code>
     * $worksheetstyles = $objPHPExceller->getActiveSheet()->getConditionalStyles();
     * if ($worksheetstyles[0]->getConditionType() == PHPExceller_Style_Conditional::CONDITION_IconSet) { $color = $worksheetstyles[0]->getshowValue() ; }
     * </code>
     *
     * @return integer
    */
    public function getShowValue()
    {
        return $this->showValue;
    }

    /*
     * Set whether the value should be shown in the cell  (0=hide; 1=show (default))
     *
     * <code>
     * $worksheetstyles = $objPHPExceller->getActiveSheet()->getConditionalStyles();
     * if ($worksheetstyles[0]->getConditionType() == PHPExceller_Style_Conditional::CONDITION_IconSet)
     * {
     *      $worksheetstyles[0]->setShowValue(0);
     * }
     * </code>
     *
     * @param    integer    1=show value;0=hide value
     * @returns PHPExceller_Style_IconSet
    */
    public function setShowValue($show)
    {
           $this->showValue = $show;
           return $this;
    }

    /*
     * get whether a custom set of icons is used (1=custom; 0=no custom icons)
     *
     * <code>
     * $worksheetstyles = $objPHPExceller->getActiveSheet()->getConditionalStyles();
     * if ($worksheetstyles[0]->getConditionType() == PHPExceller_Style_Conditional::CONDITION_IconSet) { $color = $worksheetstyles[0]->getCustom() ; }
     * </code>
     *
     * @return integer
    */
    public function getCustom()
    {
        return $this->custom;
    }

    /*
     * Set the usage of custom set of icons (note : cfIcons must also be defined to make this work)
     *
     * <code>
     * $worksheetstyles = $objPHPExceller->getActiveSheet()->getConditionalStyles();
     * if ($worksheetstyles[0]->getConditionType() == PHPExceller_Style_Conditional::CONDITION_IconSet)
     * {
     *      $worksheetstyles[0]->setCustom(1);
     * }
     * </code>
     *
     * @param    integer    1=custom settings
     * @returns PHPExceller_Style_IconSet
    */
    public function setCustom($custom)
    {
           $this->custom = $custom;
           return $this;
    }


    /*
     * Get the cfvo settings
     *
     * @params    boolean    set to true if the extlst cfvo settings must be used
     * <code>
     * $worksheetstyles = $objPHPExceller->getActiveSheet()->getConditionalStyles();
     * if ($worksheetstyles[0]->getConditionType() == PHPExceller_Style_Conditional::CONDITION_IconSet) { $cfvo = $worksheetstyles[0]->getCfvos() ; }
     * </code>
     *
     * @return array of CFVOType    all CFVOTypes for this IconSet
    */
    public function getCfvos($use_extlst_cfvos=false)
    {
        if ($use_extlst_cfvos)
        {
            return $this->extlst_cfvos;
        }
        else
        {
            return $this->cfvos;
        }
    }

    /*
    * add a cfvo type
    *
    * @param    cfvotype
    * @param    boolean    set to true is the data must be entered to the extlst cvfo array
    * <code>
    * $worksheetstyles = $objPHPExceller->getActiveSheet()->getConditionalStyles();
    * if ($worksheetstyles[0]->getConditionType() == PHPExceller_Style_Conditional::CONDITION_IconSet)
    * {
    *    $worksheetstyles[0]->addCfvoType(PHPExceller_Style_CFVOType::fromString('min'));
    *   $worksheetstyles[0]->addCfvoType(PHPExceller_Style_CFVOType::fromString('max'));
    * }
    * </code>
    */
    public function addCfvo($cfvotype = null, $use_extlst_cfvos=false)
    {
        if ($cfvotype)
        {
            if (!$use_extlst_cfvos)
            {
                array_push($this->cfvos,$cfvotype);
            }
            else
            {
                if (!$this->extlst_cfvos) { $this->extlst_cfvos = array();}
                array_push($this->extlst_cfvos,$cfvotype);
            }
        }
           return $this;
    }
    /*
     * Add a list of cfvos (Note: existing entries will be destroyed
     *
     * @param    array    list of cfvotypes
     * @param    boolean    set to true is the data must be entered to the extlst cvfo array
     *
     *
     */
    public function addCfvos($cfvos = null, $use_extlst_cfvos=false)
    {
        if ($use_extlst_cfvos)
        {
            $this->extlst_cfvos = $cfvos;
        }
        else
        {
            $this->cfvos = $cfvos;
        }
           return $this;
    }

    /*
     * Get the cfIcon settings
     *
     * <code>
     * $worksheetstyles = $objPHPExceller->getActiveSheet()->getConditionalStyles();
     * if ($worksheetstyles[0]->getConditionType() == PHPExceller_Style_Conditional::CONDITION_IconSet) { $cfvo = $worksheetstyles[0]->getCfIcons() ; }
     * </code>
     *
     * @return array of CFIcon    all cfIcons for this IconSet
    */
    public function getCfIcons()
    {
        return $this->cfIcons;
   }

    /*
    * add a cfIcon object (note : cfvo's must match up and custom must be set to 1)
    *
    * @param    cfIcon
    * <code>
    * $worksheetstyles = $objPHPExceller->getActiveSheet()->getConditionalStyles();
    * if ($worksheetstyles[0]->getConditionType() == PHPExceller_Style_Conditional::CONDITION_IconSet)
    * {
    *    $worksheetstyles[0]->addCfIcon(PHPExceller_Style_IconSet_CFIcon::fromArray(array('iconSet'=>'3Triangles','iconId'=>0));
    *   $worksheetstyles[0]->addCfIcon(PHPExceller_Style_IconSet_CFIcon::fromArray(array('iconSet'=>'3TrafficLights1','iconId'=>1)));
    *   $worksheetstyles[0]->addCfIcon(PHPExceller_Style_IconSet_CFIcon::fromArray(array('iconSet'=>'3Arrows','iconId'=>2)));
    * }
    * </code>
    */
    public function addCfIcon($cfIcon)
    {
        if (is_null($this->cfIcons)) { $this->cfIcons = array(); }
        array_push($this->cfIcons,$cfIcon);
           return $this;
    }

    /*
     * Add a list of cfIcons (Note: existing entries will be destroyed
     *
     * @param    array    list of cfIcons
     * @param    boolean    set to true is the data must be entered to the extlst cvfo array
     *
     *
     */
    public function addCfIcons($cfIcons)
    {
        $this->cfIcons = $cfIcons;
    }

    /*
     * create a datastructure for creating the IconSet element with default properties
     *
     * @param    array    array with properties
     */
    public function getDefaultData()
    {
        if ((!is_null($this->custom) && ($this->custom == '1')) ||
            ($this->needsExtLstEntry()))
        {
            // these types need to be inserted in the extlst block
            $result = null;
        }
        else
        {
            // add cfvo
            $cfvos = $this->getCfvos();
            $result = array('name' => 'iconSet', 'attributes' => array());
            foreach ($cfvos as $cfvo)
            {
                    $result['attributes'][] = $cfvo->toArray();
            }

            if (!is_null($this->iconSetName))
            {
                $result['attributes'][] = array('name'=>'iconSet', 'attributes' => $this->iconSetName);
            }

            if (!is_null($this->reverse))
            {
                $result['attributes'][] = array('name' => 'reverse', 'attributes' => $this->reverse);
            }

            if (!is_null($this->showValue))
            {
                $result['attributes'][] = array('name' => 'showValue', 'attributes' => $this->showValue);
            }
        }
        return $result;
    }

    /*
     *  Indicate whether this object needs a extlst data entry
     *
     * @return    bool    4 and 5 type icons need to be added to the extlst block
     */
    public function needsExtLstEntry()
    {
        if (
            ($this->iconSetName == PHPExceller_Style_IconSet::ICONSET_NAME_5Boxes) ||
            ($this->custom == 1))
        {
            return true;
        }
        else
        {
            return false;
        }
    }

    /*
     * Get the namespace setting for creating the extLst entry
     *
     * <code>
     * $worksheetstyles = $objPHPExceller->getActiveSheet()->getConditionalStyles();
     * $namespace = $worksheetstyles[0]->getNamespace();
     * </code>
     *
     * @return string containing namespace setting
    */
    public function getNamespace()
    {
        return $this->namespace;
    }

        /*
     * Indicates whether this object needs a reference to the entry in the extLst section
     *
     * @returns    bool    true is such a reference is needed
     *
     */
     public function needsExtLstReference()
     {
        return false;
     }
         /*
     * Get an array containing the iconset data that must be written to the extLst entry of the worksheet
     *
     * @param    int    priority for this element
     * @return    array    array containing all the data that must be written to the extLst entry of the worksheert
     *
     */
    public function getExtLstData($priority)
    {
        $forExtLst = true;
        $data = $this->getElementsAsArray($forExtLst);
        $result = array('name' => 'cfRule',
                        'cellReference' => $this->getCellReference(),
                        'attributes' => array(array('name' => 'type',   'attributes' => 'iconSet'),
                                                array('name' => 'priority',  'attributes' => $priority),
                                                array('name' => 'id',        'attributes' => $this->getClassID()),
                                                array('name' => 'iconSet',   'attributes' => $data)));
        return $result;
    }

}

<?php
/*
TODO :

- check : if priority of 2 databars work correctly
- check if auto="1" is handled for reading+writing existing sheets for the 'negativeFillColor' and 'axisColor is handled properly

/*<code>
       $conditional = new PHPExceller_Style_Conditional();
       $conditional->setPriority(0);
       $conditionType = PHPExceller_Style_Conditional::CONDITION_DATABAR;
       $settingsArray =  array(
                          'cellReference' => 'A1:A5',
                          'color' => ['rgb' => 'FF557DBC'],
                          //'fillColor' => ['rgb' =>'FF00FF00'],
                          'border' => true,
                          'borderColor' => ['rgb' =>'FF5279BA'],
                          'negativeFillColor' => ['auto' => '1'],
                          'axisColor' => ['auto' => '1'],
                          'cfvos' => array(array('type'=>'min'),array('type'=>'max'))
                          'showValue' => true,
                          //'negativeBarColorSameAsPositive' => 0,
                          //'negativeFillColor' => ['rgb' =>'FFFF0000'],
                          //'negativeBarBorderColorSameAsPositive' => 0,
                          //'negativeBorderColor' => ['rgb' =>'FFFF00FF'],
                          //'axisColor' => ['rgb' =>'FF7F7F7F'],
                          //'minLength' => 20,
                          //'maxLength' => 70,
                          //'direction' => 'context',
                          //'axisPosition' => 'middle',

                      );
       $conditional->setConditionObjectFromArray($conditionType, $settingsArray);
       $PHPExcellerObj->getActiveSheet->getStyle('A1')->addConditionalStyle($conditional);
</code>
<code>

</code>
*/

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
 * PHPExceller_Style_DataBar
 *
 * @category   PHPExceller
 * @package    PHPExceller_Style
 * @copyright  Copyright (c) 2021 PHPExceller
 * @author    Bas Vijfwinkel
 */
class PHPExceller_Style_DataBar extends PHPExceller_Style_GroupedConditional implements PHPExceller_IComparable
{

    /**
    * CFVO Array for Databar
    *
    * @var array
    */

    protected $cfvos;

    /**
    * ExtLstCFVO Array for Databar (cfvo data in the default databar block and the extlst data block might differ)
    *
    * @var array
    */

    protected $extlst_cfvos;

    /**
    * color for Databar
    *
    * @var PHPExceller_Style_Color
    */

    protected $color;

    /**
    * fill color
    *
    * @var PHPExceller_Style_Color
    */
    protected $fillColor;

    /**
    * border color
    *
    * @var PHPExceller_Style_Color
    */
    protected $borderColor;

    /**
    *     negative fill color
    *
    * @var PHPExceller_Style_Color
    */
    protected $negativeFillColor;
    protected $negativeFillColorAuto;

    /**
    * negative border color
    *
    * @var PHPExceller_Style_Color
    */
    protected $negativeBorderColor;

    /**
    * axis color
    *
    * @var PHPExceller_Style_Color
    */
    protected $axisColor;
    protected $axisColorAuto;

    /**
    * min langth
    *
    * @var unsigned int (default : 10)
    */
    protected $minLength;

    /**
    * max length
    *
    * @var unsigned int (default : 90)
    */
    protected $maxLength;

    /**
    * show value in the cell (0=false;1=true)
    *
    * @var Integer (default : 1)
    */
    protected $showValue;

    /**
    * show border (Note : 0=false; 1=true)
    *
    * @var Integer    (default : 0)
    */
    protected $border;

    /**
    * show gradient (Note : 0=false; 1=true)
    *
    * @var Integer (default : 0)
    */
    protected $gradient;

    /**
    * direction of the databar
    *
    * @var PHPExceller_Style_DataBarDirection (default 'context')
    */
    protected $direction;

    /**
    * negativeBarColorSameAsPositive  (0=false;1=true)
    *
    * @var Integer (default : 0)
    */
    protected $negativeBarColorSameAsPositive;

    /**
    * negativeBarBorderColorSameAsPositive
    *
    * @var Boolean (default : true)
    */
    protected $negativeBarBorderColorSameAsPositive;

    /**
    * axisPosition
    *
    * @var PHPExceller_Style_DataBarAxisPosition (default : automatic)
    */
    protected $axisPosition;

    /**
    * namespace for extLst entry (optional)
    *
    * @var string (default : http://schemas.microsoft.com/office/spreadsheetml/2009/9/main)
    */
    protected $namespace;

    /**
    * uniq id for this object
    *
    * @var string
    */
    protected $id;

    /**
     * Create a new PHPExceller_Style_Border
     *
     */
    public function __construct()
    {
        parent::__construct();
        // default namespace
        $this->namespace = "http://schemas.microsoft.com/office/spreadsheetml/2009/9/main";
    }

    /**
     * Check if the object needs a extLst entry
     * if besides color and cfvo any other property is used, Excel2010 only needs to add a conditional format entry in the extLst structure.
     *
     *  @return : boolean    true if an extLst entry needs to be generated.
     */
     public function needsExtLstEntry()
     {
        return ($this->extlst_cfvos ||
                $this->fillColor ||
                $this->borderColor ||
                $this->negativeFillColor ||
                $this->negativeFillColorAuto ||
                $this->negativeBorderColor ||
                $this->axisColor ||
                $this->axisColorAuto ||
                !is_null($this->minLength) ||
                !is_null($this->maxLength) ||
                //!is_null($this->showValue) ||
                !is_null($this->border) ||
                !is_null($this->gradient) ||
                $this->direction ||
                !is_null($this->negativeBarColorSameAsPositive) ||
                !is_null($this->negativeBarBorderColorSameAsPositive) ||
                $this->axisPosition)?true:false;
     }

    /**
     * Apply styles from array
     *
     * @param    array    $pStyles    Array containing style information
     * @param   boolean    $checkInput    set to true if the validity of the data must be checked (Excel2010 does not seem to follow the specifications...)
     * @param    boolean    $isExtLstData    set to true if the data is from an extlst block (cfvo data will be stored separately)
     * @return void
     * Note : cfvos parameter will override all existing cfvo values; If you want to preserve them, add them manually with addCfvo
     */
    public function applyFromArray($pStyles = null, $checkInput=true, $isExtLstData= false)
        {
        if (!is_array($pStyles)) { throw new PHPExceller_Exception("DataBar : invalid input applyFromArray :".var_export($pStyles,true)); }

            // checks
            if ($checkInput)
            {
                if ((array_key_exists('minLength', $pStyles)) && (array_key_exists('maxLength', $pStyles)) && ($pStyles['minLength'] > $pStyles['maxLength']))
                {
                   throw new PHPExceller_Exception("DataBar : minLength should be smaller or equal to maxLength");
                }
                if ((array_key_exists('negativeBarColorSameAsPositive', $pStyles)) && (!(array_key_exists('negativeFillColor', $pStyles))))
                {
                   throw new PHPExceller_Exception("DataBar : negativeFillColor should be set");
                }
                if ((array_key_exists('negativeBarBorderColorSameAsPositive', $pStyles)) && (!(array_key_exists('negativeBorderColor', $pStyles))))
                {
                   throw new PHPExceller_Exception("DataBar : negativeBorderColor should be set");
                }
                if (array_key_exists('axisColor', $pStyles))
                {
                    if ((array_key_exists('axisPosition', $pStyles)) && ($pStyles['axisPosition'] == PHPExceller_Style_DataBar_DataBarAxisPosition::NONE))
                    {
                        throw new PHPExceller_Exception("DataBar : in order to set axisColor, axisPosition must be defined and not set to 'NONE'");
                    }
                }
                if (array_key_exists('border', $pStyles) && ($pStyles['border']) &&
                    (!array_key_exists('negativeFillColor', $pStyles) || !array_key_exists('axisColor', $pStyles) ))
                {
                    throw new PHPExceller_Exception("DataBar : Set both negativeFillColor and axisColor to array('auto' => '1') when using borderColor without setting negativeFillColor and axisColor");
                }
                if (((array_key_exists('fillColor', $pStyles)) ||
                         (array_key_exists('borderColor', $pStyles)) ||
                         (array_key_exists('negativeFillColor', $pStyles)) ||
                         (array_key_exists('negativeBorderColor', $pStyles)) ||
                         (array_key_exists('minLength', $pStyles)) ||
                         (array_key_exists('maxLength', $pStyles)) ||
                         (array_key_exists('border', $pStyles)) ||
                         (array_key_exists('gradient', $pStyles)) ||
                         (array_key_exists('negativeBarColorSameAsPositive', $pStyles)) ||
                         (array_key_exists('negativeBarBorderColorSameAsPositive', $pStyles)) ||
                         (array_key_exists('direction', $pStyles)) ||
                         (array_key_exists('axisPosition', $pStyles)) ||
                         (array_key_exists('axisColor', $pStyles))
                       ) && (!array_key_exists('cfvos', $pStyles)))
                {
                    throw new PHPExceller_Exception("DataBar : set the cfvos property to array(array('type'=>'min'),array('type'=>'max')) in order to use optional properties");
                }

            }

            // default properties
            if (array_key_exists('cellReference', $pStyles))       { $this->setCellReference($pStyles['cellReference']); }
            if (array_key_exists('color', $pStyles))               { $this->setColor(new PHPExceller_Style_Color($pStyles['color']['rgb'])); }
            if (array_key_exists('cfvos', $pStyles))
            {

                $resultcfvos = array();
                foreach ($pStyles['cfvos'] as $cfvotype) { $resultcfvos[] = PHPExceller_Style_CFVOType::fromArray($cfvotype); }
                $this->addCfvos($resultcfvos,$isExtLstData); // add cfvo to ext_cfvo to preserve both
                if ((!$isExtLstData) &&
                        ((array_key_exists('fillColor', $pStyles)) ||
                         (array_key_exists('borderColor', $pStyles)) ||
                         (array_key_exists('negativeFillColor', $pStyles)) ||
                         (array_key_exists('negativeBorderColor', $pStyles)) ||
                         (array_key_exists('minLength', $pStyles)) ||
                         (array_key_exists('maxLength', $pStyles)) ||
                         (array_key_exists('border', $pStyles)) ||
                         (array_key_exists('gradient', $pStyles)) ||
                         (array_key_exists('negativeBarColorSameAsPositive', $pStyles)) ||
                         (array_key_exists('negativeBarBorderColorSameAsPositive', $pStyles)) ||
                         (array_key_exists('direction', $pStyles)) ||
                         (array_key_exists('axisPosition', $pStyles)) ||
                         (array_key_exists('axisColor', $pStyles))
                       )
                    )
                {
                    $this->addCfvos($resultcfvos,true);
                }
            }

            if (array_key_exists('showValue', $pStyles))           { $this->setShowValue((int)$pStyles['showValue']); }

            // additional properties
            if (array_key_exists('fillColor', $pStyles))           { $this->setFillColor(new PHPExceller_Style_Color($pStyles['fillColor']['rgb'])); }
            if (array_key_exists('borderColor', $pStyles))         { $this->setBorderColor(new PHPExceller_Style_Color($pStyles['borderColor']['rgb'])); }
            if (array_key_exists('negativeFillColor', $pStyles))
            {
                if (isset($pStyles['negativeFillColor']['auto']) && ($pStyles['negativeFillColor']['auto']=='1'))
                {
                    $this->negativeFillColorAuto = true;
                }
                else
                {
                    $this->setNegativeFillColor(new PHPExceller_Style_Color($pStyles['negativeFillColor']['rgb']));
                }
            }
            if (array_key_exists('negativeBorderColor', $pStyles)) { $this->setNegativeBorderColor(new PHPExceller_Style_Color($pStyles['negativeBorderColor']['rgb'])); }
            if (array_key_exists('minLength', $pStyles))           { $this->setMinLength((int)$pStyles['minLength']); }
            if (array_key_exists('maxLength', $pStyles))           { $this->setMaxLength((int)$pStyles['maxLength']); }
            if (array_key_exists('border', $pStyles))              { $this->setBorder((int)$pStyles['border']); }
            if (array_key_exists('gradient', $pStyles))            { $this->setGradient((int)$pStyles['gradient']); }
            if (array_key_exists('negativeBarColorSameAsPositive', $pStyles)) { $this->setNegativeBarColorSameAsPositive((int)$pStyles['negativeBarColorSameAsPositive']); }
            if (array_key_exists('negativeBarBorderColorSameAsPositive', $pStyles)) { $this->setNegativeBarBorderColorSameAsPositive((int)$pStyles['negativeBarBorderColorSameAsPositive']); }
            if (array_key_exists('direction', $pStyles))          { $this->setDirection(PHPExceller_Style_DataBar_DataBarDirection::fromString($pStyles['direction'])); }
            if (array_key_exists('axisPosition', $pStyles))       { $this->setAxisPosition(PHPExceller_Style_DataBar_DataBarAxisPosition::fromString($pStyles['axisPosition'])); }
            if (array_key_exists('axisColor', $pStyles))
            {
                if (isset($pStyles['axisColor']['auto']) && ($pStyles['axisColor']['auto']=='1'))
                {
                     $this->axisColorAuto = true;
                }
                if ((array_key_exists('axisPosition', $pStyles)) && ($this->axisPosition->toString() != PHPExceller_Style_DataBar_DataBarAxisPosition::NONE))
                {
                    $this->setAxisColor(new PHPExceller_Style_Color($pStyles['axisColor']['rgb']));
                }
            }
        }

    /*
    * Create an array of elements containing all the information of the input xml object
    *
    * @param    string    $ref    cell reference (e.g A1:A5)
    * @param    SimpleXML $cfRule    cfRule xml structure containing the databar section
    * @param    SimpleXML    $extLst    extLst structure of the worksheet the databar object is defined for
    * @return   void
    * @throws PHPExceller_Exception
    */
    public function applyFromXML($ref, $cfRule, $extLst)
    {
        // these default properties must exist
        if (isset($cfRule->dataBar->color[0]['rgb']) &&
            isset($cfRule->dataBar->cfvo[0]['type']) &&
            isset($cfRule->dataBar->cfvo[1]['type'])
            ) 
        {
            // add default properties (if they are defined)
            $this->setCellReference($ref);
            $this->setColor(new PHPExceller_Style_Color((string)$cfRule->dataBar->color[0]['rgb']));
            $this->cfvos = array(); // clear our the list before adding new ones
            $this->addCfvo(PHPExceller_Style_CFVOType::fromXML($cfRule->dataBar->cfvo[0]));
            $this->addCfvo(PHPExceller_Style_CFVOType::fromXML($cfRule->dataBar->cfvo[1]));
            if (isset($cfRule->dataBar['showValue']))
            {
                $this->setShowValue((int)$cfRule->dataBar['showValue']) ;
            }

            // check if an extLst object is used to mark up this databar
            if (isset($cfRule->extLst) &&
                isset($cfRule->extLst->ext['uri']) &&
                ($cfRule->extLst->ext['uri'] == "{B025F937-C7B1-47D3-B67F-A62EFF666E3E}") // ID for ext uri : http://msdn.microsoft.com/en-us/library/dd905242%28v=office.12%29.aspx
                )
            {
                // extract all data
                foreach ($cfRule->extLst->ext[0]->getNamespaces(true) as $ns_name => $ns_uri)
                {
                    if ($ns_name)
                    {
                        // save this namespace
                        $this->setNamespace($ns_uri);
                        // look up the id of this extLst entry
                        $children = $cfRule->extLst->ext[0]->children($ns_name,TRUE);
                        $linkid = (string)$children->id;
                        // search the data in the extLst structure
                        if ($extLst)
                        {
                            // find the object that matches the linkid
                            $linkeddatabarelements = $extLst->xpath('//*[@id="'.$linkid.'"]');
                            if ($linkeddatabarelements)
                            {
                                // create an array with all the relevant information and
                                // add it to the databar object
                                $linkeddatabarelement = $linkeddatabarelements[0]->children($ns_name,TRUE);
                                $databar = $linkeddatabarelement->dataBar;
                                $databar_array = $this->xml2array($databar,$ns_name);
                                // apply all the setting from the array
                                $isExtLstData = true;
                                $this->applyFromArray($databar_array, false, $isExtLstData);
                            }
                            else
                            {
                                // missing databar element with ID
                                throw new PHPExceller_Exception("DataBar : missing databar element entry with id ".$linkid." in extLst");
                            }
                        }
                        else
                        {
                            throw new PHPExceller_Exception("DataBar : missing extLst entry");
                        }
                    }
                }

            }
        }
        else
        {
            // missing property
            throw new PHPExceller_Exception("DataBar : missing color or cvfo setting");
        }
    }

    /*
    * Create an array of elements containing all the information of the input xml object
    *
    * @param    SimpleXML    xml object to convert to an array
    * @param    String    namespace of the elements
    * @return    Array    array containing all the information of the input xml object
    */
    protected function xml2array($inputxml, $namespace=NULL)
    {
        $result = array();
        // add attributes
        foreach ($inputxml->attributes() as $attr_name => $attr_value)
        {
            $result[$attr_name] = (string)$attr_value;
        }
        // add child objects
        if ($namespace)
        {
            $children = $inputxml->children($namespace, TRUE);
        }
        else
        {
            $children = $inputxml->children();
        }
 
        foreach ($children as $prop_name => $prop)
        {
            $prop_array = array();
            foreach ($prop->attributes() as $attr_name => $attr_value)
            {
                $prop_array[$attr_name] = (string)$attr_value;
            }

            if ($prop_name != 'cfvo')
            {
                $result[$prop_name] = $prop_array;
            }
            else
            {
                if (!isset($result['cfvos'])) { $result['cfvos'] = array(); }

                // there might be a child element with the value
                $cfvo_formula = $prop->children('xm', TRUE);
                if ($cfvo_formula)
                {
                    $prop_array['xm:f'] = (int)$cfvo_formula->f;
                }
                $result['cfvos'][] = $prop_array;
            }
        } 
        return $result;
    }

    /*
    * Create an array of elements containing all the information of the object
    *  in order to pass it on to the xmlwriter
    *
    * @return    Array    array containing all the information of this databar
    * NOTE : The order in which the elements are written into the array should not matter but for Microsoft Excel it apparently _DOES MATTER_
    */
    public function getElementsAsArray($forExtLst=false)
    {
        // 1. cfvo's (do not add to the extlst ouput)
        $result = array();
        if (($this->cfvos) && (!$forExtLst))
        {
            foreach($this->cfvos as $cfvotype)
            {
                $result[] = $cfvotype->toArray($forExtLst);
            }
        }
        // 2. extlst_cfvo's (add to the extlst ouput with name 'cfvo'; else 'extlst_cfvo')
        if ($this->extlst_cfvos)
        {
            foreach($this->extlst_cfvos as $cfvotype)
            {
                $result[] = $cfvotype->toArray($forExtLst,(($forExtLst)?'cfvo':'extlst_cfvo'));
            }
        }
        // 3. color (do not add to the extlst ouput)
        if (($this->color) && (!$forExtLst)) { $result[] = array('name' => 'color', 'attributes' => array(array('name' => 'rgb', 'attributes' => $this->color->getARGB()))); }
        // 4. fillColor
        if ($this->fillColor)          { $result[] = array('name' => 'fillColor', 'attributes' => array(array('name' => 'rgb', 'attributes' => $this->fillColor->getARGB()))); }
        // 5. borderColor (only if border = true)
        if ((!is_null($this->border)) && ($this->borderColor))
        {
            $result[] = array('name' => 'borderColor', 'attributes' => array(array('name' => 'rgb', 'attributes' => $this->borderColor->getARGB())));
        }
        // 6. negativeFillColor (only if negativeBarColorSameAsPositive = true)
        //if (($this->negativeBarColorSameAsPositive === 0) && ($this->negativeFillColor)) // Excel ignores specification : negativeBarColorSameAsPositive not needed ?
        if ($this->negativeFillColorAuto)
        {
            $result[] = array('name' => 'negativeFillColor', 'attributes' => array(array('name' => 'auto', 'attributes' => '1')));
        }
        else if ($this->negativeFillColor)
        {
            $result[] = array('name' => 'negativeFillColor', 'attributes' => array(array('name' => 'rgb', 'attributes' => $this->negativeFillColor->getARGB())));
        }

        // 7. negativeBorderColor (only if negativeBarBorderColorSameAsPositive and border are true)
        if (($this->negativeBarBorderColorSameAsPositive === 0) && ($this->negativeBorderColor) && (!is_null($this->border)))
        {
            $result[] = array('name' => 'negativeBorderColor', 'attributes' => array(array('name' => 'rgb', 'attributes' => $this->negativeBorderColor->getARGB())));
        }
        // 8. axisColor ( only if axis position is not NONE)
        //if ($this->axisPosition)
        //{
            //$axisPosition = $this->axisPosition->toString();
            //if (($axisPosition != PHPExceller_Style_DataBar_DataBarAxisPosition::NONE) && ($this->axisColor))
            if ($this->axisColorAuto)
            {
                $result[] = array('name' => 'axisColor', 'attributes' => array(array('name' => 'auto', 'attributes' => '1')));
            }
            elseif ($this->axisColor) // it seems that excel does not follow the definition here and defines an axisColor without an axisPosition
            {
                $result[] = array('name' => 'axisColor', 'attributes' => array(array('name' => 'rgb', 'attributes' => $this->axisColor->getARGB())));
            }
        //}
        // 9. minLength
        if (!is_null($this->minLength))
        {
            $result[] = array('name' => 'minLength', 'attributes' => $this->minLength);
        }
        // 10. maxLength
        if (!is_null($this->maxLength))
        {
            $result[] = array('name' => 'maxLength', 'attributes' => $this->maxLength);
        }
        // 11. showValue
        // Excel stores this attribute in the dataBar element
        /*if (!is_null($this->showValue))
        {
            $result[] = array('name' => 'showValue', 'attributes' => $this->showValue);
        }*/
        // 12. borderColor (only id borderColor also exists)
        if (($this->border) && ($this->borderColor))
        {
            $result[] = array('name' => 'border', 'attributes' => $this->border);
        }
        // 13. gradient
        if (!is_null($this->gradient))
        {
            $result[] = array('name' => 'gradient', 'attributes' => $this->gradient);
        }
        // 14. direction
        if ($this->direction)
        {
            $result[] = array('name' => 'direction', 'attributes' => $this->getDirection()->toString());
        }
        // 14. negativeBarColorSameAsPositive (only if negativeFillColor is set )
        if (!is_null($this->negativeBarColorSameAsPositive) && ($this->negativeFillColor))
        {
            $result[] = array('name' => 'negativeBarColorSameAsPositive', 'attributes' => $this->negativeBarColorSameAsPositive);
        }
        // 15. negativeBarBorderColorSameAsPositive (only if negativeBorderColor and border are set )
        if (!is_null($this->negativeBarBorderColorSameAsPositive) && ($this->negativeBorderColor) && ($this->border))
        {
            $result[] = array('name' => 'negativeBarBorderColorSameAsPositive', 'attributes' => $this->negativeBarBorderColorSameAsPositive);
        }
        // 16. axisPosition
        if ($this->axisPosition)
        {
            $axisPosition = $this->axisPosition->toString();
            if ($axisPosition == PHPExceller_Style_DataBar_DataBarAxisPosition::NONE)
            {
                $result[] = array('name' => 'axisPosition', 'attributes' => $axisPosition);
            }
            else
            {
                $result[] = array('name' => 'axisPosition', 'attributes' => $axisPosition);
            }
        }
 
        // return the resulting array
        return $result;
    }

    /*
     * get the color for the databar
     *
     * <code>
     * $worksheetstyles = $objPHPExceller->getActiveSheet()->getConditionalStyles();
     * if ($worksheetstyles[0]->getConditionType() == PHPExceller_Style_Conditional::CONDITION_DATABAR) { $color = $worksheetstyles[0]->getColor() ; }
     * </code>
     *
     * @return PHPExceller_Style_Color
    */
    public function getColor()
    {
        return $this->color;
    }

    /*
     * Set the fill color of the databar
     *
     * <code>
     * $worksheetstyles = $objPHPExceller->getActiveSheet()->getConditionalStyles();
     * if ($worksheetstyles[0]->getConditionType() == PHPExceller_Style_Conditional::CONDITION_DATABAR)
     * {
     *      $worksheetstyles[0]->setColor(new PHPExceller_Style_Color(PHPExceller_Style_Color::COLOR_RED));
     * }
     * </code>
     *
     * @param    PHPExceller_Style_Color    fill color of the databar
     * @return void
    */
    public function setColor($color = null)
    {
           $this->color = $color;
    }


    /*
     * Get the cfvo settings
     *
     * @params    boolean    set to true if the extlst cfvo settings must be used
     * <code>
     * $worksheetstyles = $objPHPExceller->getActiveSheet()->getConditionalStyles();
     * if ($worksheetstyles[0]->getConditionType() == PHPExceller_Style_Conditional::CONDITION_DATABAR) { $cfvo = $worksheetstyles[0]->getCfvo() ; }
     * </code>
     *
     * @return array of CFVOType    all CFVOTypes for this databar
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
    * @return void
    * <code>
    * $worksheetstyles = $objPHPExceller->getActiveSheet()->getConditionalStyles();
    * if ($worksheetstyles[0]->getConditionType() == PHPExceller_Style_Conditional::CONDITION_DATABAR)
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
    }
    /*
     * Add a list of cfvos (Note: existing entries will be destroyed
     *
     * @param    array    list of cfvotypes
     * @param    boolean    set to true is the data must be entered to the extlst cvfo array
     * @return void
     *
     */
    public function addCfvos($cfvo = null, $use_extlst_cfvos=false)
    {
        if ($use_extlst_cfvos)
        {
            $this->extlst_cfvos = $cfvo;
        }
        else
        {
            $this->cfvos = $cfvo;
        }
    }



    /*
     * Get the fill color
     *
     *
     * <code>
     * $worksheetstyles = $objPHPExceller->getActiveSheet()->getConditionalStyles();
     * if ($worksheetstyles[0]->getConditionType() == PHPExceller_Style_Conditional::CONDITION_DATABAR) { $fillColor = $worksheetstyles[0]->getFillColor() ; }
     * </code>
     *
     * @return PHPExceller_Style_Color
    */
    public function getFillColor()
    {
        return $this->fillColor;
    }

    /*
     * Set the fill color of the databar
     *
     * <code>
     * $worksheetstyles = $objPHPExceller->getActiveSheet()->getConditionalStyles();
     * if ($worksheetstyles[0]->getConditionType() == PHPExceller_Style_Conditional::CONDITION_DATABAR)
     * {
     *      $worksheetstyles[0]->setFillColor(new PHPExceller_Style_Color(PHPExceller_Style_Color::COLOR_GREEN));
     * }
     * </code>
     *
     * @param    PHPExceller_Style_Color    fill color of the databar
     * @return void
    */
    public function setFillColor($value)
    {
        $this->fillColor = $value;
    }

    /*
     * Get the border color
     *
     *
     * <code>
     * $worksheetstyles = $objPHPExceller->getActiveSheet()->getConditionalStyles();
     * if ($worksheetstyles[0]->getConditionType() == PHPExceller_Style_Conditional::CONDITION_DATABAR) { $borderColor = $worksheetstyles[0]->getBorderColor() ; }
     * </code>
     *
     * @return PHPExceller_Style_Color
    */
    public function getBorderColor()
    {
        return $this->borderColor;
    }

    /*
     * Set the border color of the databar
     *
     * <code>
     * $worksheetstyles = $objPHPExceller->getActiveSheet()->getConditionalStyles();
     * if ($worksheetstyles[0]->getConditionType() == PHPExceller_Style_Conditional::CONDITION_DATABAR)
     * {
     *      $worksheetstyles[0]->setBorderColor(new PHPExceller_Style_Color(PHPExceller_Style_Color::COLOR_BLUE));
     * }
     * </code>
     *
     * @param    PHPExceller_Style_Color    border color of the databar
     * @return void
    */
    public function setBorderColor($value)
    {
        $this->borderColor = $value;
    }

    /*
     * Get the negative fill color
     *
     *
     * <code>
     * $worksheetstyles = $objPHPExceller->getActiveSheet()->getConditionalStyles();
     * if ($worksheetstyles[0]->getConditionType() == PHPExceller_Style_Conditional::CONDITION_DATABAR) { $negativeFillColor = $worksheetstyles[0]->getNegativeFillColor() ; }
     * </code>
     *
     * @return PHPExceller_Style_Color
    */
    public function getNegativeFillColor()
    {
        return $this->negativeFillColor;
    }

    /*
     * Set the fill color of the negative part of databar
     *
     * <code>
     * $worksheetstyles = $objPHPExceller->getActiveSheet()->getConditionalStyles();
     * if ($worksheetstyles[0]->getConditionType() == PHPExceller_Style_Conditional::CONDITION_DATABAR)
     * {
     *      $worksheetstyles[0]->setNegativeFillColor(new PHPExceller_Style_Color(PHPExceller_Style_Color::COLOR_GREEN));
     * }
     * </code>
     *
     * @param    PHPExceller_Style_Color    fill color of the negative part of databar
     * @return void
    */
    public function setNegativeFillColor($value)
    {
        $this->negativeFillColor = $value;
        $this->negativeFillColorAuto = false;
    }

    /*
     * Get the negative border color
     *
     *
     * <code>
     * $worksheetstyles = $objPHPExceller->getActiveSheet()->getConditionalStyles();
     * if ($worksheetstyles[0]->getConditionType() == PHPExceller_Style_Conditional::CONDITION_DATABAR) { $negativeBorderColor = $worksheetstyles[0]->getNegativeBorderColor() ; }
     * </code>
     *
     * @return PHPExceller_Style_Color
    */
    public function getNegativeBorderColor()
    {
        return $this->negativeBorderColor;
    }

    /*
     * Set the axis color of the negative part of databar
     *
     * <code>
     * $worksheetstyles = $objPHPExceller->getActiveSheet()->getConditionalStyles();
     * if ($worksheetstyles[0]->getConditionType() == PHPExceller_Style_Conditional::CONDITION_DATABAR)
     * {
     *      $worksheetstyles[0]->setNegativeBorderColor(new PHPExceller_Style_Color(PHPExceller_Style_Color::COLOR_BLUE));
     * }
     * </code>
     *
     * @param    PHPExceller_Style_Color    axis color of the negative part of databar
     * @return void
    */
    public function setNegativeBorderColor($value)
    {
        $this->negativeBorderColor = $value;
    }

    /*
     * Get the axis color
     *
     * <code>
     * $worksheetstyles = $objPHPExceller->getActiveSheet()->getConditionalStyles();
     * if ($worksheetstyles[0]->getConditionType() == PHPExceller_Style_Conditional::CONDITION_DATABAR) { $axisColor = $worksheetstyles[0]->getAxisColor() ; }
     * </code>
     *
     * @return PHPExceller_Style_Color
    */
    public function getAxisColor()
    {
        return $this->axisColor;
    }

    /*
     * Set the axis color of the databar
     *
     * <code>
     * $worksheetstyles = $objPHPExceller->getActiveSheet()->getConditionalStyles();
     * if ($worksheetstyles[0]->getConditionType() == PHPExceller_Style_Conditional::CONDITION_DATABAR)
     * {
     *      $worksheetstyles[0]->setAxisColor(new PHPExceller_Style_Color(PHPExceller_Style_Color::COLOR_BLUE));
     * }
     * </code>
     *
     * @param    PHPExceller_Style_Color    axis color
     * @return void
    */
    public function setAxisColor($value)
    {
        $this->axisColor = $value;
        $this->axiscolorAuto = false;
    }

    /*
     * Get the minimum length of the bar
     *
     *
     * <code>
     * $worksheetstyles = $objPHPExceller->getActiveSheet()->getConditionalStyles();
     * if ($worksheetstyles[0]->getConditionType() == PHPExceller_Style_Conditional::CONDITION_DATABAR) { $minLength = $worksheetstyles[0]->getMinLength() ; }
     * </code>
     *
     * @return unsigned int
    */
    public function getMinLength()
    {
        return $this->minLength;
    }

    /*
     * Set the minimum length of the databar
     *
     * <code>
     * $worksheetstyles = $objPHPExceller->getActiveSheet()->getConditionalStyles();
     * if ($worksheetstyles[0]->getConditionType() == PHPExceller_Style_Conditional::CONDITION_DATABAR)
     * {
     *      $worksheetstyles[0]->setMinLength(50);
     * }
     * </code>
     *
     * @param    unigned int    minimum length of the databar
     * @return void
    */
    public function setMinLength($value)
    {
        if (($value >= 0)&&($value > 100))
        {
            throw new PHPExceller_Exception("DataBar : minLength should be in the range of 0 to 100");
        }
        $this->minLength = $value;
    }

    /*
     * Get the maximum length of the bar
     *
     *
     * <code>
     * $worksheetstyles = $objPHPExceller->getActiveSheet()->getConditionalStyles();
     * if ($worksheetstyles[0]->getConditionType() == PHPExceller_Style_Conditional::CONDITION_DATABAR) { $maxLength = $worksheetstyles[0]->getMaxLength() ; }
     * </code>
     *
     * @return unsigned int
    */
    public function getMaxLength()
    {
        return $this->maxLength;
    }

    /*
     * Set the maximum length of the databar
     *
     * <code>
     * $worksheetstyles = $objPHPExceller->getActiveSheet()->getConditionalStyles();
     * if ($worksheetstyles[0]->getConditionType() == PHPExceller_Style_Conditional::CONDITION_DATABAR)
     * {
     *      $worksheetstyles[0]->setMaxLength(50);
     * }
     * </code>
     *
     * @param    unigned int    maximum length of the databar
     * @return void
    */
    public function setMaxLength($value)
    {
        if (($value >= 0)&&($value > 100))
        {
            throw new PHPExceller_Exception("DataBar : maxLength should be in the range of 0 to 100");
        }
        $this->maxLength = $value;
    }

    /*
     * Get whether the cell value will be shown
     *
     *
     * <code>
     * $worksheetstyles = $objPHPExceller->getActiveSheet()->getConditionalStyles();
     * if ($worksheetstyles[0]->getConditionType() == PHPExceller_Style_Conditional::CONDITION_DATABAR) { $showValue = $worksheetstyles[0]->getShowValue() ; }
     * </code>
     *
     * @return integer (0=false;1=true)
    */
    public function getShowValue()
    {
        return $this->showValue;
    }

    /*
     * Indicate whether the to show the value of the cell
     *
     * <code>
     * $worksheetstyles = $objPHPExceller->getActiveSheet()->getConditionalStyles();
     * if ($worksheetstyles[0]->getConditionType() == PHPExceller_Style_Conditional::CONDITION_DATABAR)
     * {
     *      $worksheetstyles[0]->setShowValue(1);
     * }
     * </code>
     *
     * @param    integer    1 if the value of the cell must be shown
     * @return void
    */
    public function setShowValue($value)
    {
        $this->showValue = $value;
    }

    /*
     * Get whether the show the border
     *
     *
     * <code>
     * $worksheetstyles = $objPHPExceller->getActiveSheet()->getConditionalStyles();
     * if ($worksheetstyles[0]->getConditionType() == PHPExceller_Style_Conditional::CONDITION_DATABAR) { $showBorder = $worksheetstyles[0]->getBorder() ; }
     * </code>
     *
     * @return integer
    */
    public function getBorder()
    {
        return $this->border;
    }

    /*
     * Indicate whether the to use a border for the databar
     *
     * <code>
     * $worksheetstyles = $objPHPExceller->getActiveSheet()->getConditionalStyles();
     * if ($worksheetstyles[0]->getConditionType() == PHPExceller_Style_Conditional::CONDITION_DATABAR)
     * {
     *      $worksheetstyles[0]->setBorder(1);
     * }
     * </code>
     *
     * @param    integer    1 if a border must be used for the databar
     * @return void
    */
    public function setBorder($value)
    {
        $this->border = $value;
    }

    /*
     * Get whether to use a gradient
     *
     *
     * <code>
     * $worksheetstyles = $objPHPExceller->getActiveSheet()->getConditionalStyles();
     * if ($worksheetstyles[0]->getConditionType() == PHPExceller_Style_Conditional::CONDITION_DATABAR) { $showGradient = $worksheetstyles[0]->getGradient() ; }
     * </code>
     *
     * @return bool
    */
    public function getGradient()
    {
        return $this->gradient;
    }

    /*
     * Indicate whether the to use a gradient for the databar
     *
     * <code>
     * $worksheetstyles = $objPHPExceller->getActiveSheet()->getConditionalStyles();
     * if ($worksheetstyles[0]->getConditionType() == PHPExceller_Style_Conditional::CONDITION_DATABAR)
     * {
     *      $worksheetstyles[0]->setGradient(0);
     * }
     * </code>
     *
     * @param    integer    0 if no gradient must be used
     * @return void
    */
    public function setGradient($value)
    {
        $this->gradient = $value;
    }

    /*
     * Get the direction of the databar
     *
     *
     * <code>
     * $worksheetstyles = $objPHPExceller->getActiveSheet()->getConditionalStyles();
     * if ($worksheetstyles[0]->getConditionType() == PHPExceller_Style_Conditional::CONDITION_DATABAR) { $direction = $worksheetstyles[0]->getDirection() ; }
     * </code>
     *
     * @return PHPExceller_Style_DataBar_DataBarAxisDirection
    */
    public function getDirection()
    {
        return $this->direction;
    }

    /*
     * Set the axis direction setting
     *
     *
     * <code>
     * $worksheetstyles = $objPHPExceller->getActiveSheet()->getConditionalStyles();
     * if ($worksheetstyles[0]->getConditionType() == PHPExceller_Style_Conditional::CONDITION_DATABAR)
     * {
     *      $worksheetstyles[0]->setDirection(new PHPExceller_Style_DataBar_DataBarDirection(PHPExceller_Style_DataBar_DataBarDirection::fromString('context')));
     * }
     * </code>
     *
     * @param PHPExceller_Style_DataBar_DataBarDirection
     * @return void
    */
    public function setDirection($value)
    {
        $this->direction = $value;
    }

    /*
     * Get whether to use the same color for the negative part of the databar
     *
     *
     * <code>
     * $worksheetstyles = $objPHPExceller->getActiveSheet()->getConditionalStyles();
     * if ($worksheetstyles[0]->getConditionType() == PHPExceller_Style_Conditional::CONDITION_DATABAR)
     * {
     *   $useSameColorForNegative = $worksheetstyles[0]->getNegativeBarColorSameAsPositive() ;
     * }
     * </code>
     *
     * @return integer 0=false; 1=true
    */
    public function getNegativeBarColorSameAsPositive()
    {
        return $this->negativeBarColorSameAsPositive;
    }

    /*
     * Indicate whether the negative bar color needs to be the same as the positive bar color
     *
     *
     * <code>
     * $worksheetstyles = $objPHPExceller->getActiveSheet()->getConditionalStyles();
     * if ($worksheetstyles[0]->getConditionType() == PHPExceller_Style_Conditional::CONDITION_DATABAR)
     * {
     *      $worksheetstyles[0]->setNegativeBarColorSameAsPositive(true);
     * }
     * </code>
     *
     * @param    integer    1=true if the negative bar color needs to be the same as the positive bar color
     * @return void
    */
    public function setNegativeBarColorSameAsPositive($value)
    {
        $this->negativeBarColorSameAsPositive = $value;
    }

    /*
     * Get whether to use the same color for the border of the negative part of the databar
     *
     *
     * <code>
     * $worksheetstyles = $objPHPExceller->getActiveSheet()->getConditionalStyles();
     * if ($worksheetstyles[0]->getConditionType() == PHPExceller_Style_Conditional::CONDITION_DATABAR)
     * {
     *   $useSameColorForNegativeBorder = $worksheetstyles[0]->getNegativeBarBorderColorSameAsPositive() ;
     * }
     * </code>
     *
     * @return integer (0=false;1=true)
    */
    public function getNegativeBarBorderColorSameAsPositive()
    {
        return $this->negativeBarBorderColorSameAsPositive;
    }

    /*
     * Indicate whether the negative bar border color needs to be the same as the positive bar border color
     *
     *
     * <code>
     * $worksheetstyles = $objPHPExceller->getActiveSheet()->getConditionalStyles();
     * if ($worksheetstyles[0]->getConditionType() == PHPExceller_Style_Conditional::CONDITION_DATABAR)
     * {
     *      $worksheetstyles[0]->setNegativeBarBorderColorSameAsPositive(1);
     * }
     * </code>
     *
     * @param    integer    1(=true) if the negative bar border color needs to be the same as the positive bar border color
     * @return void
    */
    public function setNegativeBarBorderColorSameAsPositive($value)
    {
        $this->negativeBarBorderColorSameAsPositive = $value;
    }

    /*
     * Get the axis position setting
     *
     * <code>
     * $worksheetstyles = $objPHPExceller->getActiveSheet()->getConditionalStyles();
     * if ($worksheetstyles[0]->getConditionType() == PHPExceller_Style_Conditional::CONDITION_DATABAR) { $axisPosition = $worksheetstyles[0]->getAxisPosition(); }
     * </code>
     *
     * @return PHPExceller_Style_DataBar_DataBarAxisPosition
    */
    public function getAxisPosition()
    {
        return $this->axisPosition;
    }

    /*
     * Set the axis position setting
     *
     * <code>
     * $worksheetstyles = $objPHPExceller->getActiveSheet()->getConditionalStyles();
     * if ($worksheetstyles[0]->getConditionType() == PHPExceller_Style_Conditional::CONDITION_DATABAR)
     * {
     *      $worksheetstyles[0]->setAxisPosition(new PHPExceller_Style_DataBar_DataBarAxisPosition(PHPExceller_Style_DataBar_DataBarAxisPosition::fromString('automatic')));
     * }
     * </code>
     *
     * @param PHPExceller_Style_DataBar_DataBarAxisPosition
     * @return void
    */
    public function setAxisPosition($value)
    {
        $this->axisPosition = $value;
    }

    /*
     * Set the namespace setting for creating the extLst entry
     * Default already set to http://schemas.microsoft.com/office/spreadsheetml/2009/9/main
     *
     * <code>
     * $worksheetstyles = $objPHPExceller->getActiveSheet()->getConditionalStyles();
     * $worksheetstyles[0]->setNamespace('http://schemas.microsoft.com/office/spreadsheetml/2009/9/main)');
     * </code>
     *
     * @return string containing namespace setting
    */
    public function setNamespace($namespace)
    {
        $this->namespace = $namespace;
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
     * Get an array containing the databar data that must be written to the extLst entry of the worksheet
     *
     * @params priority not used
     * @return    array    array containing all the data that must be written to the extLst entry of the worksheert
     *
     */
    public function getExtLstData($priority)
    {
        $forExtLst = true;
        $data = $this->getElementsAsArray($forExtLst);
            // add id andtype to
        $result = array('name' => 'cfRule',
                        'cellReference' => $this->getCellReference(),
                        'attributes' => array(array('name' => 'type',   'attributes' => 'dataBar'),
                                               array('name' => 'id',     'attributes' => $this->getClassID()),
                                               array('name' => 'dataBar','attributes' => $data)));
 
        return $result;
    }

    /*
     * create a datastructure for creating the databar element with default properties
     *
     * @param    array    array with properties
     */
    public function getDefaultData()
    {
        $cfvos = $this->getCfvos();
        $result = array('name' => 'dataBar',
                        'attributes' => array($cfvos[0]->toArray(),
                                              $cfvos[1]->toArray(),
                                              array('name' => 'color', 'attributes' => array(array('name' => 'rgb', 'attributes' => $this->getColor()->getARGB())))
                                              )
                        );
        if (!is_null($this->showValue)) { $result['attributes'][] = array('name' => 'showValue' , 'attributes' => $this->showValue); }
        return $result;
    }

    /*
     * Indicates whether this object needs a reference to the entry in the extLst section
     *
     * @returns    bool    true is such a reference is needed
     *
     */
     public function needsExtLstReference()
     {
        return true;

 }

}

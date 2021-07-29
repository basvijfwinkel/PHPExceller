<?php
namespace PHPExceller\Calculation;

use PHPExceller\Calculation\Exception;

/**
 * PHPExceller_Calculation_Function
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
class Function
{
    /* Function categories */
    const CATEGORY_CUBE                 = 'Cube';
    const CATEGORY_DATABASE             = 'Database';
    const CATEGORY_DATE_AND_TIME        = 'Date and Time';
    const CATEGORY_ENGINEERING          = 'Engineering';
    const CATEGORY_FINANCIAL            = 'Financial';
    const CATEGORY_INFORMATION          = 'Information';
    const CATEGORY_LOGICAL              = 'Logical';
    const CATEGORY_LOOKUP_AND_REFERENCE = 'Lookup and Reference';
    const CATEGORY_MATH_AND_TRIG        = 'Math and Trig';
    const CATEGORY_STATISTICAL          = 'Statistical';
    const CATEGORY_TEXT_AND_DATA        = 'Text and Data';

    /**
     * Category (represented by CATEGORY_*)
     *
     * @var string
     */
    private $category;

    /**
     * Excel name
     *
     * @var string
     */
    private $excelName;

    /**
     * PHPExceller name
     *
     * @var string
     */
    private $PHPExcellerName;

    /**
     * Create a new PHPExceller_Calculation_Function
     *
     * @param     string        $pCategory         Category (represented by CATEGORY_*)
     * @param     string        $pExcelName        Excel function name
     * @param     string        $pPHPExcellerName    PHPExceller function mapping
     * @throws     PHPExceller_Calculation_Exception
     */
    public function __construct($pCategory = null, $pExcelName = null, $pPHPExcellerName = null)
    {
        if (($pCategory !== null) && ($pExcelName !== null) && ($pPHPExcellerName !== null)) {
            // Initialise values
            $this->category     = $pCategory;
            $this->excelName    = $pExcelName;
            $this->PHPExcellerName = $pPHPExcellerName;
        } else {
            throw new PHPExceller_Calculation_Exception("Invalid parameters passed.");
        }
    }

    /**
     * Get Category (represented by CATEGORY_*)
     *
     * @return string
     */
    public function getCategory()
    {
        return $this->category;
    }

    /**
     * Set Category (represented by CATEGORY_*)
     *
     * @param     string        $value
     * @throws     PHPExceller_Calculation_Exception
     */
    public function setCategory($value = null)
    {
        if (!is_null($value)) {
            $this->category = $value;
        } else {
            throw new PHPExceller_Calculation_Exception("Invalid parameter passed.");
        }
    }

    /**
     * Get Excel name
     *
     * @return string
     */
    public function getExcelName()
    {
        return $this->excelName;
    }

    /**
     * Set Excel name
     *
     * @param string    $value
     */
    public function setExcelName($value)
    {
        $this->excelName = $value;
    }

    /**
     * Get PHPExceller name
     *
     * @return string
     */
    public function getPHPExcellerName()
    {
        return $this->PHPExcellerName;
    }

    /**
     * Set PHPExceller name
     *
     * @param string    $value
     */
    public function setPHPExcellerName($value)
    {
        $this->PHPExcellerName = $value;
    }
}

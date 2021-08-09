<?php
namespace PHPExceller\Calculation;

use PHPExceller\Calculation\Exception;
use PHPExceller\Calculation\Calculation_Exception;

/**
 * Based in PHPExcel_Calculation_Function
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
     * Create a new Function
     *
     * @param     string        $pCategory         Category (represented by CATEGORY_*)
     * @param     string        $pExcelName        Excel function name
     * @param     string        $pPHPExcellerName    PHPExceller function mapping
     * @throws    Calculation_Exception
     */
    public function __construct($pCategory = null, $pExcelName = null, $pPHPExcellerName = null)
    {
        if (($pCategory !== null) && ($pExcelName !== null) && ($pPHPExcellerName !== null)) {
            // Initialise values
            $this->category     = $pCategory;
            $this->excelName    = $pExcelName;
            $this->PHPExcellerName = $pPHPExcellerName;
        } else {
            throw new Calculation_Exception("Invalid parameters passed.");
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
     * @throws     Calculation_Exception
     */
    public function setCategory($value = null)
    {
        if (!is_null($value)) {
            $this->category = $value;
        } else {
            throw new Calculation_Exception("Invalid parameter passed.");
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

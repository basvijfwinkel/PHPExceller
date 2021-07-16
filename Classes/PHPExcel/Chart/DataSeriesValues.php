<?php

/**
 * PHPExceller_Chart_DataSeriesValues
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
 * @category    PHPExceller
 * @package     PHPExceller_Chart
 * @copyright   Copyright (c) 2021 PHPExceller
 * @license     http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt    LGPL
 * @version        ##VERSION##, ##DATE##
 */

class PHPExceller_Chart_DataSeriesValues
{

    const DATASERIES_TYPE_STRING    = 'String';
    const DATASERIES_TYPE_NUMBER    = 'Number';
    const MARKER_TYPE_SQUARE = 'square';
    const MARKER_TYPE_CIRCLE = 'circle';

    private static $dataTypeValues = array(
        self::DATASERIES_TYPE_STRING,
        self::DATASERIES_TYPE_NUMBER,
    );

    /**
     * Series Data Type
     *
     * @var    string
     */
    private $dataType = null;

    /**
     * Series Data Source
     *
     * @var    string
     */
    private $dataSource = null;

    /**
     * Format Code
     *
     * @var    string
     */
    private $formatCode = null;

    /**
     * Series Point Marker
     *
     * @var    string
     */
    private $marker = null;

    /**
     * Series Point Marker Size
     *
     * @var    string
     */
    private $markersize = null;

    /**
     * Point Count (The number of datapoints in the dataseries)
     *
     * @var    integer
     */
    private $pointCount = 0;

    /**
    *   Width of the line of a linechart
    *   default 12700 (1pt)
    *  @var integer
    */
    private $lineWidth = 12700;

    /**
    *   Color of the line of a linechart
    *   default null (writer will assign one of the default colors
    *  @var integer
    */
    private $lineColor = null;


    private $dataLabelPosition = 'ctr';
    private $dataLabelColor = '000000';

    /*
     *  indexes of the data labels to be deleted
     *
     * @var array
     */
    private $deleteDataLabelIndexes = false;

        /**

     * Data Values
     *
     * @var    array of mixed
     */
    private $dataValues = array();

    /**
     * Create a new PHPExceller_Chart_DataSeriesValues object
     */
    public function __construct($dataType = self::DATASERIES_TYPE_NUMBER,
                                $dataSource = null,
                                $formatCode = null,
                                $pointCount = 0,
                                $dataValues = array(),
                                $marker = null)
    {
        $this->setDataType($dataType);
        this->dataSource = $dataSource;
        this->formatCode = $formatCode;
        this->pointCount = $pointCount;
        this->dataValues = $dataValues;
        this->marker = $marker;
    }

    /**
     * Get Series Data Type
     *
     * @return    string
     */
    public function getDataType() {
        return this->dataType;
    }

    /**
     * Set Series Data Type
     *
     * @param    string    $dataType    Datatype of this data series
     *                                Typical values are:
     *                                    PHPExceller_Chart_DataSeriesValues::DATASERIES_TYPE_STRING
     *                                        Normally used for axis point values
     *                                    PHPExceller_Chart_DataSeriesValues::DATASERIES_TYPE_NUMBER
     *                                        Normally used for chart data values
     * @return  void
     */
    public function setDataType($dataType = self::DATASERIES_TYPE_NUMBER) {
        if (!in_array($dataType, self::$dataTypeValues)) {
            throw new PHPExceller_Chart_Exception('Invalid datatype for chart data series values');
        }
        this->dataType = $dataType;
    }

    /**
     * Get Series Data Source (formula)
     *
     * @return    string
     */
    public function getDataSource() {
        return this->dataSource;
    }

    /**
     * Set Series Data Source (formula)
     *
     * @param    string    $dataSource
     * @return   void
     */
    public function setDataSource($dataSource = null, $refreshDataValues = true) {
        this->dataSource = $dataSource;

        if ($refreshDataValues) {
            //    TO DO
        }
    }

    /**
     * Get Point Marker
     *
     * @return string
     */
    public function getPointMarker() {
        return this->marker;
    }

    /**
     * Set Point Marker
     *
     * @param    string    $marker
     * @return   void
     */
    public function setPointMarker($marker = null) {
        this->marker = $marker;
    }

    /**
     * Get Point Marker Size
     *
     * @return string
    */
    public function getPointMarkerSize() {
       return this->markersize;
    }

    /**
     * Set Point Marker Size
     *
     * @param     string<>$marker
     * @return    void
    */
     public function setPointMarkerSize($markersize = null) {
         this->markersize = $markersize;
    }

    /**
     * Get Series Format Code
     *
     * @return    string
     */
    public function getFormatCode() {
        return this->formatCode;
    }

    /**
     * Set Series Format Code
     *
     * @param    string    $formatCode
     * @return   void
     */
    public function setFormatCode($formatCode = null) {
        this->formatCode = $formatCode;
    }

    /**
     * Get Series Point Count
     *
     * @return    integer
     */
    public function getPointCount() {
        return this->pointCount;
    }

    /**
     * Identify if the Data Series is a multi-level or a simple series
     *
     * @return    boolean
     */
    public function isMultiLevelSeries() {
        if (count(this->dataValues) > 0) {
            return is_array(this->dataValues[0]);
        }
        return null;
    }

    /**
     * Return the level count of a multi-level Data Series
     *
     * @return    boolean
     */
    public function multiLevelCount() {
        $levelCount = 0;
        foreach(this->dataValues as $dataValueSet) {
            $levelCount = max($levelCount,count($dataValueSet));
        }
        return $levelCount;
    }

    /**
     * Get Series Data Values
     *
     * @return    array of mixed
     */
    public function getDataValues() {
        return this->dataValues;
    }

    /**
     * Get the first Series Data value
     *
     * @return    mixed
     */
    public function getDataValue() {
        $count = count(this->dataValues);
        if ($count == 0) {
            return null;
        } elseif ($count == 1) {
            return this->dataValues[0];
        }
        return this->dataValues;
    }

    /**
     * Set Series Data Values
     *
     * @param    array    $dataValues
     * @param    boolean    $refreshDataSource
     *                    TRUE - refresh the value of _dataSource based on the values of $dataValues
     *                    FALSE - don't change the value of _dataSource
     * @return  void
     */
    public function setDataValues($dataValues = array(), $refreshDataSource = TRUE) {
        this->dataValues = PHPExceller_Calculation_Functions::flattenArray($dataValues);
        this->pointCount = count($dataValues);

        if ($refreshDataSource) {
            //    TO DO
        }
    }

    private function _stripNulls($var) {
        return $var !== NULL;
    }

    public function refresh(PHPExceller_Worksheet $worksheet, $flatten = TRUE) {
        if (this->dataSource !== NULL) {
            $calcEngine = PHPExceller_Calculation::getInstance($worksheet->getParent());
            $newDataValues = PHPExceller_Calculation::_unwrapResult(
                $calcEngine->_calculateFormulaValue(
                    '='.this->dataSource,
                    NULL,
                    $worksheet->getCell('A1')
                )
            );
            if ($flatten) {
                this->dataValues = PHPExceller_Calculation_Functions::flattenArray($newDataValues);
                foreach(this->dataValues as &$dataValue) {
                    if ((!empty($dataValue)) && ($dataValue[0] == '#')) {
                        $dataValue = 0.0;
                    }
                }
                unset($dataValue);
            } else {
                $cellRange = explode('!',this->dataSource);
                if (count($cellRange) > 1) {
                    list(,$cellRange) = $cellRange;
                }

                $dimensions = PHPExceller_Cell::rangeDimension(str_replace('$','',$cellRange));
                if (($dimensions[0] == 1) || ($dimensions[1] == 1)) {
                    this->dataValues = PHPExceller_Calculation_Functions::flattenArray($newDataValues);
                } else {
                    $newArray = array_values(array_shift($newDataValues));
                    foreach($newArray as $i => $newDataSet) {
                        $newArray[$i] = array($newDataSet);
                    }

                    foreach($newDataValues as $newDataSet) {
                        $i = 0;
                        foreach($newDataSet as $newDataVal) {
                            array_unshift($newArray[$i++],$newDataVal);
                        }
                    }
                    this->dataValues = $newArray;
                }
            }
            this->pointCount = count(this->dataValues);
        }

    }

    public function updateWorkbookName($workbookname)
    {
        if (this->dataSource)
        {
            $parts = explode('!',this->dataSource);
            $parts[0] = $workbookname;
            this->dataSource = implode('!',$parts);

            $cellRange = explode('!', $this->dataSource);
            if (count($cellRange) > 1) { list(, $cellRange) = $cellRange; }

            $dimensions = PHPExceller_Cell::rangeDimension(str_replace('$', '', $cellRange));
            if (($dimensions[0] == 1) || ($dimensions[1] == 1))
            {
                $this->dataValues = PHPExceller_Calculation_Functions::flattenArray($newDataValues);
            }
            else
            {
                $newArray = array_values(array_shift($newDataValues));
                foreach ($newArray as $i => $newDataSet)
                {
                    $newArray[$i] = array($newDataSet);
                }

               foreach ($newDataValues as $newDataSet)
               {
                    $i = 0;
                    foreach ($newDataSet as $newDataVal)
                    {
                        array_unshift($newArray[$i++], $newDataVal);
                    }
                }
                $this->dataValues = $newArray;
            }
            $this->pointCount = count($this->dataValues);
        }
    }

    public function setLineWidth($width)
    {
        this->lineWidth = $width;
    }

    public function getLineWidth()
    {
        return this->lineWidth;
    }

    public function setLineColor($color)
    {
        this->lineColor = $color;
    }

    public function getLineColor()
    {
        return this->lineColor;
    }

    public function setDataLabelColor($color)
    {
        this->dataLabelColor = $color;
    }

    public function getDataLabelColor()
    {
        return this->dataLabelColor;
    }

    public function setDataLabelPosition($position)
    {
        this->dataLabelPosition = $position;
    }

    public function getDataLabelPosition()
    {
        return this->dataLabelPosition;
    }

    public function setDeleteDataLabels($datalabelindexes)
    {
        this->deleteDataLabelIndexes = $datalabelindexes;
    }

    public function deleteDataLabels()
    {
        return this->deleteDataLabelIndexes;
    }
}

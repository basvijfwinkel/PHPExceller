<?php

/**
 * PHPExceller_Chart
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
 * @copyright   Copyright (c) 2021
 * @license     http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt    LGPL
 * @version     ##VERSION##, ##DATE##
 */
class PHPExceller_Chart
{
    /**
     * Chart Name
     *
     * @var string
     */
    private $name = '';

    /**
     * Worksheet
     *
     * @var PHPExceller_Worksheet
     */
    private $worksheet = null;

    /**
     * Chart Title
     *
     * @var PHPExceller_Chart_Title
     */
    private $title = null;

    /**
     * Chart Legend
     *
     * @var PHPExceller_Chart_Legend
     */
    private $legend = null;

    /**
     * X-Axis Label
     *
     * @var PHPExceller_Chart_Title
     */
    private $xAxisLabel = null;

    /**
     * Y-Axis Label
     *
     * @var PHPExceller_Chart_Title
     */
    private $yAxisLabel = null;

    /**
     * Chart Plot Area
     *
     * @var PHPExceller_Chart_PlotArea
     */
    private $plotArea = null;

    /**
     * Plot Visible Only
     *
     * @var boolean
     */
    private $plotVisibleOnly = true;

    /**
     * Display Blanks as
     *
     * @var string
     */
    private $displayBlanksAs = '0';

  /**
   * Chart Asix Y as
   *
   * @var PHPExceller_Chart_Axis
   */
  private $yAxis = null;

  /**
   * Chart Asix X as
   *
   * @var PHPExceller_Chart_Axis
   */
  private $xAxis = null;

  /**
   * Chart Major Gridlines as
   *
   * @var PHPExceller_Chart_GridLines
   */
  private $majorGridlines = null;

  /**
   * Chart Minor Gridlines as
   *
   * @var PHPExceller_Chart_GridLines
   */
  private $minorGridlines = null;

    /**
     * Top-Left Cell Position
     *
     * @var string
     */
    private $topLeftCellRef = 'A1';


    /**
     * Top-Left X-Offset
     *
     * @var integer
     */
    private $topLeftXOffset = 0;


    /**
     * Top-Left Y-Offset
     *
     * @var integer
     */
    private $topLeftYOffset = 0;


    /**
     * Bottom-Right Cell Position
     *
     * @var string
     */
    private $bottomRightCellRef = 'A1';


    /**
     * Bottom-Right X-Offset
     *
     * @var integer
     */
    private $bottomRightXOffset = 10;


    /**
     * Bottom-Right Y-Offset
     *
     * @var integer
     */
    private $bottomRightYOffset = 10;

    /**
     * Secondary Y-axis
     * Note : no checks are performed if this model contains 2 dataseries or is the right type
     * Set to NULL if not used (default)
     * @var integer
     */
    private $useSecondaryYAxis = false;

        /**
         * Secondary Y-axis
         * Note : no checks are performed if this model contains 2 dataseries or is the right type
         * Set to false if not used (default)
         * @var integer
        */
        private $secondaryYAxisLabel = NULL;
    /**
     * Create a new PHPExceller_Chart
     */
    public function __construct($name, 
                                PHPExceller_Chart_Title $title = null, 
                                PHPExceller_Chart_Legend $legend = null, 
                                PHPExceller_Chart_PlotArea $plotArea = null, 
                                $plotVisibleOnly = true, 
                                $displayBlanksAs = '0', 
                                PHPExceller_Chart_Title $xAxisLabel = null, 
                                PHPExceller_Chart_Title $yAxisLabel = null, 
                                PHPExceller_Chart_Axis $xAxis = null, 
                                PHPExceller_Chart_Axis $yAxis = null, 
                                PHPExceller_Chart_GridLines $majorGridlines = null, 
                                PHPExceller_Chart_GridLines $minorGridlines = null,
                                PHPExceller_Chart_Axis $secondaryYAxis = NULL,
                                PHPExceller_Chart_Axis $secondaryXAxis = NULL)
    {
        $this->name = $name;
        $this->title = $title;
        $this->legend = $legend;
        $this->xAxisLabel = $xAxisLabel;
        $this->yAxisLabel = $yAxisLabel;
        $this->plotArea = $plotArea;
        $this->plotVisibleOnly = $plotVisibleOnly;
        $this->displayBlanksAs = $displayBlanksAs;
        $this->xAxis = $xAxis;
        $this->yAxis = $yAxis;
        $this->majorGridlines = $majorGridlines;
        $this->minorGridlines = $minorGridlines;
        $this->secondaryYAxis = $secondaryYAxis;
        $this->secondaryXAxis = $secondaryXAxis;
    }

    /**
    * Set the secondary Y-Axis
    * @param    Boolean $use 
    * @return    PHPExceller_Chart
    *
    */
    public function setSecondaryYAxis($secondaryYAxis)
    {
        $this->secondaryYAxis = $secondaryYAxis;
                $this->useSecondaryYAxis = true;
    }
    
    /**
    * Set the secondary X-Axis
    * @param    Boolean $use 
    * @return    PHPExceller_Chart
    *
    */
    public function setSecondaryXAxis($secondaryXAxis)
    {
        $this->secondaryXAxis = $secondaryXAxis;
                $this->useSecondaryYAxis = true;
    }

        /**
        * Set the secondary Y-Axis label
        * @param PHPExceller_Chart_Title $secondaryYAxisLabel
        * @return void
        *
        */
        public function setSecondaryYAxisLabel($secondaryYAxisLabel)
        {
            $this->secondaryYAxisLabel = $secondaryYAxisLabel;
        }

    
    /**
    * Get the secondary Y-Axis
    * @return void
    *
    */
    public function getSecondaryYAxis()
    {
             if (!is_null($this->secondaryYAxis))
             {
        return $this->secondaryYAxis;
             }
             else
             {
                 return $this->getChartAxisY();
             }
             
    }

       /**
       * Get the secondary Y-Axis Label
       * @return void
       *
       */
       public function getSecondaryYAxisLabel()
       {
           return $this->secondaryYAxisLabel;
        }

    
    /**
    * Get the secondary X-Axis
    * @return void
    *
    */
    public function getSecondaryXAxis()
    {
            if (!is_null($this->secondaryXAxis))
            {
                return $this->secondaryXAxis;
            }
            else
            {
                // no specific secondary xAxis defined -> use the same as the primary xAxis
                return $this->getChartAxisX();
            }
        }

        /**
        *   Get whether a seconday Axis is used
        *  @return Boolean
        */
        public function usesSecondaryAxis()
        {
            return $this->useSecondaryYAxis;
        }
    
    /**
     * Get Name
     *
     * @return string
     */
    public function getName() {
        return $this->name;
    }

    /**
     * Get Worksheet
     *
     * @return PHPExceller_Worksheet
     */
    public function getWorksheet() {
        return $this->worksheet;
    }

    /**
     * Set Worksheet
     *
     * @param    PHPExceller_Worksheet    $pValue
     * @throws    PHPExceller_Chart_Exception
     * @return void
     */
    public function setWorksheet(PHPExceller_Worksheet $pValue = null) {
        $this->worksheet = $pValue;

    }

    /**
     * Get Title
     *
     * @return PHPExceller_Chart_Title
     */
    public function getTitle() {
        return $this->title;
    }

    /**
     * Set Title
     *
     * @param    PHPExceller_Chart_Title $title
     * @return void
     */
    public function setTitle(PHPExceller_Chart_Title $title) {
        $this->title = $title;
    }

    /**
     * Get Legend
     *
     * @return void
     */
    public function getLegend() {
        return $this->legend;
    }

    /**
     * Set Legend
     *
     * @param    PHPExceller_Chart_Legend $legend
     * @return    PHPExceller_Chart
     */
    public function setLegend(PHPExceller_Chart_Legend $legend) {
        $this->legend = $legend;
    }

    /**
     * Get X-Axis Label
     *
     * @return PHPExceller_Chart_Title
     */
    public function getXAxisLabel() {
        return $this->xAxisLabel;
    }

    /**
     * Set X-Axis Label
     *
     * @param    PHPExceller_Chart_Title $label
     * @return    PHPExceller_Chart
     */
    public function setXAxisLabel(PHPExceller_Chart_Title $label) {
        $this->xAxisLabel = $label;
    }

    /**
     * Get Y-Axis Label
     *
     * @return PHPExceller_Chart_Title
     */
    public function getYAxisLabel() {
        return $this->yAxisLabel;
    }

    /**
     * Set Y-Axis Label
     *
     * @param    PHPExceller_Chart_Title $label
     * @return    PHPExceller_Chart
     */
    public function setYAxisLabel(PHPExceller_Chart_Title $label) {
        $this->yAxisLabel = $label;
    }

    /**
     * Get Plot Area
     *
     * @return PHPExceller_Chart_PlotArea
     */
    public function getPlotArea() {
        return $this->plotArea;
    }

    /**
     * Get Plot Visible Only
     *
     * @return boolean
     */
    public function getPlotVisibleOnly() {
        return $this->plotVisibleOnly;
    }

    /**
     * Set Plot Visible Only
     *
     * @param boolean $plotVisibleOnly
     * @return void
     */
    public function setPlotVisibleOnly($plotVisibleOnly = true) {
        $this->plotVisibleOnly = $plotVisibleOnly;
    }

    /**
     * Get Display Blanks as
     *
     * @return string
     */
    public function getDisplayBlanksAs() {
        return $this->displayBlanksAs;
    }

    /**
     * Set Display Blanks as
     *
     * @param string $displayBlanksAs
     * @return void
     */
    public function setDisplayBlanksAs($displayBlanksAs = '0') {
        $this->displayBlanksAs = $displayBlanksAs;
    }


  /**
   * Get yAxis
   *
   * @return PHPExceller_Chart_Axis
   */
  public function getChartAxisY() {
    if($this->yAxis !== NULL){
      return $this->yAxis;
    }

    return new PHPExceller_Chart_Axis();
  }

  /**
   * Get xAxis
   *
   * @return PHPExceller_Chart_Axis
   */
  public function getChartAxisX() {
    if($this->xAxis !== NULL){
      return $this->xAxis;
    }

    return new PHPExceller_Chart_Axis();
  }

  /**
   * Get Major Gridlines
   *
   * @return PHPExceller_Chart_GridLines
   */
  public function getMajorGridlines() {
    if($this->majorGridlines !== NULL){
      return $this->majorGridlines;
    }

    return new PHPExceller_Chart_GridLines();
  }

  /**
   * Get Minor Gridlines
   *
   * @return PHPExceller_Chart_GridLines
   */
  public function getMinorGridlines() {
    if($this->minorGridlines !== NULL){
      return $this->minorGridlines;
    }

    return new PHPExceller_Chart_GridLines();
  }


    /**
     * Set the Top Left position for the chart
     *
     * @param    string    $cell
     * @param    integer    $xOffset
     * @param    integer    $yOffset
     * @return void
     */
    public function setTopLeftPosition($cell, $xOffset=null, $yOffset=null) {
        $this->topLeftCellRef = $cell;
        if (!is_null($xOffset))
            $this->setTopLeftXOffset($xOffset);
        if (!is_null($yOffset))
            $this->setTopLeftYOffset($yOffset);
    }

    /**
     * Get the top left position of the chart
     *
     * @return array    an associative array containing the cell address, X-Offset and Y-Offset from the top left of that cell
     */
    public function getTopLeftPosition() {
        return array( 'cell'    => $this->topLeftCellRef,
                      'xOffset'    => $this->topLeftXOffset,
                      'yOffset'    => $this->topLeftYOffset
                    );
    }

    /**
     * Get the cell address where the top left of the chart is fixed
     *
     * @return string
     */
    public function getTopLeftCell() {
        return $this->topLeftCellRef;
    }

    /**
     * Set the Top Left cell position for the chart
     *
     * @param    string    $cell
     * @return void
     */
    public function setTopLeftCell($cell) {
        $this->topLeftCellRef = $cell;
    }

    /**
     * Set the offset position within the Top Left cell for the chart
     *
     * @param    integer    $xOffset
     * @param    integer    $yOffset
     * @return void
     */
    public function setTopLeftOffset($xOffset=null,$yOffset=null) {
        if (!is_null($xOffset))
            $this->setTopLeftXOffset($xOffset);
        if (!is_null($yOffset))
            $this->setTopLeftYOffset($yOffset);
    }

    /**
     * Get the offset position within the Top Left cell for the chart
     *
     * @return integer[]
     */
    public function getTopLeftOffset() {
        return array( 'X' => $this->topLeftXOffset,
                      'Y' => $this->topLeftYOffset
                    );
    }

    public function setTopLeftXOffset($xOffset) {
        $this->topLeftXOffset = $xOffset;
    }

    public function getTopLeftXOffset() {
        return $this->topLeftXOffset;
    }

    public function setTopLeftYOffset($yOffset) {
        $this->topLeftYOffset = $yOffset;
    }

    public function getTopLeftYOffset() {
        return $this->topLeftYOffset;
    }

    /**
     * Set the Bottom Right position of the chart
     *
     * @param    string    $cell
     * @param    integer    $xOffset
     * @param    integer    $yOffset
     * @return void
     */
    public function setBottomRightPosition($cell, $xOffset=null, $yOffset=null) {
        $this->bottomRightCellRef = $cell;
        if (!is_null($xOffset))
            $this->setBottomRightXOffset($xOffset);
        if (!is_null($yOffset))
            $this->setBottomRightYOffset($yOffset);
    }

    /**
     * Get the bottom right position of the chart
     *
     * @return array    an associative array containing the cell address, X-Offset and Y-Offset from the top left of that cell
     */
    public function getBottomRightPosition() {
        return array( 'cell'    => $this->bottomRightCellRef,
                      'xOffset'    => $this->bottomRightXOffset,
                      'yOffset'    => $this->bottomRightYOffset
                    );
    }

    public function setBottomRightCell($cell) {
        $this->bottomRightCellRef = $cell;
    }

    /**
     * Get the cell address where the bottom right of the chart is fixed
     *
     * @return string
     */
    public function getBottomRightCell() {
        return $this->bottomRightCellRef;
    }

    /**
     * Set the offset position within the Bottom Right cell for the chart
     *
     * @param    integer    $xOffset
     * @param    integer    $yOffset
     * @return void
     */
    public function setBottomRightOffset($xOffset=null,$yOffset=null) {
        if (!is_null($xOffset))
            $this->setBottomRightXOffset($xOffset);
        if (!is_null($yOffset))
            $this->setBottomRightYOffset($yOffset);
    }

    /**
     * Get the offset position within the Bottom Right cell for the chart
     *
     * @return integer[]
     */
    public function getBottomRightOffset() {
        return array( 'X' => $this->bottomRightXOffset,
                      'Y' => $this->bottomRightYOffset
                    );
    }

    public function setBottomRightXOffset($xOffset) {
        $this->bottomRightXOffset = $xOffset;
    }

    public function getBottomRightXOffset() {
        return $this->bottomRightXOffset;
    }

    public function setBottomRightYOffset($yOffset) {
        $this->bottomRightYOffset = $yOffset;
    }

    public function getBottomRightYOffset() {
        return $this->bottomRightYOffset;
    }


    public function refresh() {
        if ($this->worksheet !== NULL) {
            $this->plotArea->refresh($this->worksheet);
        }
    }

    public function render($outputDestination = null) {
        $libraryName = PHPExceller_Settings::getChartRendererName();
        if (is_null($libraryName)) {
            return false;
        }
        //    Ensure that data series values are up-to-date before we render
        $this->refresh();

        $libraryPath = PHPExceller_Settings::getChartRendererPath();
        $includePath = str_replace('\\','/',get_include_path());
        $rendererPath = str_replace('\\','/',$libraryPath);
        if (strpos($rendererPath,$includePath) === false) {
            set_include_path(get_include_path() . PATH_SEPARATOR . $libraryPath);
        }

        $rendererName = 'PHPExceller_Chart_Renderer_'.$libraryName;
        $renderer = new $rendererName($this);

        if ($outputDestination == 'php://output') {
            $outputDestination = null;
        }
        return $renderer->render($outputDestination);
    }
    
    public function updateWorkbookName($workbookname)
    {
        if ($this->plotArea)
        {
            $this->plotArea->updateWorkbookName($workbookname);
        }
    }

    /**
     * Set the Top Left position for the chart
     *
     * @param    string    $cell
     * @param    integer    $xOffset
     * @param    integer    $yOffset
     * @return void
     */
    public function setTopLeftPosition($cell, $xOffset = null, $yOffset = null)
    {
        $this->topLeftCellRef = $cell;
        if (!is_null($xOffset)) {
            $this->setTopLeftXOffset($xOffset);
        }
        if (!is_null($yOffset)) {
            $this->setTopLeftYOffset($yOffset);
        }
    }

    /**
     * Get the top left position of the chart
     *
     * @return array    an associative array containing the cell address, X-Offset and Y-Offset from the top left of that cell
     */
    public function getTopLeftPosition()
    {
        return array(
            'cell'    => $this->topLeftCellRef,
            'xOffset' => $this->topLeftXOffset,
            'yOffset' => $this->topLeftYOffset
        );
    }

    /**
     * Get the cell address where the top left of the chart is fixed
     *
     * @return string
     */
    public function getTopLeftCell()
    {
        return $this->topLeftCellRef;
    }

    /**
     * Set the Top Left cell position for the chart
     *
     * @param    string    $cell
     * @return void
     */
    public function setTopLeftCell($cell)
    {
        $this->topLeftCellRef = $cell;
    }

    /**
     * Set the offset position within the Top Left cell for the chart
     *
     * @param    integer    $xOffset
     * @param    integer    $yOffset
     * @return void
     */
    public function setTopLeftOffset($xOffset = null, $yOffset = null)
    {
        if (!is_null($xOffset)) {
            $this->setTopLeftXOffset($xOffset);
        }
        if (!is_null($yOffset)) {
            $this->setTopLeftYOffset($yOffset);
        }
    }

    /**
     * Get the offset position within the Top Left cell for the chart
     *
     * @return integer[]
     */
    public function getTopLeftOffset()
    {
        return array(
            'X' => $this->topLeftXOffset,
            'Y' => $this->topLeftYOffset
        );
    }

    public function setTopLeftXOffset($xOffset)
    {
        $this->topLeftXOffset = $xOffset;
    }

    public function getTopLeftXOffset()
    {
        return $this->topLeftXOffset;
    }

    public function setTopLeftYOffset($yOffset)
    {
        $this->topLeftYOffset = $yOffset;
    }

    public function getTopLeftYOffset()
    {
        return $this->topLeftYOffset;
    }

    /**
     * Set the Bottom Right position of the chart
     *
     * @param    string    $cell
     * @param    integer    $xOffset
     * @param    integer    $yOffset
     * @return void
     */
    public function setBottomRightPosition($cell, $xOffset = null, $yOffset = null)
    {
        $this->bottomRightCellRef = $cell;
        if (!is_null($xOffset)) {
            $this->setBottomRightXOffset($xOffset);
        }
        if (!is_null($yOffset)) {
            $this->setBottomRightYOffset($yOffset);
        }
    }

    /**
     * Get the bottom right position of the chart
     *
     * @return array    an associative array containing the cell address, X-Offset and Y-Offset from the top left of that cell
     */
    public function getBottomRightPosition()
    {
        return array(
            'cell'    => $this->bottomRightCellRef,
            'xOffset' => $this->bottomRightXOffset,
            'yOffset' => $this->bottomRightYOffset
        );
    }

    public function setBottomRightCell($cell)
    {
        $this->bottomRightCellRef = $cell;
    }

    /**
     * Get the cell address where the bottom right of the chart is fixed
     *
     * @return string
     */
    public function getBottomRightCell()
    {
        return $this->bottomRightCellRef;
    }

    /**
     * Set the offset position within the Bottom Right cell for the chart
     *
     * @param    integer    $xOffset
     * @param    integer    $yOffset
     * @return void
     */
    public function setBottomRightOffset($xOffset = null, $yOffset = null)
    {
        if (!is_null($xOffset)) {
            $this->setBottomRightXOffset($xOffset);
        }
        if (!is_null($yOffset)) {
            $this->setBottomRightYOffset($yOffset);
        }
    }

    /**
     * Get the offset position within the Bottom Right cell for the chart
     *
     * @return integer[]
     */
    public function getBottomRightOffset()
    {
        return array(
            'X' => $this->bottomRightXOffset,
            'Y' => $this->bottomRightYOffset
        );
    }

    public function setBottomRightXOffset($xOffset)
    {
        $this->bottomRightXOffset = $xOffset;
    }

    public function getBottomRightXOffset()
    {
        return $this->bottomRightXOffset;
    }

    public function setBottomRightYOffset($yOffset)
    {
        $this->bottomRightYOffset = $yOffset;
    }

    public function getBottomRightYOffset()
    {
        return $this->bottomRightYOffset;
    }


    public function refresh()
    {
        if ($this->worksheet !== null) {
            $this->plotArea->refresh($this->worksheet);
        }
    }

    public function render($outputDestination = null)
    {
        $libraryName = PHPExceller_Settings::getChartRendererName();
        if (is_null($libraryName)) {
            return false;
        }
        //    Ensure that data series values are up-to-date before we render
        $this->refresh();

        $libraryPath = PHPExceller_Settings::getChartRendererPath();
        $includePath = str_replace('\\', '/', get_include_path());
        $rendererPath = str_replace('\\', '/', $libraryPath);
        if (strpos($rendererPath, $includePath) === false) {
            set_include_path(get_include_path() . PATH_SEPARATOR . $libraryPath);
        }

        $rendererName = 'PHPExceller_Chart_Renderer_'.$libraryName;
        $renderer = new $rendererName($this);

        if ($outputDestination == 'php://output') {
            $outputDestination = null;
        }
        return $renderer->render($outputDestination);
    }
}

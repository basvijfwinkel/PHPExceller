<?php
namespace PHPExceller\Chart;

use PHPExceller\Worksheet;
use PHPExceller\Chart\DataSeriesValues;

/**
 * Based on PHPExcel_Chart_DataSeries
 */

class DataSeries
{
    const TYPE_BARCHART        = 'barChart';
    const TYPE_BARCHART_3D     = 'bar3DChart';
    const TYPE_LINECHART       = 'lineChart';
    const TYPE_LINECHART_3D    = 'line3DChart';
    const TYPE_AREACHART       = 'areaChart';
    const TYPE_AREACHART_3D    = 'area3DChart';
    const TYPE_PIECHART        = 'pieChart';
    const TYPE_PIECHART_3D     = 'pie3DChart';
    const TYPE_DOUGHTNUTCHART  = 'doughnutChart';
    const TYPE_DONUTCHART      = self::TYPE_DOUGHTNUTCHART;    //    Synonym
    const TYPE_SCATTERCHART    = 'scatterChart';
    const TYPE_SURFACECHART    = 'surfaceChart';
    const TYPE_SURFACECHART_3D = 'surface3DChart';
    const TYPE_RADARCHART      = 'radarChart';
    const TYPE_BUBBLECHART     = 'bubbleChart';
    const TYPE_STOCKCHART      = 'stockChart';
    const TYPE_CANDLECHART     = self::TYPE_STOCKCHART;       //    Synonym

    const GROUPING_CLUSTERED       = 'clustered';
    const GROUPING_STACKED         = 'stacked';
    const GROUPING_PERCENT_STACKED = 'percentStacked';
    const GROUPING_STANDARD        = 'standard';

    const DIRECTION_BAR        = 'bar';
    const DIRECTION_HORIZONTAL = self::DIRECTION_BAR;
    const DIRECTION_COL        = 'col';
    const DIRECTION_COLUMN     = self::DIRECTION_COL;
    const DIRECTION_VERTICAL   = self::DIRECTION_COL;

    const STYLE_LINEMARKER   = 'lineMarker';
    const STYLE_SMOOTHMARKER = 'smoothMarker';
    const STYLE_MARKER       = 'marker';
    const STYLE_FILLED       = 'filled';

    /**
     * Series Plot Type
     *
     * @var string
     */
    private $plotType = null;

    /**
     * Plot Grouping Type
     *
     * @var boolean
     */
    private $plotGrouping = null;

    /**
     * Plot Direction
     *
     * @var boolean
     */
    private $plotDirection = null;

    /**
     * Plot Style
     *
     * @var string
     */
    private $plotStyle = null;

    /**
     * Order of plots in Series
     *
     * @var array of integer
     */
    private $plotOrder = array();

    /**
     * Plot Label
     *
     * @var array of DataSeriesValues
     */
    private $plotLabel = array();

    /**
     * Plot Category
     *
     * @var array of DataSeriesValues
     */
    private $plotCategory = array();

    /**
     * Smooth Line
     *
     * @var string
     */
    private $smoothLine = null;

    /**
     * Plot Values
     *
     * @var array of DataSeriesValues
     */
    private $plotValues = array();

    /**
     * Plot Bubble Sizes
     *
     * @var array of DataSeriesValues
     */
    private $plotBubbleChartSizes = array();

    /**
     * Plot Bubble Labels
     *
     * @var array of DataSeriesValues
     */
    private $plotBubbleChartLabels = array();

    private $plotBubbleScale = 100;
    /**
     * Create a new DataSeries
     */
    public function __construct($plotType = null,
                                    $plotGrouping = null,
                                    $plotOrder = array(),
                                    $plotLabel = array(),
                                    $plotCategory = array(),
                                    $plotValues = array(),
                                    $plotDirection = null,
                                    $smoothLine = null,
                                    $plotStyle = null,
                                    $plotBubbleChartSizes = array(),
                                    $plotBubbleChartLabels = array(),
                                    $plotBubbleScale = 100
                                   )
    {
        $this->plotType = $plotType;
        $this->plotGrouping = $plotGrouping;
        $this->plotOrder = $plotOrder;
        $keys = array_keys($plotValues);
        $this->plotValues = $plotValues;
        if ((count($plotLabel) == 0) || (is_null($plotLabel[$keys[0]]))) {
            $plotLabel[$keys[0]] = new DataSeriesValues();
        }

        $this->plotLabel = $plotLabel;
        if ((count($plotCategory) == 0) || (is_null($plotCategory[$keys[0]]))) {
            $plotCategory[$keys[0]] = new DataSeriesValues();
        }
        $this->plotCategory = $plotCategory;
        $this->smoothLine = $smoothLine;
        $this->plotStyle = $plotStyle;

        if (is_null($plotDirection)) {
            $plotDirection = self::DIRECTION_COL;
        }
        $this->plotDirection = $plotDirection;
        $this->plotBubbleChartSizes = $plotBubbleChartSizes;
        $this->plotBubbleChartLabels = $plotBubbleChartLabels;
        $this->plotBubbleScale = $plotBubbleScale;
    }

    /**
     * Get Plot Type
     *
     * @return string
     */
    public function getPlotType() {
        return $this->plotType;
    }

    /**
     * Set Plot Type
     *
     * @param string $plotType
     * @return void
     */
    public function setPlotType($plotType = '') {
        $this->plotType = $plotType;
    }

    /**
     * Get Bubble Scale
     *
     * @return string
     */
    public function getBubbleScale() {
        return $this->plotBubbleScale;
    }

    /**
     * Set BubbleScale
     *
     * @param string $plotBubbleScale
     * @return void
     */
    public function setBubbleScale($plotBubbleScale = 100) {
        $this->plotBubbleScale = $plotBubbleScale;
    }


    /**
     * Get Plot Grouping Type
     *
     * @return string
     */
    public function getPlotGrouping() {
        return $this->plotGrouping;
    }

    /**
     * Set Plot Grouping Type
     *
     * @param string $groupingType
     * @return DataSeries
     */
    public function setPlotGrouping($groupingType = null) {
        $this->plotGrouping = $groupingType;
    }

    /**
     * Get Plot Direction
     *
     * @return string
     */
    public function getPlotDirection()
    {
        return $this->plotDirection;
    }

    /**
     * Set Plot Direction
     *
     * @param string $plotDirection
     * @return void
     */
    public function setPlotDirection($plotDirection = null)
    {
        $this->plotDirection = $plotDirection;
    }

    /**
     * Get Plot Style
     *
     * @return string
     */
    public function getPlotStyle()
    {
        return $this->plotStyle;
    }

    /**
     * Get Plot Order
     *
     * @return string
     */
    public function getPlotOrder() {
        return $this->plotOrder;
    }

    /**
     * Get Plot Labels
     *
     * @return array of DataSeriesValues
     */
    public function getPlotLabels() {
        return $this->plotLabel;
    }

    /**
     * Get Plot Label by Index
     *
     * @return DataSeriesValues
     */
    public function getPlotLabelByIndex($index) {
        $keys = array_keys($this->plotLabel);
        if (in_array($index,$keys)) {
            return $this->plotLabel[$index];
        } elseif(isset($keys[$index])) {
            return $this->plotLabel[$keys[$index]];
        }
        return false;
    }

    /**
     * Get Plot Categories
     *
     * @return array of DataSeriesValues
     */
    public function getPlotCategories() {
        return $this->plotCategory;
    }

    /**
     * Get Plot Category by Index
     *
     * @return DataSeriesValues
     */
    public function getPlotCategoryByIndex($index) {
        $keys = array_keys($this->plotCategory);
        if (in_array($index,$keys)) {
            return $this->plotCategory[$index];
        } elseif(isset($keys[$index])) {
            return $this->plotCategory[$keys[$index]];
        }
        return false;
    }

    /**
     * Set Plot Style
     *
     * @param string $plotStyle
     * @return void
     */
    public function setPlotStyle($plotStyle = null) {
        $this->plotStyle = $plotStyle;
    }

    /**
     * Get Plot Values
     *
     * @return array of DataSeriesValues
     */
    public function getPlotValues() {
        return $this->plotValues;
    }

    /**
     * Get Plot Bubble Chart Sizes
     *
     * @return array of DataSeriesValues
     */
    public function getBubbleChartSizes() {
        return $this->plotBubbleChartSizes;
    }

    /**
     * Get Plot Bubble Chart Labels
     *
     * @return array of DataSeriesValues
     */
    public function getBubbleChartLabels() {
        return $this->plotBubbleChartLabels;
    }

    /**
     * Get Plot Values by Index
     *
     * @return DataSeriesValues
     */
    public function getPlotValuesByIndex($index) {
        $keys = array_keys($this->plotValues);
        if (in_array($index,$keys)) {
            return $this->plotValues[$index];
        } elseif(isset($keys[$index])) {
            return $this->plotValues[$keys[$index]];
        }
        return false;
    }

    /**
     * Get Plot Bubble Chart Sizes by Index
     *
     * @return DataSeriesValues
     */
    public function getBubbleChartSizesByIndex($index) {
        $keys = array_keys($this->plotBubbleChartSizes);
        if (in_array($index,$keys)) {
            return $this->plotBubbleChartSizes[$index];
        } elseif(isset($keys[$index])) {
            return $this->plotBubbleChartSizes[$keys[$index]];
        }
        return false;
    }

    /**
     * Get Plot Bubble Chart Labels by Index
     *
     * @return DataSeriesValues
     */
    public function getBubbleChartLabelsByIndex($index) {
        $keys = array_keys($this->plotBubbleChartLabels);
        if (in_array($index,$keys)) {
            return $this->plotBubbleChartLabels[$index];
        } elseif(isset($keys[$index])) {
            return $this->plotBubbleChartLabels[$keys[$index]];
        }
        return false;
    }

    /**
     * Get Number of Plot Series
     *
     * @return integer
     */
    public function getPlotSeriesCount() {
        return count($this->plotValues);
    }

    /**
     * Get Number of Plot Bubble Chart Sizes
     *
     * @return integer
     */
    public function getPlotBubbleChartSizesSeriesCount() {
        return count($this->plotBubbleChartSizes);
    }

    /**
     * Get Number of Plot Bubble Chart Labels
     *
     * @return integer
     */
    public function getPlotBubbleChartLabelsSeriesCount() {
        return count($this->plotBubbleChartLabels);
    }

    /**
     * Get Smooth Line
     *
     * @return boolean
     */
    public function getSmoothLine() {
        return $this->smoothLine;
    }

    /**
     * Set Smooth Line
     *
     * @param boolean $smoothLine
     * @return void
     */
    public function setSmoothLine($smoothLine = TRUE) {
        $this->smoothLine = $smoothLine;
    }

    public function refresh(Worksheet $worksheet) {
        foreach($this->plotValues as $plotValues) {
            if ($plotValues !== NULL)
                $plotValues->refresh($worksheet, TRUE);
        }
        foreach($this->plotBubbleChartSizes as $plotBubbleChartSizes) {
            if ($plotBubbleChartSizes !== NULL)
                $plotBubbleChartSizes->refresh($worksheet, TRUE);
        }
        foreach($this->plotBubbleChartLabels as $plotBubbleChartLabels) {
            if ($plotBubbleChartLabels !== NULL)
                $plotBubbleChartLabels->refresh($worksheet, TRUE);
        }
        foreach($this->plotLabel as $plotValues) {
            if ($plotValues !== NULL)
                $plotValues->refresh($worksheet, TRUE);
        }
        foreach($this->plotCategory as $plotValues) {
            if ($plotValues !== NULL)
                $plotValues->refresh($worksheet, FALSE);
        }
    }

    public function updateWorkbookName($workbookname)
    {
        // update plotlabels
        foreach($this->plotLabel as $plotlabel)
        {
            $plotlabel->updateWorkbookName($workbookname);
        }
        // update plotvalues
        foreach($this->plotValues as $plotvalues)
        {
            $plotvalues->updateWorkbookName($workbookname);
        }
        // update bubble chart sizes
        foreach($this->plotBubbleChartSizes as $plotbubblechartsizes)
        {
            $plotbubblechartsizes->updateWorkbookName($workbookname);
        }
        // update bubble chart labels
        foreach($this->plotBubbleChartLabels as $plotbubblechartlabels)
        {
            $plotbubblechartlabels->updateWorkbookName($workbookname);
        }
        //update plotcategory
        foreach($this->plotCategory as $plotcategory)
        {
            $plotcategory->updateWorkbookName($workbookname);
        }
    }
}

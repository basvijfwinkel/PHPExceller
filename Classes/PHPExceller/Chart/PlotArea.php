<?php
namespace PHPExceller\Chart;

use PHPExceller\Chart\Layout;
use PHPExceller\Worksheet;

/**
 * Based on PHPExcel_Chart_PlotArea
 */

class PlotArea
{
    /**
     * PlotArea Layout
     *
     * @var Layout
     */
    private $layout = null;

    /**
     * Plot Series
     *
     * @var array of DataSeries
     */
    private $plotSeries = array();

    /**
     * Create a new PlotArea
     */
    public function __construct(Layout $layout = null, $plotSeries = array())
    {
        this->layout = $layout;
        this->plotSeries = $plotSeries;
    }

    /**
     * Get Layout
     *
     * @return Layout
     */
    public function getLayout() {
        return this->layout;
    }

    /**
     * Get Number of Plot Groups
     *
     * @return array of DataSeries
     */
    public function getPlotGroupCount() {
        return count(this->plotSeries);
    }

    /**
     * Get Number of Plot Series
     *
     * @return integer
     */
    public function getPlotSeriesCount() {
        $seriesCount = 0;
        foreach(this->plotSeries as $plot) {
            $seriesCount += $plot->getPlotSeriesCount();
        }
        return $seriesCount;
    }

    /**
     * Get Plot Series
     *
     * @return array of DataSeries
     */
    public function getPlotGroup() {
        return this->plotSeries;
    }

    /**
     * Get Plot Series by Index
     *
     * @return DataSeries
     */
    public function getPlotGroupByIndex($index) {
        return this->plotSeries[$index];
    }

    /**
     * Set Plot Series
     *
     * @param [DataSeries]
     * @return void
     */
    public function setPlotSeries($plotSeries = array()) {
        this->plotSeries = $plotSeries;
    }

    public function refresh(Worksheet $worksheet) {
        foreach(this->plotSeries as $plotSeries) {
            $plotSeries->refresh($worksheet);
        }
    }

    public function updateWorkbookName($workbookname)
    {
        if (this->plotSeries)
        {
            foreach(this->plotSeries as $dataSerie)
            {
                $dataSerie->updateWorkbookName($workbookname);
            }
        }
    }
}

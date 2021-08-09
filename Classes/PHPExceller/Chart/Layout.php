<?php
namespace PHPExceller\Chart;

/**
 * Based on PHPExcel_Chart_Layout
 */
class Layout
{
        // data label positions
        const LABEL_POS_CENTER = 'ctr';
        const LABEL_POS_TOP    = 't';
        const LABEL_POS_BOTTOM = 'b';
        const LABEL_POS_LEFT   = 'l';
        const LABEL_POS_RIGHT  = 'r';

    /**
     * layoutTarget
     *
     * @var string
     */
    private $layoutTarget = NULL;

    /**
     * X Mode
     *
     * @var string
     */
    private $xMode        = NULL;

    /**
     * Y Mode
     *
     * @var string
     */
    private $yMode        = NULL;

    /**
     * X-Position
     *
     * @var float
     */
    private $xPos        = NULL;

    /**
     * Y-Position
     *
     * @var float
     */
    private $yPos        = NULL;

    /**
     * width
     *
     * @var float
     */
    private $width        = NULL;

    /**
     * height
     *
     * @var float
     */
    private $height    = NULL;

    /**
     * show legend key
     * Specifies that legend keys should be shown in data labels
     *
     * @var boolean
     */
    private $showLegendKey    = NULL;

    /**
     * show value
     * Specifies that the value should be shown in a data label.
     *
     * @var boolean
     */
    private $showVal    = NULL;

    /**
     * show category name
     * Specifies that the category name should be shown in the data label.
     *
     * @var boolean
     */
    private $showCatName    = NULL;

    /**
     * show data series name
     * Specifies that the series name should be shown in the data label.
     *
     * @var boolean
     */
    private $showSerName    = NULL;

    /**
     * show percentage
     * Specifies that the percentage should be shown in the data label.
     *
     * @var boolean
     */
    private $showPercent    = NULL;

    /**
     * show bubble size
     *
     * @var boolean
     */
    private $showBubbleSize    = NULL;

    /**
     * show leader lines
     * Specifies that leader lines should be shown for the data label.
     *
     * @var boolean
     */
    private $showLeaderLines    = NULL;

    /**
     * data label position
     * Specifies the position for the label
     *
     * @var string ctr/t/b/l/r
     */
    private $dataLabelPosition = self::LABEL_POS_CENTER;

        /**
         * data label color
         * default : black
         * @var string hex RGB color code
         */
        private $dataLabelColor = "000000";

    /**
     * Create a new Layout
     */
    public function __construct($layout=array())
    {
        if (isset($layout['layoutTarget'])) { this->layoutTarget    = $layout['layoutTarget'];    }
        if (isset($layout['xMode']))        { this->xMode            = $layout['xMode'];          }
        if (isset($layout['yMode']))        { this->yMode            = $layout['yMode'];          }
        if (isset($layout['x']))            { this->xPos            = (float) $layout['x'];       }
        if (isset($layout['y']))            { this->yPos            = (float) $layout['y'];       }
        if (isset($layout['w']))            { this->width            = (float) $layout['w'];      }
        if (isset($layout['h']))            { this->height        = (float) $layout['h'];         }
    }

    /**
     * Get Layout Target
     *
     * @return string
     */
    public function getLayoutTarget() {
        return this->layoutTarget;
    }

    /**
     * Set Layout Target
     *
     * @param Layout Target $value
     * @return void
     */
    public function setLayoutTarget($value) {
        this->layoutTarget = $value;
    }

    /**
     * Get X-Mode
     *
     * @return string
     */
    public function getXMode() {
        return this->xMode;
    }

    /**
     * Set X-Mode
     *
     * @param X-Mode $value
     * @return void
     */
    public function setXMode($value) {
        this->xMode = $value;
    }

    /**
     * Get Y-Mode
     *
     * @return string
     */
    public function getYMode() {
        return this->yMode;
    }

    /**
     * Set Y-Mode
     *
     * @param Y-Mode $value
     * @return void
     */
    public function setYMode($value) {
        this->yMode = $value;
    }

    /**
     * Get X-Position
     *
     * @return number
     */
    public function getXPosition() {
        return this->xPos;
    }

    /**
     * Set X-Position
     *
     * @param X-Position $value
     * @return void
     */
    public function setXPosition($value) {
        this->xPos = $value;
    }

    /**
     * Get Y-Position
     *
     * @return number
     */
    public function getYPosition() {
        return this->yPos;
    }

    /**
     * Set Y-Position
     *
     * @param Y-Position $value
     * @return void
     */
    public function setYPosition($value) {
        this->yPos = $value;
    }

    /**
     * Get Width
     *
     * @return number
     */
    public function getWidth() {
        return this->width;
    }

    /**
     * Set Width
     *
     * @param Width $value
     * @return void
     */
    public function setWidth($value) {
        this->width = $value;
    }

    /**
     * Get Height
     *
     * @return number
     */
    public function getHeight() {
        return this->height;
    }

    /**
     * Set Height
     *
     * @param Height $value
     * @return void
     */
    public function setHeight($value) {
        this->height = $value;
    }


    /**
     * Get show legend key
     *
     * @return boolean
     */
    public function getShowLegendKey() {
        return this->showLegendKey;
    }

    /**
     * Set show legend key
     * Specifies that legend keys should be shown in data labels.
     *
     * @param boolean $value        Show legend key
     * @return void
     */
    public function setShowLegendKey($value) {
        this->showLegendKey = $value;
    }

    /**
     * Get show value
     *
     * @return boolean
     */
    public function getShowVal() {
        return this->showVal;
    }

    /**
     * Set show val
     * Specifies that the value should be shown in data labels.
     *
     * @param boolean $value        Show val
     * @return void
     */
    public function setShowVal($value) {
        this->showVal = $value;
    }

    /**
     * Get show category name
     *
     * @return boolean
     */
    public function getShowCatName() {
        return this->showCatName;
    }

    /**
     * Set show cat name
     * Specifies that the category name should be shown in data labels.
     *
     * @param boolean $value        Show cat name
     * @return void
     */
    public function setShowCatName($value) {
        this->showCatName = $value;
    }

    /**
     * Get show data series name
     *
     * @return boolean
     */
    public function getShowSerName() {
        return this->showSerName;
    }

    /**
     * Set show ser name
     * Specifies that the series name should be shown in data labels.
     *
     * @param boolean $value        Show series name
     * @return void
     */
    public function setShowSerName($value) {
        this->showSerName = $value;
    }

    /**
     * Get show percentage
     *
     * @return boolean
     */
    public function getShowPercent() {
        return this->showPercent;
    }

    /**
     * Set show percentage
     * Specifies that the percentage should be shown in data labels.
     *
     * @param boolean $value        Show percentage
     * @return void
     */
    public function setShowPercent($value) {
        this->showPercent = $value;
    }

    /**
     * Get show bubble size
     *
     * @return boolean
     */
    public function getShowBubbleSize() {
        return this->showBubbleSize;
    }

    /**
     * Set show bubble size
     * Specifies that the bubble size should be shown in data labels.
     *
     * @param boolean $value        Show bubble size
     * @return void
     */
    public function setShowBubbleSize($value) {
        this->showBubbleSize = $value;
    }

    /**
     * Get show leader lines
     *
     * @return boolean
     */
    public function getShowLeaderLines() {
        return this->showLeaderLines;
    }

    /**
     * Set show leader lines
     * Specifies that leader lines should be shown in data labels.
     *
     * @param boolean $value        Show leader lines
     * @return void
     */
    public function setShowLeaderLines($value) {
        this->showLeaderLines = $value;
    }

    /**
     * Get the data label position
     *
     * @return string / const
     */
     public function getDataLabelPosition() {
        return this->dataLabelPosition;
     }

     /**
      * Set the data label position
      *  Specifies that the data label position
      *
      * @param string / Layout::LABEL_POS_CENTER LABEL_POS_TOP LABEL_POS_BOTTOM LABEL_POS_RIGHT LABEL_POS_LEFT
      * @return void
     */
     public function setDataLabelPosition($value) {
         this->dataLabelPosition = $value;
    }

    public function getDataLabelColor() 
    {
        return this->dataLabelColor;
    }

    public function setDataLabelColor($value) 
    {
         this->dataLabelColor = $value;
    }
}

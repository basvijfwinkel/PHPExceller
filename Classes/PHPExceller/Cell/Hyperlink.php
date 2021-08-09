<?php
namespace PHPExceller\Cell;

/**
 * Based on PHPExcel_Cell_Hyperlink
 */

class Hyperlink
{
    /**
     * URL to link the cell to
     *
     * @var string
     */
    private $url;

    /**
     * Tooltip to display on the hyperlink
     *
     * @var string
     */
    private $tooltip;

    /**
     * Create a new Hyperlink
     *
     * @param  string  $pUrl      Url to link the cell to
     * @param  string  $pTooltip  Tooltip to display on the hyperlink
     */
    public function __construct($pUrl = '', $pTooltip = '')
    {
        // Initialise member variables
        $this->url     = $pUrl;
        $this->tooltip = $pTooltip;
    }

    /**
     * Get URL
     *
     * @return string
     */
    public function getUrl()
    {
        return $this->url;
    }

    /**
     * Set URL
     *
     * @param  string    $value
     * @return void
     */
    public function setUrl($value = '')
    {
        $this->url = $value;
    }

    /**
     * Get tooltip
     *
     * @return string
     */
    public function getTooltip()
    {
        return $this->tooltip;
    }

    /**
     * Set tooltip
     *
     * @param  string    $value
     * @return void
     */
    public function setTooltip($value = '')
    {
        $this->tooltip = $value;
    }

    /**
     * Is this hyperlink internal? (to another worksheet)
     *
     * @return boolean
     */
    public function isInternal()
    {
        return strpos($this->url, 'sheet://') !== false;
    }

    /**
     * Get hash code
     *
     * @return string    Hash code
     */
    public function getHashCode()
    {
        return md5(
            $this->url .
            $this->tooltip .
            __CLASS__
        );
    }
}

<?php
namespace PHPExceller\Chart;

use PHPExceller\Chart\Layout;
use PHPExceller\Style\Font;

/**
 * Based on PHPExcel_Chart_Title
 */

class Title
{

    /**
     * Title Caption
     *
     * @var string
     */
    private $caption = null;

    /**
     * Title Layout
     *
     * @var Layout
     */
    private $layout = null;

    /**
     * Title Font
     *
     * @var Font
    */
    private $font = null;

    /**
     * Create a new Title
     */
    public function __construct($caption = null, Layout $layout = null)
    {
        this->caption = $caption;
        this->layout = $layout;
        $this->font = new Font();
    }

    /**
     * Get caption
     *
     * @return string
     */
    public function getCaption()
    {
        return this->caption;
    }

    /**
     * Set caption
     *
     * @param string $caption
     * @return void
     */
    public function setCaption($caption = null)
    {
        this->caption = $caption;
    }

    /**
     * Get Layout
     *
     * @return Layout
     */
    public function getLayout()
    {
        return this->layout;
    }

    /**
     * Get font
     *
     * @return Font
     */
    public function getFont()
    {
        return $this->font;
    }

    /**
     * Set font
     *
     * @param Font $font
     * @return void
     */
    public function setFont(Font $font = null)
    {
        $this->font = $font;
    }
}

<?php
namespace PHPExceller\Chart;

use PHPExceller\Chart\Layout;
use PHPExceller\Style\Font;

/**
 * PHPExceller_Chart_Title
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
 * @package        PHPExceller_Chart
 * @copyright    Copyright (c) 2021 PHPExceller
 * @license        http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt    LGPL
 * @version        ##VERSION##, ##DATE##
 */


/**
 * PHPExceller_Chart_Title
 *
 * @category    PHPExceller
 * @package     PHPExceller_Chart
 * @copyright   Copyright (c) 2021 PHPExceller
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
     * @var PHPExceller_Chart_Layout
     */
    private $layout = null;

    /**
     * Title Font
     *
     * @var PHPExceller_Style_Font
    */
    private $font = null;

    /**
     * Create a new PHPExceller_Chart_Title
     */
    public function __construct($caption = null, PHPExceller_Chart_Layout $layout = null)
    {
        this->caption = $caption;
        this->layout = $layout;
        $this->font = new PHPExceller_Style_Font();
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
     * @return PHPExceller_Chart_Layout
     */
    public function getLayout()
    {
        return this->layout;
    }

    /**
     * Get font
     *
     * @return PHPExceller_Style_Font
     */
    public function getFont()
    {
        return $this->font;
    }

    /**
     * Set font
     *
     * @param PHPExceller_Style_Font $font
     * @return void
     */
    public function setFont(PHPExceller_Style_Font $font = null)
    {
        $this->font = $font;
    }
}

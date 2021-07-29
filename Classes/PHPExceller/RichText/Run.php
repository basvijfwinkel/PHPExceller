<?php
namespace PHPExceller\RichText;

use PHPExceller\RichText\TextElement;
use PHPExceller\RichText\ITextElement;
use PHPExceller\Style\Font;

/**
 * PHPExceller_RichText_Run
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
 * @package    PHPExceller_RichText
 * @copyright  Copyright (c) 2021 PHPExceller
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt    LGPL
 * @version    ##VERSION##, ##DATE##
 */
class Run extends TextElement implements ITextElement
{
    /**
     * Font
     *
     * @var PHPExceller_Style_Font
     */
    private $font;

    /**
     * Create a new PHPExceller_RichText_Run instance
     *
     * @param     string        $pText        Text
     */
    public function __construct($pText = '')
    {
        // Initialise variables
        $this->setText($pText);
        $this->font = new PHPExceller_Style_Font();
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
     * @param    PHPExceller_Style_Font        $pFont        Font
     * @throws     PHPExceller_Exception
     * @return PHPExceller_RichText_ITextElement
     */
    public function setFont(PHPExceller_Style_Font $pFont = null)
    {
        $this->font = $pFont;
        return $this;
    }

    /**
     * Get hash code
     *
     * @return string    Hash code
     */
    public function getHashCode()
    {
        return md5(
            $this->getText() .
            $this->font->getHashCode() .
            __CLASS__
        );
    }

    /**
     * Implement PHP __clone to create a deep clone, not just a shallow copy.
     */
    public function __clone()
    {
        $vars = get_object_vars($this);
        foreach ($vars as $key => $value) {
            if (is_object($value)) {
                $this->$key = clone $value;
            } else {
                $this->$key = $value;
            }
        }
    }
}

<?php
namespace PHPExceller;

use PHPExceller\PHPExceller_IComparable;
use PHPExceller\PHPExceller_Cell;
use PHPExceller\RichText\PHPExceller_RichText_Run;
use PHPExceller\Cell\PHPExceller_Cell_DataType;
use PHPExceller\RichText\PHPExceller_RichText_ITextElement;
use PHPExceller\RichText\PHPExceller_RichText_TextElement;
use PHPExceller\PHPExceller_Exception;

/**
 * PHPExceller_RichText
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
 * @category   PHPExceller
 * @package    PHPExceller_RichText
 * @copyright  Copyright (c) 2021
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt    LGPL
 * @version    ##VERSION##, ##DATE##
 */
class PHPExceller_RichText implements PHPExceller_IComparable
{
    /**
     * Rich text elements
     *
     * @var PHPExceller_RichText_ITextElement[]
     */
    private $richTextElements;

    /**
     * Create a new PHPExceller_RichText instance
     *
     * @param PHPExceller_Cell $pCell
     * @throws PHPExceller_Exception
     */
    public function __construct(PHPExceller_Cell $pCell = null)
    {
        // Initialise variables
        $this->richTextElements = array();

        // Rich-Text string attached to cell?
        if ($pCell !== null) {
            // Add cell text and style
            if ($pCell->getValue() != "") {
                $objRun = new PHPExceller_RichText_Run($pCell->getValue());
                $objRun->setFont(clone $pCell->getParent()->getStyle($pCell->getCoordinate())->getFont());
                $this->addText($objRun);
            }

            // Set parent value
            $pCell->setValueExplicit($this, PHPExceller_Cell_DataType::TYPE_STRING);
        }
    }

    /**
     * Add text
     *
     * @param PHPExceller_RichText_ITextElement $pText Rich text element
     * @throws PHPExceller_Exception
     * @return void
     */
    public function addText(PHPExceller_RichText_ITextElement $pText = null)
    {
        $this->richTextElements[] = $pText;
    }

    /**
     * Create text
     *
     * @param string $pText Text
     * @return PHPExceller_RichText_TextElement
     * @throws PHPExceller_Exception
     */
    public function createText($pText = '')
    {
        $objText = new PHPExceller_RichText_TextElement($pText);
        $this->addText($objText);
        return $objText;
    }

    /**
     * Create text run
     *
     * @param string $pText Text
     * @return PHPExceller_RichText_Run
     * @throws PHPExceller_Exception
     */
    public function createTextRun($pText = '')
    {
        $objText = new PHPExceller_RichText_Run($pText);
        $this->addText($objText);
        return $objText;
    }

    /**
     * Get plain text
     *
     * @return string
     */
    public function getPlainText()
    {
        // Return value
        $returnValue = '';

        // Loop through all PHPExceller_RichText_ITextElement
        foreach ($this->richTextElements as $text) {
            $returnValue .= $text->getText();
        }

        // Return
        return $returnValue;
    }

    /**
     * Convert to string
     *
     * @return string
     */
    public function __toString()
    {
        return $this->getPlainText();
    }

    /**
     * Get Rich Text elements
     *
     * @return PHPExceller_RichText_ITextElement[]
     */
    public function getRichTextElements()
    {
        return $this->richTextElements;
    }

    /**
     * Set Rich Text elements
     *
     * @param PHPExceller_RichText_ITextElement[] $pElements Array of elements
     * @throws PHPExceller_Exception
     * @return void
     */
    public function setRichTextElements($pElements = null)
    {
        if (is_array($pElements)) {
            $this->richTextElements = $pElements;
        } else {
            throw new PHPExceller_Exception("Invalid PHPExceller_RichText_ITextElement[] array passed.");
        }
    }

    /**
     * Get hash code
     *
     * @return string    Hash code
     */
    public function getHashCode()
    {
        $hashElements = '';
        foreach ($this->richTextElements as $element) {
            $hashElements .= $element->getHashCode();
        }

        return md5(
            $hashElements .
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

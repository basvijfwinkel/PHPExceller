<?php

/**
 * PHPExceller_Style_Border
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
 * @package    PHPExceller_Style
 * @copyright  Copyright (c) 2021 PHPExceller
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt    LGPL
 * @version    ##VERSION##, ##DATE##
 */
class PHPExceller_Style_Border extends PHPExceller_Style_Supervisor implements PHPExceller_IComparable
{
    /* Border style */
    const BORDER_NONE             = 'none';
    const BORDER_DASHDOT          = 'dashDot';
    const BORDER_DASHDOTDOT       = 'dashDotDot';
    const BORDER_DASHED           = 'dashed';
    const BORDER_DOTTED           = 'dotted';
    const BORDER_DOUBLE           = 'double';
    const BORDER_HAIR             = 'hair';
    const BORDER_MEDIUM           = 'medium';
    const BORDER_MEDIUMDASHDOT    = 'mediumDashDot';
    const BORDER_MEDIUMDASHDOTDOT = 'mediumDashDotDot';
    const BORDER_MEDIUMDASHED     = 'mediumDashed';
    const BORDER_SLANTDASHDOT     = 'slantDashDot';
    const BORDER_THICK            = 'thick';
    const BORDER_THIN             = 'thin';

    /**
     * Border style
     *
     * @var string
     */
    protected $borderStyle = PHPExceller_Style_Border::BORDER_NONE;

    /**
     * Border color
     *
     * @var PHPExceller_Style_Color
     */
    protected $color;

    /**
     * Parent property name
     *
     * @var string
     */
    protected $parentPropertyName;

    /**
     * Create a new PHPExceller_Style_Border
     *
     * @param    boolean    $isSupervisor    Flag indicating if this is a supervisor or not
     *                                    Leave this value at default unless you understand exactly what
     *                                        its ramifications are
     * @param    boolean    $isConditional    Flag indicating if this is a conditional style or not
     *                                    Leave this value at default unless you understand exactly what
     *                                        its ramifications are
     */
    public function __construct($isSupervisor = false, $isConditional = false)
    {
        // Supervisor?
        parent::__construct($isSupervisor);

        // Initialise values
        $this->color    = new PHPExceller_Style_Color(PHPExceller_Style_Color::COLOR_BLACK, $isSupervisor);

        // bind parent if we are a supervisor
        if ($isSupervisor) {
            $this->color->bindParent($this, 'color');
        }
    }

    /**
     * Bind parent. Only used for supervisor
     *
     * @param PHPExceller_Style_Borders $parent
     * @param string $parentPropertyName
     * @return void
     */
    public function bindParent($parent, $parentPropertyName = null)
    {
        $this->parent = $parent;
        $this->parentPropertyName = $parentPropertyName;
    }

    /**
     * Get the shared style component for the currently active cell in currently active sheet.
     * Only used for style supervisor
     *
     * @return PHPExceller_Style_Border
     * @throws PHPExceller_Exception
     */
    public function getSharedComponent()
    {
        switch ($this->parentPropertyName) {
            case 'allBorders':
            case 'horizontal':
            case 'inside':
            case 'outline':
            case 'vertical':
                throw new PHPExceller_Exception('Cannot get shared component for a pseudo-border.');
                break;
            case 'bottom':
                return $this->parent->getSharedComponent()->getBottom();
            case 'diagonal':
                return $this->parent->getSharedComponent()->getDiagonal();
            case 'left':
                return $this->parent->getSharedComponent()->getLeft();
            case 'right':
                return $this->parent->getSharedComponent()->getRight();
            case 'top':
                return $this->parent->getSharedComponent()->getTop();
        }
    }

    /**
     * Build style array from subcomponents
     *
     * @param array $array
     * @return array
     */
    public function getStyleArray($array)
    {
        switch ($this->parentPropertyName) {
            case 'allBorders':
            case 'bottom':
            case 'diagonal':
            case 'horizontal':
            case 'inside':
            case 'left':
            case 'outline':
            case 'right':
            case 'top':
            case 'vertical':
                $key = strtolower('vertical');
                break;
        }
        return $this->parent->getStyleArray(array($key => $array));
    }

    /**
     * Apply styles from array
     *
     * <code>
     * $objPHPExceller->getActiveSheet()->getStyle('B2')->getBorders()->getTop()->applyFromArray(
     *        array(
     *            'style' => PHPExceller_Style_Border::BORDER_DASHDOT,
     *            'color' => array(
     *                'rgb' => '808080'
     *            )
     *        )
     * );
     * </code>
     *
     * @param    array    $pStyles    Array containing style information
     * @throws    PHPExceller_Exception
     * @return void
     */
    public function applyFromArray($pStyles = null)
    {
        if (is_array($pStyles)) {
            if ($this->isSupervisor) {
                $this->getActiveSheet()->getStyle($this->getSelectedCells())->applyFromArray($this->getStyleArray($pStyles));
            } else {
                if (isset($pStyles['style'])) {
                    $this->setBorderStyle($pStyles['style']);
                }
                if (isset($pStyles['color'])) {
                    $this->getColor()->applyFromArray($pStyles['color']);
                }
            }
        } else {
            throw new PHPExceller_Exception("Invalid style array passed.");
        }
    }

    /**
     * Get Border style
     *
     * @return string
     */
    public function getBorderStyle()
    {
        if ($this->isSupervisor) {
            return $this->getSharedComponent()->getBorderStyle();
        }
        return $this->borderStyle;
    }

    /**
     * Set Border style
     *
     * @param string|boolean    $pValue
     *                            When passing a boolean, FALSE equates PHPExceller_Style_Border::BORDER_NONE
     *                                and TRUE to PHPExceller_Style_Border::BORDER_MEDIUM
     * @return void
     */
    public function setBorderStyle($pValue = PHPExceller_Style_Border::BORDER_NONE)
    {

        if (empty($pValue)) {
            $pValue = PHPExceller_Style_Border::BORDER_NONE;
        } elseif (is_bool($pValue) && $pValue) {
            $pValue = PHPExceller_Style_Border::BORDER_MEDIUM;
        }
        if ($this->isSupervisor) {
            $styleArray = $this->getStyleArray(array('style' => $pValue));
            $this->getActiveSheet()->getStyle($this->getSelectedCells())->applyFromArray($styleArray);
        } else {
            $this->borderStyle = $pValue;
        }
    }

    /**
     * Get Border Color
     *
     * @return PHPExceller_Style_Color
     */
    public function getColor()
    {
        return $this->color;
    }

    /**
     * Set Border Color
     *
     * @param    PHPExceller_Style_Color $pValue
     * @throws    PHPExceller_Exception
     * @return void
     */
    public function setColor(PHPExceller_Style_Color $pValue = null)
    {
        // make sure parameter is a real color and not a supervisor
        $color = $pValue->getIsSupervisor() ? $pValue->getSharedComponent() : $pValue;

        if ($this->isSupervisor) {
            $styleArray = $this->getColor()->getStyleArray(array('argb' => $color->getARGB()));
            $this->getActiveSheet()->getStyle($this->getSelectedCells())->applyFromArray($styleArray);
        } else {
            $this->color = $color;
        }
    }

    /**
     * Get hash code
     *
     * @return string    Hash code
     */
    public function getHashCode()
    {
        if ($this->isSupervisor) {
            return $this->getSharedComponent()->getHashCode();
        }
        return md5(
            $this->borderStyle .
            $this->color->getHashCode() .
            __CLASS__
        );
    }
}

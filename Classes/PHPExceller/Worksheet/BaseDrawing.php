<?php
namespace PHPExceller\Worksheet;

use PHPExceller\PHPExceller_IComparable;
use PHPExceller\Worksheet\Drawing\PHPExceller_Worksheet_Drawing_Shadow;
use PHPExceller\PHPExceller_Worksheet;

/**
 * PHPExceller_Worksheet_BaseDrawing
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
 * @package    PHPExceller_Worksheet
 * @copyright  Copyright (c) 2021 PHPExceller
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt    LGPL
 * @version    ##VERSION##, ##DATE##
 */
class PHPExceller_Worksheet_BaseDrawing implements PHPExceller_IComparable
{
    /**
     * Image counter
     *
     * @var int
     */
    private static $imageCounter = 0;

    /**
     * Image index
     *
     * @var int
     */
    private $imageIndex = 0;

    /**
     * Name
     *
     * @var string
     */
    protected $name;

    /**
     * Description
     *
     * @var string
     */
    protected $description;

    /**
     * Worksheet
     *
     * @var PHPExceller_Worksheet
     */
    protected $worksheet;

    /**
     * Coordinates
     *
     * @var string
     */
    protected $coordinates;

    /**
     * Offset X
     *
     * @var int
     */
    protected $offsetX;

    /**
     * Offset Y
     *
     * @var int
     */
    protected $offsetY;

    /**
     * Width
     *
     * @var int
     */
    protected $width;

    /**
     * Height
     *
     * @var int
     */
    protected $height;

    /**
     * Proportional resize
     *
     * @var boolean
     */
    protected $resizeProportional;

    /**
     * Rotation
     *
     * @var int
     */
    protected $rotation;

    /**
     * Shadow
     *
     * @var PHPExceller_Worksheet_Drawing_Shadow
     */
    protected $shadow;

    /**
     * Create a new PHPExceller_Worksheet_BaseDrawing
     */
    public function __construct()
    {
        // Initialise values
        $this->name                = '';
        $this->description        = '';
        $this->worksheet          = null;
        $this->coordinates        = 'A1';
        $this->offsetX            = 0;
        $this->offsetY            = 0;
        $this->width              = 0;
        $this->height             = 0;
        $this->resizeProportional = true;
        $this->rotation           = 0;
        $this->shadow             = new PHPExceller_Worksheet_Drawing_Shadow();

        // Set image index
        self::$imageCounter++;
        $this->imageIndex             = self::$imageCounter;
    }

    /**
     * Get image index
     *
     * @return int
     */
    public function getImageIndex()
    {
        return $this->imageIndex;
    }

    /**
     * Get Name
     *
     * @return string
     */
    public function getName()
    {
        return $this->name;
    }

    /**
     * Set Name
     *
     * @param string $pValue
     * @return void
     */
    public function setName($pValue = '')
    {
        $this->name = $pValue;
    }

    /**
     * Get Description
     *
     * @return string
     */
    public function getDescription()
    {
        return $this->description;
    }

    /**
     * Set Description
     *
     * @param string $pValue
     * @return void
     */
    public function setDescription($pValue = '')
    {
        $this->description = $pValue;
    }

    /**
     * Get Worksheet
     *
     * @return PHPExceller_Worksheet
     */
    public function getWorksheet()
    {
        return $this->worksheet;
    }

    /**
     * Set Worksheet
     *
     * @param     PHPExceller_Worksheet     $pValue
     * @param     bool                $pOverrideOld    If a Worksheet has already been assigned, overwrite it and remove image from old Worksheet?
     * @throws     PHPExceller_Exception
     * @return void
     */
    public function setWorksheet(PHPExceller_Worksheet $pValue = null, $pOverrideOld = false)
    {
        if (is_null($this->worksheet)) {
            // Add drawing to PHPExceller_Worksheet
            $this->worksheet = $pValue;
            $this->worksheet->getCell($this->coordinates);
            $this->worksheet->getDrawingCollection()->append($this);
        } else {
            if ($pOverrideOld) {
                // Remove drawing from old PHPExceller_Worksheet
                $iterator = $this->worksheet->getDrawingCollection()->getIterator();

                while ($iterator->valid()) {
                    if ($iterator->current()->getHashCode() == $this->getHashCode()) {
                        $this->worksheet->getDrawingCollection()->offsetUnset($iterator->key());
                        $this->worksheet = null;
                        break;
                    }
                }

                // Set new PHPExceller_Worksheet
                $this->setWorksheet($pValue);
            } else {
                throw new PHPExceller_Exception("A PHPExceller_Worksheet has already been assigned. Drawings can only exist on one PHPExceller_Worksheet.");
            }
        }
    }

    /**
     * Get Coordinates
     *
     * @return string
     */
    public function getCoordinates()
    {
        return $this->coordinates;
    }

    /**
     * Set Coordinates
     *
     * @param string $pValue
     * @return void
     */
    public function setCoordinates($pValue = 'A1')
    {
        $this->coordinates = $pValue;
    }

    /**
     * Get OffsetX
     *
     * @return int
     */
    public function getOffsetX()
    {
        return $this->offsetX;
    }

    /**
     * Set OffsetX
     *
     * @param int $pValue
     * @return void
     */
    public function setOffsetX($pValue = 0)
    {
        $this->offsetX = $pValue;
    }

    /**
     * Get OffsetY
     *
     * @return int
     */
    public function getOffsetY()
    {
        return $this->offsetY;
    }

    /**
     * Set OffsetY
     *
     * @param int $pValue
     * @return void
     */
    public function setOffsetY($pValue = 0)
    {
        $this->offsetY = $pValue;
    }

    /**
     * Get Width
     *
     * @return int
     */
    public function getWidth()
    {
        return $this->width;
    }

    /**
     * Set Width
     *
     * @param int $pValue
     * @return void
     */
    public function setWidth($pValue = 0)
    {
        // Resize proportional?
        if ($this->resizeProportional && $pValue != 0) {
            $ratio = $this->height / ($this->width != 0 ? $this->width : 1);
            $this->height = round($ratio * $pValue);
        }

        // Set width
        $this->width = $pValue;
    }

    /**
     * Get Height
     *
     * @return int
     */
    public function getHeight()
    {
        return $this->height;
    }

    /**
     * Set Height
     *
     * @param int $pValue
     * @return void
     */
    public function setHeight($pValue = 0)
    {
        // Resize proportional?
        if ($this->resizeProportional && $pValue != 0) {
            $ratio = $this->width / ($this->height != 0 ? $this->height : 1);
            $this->width = round($ratio * $pValue);
        }

        // Set height
        $this->height = $pValue;
    }

    /**
     * Set width and height with proportional resize
     * Example:
     * <code>
     * $objDrawing->setResizeProportional(true);
     * $objDrawing->setWidthAndHeight(160,120);
     * </code>
     *
     * @author Vincent@luo MSN:kele_100@hotmail.com
     * @param int $width
     * @param int $height
     * @return void
     */
    public function setWidthAndHeight($width = 0, $height = 0)
    {
        $xratio = $width / ($this->width != 0 ? $this->width : 1);
        $yratio = $height / ($this->height != 0 ? $this->height : 1);
        if ($this->resizeProportional && !($width == 0 || $height == 0)) {
            if (($xratio * $this->height) < $height) {
                $this->height = ceil($xratio * $this->height);
                $this->width  = $width;
            } else {
                $this->width    = ceil($yratio * $this->width);
                $this->height    = $height;
            }
        } else {
            $this->width = $width;
            $this->height = $height;
        }
    }

    /**
     * Get ResizeProportional
     *
     * @return boolean
     */
    public function getResizeProportional()
    {
        return $this->resizeProportional;
    }

    /**
     * Set ResizeProportional
     *
     * @param boolean $pValue
     * @return void
     */
    public function setResizeProportional($pValue = true)
    {
        $this->resizeProportional = $pValue;
    }

    /**
     * Get Rotation
     *
     * @return int
     */
    public function getRotation()
    {
        return $this->rotation;
    }

    /**
     * Set Rotation
     *
     * @param int $pValue
     * @return void
     */
    public function setRotation($pValue = 0)
    {
        $this->rotation = $pValue;
    }

    /**
     * Get Shadow
     *
     * @return PHPExceller_Worksheet_Drawing_Shadow
     */
    public function getShadow()
    {
        return $this->shadow;
    }

    /**
     * Set Shadow
     *
     * @param     PHPExceller_Worksheet_Drawing_Shadow $pValue
     * @throws     PHPExceller_Exception
     * @return void
     */
    public function setShadow(PHPExceller_Worksheet_Drawing_Shadow $pValue = null)
    {
           $this->shadow = $pValue;
    }

    /**
     * Get hash code
     *
     * @return string    Hash code
     */
    public function getHashCode()
    {
        return md5(
            $this->name .
            $this->description .
            $this->worksheet->getHashCode() .
            $this->coordinates .
            $this->offsetX .
            $this->offsetY .
            $this->width .
            $this->height .
            $this->rotation .
            $this->shadow->getHashCode() .
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

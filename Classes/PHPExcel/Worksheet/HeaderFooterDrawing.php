<?php
namespace PHPExceller\Worksheet;

/**
 * PHPExceller_Worksheet_HeaderFooterDrawing
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
class PHPExceller_Worksheet_HeaderFooterDrawing extends PHPExceller_Worksheet_Drawing implements PHPExceller_IComparable
{
    /**
     * Path
     *
     * @var string
     */
    private $path;

    /**
     * Name
     *
     * @var string
     */
    protected $name;

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
     * Create a new PHPExceller_Worksheet_HeaderFooterDrawing
     */
    public function __construct()
    {
        // Initialise values
        $this->path                = '';
        $this->name                = '';
        $this->offsetX             = 0;
        $this->offsetY             = 0;
        $this->width               = 0;
        $this->height              = 0;
        $this->resizeProportional  = true;
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
            $ratio = $this->width / $this->height;
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
            $ratio = $this->width / $this->height;
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
        $xratio = $width / $this->width;
        $yratio = $height / $this->height;
        if ($this->resizeProportional && !($width == 0 || $height == 0)) {
            if (($xratio * $this->height) < $height) {
                $this->height = ceil($xratio * $this->height);
                $this->width  = $width;
            } else {
                $this->width    = ceil($yratio * $this->width);
                $this->height    = $height;
            }
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
     * Get Filename
     *
     * @return string
     */
    public function getFilename()
    {
        return basename($this->path);
    }

    /**
     * Get Extension
     *
     * @return string
     */
    public function getExtension()
    {
        $parts = explode(".", basename($this->path));
        return end($parts);
    }

    /**
     * Get Path
     *
     * @return string
     */
    public function getPath()
    {
        return $this->path;
    }

    /**
     * Set Path
     *
     * @param     string         $pValue            File path
     * @param     boolean        $pVerifyFile    Verify file
     * @throws     PHPExceller_Exception
     * @return void
     */
    public function setPath($pValue = '', $pVerifyFile = true)
    {
        if ($pVerifyFile) {
            if (file_exists($pValue)) {
                $this->path = $pValue;

                if ($this->width == 0 && $this->height == 0) {
                    // Get width/height
                    list($this->width, $this->height) = getimagesize($pValue);
                }
            } else {
                throw new PHPExceller_Exception("File $pValue not found!");
            }
        } else {
            $this->path = $pValue;
        }
    }

    /**
     * Get hash code
     *
     * @return string    Hash code
     */
    public function getHashCode()
    {
        return md5(
            $this->path .
            $this->name .
            $this->offsetX .
            $this->offsetY .
            $this->width .
            $this->height .
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

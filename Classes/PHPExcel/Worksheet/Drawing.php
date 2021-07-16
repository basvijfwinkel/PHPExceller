<?php

/**
 * PHPExceller_Worksheet_Drawing
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
 * @package    PHPExceller_Worksheet_Drawing
 * @copyright  Copyright (c) 2021 PHPExceller
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt    LGPL
 * @version    ##VERSION##, ##DATE##
 */
class PHPExceller_Worksheet_Drawing extends PHPExceller_Worksheet_BaseDrawing implements PHPExceller_IComparable
{
    /**
     * Path
     *
     * @var string
     */
    private $path;

    /**
     * rId of the drawing
     *
     * @var string
     */
    protected $_rId;
    
    protected $_referenceHashTag;
    
    /**
     * Create a new PHPExceller_Worksheet_Drawing
     */
    public function __construct()
    {
        // Initialise values
        $this->_path                = '';
        $this->_referenceHashTag    = null;

        // Initialize parent
        parent::__construct();
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
     * Get indexed filename (using image index)
     *
     * @return string
     */
    public function getIndexedFilename($index = null) {
        $fileName = $this->getFilename();
        $fileName = str_replace(' ', '_', $fileName);
        if (is_null($index)) 
        {
            if (!is_null($this->_rId))
            {
                $index = $this->_rId;
            }
            else
            {
                die('Media should already have been saved');
            }
        }
        return str_replace('.' . $this->getExtension(), '', $fileName) . $index . '.' . $this->getExtension();
    }

    /**
     * Get Extension
     *
     * @return string
     */
    public function getExtension()
    {
        $exploded = explode(".", basename($this->path));
        return $exploded[count($exploded) - 1];
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
            parent::getHashCode() .
            __CLASS__
        );
    }

    /**
     * Implement PHP __clone to create a deep clone, not just a shallow copy.
     */
    public function __clone() {
        $vars = get_object_vars($this);
        foreach ($vars as $key => $value) {
            if (is_object($value)) {
                $this->$key = clone $value;
            } else {
                $this->$key = $value;
            }
        }
    }
    
    public function setMediaReferenceId($rId)
    {
        $this->_rId = $rId;
    }
    
    public function getMediaReferenceId()
    {
        return $this->_rId;
    }
    
    public function getReferenceHashTag()
    {
        return $this->_referenceHashTag ;
    }
}

<?php
namespace PHPExceller\Worksheet;

use PHPExceller\Worksheet\BaseDrawing;
use PHPExceller\IComparable;

/**
 * PHPExceller_Worksheet_MemoryDrawing
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
class MemoryDrawing extends BaseDrawing implements IComparable
{
    /* Rendering functions */
    const RENDERING_DEFAULT                    = 'imagepng';
    const RENDERING_PNG                        = 'imagepng';
    const RENDERING_GIF                        = 'imagegif';
    const RENDERING_JPEG                    = 'imagejpeg';

    /* MIME types */
    const MIMETYPE_DEFAULT                    = 'image/png';
    const MIMETYPE_PNG                        = 'image/png';
    const MIMETYPE_GIF                        = 'image/gif';
    const MIMETYPE_JPEG                        = 'image/jpeg';

    /**
     * Image resource
     *
     * @var resource
     */
    private $_imageResource;

    /**
     * Rendering function
     *
     * @var string
     */
    private $_renderingFunction;

    /**
     * Mime type
     *
     * @var string
     */
    private $_mimeType;

    /**
     * Unique name
     *
     * @var string
     */
    private $_uniqueName;
    
    /**
     * Hashtag for the image we referencing
     *
     * @var string
     */
    protected $_referenceHashTag;

    /**
     * rId of the image
     *
     * @var string
     */
    protected $_rId;
    
    /**
     * Create a new PHPExceller_Worksheet_MemoryDrawing
     */
    public function __construct()
    {
        // Initialise values
        $this->_imageResource        = null;
        $this->_renderingFunction     = self::RENDERING_DEFAULT;
        $this->_mimeType            = self::MIMETYPE_DEFAULT;
        $this->_uniqueName            = md5(rand(0, 9999). time() . rand(0, 9999));
        $this->_referenceHashTag    = null;
        $this->_rId                    = null;

        // Initialize parent
        parent::__construct();
    }

    /**
     * Get image resource
     *
     * @return resource
     */
    public function getImageResource()
    {
        return $this->imageResource;
    }

    /**
     * Set image resource
     *
     * @param    $value resource
     * @return void
     */
    public function setImageResource($value = null)
    {
        $this->imageResource = $value;

        if (!is_null($this->imageResource)) {
            // Get width/height
            $this->width  = imagesx($this->imageResource);
            $this->height = imagesy($this->imageResource);
        }
    }

    /**
     * Get rendering function
     *
     * @return string
     */
    public function getRenderingFunction()
    {
        return $this->renderingFunction;
    }

    /**
     * Set rendering function
     *
     * @param string $value
     * @return PHPExceller_Worksheet_MemoryDrawing
     */
    public function setRenderingFunction($value = PHPExceller_Worksheet_MemoryDrawing::RENDERING_DEFAULT)
    {
        $this->renderingFunction = $value;
    }

    /**
     * Get mime type
     *
     * @return void
     */
    public function getMimeType()
    {
        return $this->mimeType;
    }

    /**
     * Set mime type
     *
     * @param string $value
     * @return void
     */
    public function setMimeType($value = PHPExceller_Worksheet_MemoryDrawing::MIMETYPE_DEFAULT)
    {
        $this->mimeType = $value;
    }

    /**
     * Get indexed filename (using image index)
     *
     * @return string
     */
    public function getIndexedFilename($index=null) 
    {
        $extension     = strtolower($this->getMimeType());
        $extension     = explode('/', $extension);
        $extension     = $extension[1];
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
        return $this->_uniqueName . $index . '.' . $extension;
    }

    /**
     * Get hash code
     *
     * @return string    Hash code
     */
    public function getHashCode()
    {
        return md5(
            $this->renderingFunction .
            $this->mimeType .
            $this->uniqueName .
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
    
    /*
    * set the hashcode of the image we are referencing
    */

    public function setReferenceHashTag($hashTag)
    {
        $this->_referenceHashTag = $hashTag;
    }
    
    /*
    * get the hashcode of the image we are referencing
    */
    public function getReferenceHashTag()
    {
        return $this->_referenceHashTag ;
    }
    
    /*
    * set the rId of the image
    */
    
    public function setMediaReferenceId($rId)
    {
        $this->_rId = $rId;
    }
    
    /*
    * get the rId of the image
    */
    public function getMediaReferenceId()
    {
        return $this->_rId;
    }
    
    /*
    * explicitly set the width and height of the resource
    * This is only used by images that reference another image
    */
    public function setWidthAndHeightForced($width, $height)
    {
        $this->_width = $width;
        $this->_height = $height;
    }
}

<?php
namespace PHPExceller\Shared\Escher\DggContainer\BstoreContainer\BSE;

/**
 * PHPExceller_Shared_Escher_DggContainer_BstoreContainer_BSE_Blip
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
 * @package    PHPExceller_Shared_Escher
 * @copyright  Copyright (c) 2021 PHPExceller
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt    LGPL
 * @version    ##VERSION##, ##DATE##
 */
class PHPExceller_Shared_Escher_DggContainer_BstoreContainer_BSE_Blip
{
    /**
     * The parent BSE
     *
     * @var PHPExceller_Shared_Escher_DggContainer_BstoreContainer_BSE
     */
    private $parent;

    /**
     * Raw image data
     *
     * @var string
     */
    private $data;

    /**
     * Get the raw image data
     *
     * @return string
     */
    public function getData()
    {
        return $this->data;
    }

    /**
     * Set the raw image data
     *
     * @param string
     */
    public function setData($data)
    {
        $this->data = $data;
    }

    /**
     * Set parent BSE
     *
     * @param PHPExceller_Shared_Escher_DggContainer_BstoreContainer_BSE $parent
     */
    public function setParent($parent)
    {
        $this->parent = $parent;
    }

    /**
     * Get parent BSE
     *
     * @return PHPExceller_Shared_Escher_DggContainer_BstoreContainer_BSE $parent
     */
    public function getParent()
    {
        return $this->parent;
    }
}

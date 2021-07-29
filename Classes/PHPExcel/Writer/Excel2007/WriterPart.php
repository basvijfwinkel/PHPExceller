<?php
namespace PHPExceller\Writer;

use PHPExceller\Writer\PHPExceller_Writer_IWriter;
use PHPExceller\Writer\PHPExceller_Writer_Exception;

/**
 * PHPExceller_Writer_Excel2007_WriterPart
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
 * @package    PHPExceller_Writer_Excel2007
 * @copyright  Copyright (c) 2021 PHPExceller
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt    LGPL
 * @version    ##VERSION##, ##DATE##
 */
abstract class PHPExceller_Writer_Excel2007_WriterPart
{
    /**
     * Parent IWriter object
     *
     * @var PHPExceller_Writer_IWriter
     */
    private $parentWriter;

    /**
     * Set parent IWriter object
     *
     * @param PHPExceller_Writer_IWriter    $pWriter
     * @throws PHPExceller_Writer_Exception
     */
    public function setParentWriter(PHPExceller_Writer_IWriter $pWriter = null)
    {
        $this->parentWriter = $pWriter;
    }

    /**
     * Get parent IWriter object
     *
     * @return PHPExceller_Writer_IWriter
     * @throws PHPExceller_Writer_Exception
     */
    public function getParentWriter()
    {
        if (!is_null($this->parentWriter)) {
            return $this->parentWriter;
        } else {
            throw new PHPExceller_Writer_Exception("No parent PHPExceller_Writer_IWriter assigned.");
        }
    }

    /**
     * Set parent IWriter object
     *
     * @param PHPExceller_Writer_IWriter    $pWriter
     * @throws PHPExceller_Writer_Exception
     */
    public function __construct(PHPExceller_Writer_IWriter $pWriter = null)
    {
        if (!is_null($pWriter)) {
            $this->parentWriter = $pWriter;
        }
    }
}

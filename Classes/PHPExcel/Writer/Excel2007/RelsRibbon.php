<?php
namespace PHPExceller\Writer;

/**
 * PHPExceller_Writer_Excel2007_RelsRibbon
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
 * @version     ##VERSION##, ##DATE##
 */
class PHPExceller_Writer_Excel2007_RelsRibbon extends PHPExceller_Writer_Excel2007_WriterPart
{
    /**
     * Write relationships for additional objects of custom UI (ribbon)
     *
     * @param     PHPExceller    $pPHPExceller
     * @return     string         XML Output
     * @throws     PHPExceller_Writer_Exception
     */
    public function writeRibbonRelationships(PHPExceller $pPHPExceller = null)
    {
        // Create XML writer
        $objWriter = null;
        if ($this->getParentWriter()->getUseDiskCaching()) {
            $objWriter = new PHPExceller_Shared_XMLWriter(PHPExceller_Shared_XMLWriter::STORAGE_DISK, $this->getParentWriter()->getDiskCachingDirectory());
        } else {
            $objWriter = new PHPExceller_Shared_XMLWriter(PHPExceller_Shared_XMLWriter::STORAGE_MEMORY);
        }

        // XML header
        $objWriter->startDocument('1.0', 'UTF-8', 'yes');

        // Relationships
        $objWriter->startElement('Relationships');
        $objWriter->writeAttribute('xmlns', 'http://schemas.openxmlformats.org/package/2006/relationships');
        $localRels = $pPHPExceller->getRibbonBinObjects('names');
        if (is_array($localRels)) {
            foreach ($localRels as $aId => $aTarget) {
                $objWriter->startElement('Relationship');
                $objWriter->writeAttribute('Id', $aId);
                $objWriter->writeAttribute('Type', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image');
                $objWriter->writeAttribute('Target', $aTarget);
                $objWriter->endElement();
            }
        }
        $objWriter->endElement();

        return $objWriter->getData();
    }
}

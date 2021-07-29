<?php
namespace PHPExceller;

use ArrayObject;
use PHPExceller\PHPExceller;
use PHPExceller\PHPExceller_IComparable;
use PHPExceller\PHPExceller_CachedObjectStorageFactory;
use PHPExceller\Worksheet\PHPExceller_Worksheet_PageSetup;
use PHPExceller\Worksheet\PHPExceller_Worksheet_PageMargins;
use PHPExceller\Worksheet\PHPExceller_Worksheet_HeaderFooter;
use PHPExceller\Worksheet\PHPExceller_Worksheet_SheetView;
use PHPExceller\Worksheet\PHPExceller_Worksheet_Protection;
use PHPExceller\Worksheet\PHPExceller_Worksheet_RowDimension;
use PHPExceller\Worksheet\PHPExceller_Worksheet_ColumnDimension;
use PHPExceller\Worksheet\PHPExceller_Worksheet_AutoFilter;
use PHPExceller\PHPExceller_Calculation;
use PHPExceller\Shared\PHPExceller_Shared_String;
use PHPExceller\PHPExceller_Exception;
use PHPExceller\PHPExceller_Chart;
use PHPExceller\Shared\PHPExceller_Shared_Font;
use PHPExceller\Style\PHPExceller_Style_NumberFormat;
use PHPExceller\PHPExceller_ReferenceHelper;
use PHPExceller\Cell\PHPExceller_Cell_DataType;
use PHPExceller\PHPExceller_Cell;
use PHPExceller\Style\PHPExceller_Style_Conditional;
use PHPExceller\Worksheet\PHPExceller_Worksheet_RowIterator;
use PHPExceller\Worksheet\PHPExceller_Worksheet_ColumnIterator;

/**
 * PHPExceller_Worksheet
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
 * @copyright  Copyright (c) 2021
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt    LGPL
 * @version    ##VERSION##, ##DATE##
 */
class PHPExceller_Worksheet implements PHPExceller_IComparable
{
    /* Break types */
    const BREAK_NONE   = 0;
    const BREAK_ROW    = 1;
    const BREAK_COLUMN = 2;

    /* Sheet state */
    const SHEETSTATE_VISIBLE    = 'visible';
    const SHEETSTATE_HIDDEN     = 'hidden';
    const SHEETSTATE_VERYHIDDEN = 'veryHidden';

    /**
     * Parent spreadsheet
     *
     * @var PHPExceller
     */
    private $parent;

    /**
     * Cacheable collection of cells
     *
     * @var PHPExceller_CachedObjectStorage_xxx
     */
    private $cellCollection;

    /**
     * Collection of row dimensions
     *
     * @var PHPExceller_Worksheet_RowDimension[]
     */
    private $rowDimensions = array();

    /**
     * Default row dimension
     *
     * @var PHPExceller_Worksheet_RowDimension
     */
    private $defaultRowDimension;

    /**
     * Collection of column dimensions
     *
     * @var PHPExceller_Worksheet_ColumnDimension[]
     */
    private $columnDimensions = array();

    /**
     * Default column dimension
     *
     * @var PHPExceller_Worksheet_ColumnDimension
     */
    private $defaultColumnDimension = null;

    /**
     * Collection of drawings
     *
     * @var PHPExceller_Worksheet_BaseDrawing[]
     */
    private $drawingCollection = null;

    /**
     * Collection of Chart objects
     *
     * @var PHPExceller_Chart[]
     */
    private $chartCollection = array();

    /**
     * Worksheet title
     *
     * @var string
     */
    private $title;

    /**
     * Sheet state
     *
     * @var string
     */
    private $sheetState;

    /**
     * Page setup
     *
     * @var PHPExceller_Worksheet_PageSetup
     */
    private $pageSetup;

    /**
     * Page margins
     *
     * @var PHPExceller_Worksheet_PageMargins
     */
    private $pageMargins;

    /**
     * Page header/footer
     *
     * @var PHPExceller_Worksheet_HeaderFooter
     */
    private $headerFooter;

    /**
     * Sheet view
     *
     * @var PHPExceller_Worksheet_SheetView
     */
    private $sheetView;

    /**
     * Protection
     *
     * @var PHPExceller_Worksheet_Protection
     */
    private $protection;

    /**
     * Collection of styles
     *
     * @var PHPExceller_Style[]
     */
    private $styles = array();

    /**
     * Conditional styles. Indexed by cell coordinate, e.g. 'A1'
     *
     * @var array
     */
    private $conditionalStylesCollection = array();

    /**
     * Is the current cell collection sorted already?
     *
     * @var boolean
     */
    private $cellCollectionIsSorted = false;

    /**
     * Collection of breaks
     *
     * @var array
     */
    private $breaks = array();

    /**
     * Collection of merged cell ranges
     *
     * @var array
     */
    private $mergeCells = array();

    /**
     * Collection of protected cell ranges
     *
     * @var array
     */
    private $protectedCells = array();

    /**
     * Autofilter Range and selection
     *
     * @var PHPExceller_Worksheet_AutoFilter
     */
    private $autoFilter;

    /**
     * Freeze pane
     *
     * @var string
     */
    private $freezePane = '';

    /**
     * Show gridlines?
     *
     * @var boolean
     */
    private $showGridlines = true;

    /**
    * Print gridlines?
    *
    * @var boolean
    */
    private $printGridlines = false;

    /**
    * Show row and column headers?
    *
    * @var boolean
    */
    private $showRowColHeaders = true;

    /**
     * Show summary below? (Row/Column outline)
     *
     * @var boolean
     */
    private $showSummaryBelow = true;

    /**
     * Show summary right? (Row/Column outline)
     *
     * @var boolean
     */
    private $showSummaryRight = true;

    /**
     * Collection of comments
     *
     * @var PHPExceller_Comment[]
     */
    private $comments = array();

    /**
     * Active cell. (Only one!)
     *
     * @var string
     */
    private $activeCell = 'A1';

    /**
     * Selected cells
     *
     * @var string
     */
    private $selectedCells = 'A1';

    /**
     * Cached highest column
     *
     * @var string
     */
    private $cachedHighestColumn = 'A';

    /**
     * Cached highest row
     *
     * @var int
     */
    private $cachedHighestRow = 1;

    /**
     * Right-to-left?
     *
     * @var boolean
     */
    private $rightToLeft = false;

    /**
     * Hyperlinks. Indexed by cell coordinate, e.g. 'A1'
     *
     * @var array
     */
    private $hyperlinkCollection = array();

    /**
     * Data validation objects. Indexed by cell coordinate, e.g. 'A1'
     *
     * @var array
     */
    private $dataValidationCollection = array();

    /**
     * Tab color
     *
     * @var PHPExceller_Style_Color
     */
    private $tabColor;

    /**
     * Dirty flag
     *
     * @var boolean
     */
    private $dirty = true;

    /**
     * Hash
     *
     * @var string
     */
    private $hash;

    /**
    * CodeName
    *
    * @var string
    */
    private $codeName = null;

    /**
     * Create a new worksheet
     *
     * @param PHPExceller        $pParent
     * @param string        $pTitle
     */
    public function __construct(PHPExceller $pParent = null, $pTitle = 'Worksheet')
    {
        // Set parent and title
        $this->parent = $pParent;
        $this->setTitle($pTitle, false);
        // setTitle can change $pTitle
        $this->setCodeName($this->getTitle());
        $this->setSheetState(PHPExceller_Worksheet::SHEETSTATE_VISIBLE);

        $this->cellCollection         = PHPExceller_CachedObjectStorageFactory::getInstance($this);
        // Set page setup
        $this->pageSetup              = new PHPExceller_Worksheet_PageSetup();
        // Set page margins
        $this->pageMargins            = new PHPExceller_Worksheet_PageMargins();
        // Set page header/footer
        $this->headerFooter           = new PHPExceller_Worksheet_HeaderFooter();
        // Set sheet view
        $this->sheetView              = new PHPExceller_Worksheet_SheetView();
        // Drawing collection
        $this->drawingCollection      = new ArrayObject();
        // Chart collection
        $this->chartCollection        = new ArrayObject();
        // Protection
        $this->protection             = new PHPExceller_Worksheet_Protection();
        // Default row dimension
        $this->defaultRowDimension    = new PHPExceller_Worksheet_RowDimension(null);
        // Default column dimension
        $this->defaultColumnDimension = new PHPExceller_Worksheet_ColumnDimension(null);
        $this->autoFilter             = new PHPExceller_Worksheet_AutoFilter(null, $this);
    }


    /**
     * Disconnect all cells from this PHPExceller_Worksheet object,
     *    typically so that the worksheet object can be unset
     *
     */
    public function disconnectCells() {
        if ( $this->cellCollection !== NULL){
            $this->cellCollection->unsetWorksheetCells();
            $this->cellCollection = NULL;
        }
        //    detach ourself from the workbook, so that it can then delete this worksheet successfully
        $this->parent = null;
    }

    /**
     * Code to execute when this worksheet is unset()
     *
     */
    function __destruct() {
        PHPExceller_Calculation::getInstance($this->parent)
            ->clearCalculationCacheForWorksheet($this->title);
        $this->disconnectCells();
    }

   /**
     * Return the cache controller for the cell collection
     *
     * @return PHPExceller_CachedObjectStorage_xxx
     */
    public function getCellCacheController() {
        return $this->cellCollection;
    }    //    function getCellCacheController()

    /**
     * Get array of invalid characters for sheet title
     *
     * @return array
     */
    public static function getInvalidCharacters()
    {
        return self::$invalidCharacters;
    }

    /**
     * Check sheet code name for valid Excel syntax
     *
     * @param string $pValue The string to check
     * @return string The valid string
     * @throws Exception
     */
    private static function checkSheetCodeName($pValue)
    {
        $CharCount = PHPExceller_Shared_String::CountCharacters($pValue);
        if ($CharCount == 0) {
            throw new PHPExceller_Exception('Sheet code name cannot be empty.');
        }
        // Some of the printable ASCII characters are invalid:  * : / \ ? [ ] and  first and last characters cannot be a "'"
        if ((str_replace(self::$invalidCharacters, '', $pValue) !== $pValue) ||
            (PHPExceller_Shared_String::Substring($pValue, -1, 1)=='\'') ||
            (PHPExceller_Shared_String::Substring($pValue, 0, 1)=='\'')) {
            throw new PHPExceller_Exception('Invalid character found in sheet code name');
        }

        // Maximum 31 characters allowed for sheet title
        if ($CharCount > 31) {
            throw new PHPExceller_Exception('Maximum 31 characters allowed in sheet code name.');
        }

        return $pValue;
    }

   /**
     * Check sheet title for valid Excel syntax
     *
     * @param string $pValue The string to check
     * @return string The valid string
     * @throws PHPExceller_Exception
     */
    private static function checkSheetTitle($pValue)
    {
        // Some of the printable ASCII characters are invalid:  * : / \ ? [ ]
        if (str_replace(self::$invalidCharacters, '', $pValue) !== $pValue) {
            throw new PHPExceller_Exception('Invalid character found in sheet title');
        }

        // Maximum 31 characters allowed for sheet title
        if (PHPExceller_Shared_String::CountCharacters($pValue) > 31) {
            throw new PHPExceller_Exception('Maximum 31 characters allowed in sheet title.');
        }

        return $pValue;
    }

    /**
     * Get collection of cells
     *
     * @param boolean $pSorted Also sort the cell collection?
     * @return PHPExceller_Cell[]
     */
    public function getCellCollection($pSorted = true)
    {
        if ($pSorted) {
            // Re-order cell collection
            return $this->sortCellCollection();
        }
        if ($this->cellCollection !== null) {
            return $this->cellCollection->getCellList();
        }
        return array();
    }

    /**
     * Sort collection of cells
     *
     * @return void
     */
    public function sortCellCollection()
    {
        if ($this->cellCollection !== null) {
            return $this->cellCollection->getSortedCellList();
        }
    }

    /**
     * Get collection of row dimensions
     *
     * @return PHPExceller_Worksheet_RowDimension[]
     */
    public function getRowDimensions()
    {
        return $this->rowDimensions;
    }

    /**
     * Get default row dimension
     *
     * @return PHPExceller_Worksheet_RowDimension
     */
    public function getDefaultRowDimension()
    {
        return $this->defaultRowDimension;
    }

    /**
     * Get collection of column dimensions
     *
     * @return PHPExceller_Worksheet_ColumnDimension[]
     */
    public function getColumnDimensions()
    {
        return $this->columnDimensions;
    }

    /**
     * Get default column dimension
     *
     * @return PHPExceller_Worksheet_ColumnDimension
     */
    public function getDefaultColumnDimension()
    {
        return $this->defaultColumnDimension;
    }

    /**
     * Get collection of drawings
     *
     * @return PHPExceller_Worksheet_BaseDrawing[]
     */
    public function getDrawingCollection()
    {
        return $this->drawingCollection;
    }

    /**
     * Get collection of charts
     *
     * @return PHPExceller_Chart[]
     */
    public function getChartCollection()
    {
        return $this->chartCollection;
    }

    /**
     * Add chart
     *
     * @param PHPExceller_Chart $pChart
     * @param int|null $iChartIndex Index where chart should go (0,1,..., or null for last)
     * @return PHPExceller_Chart
     */
    public function addChart(PHPExceller_Chart $pChart = null, $iChartIndex = null)
    {
        $pChart->setWorksheet($this);
        if (is_null($iChartIndex)) {
            $this->chartCollection[] = $pChart;
        } else {
            // Insert the chart at the requested index
            array_splice($this->chartCollection, $iChartIndex, 0, array($pChart));
        }

        return $pChart;
    }

    /**
     * Return the count of charts on this worksheet
     *
     * @return int        The number of charts
     */
    public function getChartCount()
    {
        return count($this->chartCollection);
    }

    /**
     * Get a chart by its index position
     *
     * @param string $index Chart index position
     * @return false|PHPExceller_Chart
     * @throws PHPExceller_Exception
     */
    public function getChartByIndex($index = null)
    {
        $chartCount = count($this->chartCollection);
        if ($chartCount == 0) {
            return false;
        }
        if (is_null($index)) {
            $index = --$chartCount;
        }
        if (!isset($this->chartCollection[$index])) {
            return false;
        }

        return $this->chartCollection[$index];
    }

    /**
     * Return an array of the names of charts on this worksheet
     *
     * @return string[] The names of charts
     * @throws PHPExceller_Exception
     */
    public function getChartNames()
    {
        $chartNames = array();
        foreach ($this->chartCollection as $chart) {
            $chartNames[] = $chart->getName();
        }
        return $chartNames;
    }

    /**
     * Get a chart by name
     *
     * @param string $chartName Chart name
     * @return false|PHPExceller_Chart
     * @throws PHPExceller_Exception
     */
    public function getChartByName($chartName = '')
    {
        $chartCount = count($this->chartCollection);
        if ($chartCount == 0) {
            return false;
        }
        foreach ($this->chartCollection as $index => $chart) {
            if ($chart->getName() == $chartName) {
                return $this->chartCollection[$index];
            }
        }
        return false;
    }

    /**
     * Refresh column dimensions
     *
     * @return void
     */
    public function refreshColumnDimensions()
    {
        $currentColumnDimensions = $this->getColumnDimensions();
        $newColumnDimensions = array();

        foreach ($currentColumnDimensions as $objColumnDimension) {
            $newColumnDimensions[$objColumnDimension->getColumnIndex()] = $objColumnDimension;
        }

        $this->columnDimensions = $newColumnDimensions;
    }

    /**
     * Refresh row dimensions
     *
     * @return void
     */
    public function refreshRowDimensions()
    {
        $currentRowDimensions = $this->getRowDimensions();
        $newRowDimensions = array();

        foreach ($currentRowDimensions as $objRowDimension) {
            $newRowDimensions[$objRowDimension->getRowIndex()] = $objRowDimension;
        }

        $this->rowDimensions = $newRowDimensions;
    }

    /**
     * Calculate worksheet dimension
     *
     * @return string  String containing the dimension of this worksheet
     */
    public function calculateWorksheetDimension()
    {
        // Return
        return 'A1' . ':' .  $this->getHighestColumn() . $this->getHighestRow();
    }

    /**
     * Calculate worksheet data dimension
     *
     * @return string  String containing the dimension of this worksheet that actually contain data
     */
    public function calculateWorksheetDataDimension()
    {
        // Return
        return 'A1' . ':' .  $this->getHighestDataColumn() . $this->getHighestDataRow();
    }

    /**
     * Calculate widths for auto-size columns
     *
     * @param  boolean  $calculateMergeCells  Calculate merge cell width
     * @return void;
     */
    public function calculateColumnWidths($calculateMergeCells = false)
    {
        // initialize $autoSizes array
        $autoSizes = array();
        foreach ($this->getColumnDimensions() as $colDimension) {
            if ($colDimension->getAutoSize()) {
                $autoSizes[$colDimension->getColumnIndex()] = -1;
            }
        }

        // There is only something to do if there are some auto-size columns
        if (!empty($autoSizes)) {
            // build list of cells references that participate in a merge
            $isMergeCell = array();
            foreach ($this->getMergeCells() as $cells) {
                foreach (PHPExceller_Cell::extractAllCellReferencesInRange($cells) as $cellReference) {
                    $isMergeCell[$cellReference] = true;
                }
            }

            // loop through all cells in the worksheet
            foreach ($this->getCellCollection(false) as $cellID) {
                $cell = $this->getCell($cellID);
                if (isset($autoSizes[$this->cellCollection->getCurrentColumn()])) {
                    // Determine width if cell does not participate in a merge
                    if (!isset($isMergeCell[$this->cellCollection->getCurrentAddress()])) {
                        // Calculated value
                        // To formatted string
                        $cellValue = PHPExceller_Style_NumberFormat::toFormattedString(
                            $cell->getCalculatedValue(),
                            $this->getParent()->getCellXfByIndex($cell->getXfIndex())->getNumberFormat()->getFormatCode()
                        );

                        $autoSizes[$this->cellCollection->getCurrentColumn()] = max(
                            (float) $autoSizes[$this->cellCollection->getCurrentColumn()],
                            (float)PHPExceller_Shared_Font::calculateColumnWidth(
                                $this->getParent()->getCellXfByIndex($cell->getXfIndex())->getFont(),
                                $cellValue,
                                $this->getParent()->getCellXfByIndex($cell->getXfIndex())->getAlignment()->getTextRotation(),
                                $this->getDefaultStyle()->getFont()
                            )
                        );
                    }
                }
            }

            // adjust column widths
            foreach ($autoSizes as $columnIndex => $width) {
                if ($width == -1) {
                    $width = $this->getDefaultColumnDimension()->getWidth();
                }
                $this->getColumnDimension($columnIndex)->setWidth($width);
            }
        }
    }

    /**
     * Get parent
     *
     * @return PHPExceller
     */
    public function getParent() {
        return $this->parent;
    }

    /**
     * Re-bind parent
     *
     * @param PHPExceller $parent
     * @return void
     */
    public function rebindParent(PHPExceller $parent) {
        if ($this->parent !== null) {
            $namedRanges = $this->parent->getNamedRanges();
            foreach ($namedRanges as $namedRange) {
                $parent->addNamedRange($namedRange);
            }

            $this->parent->removeSheetByIndex(
                $this->parent->getIndex($this)
            );
        }
        $this->parent = $parent;
    }

    /**
     * Get title
     *
     * @return string
     */
    public function getTitle()
    {
        return $this->title;
    }

    /**
     * Set title
     *
     * @param string $pValue String containing the dimension of this worksheet
     * @param string $updateFormulaCellReferences boolean Flag indicating whether cell references in formulae should
     *            be updated to reflect the new sheet name.
     *          This should be left as the default true, unless you are
     *          certain that no formula cells on any worksheet contain
     *          references to this worksheet
     * @return void
     */
    public function setTitle($pValue = 'Worksheet', $updateFormulaCellReferences = true)
    {
        // Is this a 'rename' or not?
        if ($this->getTitle() == $pValue)
        {
            // nothing to do here
            return;
        }

        // Syntax check
        self::checkSheetTitle($pValue);

        // Old title
        $oldTitle = $this->getTitle();

        if ($this->parent) {
            // Is there already such sheet name?
            if ($this->parent->sheetNameExists($pValue)) {
                // Use name, but append with lowest possible integer

                if (PHPExceller_Shared_String::CountCharacters($pValue) > 29) {
                    $pValue = PHPExceller_Shared_String::Substring($pValue, 0, 29);
                }
                $i = 1;
                while ($this->parent->sheetNameExists($pValue . ' ' . $i)) {
                    ++$i;
                    if ($i == 10) {
                        if (PHPExceller_Shared_String::CountCharacters($pValue) > 28) {
                            $pValue = PHPExceller_Shared_String::Substring($pValue, 0, 28);
                        }
                    } elseif ($i == 100) {
                        if (PHPExceller_Shared_String::CountCharacters($pValue) > 27) {
                            $pValue = PHPExceller_Shared_String::Substring($pValue, 0, 27);
                        }
                    }
                }

                $altTitle = $pValue . ' ' . $i;
                return $this->setTitle($altTitle, $updateFormulaCellReferences);
            }
        }

        // Set title
        $this->title = $pValue;
        $this->dirty = true;

        if ($this->parent && $this->parent->getCalculationEngine()) {
            // New title
            $newTitle = $this->getTitle();
            PHPExceller_Calculation::getInstance($this->parent)
                ->renameCalculationCacheForWorksheet($oldTitle, $newTitle);
            if ($updateFormulaCellReferences)
                PHPExceller_ReferenceHelper::getInstance()->updateNamedFormulas($this->parent, $oldTitle, $newTitle);
        }

        // also update all graphs related to this workbook
        foreach ($this->chartCollection as $chart)
        {
            $chart->updateWorkbookName($pValue);
        }
    }

    /**
     * Get sheet state
     *
     * @return string Sheet state (visible, hidden, veryHidden)
     */
    public function getSheetState() {
        return $this->sheetState;
    }

    /**
     * Set sheet state
     *
     * @param string $value Sheet state (visible, hidden, veryHidden)
     * @return void
     */
    public function setSheetState($value = PHPExceller_Worksheet::SHEETSTATE_VISIBLE) {
        $this->sheetState = $value;
    }

    /**
     * Get page setup
     *
     * @return PHPExceller_Worksheet_PageSetup
     */
    public function getPageSetup()
    {
        return $this->pageSetup;
    }

    /**
     * Set page setup
     *
     * @param PHPExceller_Worksheet_PageSetup    $pValue
     * @return void
     */
    public function setPageSetup(PHPExceller_Worksheet_PageSetup $pValue)
    {
        $this->pageSetup = $pValue;
    }

    /**
     * Get page margins
     *
     * @return PHPExceller_Worksheet_PageMargins
     */
    public function getPageMargins()
    {
        return $this->pageMargins;
    }

    /**
     * Set page margins
     *
     * @param PHPExceller_Worksheet_PageMargins    $pValue
     * @return void
     */
    public function setPageMargins(PHPExceller_Worksheet_PageMargins $pValue)
    {
        $this->pageMargins = $pValue;
    }

    /**
     * Get page header/footer
     *
     * @return PHPExceller_Worksheet_HeaderFooter
     */
    public function getHeaderFooter()
    {
        return $this->headerFooter;
    }

    /**
     * Set page header/footer
     *
     * @param PHPExceller_Worksheet_HeaderFooter    $pValue
     * @return void
     */
    public function setHeaderFooter(PHPExceller_Worksheet_HeaderFooter $pValue)
    {
        $this->headerFooter = $pValue;
    }

    /**
     * Get sheet view
     *
     * @return PHPExceller_Worksheet_SheetView
     */
    public function getSheetView()
    {
        return $this->sheetView;
    }

    /**
     * Set sheet view
     *
     * @param PHPExceller_Worksheet_SheetView    $pValue
     * @return void
     */
    public function setSheetView(PHPExceller_Worksheet_SheetView $pValue)
    {
        $this->sheetView = $pValue;
    }

    /**
     * Get Protection
     *
     * @return PHPExceller_Worksheet_Protection
     */
    public function getProtection()
    {
        return $this->protection;
    }

    /**
     * Set Protection
     *
     * @param PHPExceller_Worksheet_Protection    $pValue
     * @return void
     */
    public function setProtection(PHPExceller_Worksheet_Protection $pValue)
    {
        $this->protection = $pValue;
        $this->dirty = true;
    }

    /**
     * Get highest worksheet column
     *
     * @param   string     $row        Return the data highest column for the specified row,
     *                                     or the highest column of any row if no row number is passed
     * @return string Highest column name
     */
    public function getHighestColumn($row = null)
    {
        if ($row == null) {
            return $this->cachedHighestColumn;
        }
        return $this->getHighestDataColumn($row);
    }

    /**
     * Get highest worksheet column that contains data
     *
     * @param   string     $row        Return the highest data column for the specified row,
     *                                     or the highest data column of any row if no row number is passed
     * @return string Highest column name that contains data
     */
    public function getHighestDataColumn($row = null)
    {
        return $this->cellCollection->getHighestColumn($row);
    }

    /**
     * Get highest worksheet row
     *
     * @param   string     $column     Return the highest data row for the specified column,
     *                                     or the highest row of any column if no column letter is passed
     * @return int Highest row number
     */
    public function getHighestRow($column = null)
    {
        if ($column == null) {
            return $this->cachedHighestRow;
        }
        return $this->getHighestDataRow($column);
    }

    /**
     * Get highest worksheet row that contains data
     *
     * @param   string     $column     Return the highest data row for the specified column,
     *                                     or the highest data row of any column if no column letter is passed
     * @return string Highest row number that contains data
     */
    public function getHighestDataRow($column = null)
    {
        return $this->cellCollection->getHighestRow($column);
    }

    /**
     * Get highest worksheet column and highest row that have cell records
     *
     * @return array Highest column name and highest row number
     */
    public function getHighestRowAndColumn()
    {
        return $this->cellCollection->getHighestRowAndColumn();
    }

    /**
     * Set a cell value
     *
     * @param string $pCoordinate Coordinate of the cell
     * @param mixed $pValue Value of the cell
     * @param bool $returnCell   Return the worksheet (false, default) or the cell (true)
     * @return void
     */
    public function setCellValue($pCoordinate = 'A1', $pValue = null, $returnCell = false)
    {
        $cell = $this->getCell(strtoupper($pCoordinate))->setValue($pValue);
    }

    /**
     * Set a cell value by using numeric cell coordinates
     *
     * @param string $pColumn Numeric column coordinate of the cell (A = 0)
     * @param string $pRow Numeric row coordinate of the cell
     * @param mixed $pValue Value of the cell
     * @param bool $returnCell Return the worksheet (false, default) or the cell (true)
     * @return void
     */
    public function setCellValueByColumnAndRow($pColumn = 0, $pRow = 1, $pValue = null, $returnCell = false)
    {
        $cell = $this->getCellByColumnAndRow($pColumn, $pRow)->setValue($pValue);
    }

    /**
     * Set a cell value
     *
     * @param string $pCoordinate Coordinate of the cell
     * @param mixed  $pValue Value of the cell
     * @param string $pDataType Explicit data type
     * @param bool $returnCell Return the worksheet (false, default) or the cell (true)
     * @return void
     */
    public function setCellValueExplicit($pCoordinate = 'A1', $pValue = null, $pDataType = PHPExceller_Cell_DataType::TYPE_STRING, $returnCell = false)
    {
        // Set value
        $cell = $this->getCell(strtoupper($pCoordinate))->setValueExplicit($pValue, $pDataType);
    }

    /**
     * Set a cell value by using numeric cell coordinates
     *
     * @param string $pColumn Numeric column coordinate of the cell
     * @param string $pRow Numeric row coordinate of the cell
     * @param mixed $pValue Value of the cell
     * @param string $pDataType Explicit data type
     * @param bool $returnCell Return the worksheet (false, default) or the cell (true)
     * @return void
     */
    public function setCellValueExplicitByColumnAndRow($pColumn = 0, $pRow = 1, $pValue = null, $pDataType = PHPExceller_Cell_DataType::TYPE_STRING, $returnCell = false)
    {
        $cell = $this->getCellByColumnAndRow($pColumn, $pRow)->setValueExplicit($pValue, $pDataType);
    }

    /**
     * Get cell at a specific coordinate
     *
     * @param string $pCoordinate    Coordinate of the cell
     * @param boolean $createIfNotExists  Flag indicating whether a new cell should be created if it doesn't
     *                                       already exist, or a null should be returned instead
     * @throws PHPExceller_Exception
     * @return null|PHPExceller_Cell Cell that was found/created or null
     */
    public function getCell($pCoordinate = 'A1', $createIfNotExists = true)
    {
        // Check cell collection
        if ($this->cellCollection->isDataSet(strtoupper($pCoordinate))) {
            return $this->cellCollection->getCacheData($pCoordinate);
        }

        // Worksheet reference?
        if (strpos($pCoordinate, '!') !== false) {
            $worksheetReference = PHPExceller_Worksheet::extractSheetTitle($pCoordinate, true);
            return $this->parent->getSheetByName($worksheetReference[0])->getCell(strtoupper($worksheetReference[1]));
        }

        // Named range?
        if ((!preg_match('/^'.PHPExceller_Calculation::CALCULATION_REGEXP_CELLREF.'$/i', $pCoordinate, $matches)) &&
            (preg_match('/^'.PHPExceller_Calculation::CALCULATION_REGEXP_NAMEDRANGE.'$/i', $pCoordinate, $matches))) {
            $namedRange = PHPExceller_NamedRange::resolveRange($pCoordinate, $this);
            if ($namedRange !== null) {
                $pCoordinate = $namedRange->getRange();
                return $namedRange->getWorksheet()->getCell($pCoordinate, $createIfNotExists);
            }
        }

        // Uppercase coordinate
        $pCoordinate = strtoupper($pCoordinate);

        if (strpos($pCoordinate, ':') !== false || strpos($pCoordinate, ',') !== false) {
            throw new PHPExceller_Exception('Cell coordinate can not be a range of cells.');
        } elseif (strpos($pCoordinate, '$') !== false) {
            throw new PHPExceller_Exception('Cell coordinate must not be absolute.');
        }

        // Create new cell object, if required
        return $createIfNotExists ? $this->createNewCell($pCoordinate) : null;
    }

    /**
     * Get cell at a specific coordinate by using numeric cell coordinates
     *
     * @param  string $pColumn Numeric column coordinate of the cell (starting from 0)
     * @param string $pRow Numeric row coordinate of the cell
     * @param boolean $createIfNotExists  Flag indicating whether a new cell should be created if it doesn't
     *                                       already exist, or a null should be returned instead
     * @return null|PHPExceller_Cell Cell that was found/created or null
     */
    public function getCellByColumnAndRow($pColumn = 0, $pRow = 1, $createIfNotExists = true)
    {
        $columnLetter = PHPExceller_Cell::stringFromColumnIndex($pColumn);
        $coordinate = $columnLetter . $pRow;

        if ($this->cellCollection->isDataSet($coordinate)) {
            return $this->cellCollection->getCacheData($coordinate);
        }

        return $this->createNewCell($coordinate);
    }

    /**
     * Create a new cell at the specified coordinate
     *
     * @param string $pCoordinate    Coordinate of the cell
     * @return PHPExceller_Cell Cell that was created
     */
    private function _createNewCell($pCoordinate)
    {
        $cell = $this->cellCollection->addCacheData(
            $pCoordinate,
            new PHPExceller_Cell(
                NULL,
                PHPExceller_Cell_DataType::TYPE_NULL,
                $this
            )
        );
        $this->cellCollectionIsSorted = false;

        // Coordinates
        $aCoordinates = PHPExceller_Cell::coordinateFromString($pCoordinate);
        if (PHPExceller_Cell::columnIndexFromString($this->cachedHighestColumn) < PHPExceller_Cell::columnIndexFromString($aCoordinates[0])) {
            $this->cachedHighestColumn = $aCoordinates[0];
        }
        $this->cachedHighestRow = max($this->cachedHighestRow, $aCoordinates[1]);

        // Cell needs appropriate xfIndex from dimensions records
        //    but don't create dimension records if they don't already exist
        $rowDimension    = $this->getRowDimension($aCoordinates[1], FALSE);
        $columnDimension = $this->getColumnDimension($aCoordinates[0], FALSE);

        if ($rowDimension !== null && $rowDimension->getXfIndex() > 0) {
            // then there is a row dimension with explicit style, assign it to the cell
            $cell->setXfIndex($rowDimension->getXfIndex());
        } elseif ($columnDimension !== null && $columnDimension->getXfIndex() > 0) {
            // then there is a column dimension, assign it to the cell
            $cell->setXfIndex($columnDimension->getXfIndex());
        }

        return $cell;
    }

    /**
     * Does the cell at a specific coordinate exist?
     *
     * @param string $pCoordinate  Coordinate of the cell
     * @throws PHPExceller_Exception
     * @return boolean
     */
    public function cellExists($pCoordinate = 'A1')
    {
       // Worksheet reference?
        if (strpos($pCoordinate, '!') !== false) {
            $worksheetReference = PHPExceller_Worksheet::extractSheetTitle($pCoordinate, true);
            return $this->parent->getSheetByName($worksheetReference[0])->cellExists(strtoupper($worksheetReference[1]));
        }

        // Named range?
        if ((!preg_match('/^'.PHPExceller_Calculation::CALCULATION_REGEXP_CELLREF.'$/i', $pCoordinate, $matches)) &&
            (preg_match('/^'.PHPExceller_Calculation::CALCULATION_REGEXP_NAMEDRANGE.'$/i', $pCoordinate, $matches))) {
            $namedRange = PHPExceller_NamedRange::resolveRange($pCoordinate, $this);
            if ($namedRange !== null) {
                $pCoordinate = $namedRange->getRange();
                if ($this->getHashCode() != $namedRange->getWorksheet()->getHashCode()) {
                    if (!$namedRange->getLocalOnly()) {
                        return $namedRange->getWorksheet()->cellExists($pCoordinate);
                    } else {
                        throw new PHPExceller_Exception('Named range ' . $namedRange->getName() . ' is not accessible from within sheet ' . $this->getTitle());
                    }
                }
            } else {
                return false;
            }
        }

        // Uppercase coordinate
        $pCoordinate = strtoupper($pCoordinate);

        if (strpos($pCoordinate, ':') !== false || strpos($pCoordinate, ',') !== false) {
            throw new PHPExceller_Exception('Cell coordinate can not be a range of cells.');
        } elseif (strpos($pCoordinate, '$') !== false) {
            throw new PHPExceller_Exception('Cell coordinate must not be absolute.');
        } else {
            // Coordinates
            $aCoordinates = PHPExceller_Cell::coordinateFromString($pCoordinate);

            // Cell exists?
            return $this->cellCollection->isDataSet($pCoordinate);
        }
    }

    /**
     * Cell at a specific coordinate by using numeric cell coordinates exists?
     *
     * @param string $pColumn Numeric column coordinate of the cell
     * @param string $pRow Numeric row coordinate of the cell
     * @return boolean
     */
    public function cellExistsByColumnAndRow($pColumn = 0, $pRow = 1)
    {
        return $this->cellExists(PHPExceller_Cell::stringFromColumnIndex($pColumn) . $pRow);
    }

    /**
     * Get row dimension at a specific row
     *
     * @param int $pRow Numeric index of the row
     * @return PHPExceller_Worksheet_RowDimension
     */
    public function getRowDimension($pRow = 1, $create = true)
    {
        // Found
        $found = null;

        // Get row dimension
        if (!isset($this->rowDimensions[$pRow])) {
            if (!$create)
                return NULL;
            $this->rowDimensions[$pRow] = new PHPExceller_Worksheet_RowDimension($pRow);

            $this->cachedHighestRow = max($this->cachedHighestRow, $pRow);
        }
        return $this->rowDimensions[$pRow];
    }

    /**
     * Get column dimension at a specific column
     *
     * @param string $pColumn String index of the column
     * @return PHPExceller_Worksheet_ColumnDimension
     */
    public function getColumnDimension($pColumn = 'A', $create = true)
    {
        // Uppercase coordinate
        $pColumn = strtoupper($pColumn);

        // Fetch dimensions
        if (!isset($this->columnDimensions[$pColumn])) {
            if (!$create)
                return NULL;
            $this->columnDimensions[$pColumn] = new PHPExceller_Worksheet_ColumnDimension($pColumn);

            if (PHPExceller_Cell::columnIndexFromString($this->cachedHighestColumn) < PHPExceller_Cell::columnIndexFromString($pColumn)) {
                $this->cachedHighestColumn = $pColumn;
            }
        }
        return $this->columnDimensions[$pColumn];
    }

    /**
     * Get column dimension at a specific column by using numeric cell coordinates
     *
     * @param string $pColumn Numeric column coordinate of the cell
     * @return PHPExceller_Worksheet_ColumnDimension
     */
    public function getColumnDimensionByColumn($pColumn = 0)
    {
        return $this->getColumnDimension(PHPExceller_Cell::stringFromColumnIndex($pColumn));
    }

    /**
     * Get styles
     *
     * @return PHPExceller_Style[]
     */
    public function getStyles()
    {
        return $this->styles;
    }

    /**
     * Get default style of workbook.
     *
     * @deprecated
     * @return PHPExceller_Style
     * @throws PHPExceller_Exception
     */
    public function getDefaultStyle()
    {
        return $this->parent->getDefaultStyle();
    }

    /**
     * Set default style - should only be used by PHPExceller_IReader implementations!
     *
     * @deprecated
     * @param PHPExceller_Style $pValue
     * @throws PHPExceller_Exception
     * @return void
     */
    public function setDefaultStyle(PHPExceller_Style $pValue)
    {
        $this->parent->getDefaultStyle()->applyFromArray(array(
            'font' => array(
                'name' => $pValue->getFont()->getName(),
                'size' => $pValue->getFont()->getSize(),
            ),
        ));
    }

    /**
     * Get style for cell
     *
     * @param string $pCellCoordinate Cell coordinate (or range) to get style for
     * @return PHPExceller_Style
     * @throws PHPExceller_Exception
     */
    public function getStyle($pCellCoordinate = 'A1')
    {
        // set this sheet as active
        $this->parent->setActiveSheetIndex($this->parent->getIndex($this));

        // set cell coordinate as active
        $this->setSelectedCells(strtoupper($pCellCoordinate));

        return $this->parent->getCellXfSupervisor();
    }

    /**
     * Get conditional styles for a cell
     *
     * @param string $pCoordinate
     * @return PHPExceller_Style_Conditional[]
     */
    public function getConditionalStyles($pCoordinate = 'A1')
    {
        $pCoordinate = strtoupper($pCoordinate);
        if (!isset($this->conditionalStylesCollection[$pCoordinate])) {
            $this->conditionalStylesCollection[$pCoordinate] = array();
        }
        return $this->conditionalStylesCollection[$pCoordinate];
    }

    /**
     * Do conditional styles exist for this cell?
     *
     * @param string $pCoordinate
     * @return boolean
     */
    public function conditionalStylesExists($pCoordinate = 'A1')
    {
        if (isset($this->conditionalStylesCollection[strtoupper($pCoordinate)])) {
            return true;
        }
        return false;
    }

    /**
     * Removes conditional styles for a cell
     *
     * @param string $pCoordinate
     * @return void
     */
    public function removeConditionalStyles($pCoordinate = 'A1')
    {
        if (isset($this->conditionalStylesCollection[$pCoordinate]))
        {
            $conditions = $this->conditionalStylesCollection[$pCoordinate];
            foreach($conditions as $condition)
            {
                $condition->removeCellReference($pCoordinate);
            }
            unset($this->conditionalStylesCollection[$pCoordinate]);
        }
    }

    public function removeConditionalStyle($pCoordinate = 'A1', $hashCode)
    {
        if (isset($this->conditionalStylesCollection[$pCoordinate]))
        {
            $conditions = $this->conditionalStylesCollection[$pCoordinate];
            foreach($conditions as $index => $condition)
            {
                if ($condition->getHashCode() == $hashCode)
                {
                    if (strpos($condition->getCellReference(),':') === false)
                    {
                        // remove single cell reference formattings (cond. formatting applied to ranges will be rebuild when writing to output)
                        $condition->removeCellReference($pCoordinate);
                    }
                    // remove the reference form the collection
                    unset($this->conditionalStylesCollection[$pCoordinate][$index]);
                }
            }
            if (count($this->conditionalStylesCollection[$pCoordinate]) == 0)
            {
                unset($this->conditionalStylesCollection[$pCoordinate]);
            }
        }
    }

    public function updateConditionalStyles($oldCoordinates,$newCoordinates)
    {
        if (isset($this->conditionalStylesCollection[$oldCoordinates]))
        {
            $conditions = $this->conditionalStylesCollection[$oldCoordinates];
            foreach($conditions as $condition)
            {
                // update the coordinates in each conditional
                $condition->updateCellReference($oldCoordinates, $newCoordinates);
            }
            // also update the index of the conditions collection so the relationship is not lost
            unset($this->conditionalStylesCollection[$oldCoordinates]);
            $this->conditionalStylesCollection[$newCoordinates] = $conditions;
        }
    }

    /**
     * Get collection of conditional styles
     *
     * @return array
     */
    public function getConditionalStylesCollection()
    {
        return $this->conditionalStylesCollection;
    }

    /**
     * Set conditional styles
     *
     * @param $pCoordinate string E.g. 'A1'
     * @param $pValue PHPExceller_Style_Conditional
     * @return void
     */
    public function setConditionalStyles($pCoordinate = 'A1', $pValue)
    {
        $this->conditionalStylesCollection[strtoupper($pCoordinate)] = $pValue;
    }

    public function addConditionalStyle($pCoordinate = 'A1', $pValue)
    {
        if (isset($this->conditionalStylesCollection[$pCoordinate]))
        {
            $existing_items = $this->conditionalStylesCollection[$pCoordinate];
            $existing_items[] = $pValue;
            $this->conditionalStylesCollection[$pCoordinate] = $existing_items;
        }
        else
        {
            $this->conditionalStylesCollection[$pCoordinate] = array($pValue);
        }
    }

    /**
     * Get style for cell by using numeric cell coordinates
     *
     * @param int $pColumn  Numeric column coordinate of the cell
     * @param int $pRow Numeric row coordinate of the cell
     * @param int pColumn2 Numeric column coordinate of the range cell
     * @param int pRow2 Numeric row coordinate of the range cell
     * @return PHPExceller_Style
     */
    public function getStyleByColumnAndRow($pColumn = 0, $pRow = 1, $pColumn2 = null, $pRow2 = null)
    {
        if (!is_null($pColumn2) && !is_null($pRow2)) {
            $cellRange = PHPExceller_Cell::stringFromColumnIndex($pColumn) . $pRow . ':' .
                PHPExceller_Cell::stringFromColumnIndex($pColumn2) . $pRow2;
            return $this->getStyle($cellRange);
        }

        if (!is_numeric($pColumn))
        {
            return $this->getStyle($pColumn . $pRow);
        }
        else
        {
            return $this->getStyle(PHPExceller_Cell::stringFromColumnIndex($pColumn) . $pRow);
        }
    }

    /**
     * Set shared cell style to a range of cells
     *
     * Please note that this will overwrite existing cell styles for cells in range!
     *
     * @deprecated
     * @param PHPExceller_Style $pSharedCellStyle Cell style to share
     * @param string $pRange Range of cells (i.e. "A1:B10"), or just one cell (i.e. "A1")
     * @throws PHPExceller_Exception
     * @return void
     */
    public function setSharedStyle(PHPExceller_Style $pSharedCellStyle = null, $pRange = '')
    {
        $this->duplicateStyle($pSharedCellStyle, $pRange);
    }

    /**
     * Duplicate cell style to a range of cells
     *
     * Please note that this will overwrite existing cell styles for cells in range!
     *
     * @param PHPExceller_Style $pCellStyle Cell style to duplicate
     * @param string $pRange Range of cells (i.e. "A1:B10"), or just one cell (i.e. "A1")
     * @throws PHPExceller_Exception
     * @return void
     */
    public function duplicateStyle(PHPExceller_Style $pCellStyle = null, $pRange = '')
    {
        // make sure we have a real style and not supervisor
        $style = $pCellStyle->getIsSupervisor() ? $pCellStyle->getSharedComponent() : $pCellStyle;

        // Add the style to the workbook if necessary
        $workbook = $this->parent;
        if ($existingStyle = $this->parent->getCellXfByHashCode($pCellStyle->getHashCode())) {
            // there is already such cell Xf in our collection
            $xfIndex = $existingStyle->getIndex();
        } else {
            // we don't have such a cell Xf, need to add
            $workbook->addCellXf($pCellStyle);
            $xfIndex = $pCellStyle->getIndex();
        }

        // Calculate range outer borders
        list($rangeStart, $rangeEnd) = PHPExceller_Cell::rangeBoundaries($pRange . ':' . $pRange);

        // Make sure we can loop upwards on rows and columns
        if ($rangeStart[0] > $rangeEnd[0] && $rangeStart[1] > $rangeEnd[1]) {
            $tmp = $rangeStart;
            $rangeStart = $rangeEnd;
            $rangeEnd = $tmp;
        }

        // Loop through cells and apply styles
        for ($col = $rangeStart[0]; $col <= $rangeEnd[0]; ++$col) {
            for ($row = $rangeStart[1]; $row <= $rangeEnd[1]; ++$row) {
                $this->getCell(PHPExceller_Cell::stringFromColumnIndex($col - 1) . $row)->setXfIndex($xfIndex);
            }
        }
    }

    /**
     * Duplicate conditional style to a range of cells
     *
     * Please note that this will overwrite existing cell styles for cells in range!
     *
     * @param    array of PHPExceller_Style_Conditional    $pCellStyle    Cell style to duplicate
     * @param string $pRange Range of cells (i.e. "A1:B10"), or just one cell (i.e. "A1")
     * @throws PHPExceller_Exception
     * @return void
     */
    public function duplicateConditionalStyle(array $pCellStyle = null, $pRange = '')
    {
        foreach ($pCellStyle as $cellStyle) {
            if (!($cellStyle instanceof PHPExceller_Style_Conditional)) {
                throw new PHPExceller_Exception('Style is not a conditional style');
            }
        }

        // Calculate range outer borders
        list($rangeStart, $rangeEnd) = PHPExceller_Cell::rangeBoundaries($pRange . ':' . $pRange);

        // Make sure we can loop upwards on rows and columns
        if ($rangeStart[0] > $rangeEnd[0] && $rangeStart[1] > $rangeEnd[1]) {
            $tmp = $rangeStart;
            $rangeStart = $rangeEnd;
            $rangeEnd = $tmp;
        }

        // Loop through cells and apply styles
        for ($col = $rangeStart[0]; $col <= $rangeEnd[0]; ++$col) {
            for ($row = $rangeStart[1]; $row <= $rangeEnd[1]; ++$row) {
                $this->setConditionalStyles(PHPExceller_Cell::stringFromColumnIndex($col - 1) . $row, $pCellStyle);
            }
        }
    }

    /**
     * Duplicate cell style array to a range of cells
     *
     * Please note that this will overwrite existing cell styles for cells in range,
     * if they are in the styles array. For example, if you decide to set a range of
     * cells to font bold, only include font bold in the styles array.
     *
     * @deprecated
     * @param array $pStyles Array containing style information
     * @param string $pRange Range of cells (i.e. "A1:B10"), or just one cell (i.e. "A1")
     * @param boolean $pAdvanced Advanced mode for setting borders.
     * @throws PHPExceller_Exception
     * @return void
     */
    public function duplicateStyleArray($pStyles = null, $pRange = '', $pAdvanced = true)
    {
        $this->getStyle($pRange)->applyFromArray($pStyles, $pAdvanced);
    }

    /**
     * Set break on a cell
     *
     * @param string $pCell Cell coordinate (e.g. A1)
     * @param int $pBreak Break type (type of PHPExceller_Worksheet::BREAK_*)
     * @throws PHPExceller_Exception
     * @return void
     */
    public function setBreak($pCell = 'A1', $pBreak = PHPExceller_Worksheet::BREAK_NONE)
    {
        // Uppercase coordinate
        $pCell = strtoupper($pCell);

        if ($pCell != '') {
            if ($pBreak == PHPExceller_Worksheet::BREAK_NONE) {
                if (isset($this->breaks[$pCell])) {
                    unset($this->breaks[$pCell]);
                }
            } else {
                $this->breaks[$pCell] = $pBreak;
            }
        } else {
            throw new PHPExceller_Exception('No cell coordinate specified.');
        }
    }

    /**
     * Set break on a cell by using numeric cell coordinates
     *
     * @param integer $pColumn Numeric column coordinate of the cell
     * @param integer $pRow Numeric row coordinate of the cell
     * @param  integer $pBreak Break type (type of PHPExceller_Worksheet::BREAK_*)
     * @return void
     */
    public function setBreakByColumnAndRow($pColumn = 0, $pRow = 1, $pBreak = PHPExceller_Worksheet::BREAK_NONE)
    {
        $this->setBreak(PHPExceller_Cell::stringFromColumnIndex($pColumn) . $pRow, $pBreak);
    }

    /**
     * Get breaks
     *
     * @return array[]
     */
    public function getBreaks()
    {
        return $this->breaks;
    }

    /**
     * Set merge on a cell range
     *
     * @param string $pRange  Cell range (e.g. A1:E1)
     * @throws PHPExceller_Exception
     * @return void
     */
    public function mergeCells($pRange = 'A1:A1')
    {
        // Uppercase coordinate
        $pRange = strtoupper($pRange);

        if (strpos($pRange, ':') !== false) {
            $this->mergeCells[$pRange] = $pRange;

            // make sure cells are created

            // get the cells in the range
            $aReferences = PHPExceller_Cell::extractAllCellReferencesInRange($pRange);

            // create upper left cell if it does not already exist
            $upperLeft = $aReferences[0];
            if (!$this->cellExists($upperLeft)) {
                $this->getCell($upperLeft)->setValueExplicit(null, PHPExceller_Cell_DataType::TYPE_NULL);
            }

            // Blank out the rest of the cells in the range (if they exist)
            $count = count($aReferences);
            for ($i = 1; $i < $count; $i++) {
                if ($this->cellExists($aReferences[$i])) {
                    $this->getCell($aReferences[$i])->setValueExplicit(null, PHPExceller_Cell_DataType::TYPE_NULL);
                }
            }
        } else {
            throw new PHPExceller_Exception('Merge must be set on a range of cells.');
        }
    }

    /**
     * Set merge on a cell range by using numeric cell coordinates
     *
     * @param int $pColumn1    Numeric column coordinate of the first cell
     * @param int $pRow1        Numeric row coordinate of the first cell
     * @param int $pColumn2    Numeric column coordinate of the last cell
     * @param int $pRow2        Numeric row coordinate of the last cell
     * @throws    PHPExceller_Exception
     * @return void
     */
    public function mergeCellsByColumnAndRow($pColumn1 = 0, $pRow1 = 1, $pColumn2 = 0, $pRow2 = 1)
    {
        $cellRange = PHPExceller_Cell::stringFromColumnIndex($pColumn1) . $pRow1 . ':' . PHPExceller_Cell::stringFromColumnIndex($pColumn2) . $pRow2;
        $this->mergeCells($cellRange);
    }

    /**
     * Remove merge on a cell range
     *
     * @param    string            $pRange        Cell range (e.g. A1:E1)
     * @throws    PHPExceller_Exception
     * @return void
     */
    public function unmergeCells($pRange = 'A1:A1')
    {
        // Uppercase coordinate
        $pRange = strtoupper($pRange);

        if (strpos($pRange, ':') !== false) {
            if (isset($this->mergeCells[$pRange])) {
                unset($this->mergeCells[$pRange]);
            } else {
                throw new PHPExceller_Exception('Cell range ' . $pRange . ' not known as merged.');
            }
        } else {
            throw new PHPExceller_Exception('Merge can only be removed from a range of cells.');
        }
    }

    /**
     * Remove merge on a cell range by using numeric cell coordinates
     *
     * @param int $pColumn1    Numeric column coordinate of the first cell
     * @param int $pRow1        Numeric row coordinate of the first cell
     * @param int $pColumn2    Numeric column coordinate of the last cell
     * @param int $pRow2        Numeric row coordinate of the last cell
     * @throws    PHPExceller_Exception
     * @return void
     */
    public function unmergeCellsByColumnAndRow($pColumn1 = 0, $pRow1 = 1, $pColumn2 = 0, $pRow2 = 1)
    {
        $cellRange = PHPExceller_Cell::stringFromColumnIndex($pColumn1) . $pRow1 . ':' . PHPExceller_Cell::stringFromColumnIndex($pColumn2) . $pRow2;
        $this->unmergeCells($cellRange);
    }

    /**
     * Get merge cells array.
     *
     * @return array[]
     */
    public function getMergeCells()
    {
        return $this->mergeCells;
    }

    /**
     * Set merge cells array for the entire sheet. Use instead mergeCells() to merge
     * a single cell range.
     *
     * @param array
     * @return void
     */
    public function setMergeCells($pValue = array())
    {
        $this->mergeCells = $pValue;
    }

    /**
     * Set protection on a cell range
     *
     * @param    string            $pRange                Cell (e.g. A1) or cell range (e.g. A1:E1)
     * @param    string            $pPassword            Password to unlock the protection
     * @param    boolean        $pAlreadyHashed    If the password has already been hashed, set this to true
     * @throws    PHPExceller_Exception
     * @return void
     */
    public function protectCells($pRange = 'A1', $pPassword = '', $pAlreadyHashed = false)
    {
        // Uppercase coordinate
        $pRange = strtoupper($pRange);

        if (!$pAlreadyHashed) {
            $pPassword = PHPExceller_Shared_PasswordHasher::hashPassword($pPassword);
        }
        $this->protectedCells[$pRange] = $pPassword;
    }

    /**
     * Set protection on a cell range by using numeric cell coordinates
     *
     * @param int  $pColumn1            Numeric column coordinate of the first cell
     * @param int  $pRow1                Numeric row coordinate of the first cell
     * @param int  $pColumn2            Numeric column coordinate of the last cell
     * @param int  $pRow2                Numeric row coordinate of the last cell
     * @param string $pPassword            Password to unlock the protection
     * @param    boolean $pAlreadyHashed    If the password has already been hashed, set this to true
     * @throws    PHPExceller_Exception
     * @return void
     */
    public function protectCellsByColumnAndRow($pColumn1 = 0, $pRow1 = 1, $pColumn2 = 0, $pRow2 = 1, $pPassword = '', $pAlreadyHashed = false)
    {
        $cellRange = PHPExceller_Cell::stringFromColumnIndex($pColumn1) . $pRow1 . ':' . PHPExceller_Cell::stringFromColumnIndex($pColumn2) . $pRow2;
        $this->protectCells($cellRange, $pPassword, $pAlreadyHashed);
    }

    /**
     * Remove protection on a cell range
     *
     * @param    string            $pRange        Cell (e.g. A1) or cell range (e.g. A1:E1)
     * @throws    PHPExceller_Exception
     * @return void
     */
    public function unprotectCells($pRange = 'A1')
    {
        // Uppercase coordinate
        $pRange = strtoupper($pRange);

        if (isset($this->protectedCells[$pRange])) {
            unset($this->protectedCells[$pRange]);
        } else {
            throw new PHPExceller_Exception('Cell range ' . $pRange . ' not known as protected.');
        }
    }

    /**
     * Remove protection on a cell range by using numeric cell coordinates
     *
     * @param int  $pColumn1            Numeric column coordinate of the first cell
     * @param int  $pRow1                Numeric row coordinate of the first cell
     * @param int  $pColumn2            Numeric column coordinate of the last cell
     * @param int $pRow2                Numeric row coordinate of the last cell
     * @param string $pPassword            Password to unlock the protection
     * @param    boolean $pAlreadyHashed    If the password has already been hashed, set this to true
     * @throws    PHPExceller_Exception
     * @return void
     */
    public function unprotectCellsByColumnAndRow($pColumn1 = 0, $pRow1 = 1, $pColumn2 = 0, $pRow2 = 1, $pPassword = '', $pAlreadyHashed = false)
    {
        $cellRange = PHPExceller_Cell::stringFromColumnIndex($pColumn1) . $pRow1 . ':' . PHPExceller_Cell::stringFromColumnIndex($pColumn2) . $pRow2;
        $this->unprotectCells($cellRange, $pPassword, $pAlreadyHashed);
    }

    /**
     * Get protected cells
     *
     * @return array[]
     */
    public function getProtectedCells()
    {
        return $this->protectedCells;
    }

    /**
     *    Get Autofilter
     *
     *    @return PHPExceller_Worksheet_AutoFilter
     */
    public function getAutoFilter()
    {
        return $this->autoFilter;
    }

    /**
     *    Set AutoFilter
     *
     *    @param    PHPExceller_Worksheet_AutoFilter|string   $pValue
     *            A simple string containing a Cell range like 'A1:E10' is permitted for backward compatibility
     *    @throws    PHPExceller_Exception
     *    @return void
     */
    public function setAutoFilter($pValue)
    {
        $pRange = strtoupper($pValue);
        if (is_string($pValue)) {
            $this->autoFilter->setRange($pValue);
        } elseif (is_object($pValue) && ($pValue instanceof PHPExceller_Worksheet_AutoFilter)) {
            $this->autoFilter = $pValue;
        }
    }

    /**
     *    Set Autofilter Range by using numeric cell coordinates
     *
     *    @param  integer  $pColumn1    Numeric column coordinate of the first cell
     *    @param  integer  $pRow1       Numeric row coordinate of the first cell
     *    @param  integer  $pColumn2    Numeric column coordinate of the second cell
     *    @param  integer  $pRow2       Numeric row coordinate of the second cell
     *    @throws    PHPExceller_Exception
     *    @return void
     */
    public function setAutoFilterByColumnAndRow($pColumn1 = 0, $pRow1 = 1, $pColumn2 = 0, $pRow2 = 1)
    {
        $this->setAutoFilter(
            PHPExceller_Cell::stringFromColumnIndex($pColumn1) . $pRow1
            . ':' .
            PHPExceller_Cell::stringFromColumnIndex($pColumn2) . $pRow2
        );
    }

    /**
     * Remove autofilter
     *
     * @return void
     */
    public function removeAutoFilter()
    {
        $this->autoFilter->setRange(null);
    }

    /**
     * Get Freeze Pane
     *
     * @return string
     */
    public function getFreezePane()
    {
        return $this->freezePane;
    }

    /**
     * Freeze Pane
     *
     * @param    string        $pCell        Cell (i.e. A2)
     *                                    Examples:
     *                                        A2 will freeze the rows above cell A2 (i.e row 1)
     *                                        B1 will freeze the columns to the left of cell B1 (i.e column A)
     *                                        B2 will freeze the rows above and to the left of cell A2
     *                                            (i.e row 1 and column A)
     * @throws    PHPExceller_Exception
     * @return void
     */
    public function freezePane($pCell = '')
    {
        // Uppercase coordinate
        $pCell = strtoupper($pCell);
        if (strpos($pCell, ':') === false && strpos($pCell, ',') === false) {
            $this->freezePane = $pCell;
        } else {
            throw new PHPExceller_Exception('Freeze pane can not be set on a range of cells.');
        }
    }

    /**
     * Freeze Pane by using numeric cell coordinates
     *
     * @param int $pColumn    Numeric column coordinate of the cell
     * @param int $pRow        Numeric row coordinate of the cell
     * @throws    PHPExceller_Exception
     * @return void
     */
    public function freezePaneByColumnAndRow($pColumn = 0, $pRow = 1)
    {
        $this->freezePane(PHPExceller_Cell::stringFromColumnIndex($pColumn) . $pRow);
    }

    /**
     * Unfreeze Pane
     *
     * @return void
     */
    public function unfreezePane()
    {
        $this->freezePane('');
    }

    /**
     * Insert a new row, updating all possible related data
     *
     * @param int $pBefore    Insert before this one
     * @param int $pNumRows    Number of rows to insert
     * @throws    PHPExceller_Exception
     * @return void
     */
    public function insertNewRowBefore($pBefore = 1, $pNumRows = 1)
    {
        if ($pBefore >= 1) {
            $objReferenceHelper = PHPExceller_ReferenceHelper::getInstance();
            $objReferenceHelper->insertNewBefore('A' . $pBefore, 0, $pNumRows, $this);
        } else {
            throw new PHPExceller_Exception("Rows can only be inserted before at least row 1.");
        }
    }

    /**
     * Insert a new column, updating all possible related data
     *
     * @param int $pBefore    Insert before this one
     * @param int $pNumCols    Number of columns to insert
     * @throws    PHPExceller_Exception
     * @return void
     */
    public function insertNewColumnBefore($pBefore = 'A', $pNumCols = 1)
    {
        if (!is_numeric($pBefore)) {
            $objReferenceHelper = PHPExceller_ReferenceHelper::getInstance();
            $objReferenceHelper->insertNewBefore($pBefore . '1', $pNumCols, 0, $this);
        } else {
            throw new PHPExceller_Exception("Column references should not be numeric.");
        }
    }

    /**
     * Insert a new column, updating all possible related data
     *
     * @param int $pBefore    Insert before this one (numeric column coordinate of the cell)
     * @param int $pNumCols    Number of columns to insert
     * @throws    PHPExceller_Exception
     * @return void
     */
    public function insertNewColumnBeforeByIndex($pBefore = 0, $pNumCols = 1)
    {
        if ($pBefore >= 0) {
            $this->insertNewColumnBefore(PHPExceller_Cell::stringFromColumnIndex($pBefore), $pNumCols);
        } else {
            throw new PHPExceller_Exception("Columns can only be inserted before at least column A (0).");
        }
    }

    /**
     * Delete a row, updating all possible related data
     *
     * @param int $pRow        Remove starting with this one
     * @param int $pNumRows    Number of rows to remove
     * @throws    PHPExceller_Exception
     * @return void
     */
    public function removeRow($pRow = 1, $pNumRows = 1)
    {
        if ($pRow >= 1) {
            $highestRow = $this->getHighestDataRow();
            $objReferenceHelper = PHPExceller_ReferenceHelper::getInstance();
            $objReferenceHelper->insertNewBefore('A' . ($pRow + $pNumRows), 0, -$pNumRows, $this);
            for ($r = 0; $r < $pNumRows; ++$r) {
                $this->getCellCacheController()->removeRow($highestRow);
                --$highestRow;
            }
        } else {
            throw new PHPExceller_Exception("Rows to be deleted should at least start from row 1.");
        }
    }

    /**
     * Remove a column, updating all possible related data
     *
     * @param string    $pColumn     Remove starting with this one
     * @param int       $pNumCols    Number of columns to remove
     * @throws    PHPExceller_Exception
     * @return void
     */
    public function removeColumn($pColumn = 'A', $pNumCols = 1)
    {
        if (!is_numeric($pColumn)) {
            $highestColumn = $this->getHighestDataColumn();
            $pColumn = PHPExceller_Cell::stringFromColumnIndex(PHPExceller_Cell::columnIndexFromString($pColumn) - 1 + $pNumCols);
            $objReferenceHelper = PHPExceller_ReferenceHelper::getInstance();
            $objReferenceHelper->insertNewBefore($pColumn . '1', -$pNumCols, 0, $this);
            for ($c = 0; $c < $pNumCols; ++$c) {
                $this->getCellCacheController()->removeColumn($highestColumn);
                $highestColumn = PHPExceller_Cell::stringFromColumnIndex(PHPExceller_Cell::columnIndexFromString($highestColumn) - 2);
            }
        } else {
            throw new PHPExceller_Exception("Column references should not be numeric.");
        }
    }

    /**
     * Remove a column, updating all possible related data
     *
     * @param int $pColumn    Remove starting with this one (numeric column coordinate of the cell)
     * @param int $pNumCols    Number of columns to remove
     * @throws    PHPExceller_Exception
     * @return void
     */
    public function removeColumnByIndex($pColumn = 0, $pNumCols = 1)
    {
        if ($pColumn >= 0) {
            $this->removeColumn(PHPExceller_Cell::stringFromColumnIndex($pColumn), $pNumCols);
        } else {
            throw new PHPExceller_Exception("Columns to be deleted should at least start from column 0");
        }
    }

    /**
     * Show gridlines?
     *
     * @return boolean
     */
    public function getShowGridlines()
    {
        return $this->showGridlines;
    }

    /**
     * Set show gridlines
     *
     * @param boolean $pValue    Show gridlines (true/false)
     * @return void
     */
    public function setShowGridlines($pValue = false)
    {
        $this->showGridlines = $pValue;
    }

    /**
    * Print gridlines?
    *
    * @return boolean
    */
    public function getPrintGridlines()
    {
        return $this->printGridlines;
    }

    /**
    * Set print gridlines
    *
    * @param boolean $pValue Print gridlines (true/false)
    * @return void
    */
    public function setPrintGridlines($pValue = false)
    {
        $this->printGridlines = $pValue;
    }

    /**
    * Show row and column headers?
    *
    * @return boolean
    */
    public function getShowRowColHeaders()
    {
        return $this->showRowColHeaders;
    }

    /**
    * Set show row and column headers
    *
    * @param boolean $pValue Show row and column headers (true/false)
    * @return void
    */
    public function setShowRowColHeaders($pValue = false)
    {
        $this->showRowColHeaders = $pValue;
    }

    /**
     * Show summary below? (Row/Column outlining)
     *
     * @return boolean
     */
    public function getShowSummaryBelow()
    {
        return $this->showSummaryBelow;
    }

    /**
     * Set show summary below
     *
     * @param boolean $pValue    Show summary below (true/false)
     * @return void
     */
    public function setShowSummaryBelow($pValue = true)
    {
        $this->showSummaryBelow = $pValue;
    }

    /**
     * Show summary right? (Row/Column outlining)
     *
     * @return boolean
     */
    public function getShowSummaryRight()
    {
        return $this->showSummaryRight;
    }

    /**
     * Set show summary right
     *
     * @param boolean $pValue    Show summary right (true/false)
     * @return void
     */
    public function setShowSummaryRight($pValue = true)
    {
        $this->showSummaryRight = $pValue;
    }

    /**
     * Get comments
     *
     * @return PHPExceller_Comment[]
     */
    public function getComments()
    {
        return $this->comments;
    }

    /**
     * Set comments array for the entire sheet.
     *
     * @param array of PHPExceller_Comment
     * @return void
     */
    public function setComments($pValue = array())
    {
        $this->comments = $pValue;
    }

    /**
     * Get comment for cell
     *
     * @param string $pCellCoordinate    Cell coordinate to get comment for
     * @return PHPExceller_Comment
     * @throws PHPExceller_Exception
     */
    public function getComment($pCellCoordinate = 'A1')
    {
        // Uppercase coordinate
        $pCellCoordinate = strtoupper($pCellCoordinate);

        if (strpos($pCellCoordinate, ':') !== false || strpos($pCellCoordinate, ',') !== false) {
            throw new PHPExceller_Exception('Cell coordinate string can not be a range of cells.');
        } elseif (strpos($pCellCoordinate, '$') !== false) {
            throw new PHPExceller_Exception('Cell coordinate string must not be absolute.');
        } elseif ($pCellCoordinate == '') {
            throw new PHPExceller_Exception('Cell coordinate can not be zero-length string.');
        } else {
            // Check if we already have a comment for this cell.
            // If not, create a new comment.
            if (isset($this->comments[$pCellCoordinate])) {
                return $this->comments[$pCellCoordinate];
            } else {
                $newComment = new PHPExceller_Comment();
                $this->comments[$pCellCoordinate] = $newComment;
                return $newComment;
            }
        }
    }

    /**
     * Get comment for cell by using numeric cell coordinates
     *
     * @param int $pColumn    Numeric column coordinate of the cell
     * @param int $pRow        Numeric row coordinate of the cell
     * @return PHPExceller_Comment
     */
    public function getCommentByColumnAndRow($pColumn = 0, $pRow = 1)
    {
        return $this->getComment(PHPExceller_Cell::stringFromColumnIndex($pColumn) . $pRow);
    }

    /**
     * Get selected cell
     *
     * @deprecated
     * @return string
     */
    public function getSelectedCell()
    {
        return $this->getSelectedCells();
    }

    /**
     * Get active cell
     *
     * @return string Example: 'A1'
     */
    public function getActiveCell()
    {
        return $this->activeCell;
    }

    /**
     * Get selected cells
     *
     * @return string
     */
    public function getSelectedCells()
    {
        return $this->selectedCells;
    }

    /**
     * Selected cell
     *
     * @param    string        $pCoordinate    Cell (i.e. A1)
     * @return void
     */
    public function setSelectedCell($pCoordinate = 'A1')
    {
        $this->setSelectedCells($pCoordinate);
    }

    /**
     * Select a range of cells.
     *
     * @param    string        $pCoordinate    Cell range, examples: 'A1', 'B2:G5', 'A:C', '3:6'
     * @throws    PHPExceller_Exception
     * @return void
     */
    public function setSelectedCells($pCoordinate = 'A1')
    {
        // Uppercase coordinate
        $pCoordinate = strtoupper($pCoordinate);

        // Convert 'A' to 'A:A'
        $pCoordinate = preg_replace('/^([A-Z]+)$/', '${1}:${1}', $pCoordinate);

        // Convert '1' to '1:1'
        $pCoordinate = preg_replace('/^([0-9]+)$/', '${1}:${1}', $pCoordinate);

        // Convert 'A:C' to 'A1:C1048576'
        $pCoordinate = preg_replace('/^([A-Z]+):([A-Z]+)$/', '${1}1:${2}1048576', $pCoordinate);

        // Convert '1:3' to 'A1:XFD3'
        $pCoordinate = preg_replace('/^([0-9]+):([0-9]+)$/', 'A${1}:XFD${2}', $pCoordinate);

        if (strpos($pCoordinate, ':') !== false || strpos($pCoordinate, ',') !== false) {
            list($first, ) = PHPExceller_Cell::splitRange($pCoordinate);
            $this->activeCell = $first[0];
        } else {
            $this->activeCell = $pCoordinate;
        }
        $this->selectedCells = $pCoordinate;
    }

    /**
     * Selected cell by using numeric cell coordinates
     *
     * @param int $pColumn Numeric column coordinate of the cell
     * @param int $pRow Numeric row coordinate of the cell
     * @throws PHPExceller_Exception
     * @return void
     */
    public function setSelectedCellByColumnAndRow($pColumn = 0, $pRow = 1)
    {
        $this->setSelectedCells(PHPExceller_Cell::stringFromColumnIndex($pColumn) . $pRow);
    }

    /**
     * Get right-to-left
     *
     * @return boolean
     */
    public function getRightToLeft()
    {
        return $this->rightToLeft;
    }

    /**
     * Set right-to-left
     *
     * @param boolean $value    Right-to-left true/false
     * @return void
     */
    public function setRightToLeft($value = false)
    {
        $this->rightToLeft = $value;
    }

    /**
     * Fill worksheet from values in array
     *
     * @param array $source Source array
     * @param mixed $nullValue Value in source array that stands for blank cell
     * @param string $startCell Insert array starting from this cell address as the top left coordinate
     * @param boolean $strictNullComparison Apply strict comparison when testing for null values in the array
     * @throws PHPExceller_Exception
     * @return void
     */
    public function fromArray($source = null, $nullValue = null, $startCell = 'A1', $strictNullComparison = false)
    {
        if (is_array($source)) {
            //    Convert a 1-D array to 2-D (for ease of looping)
            if (!is_array(end($source))) {
                $source = array($source);
            }

            // start coordinate
            list ($startColumn, $startRow) = PHPExceller_Cell::coordinateFromString($startCell);

            // Loop through $source
            foreach ($source as $rowData) {
                $currentColumn = $startColumn;
                foreach ($rowData as $cellValue) {
                    if ($strictNullComparison) {
                        if ($cellValue !== $nullValue) {
                            // Set cell value
                            $this->getCell($currentColumn . $startRow)->setValue($cellValue);
                        }
                    } else {
                        if ($cellValue != $nullValue) {
                            // Set cell value
                            $this->getCell($currentColumn . $startRow)->setValue($cellValue);
                        }
                    }
                    ++$currentColumn;
                }
                ++$startRow;
            }
        } else {
            throw new PHPExceller_Exception("Parameter \$source should be an array.");
        }
    }

    /**
     * Create array from a range of cells
     *
     * @param string $pRange Range of cells (i.e. "A1:B10"), or just one cell (i.e. "A1")
     * @param mixed $nullValue Value returned in the array entry if a cell doesn't exist
     * @param boolean $calculateFormulas Should formulas be calculated?
     * @param boolean $formatData Should formatting be applied to cell values?
     * @param boolean $returnCellRef False - Return a simple array of rows and columns indexed by number counting from zero
     *                               True - Return rows and columns indexed by their actual row and column IDs
     * @return array
     */
    public function rangeToArray($pRange = 'A1', $nullValue = null, $calculateFormulas = true, $formatData = true, $returnCellRef = false) {
        // Returnvalue
        $returnValue = array();
        //    Identify the range that we need to extract from the worksheet
        list($rangeStart, $rangeEnd) = PHPExceller_Cell::rangeBoundaries($pRange);
        $minCol = PHPExceller_Cell::stringFromColumnIndex($rangeStart[0] -1);
        $minRow = $rangeStart[1];
        $maxCol = PHPExceller_Cell::stringFromColumnIndex($rangeEnd[0] -1);
        $maxRow = $rangeEnd[1];

        $maxCol++;
        // Loop through rows
        $r = -1;
        for ($row = $minRow; $row <= $maxRow; ++$row) {
            $rRef = ($returnCellRef) ? $row : ++$r;
            $c = -1;
            // Loop through columns in the current row
            for ($col = $minCol; $col != $maxCol; ++$col) {
                $cRef = ($returnCellRef) ? $col : ++$c;
                //    Using getCell() will create a new cell if it doesn't already exist. We don't want that to happen
                //        so we test and retrieve directly against cellCollection
                if ($this->cellCollection->isDataSet($col.$row)) {
                    // Cell exists
                    $cell = $this->cellCollection->getCacheData($col.$row);
                    if ($cell->getValue() !== null) {
                        if ($cell->getValue() instanceof PHPExceller_RichText) {
                            $returnValue[$rRef][$cRef] = $cell->getValue()->getPlainText();
                        } else {
                            if ($calculateFormulas) {
                                $returnValue[$rRef][$cRef] = $cell->getCalculatedValue();
                            } else {
                                $returnValue[$rRef][$cRef] = $cell->getValue();
                            }
                        }

                        if ($formatData) {
                            $style = $this->parent->getCellXfByIndex($cell->getXfIndex());
                            $returnValue[$rRef][$cRef] = PHPExceller_Style_NumberFormat::toFormattedString(
                                $returnValue[$rRef][$cRef],
                                ($style && $style->getNumberFormat()) ?
                                    $style->getNumberFormat()->getFormatCode() :
                                    PHPExceller_Style_NumberFormat::FORMAT_GENERAL
                            );
                        }
                    } else {
                        // Cell holds a NULL
                        $returnValue[$rRef][$cRef] = $nullValue;
                    }
                } else {
                    // Cell doesn't exist
                    $returnValue[$rRef][$cRef] = $nullValue;
                }
            }
        }

        // Return
        return $returnValue;
    }


    /**
     * Create array from a range of cells
     *
     * @param  string $pNamedRange Name of the Named Range
     * @param  mixed  $nullValue Value returned in the array entry if a cell doesn't exist
     * @param  boolean $calculateFormulas  Should formulas be calculated?
     * @param  boolean $formatData  Should formatting be applied to cell values?
     * @param  boolean $returnCellRef False - Return a simple array of rows and columns indexed by number counting from zero
     *                                True - Return rows and columns indexed by their actual row and column IDs
     * @return array
     * @throws PHPExceller_Exception
     */
    public function namedRangeToArray($pNamedRange = '', $nullValue = null, $calculateFormulas = true, $formatData = true, $returnCellRef = false) {
        $namedRange = PHPExceller_NamedRange::resolveRange($pNamedRange, $this);
        if ($namedRange !== null) {
            $pWorkSheet = $namedRange->getWorksheet();
            $pCellRange = $namedRange->getRange();

            return $pWorkSheet->rangeToArray(    $pCellRange,
                                                $nullValue, $calculateFormulas, $formatData, $returnCellRef);
        }

        throw new PHPExceller_Exception('Named Range '.$pNamedRange.' does not exist.');
    }


    /**
     * Create array from worksheet
     *
     * @param mixed $nullValue Value returned in the array entry if a cell doesn't exist
     * @param boolean $calculateFormulas Should formulas be calculated?
     * @param boolean $formatData  Should formatting be applied to cell values?
     * @param boolean $returnCellRef False - Return a simple array of rows and columns indexed by number counting from zero
     *                               True - Return rows and columns indexed by their actual row and column IDs
     * @return array
     */
    public function toArray($nullValue = null, $calculateFormulas = true, $formatData = true, $returnCellRef = false) {
        // Garbage collect...
        $this->garbageCollect();

        //    Identify the range that we need to extract from the worksheet
        $maxCol = $this->getHighestColumn();
        $maxRow = $this->getHighestRow();
        // Return
        return $this->rangeToArray(    'A1:'.$maxCol.$maxRow,
                                    $nullValue, $calculateFormulas, $formatData, $returnCellRef);
    }

    /**
     * Get row iterator
     *
     * @param   integer   $startRow   The row number at which to start iterating
     * @param   integer   $endRow     The row number at which to stop iterating
     *
     * @return PHPExceller_Worksheet_RowIterator
     */
    public function getRowIterator($startRow = 1, $endRow = null) {
        return new PHPExceller_Worksheet_RowIterator($this, $startRow, $endRow);
    }

    /**
     * Get column iterator
     *
     * @param   string   $startColumn The column address at which to start iterating
     * @param   string   $endColumn   The column address at which to stop iterating
     *
     * @return PHPExceller_Worksheet_ColumnIterator
     */
    public function getColumnIterator($startColumn = 'A', $endColumn = null) {
        return new PHPExceller_Worksheet_ColumnIterator($this, $startColumn, $endColumn);
    }

    /**
     * Run PHPExceller garabage collector.
     *
     * @return void
     */
    public function garbageCollect() {
        // Flush cache
        $this->cellCollection->getCacheData('A1');
        // Build a reference table from images
//        $imageCoordinates = array();
//        $iterator = $this->getDrawingCollection()->getIterator();
//        while ($iterator->valid()) {
//            $imageCoordinates[$iterator->current()->getCoordinates()] = true;
//
//            $iterator->next();
//        }
//
        // Lookup highest column and highest row if cells are cleaned
        $colRow = $this->cellCollection->getHighestRowAndColumn();
        $highestRow = $colRow['row'];
        $highestColumn = PHPExceller_Cell::columnIndexFromString($colRow['column']);

        // Loop through column dimensions
        foreach ($this->columnDimensions as $dimension) {
            $highestColumn = max($highestColumn, PHPExceller_Cell::columnIndexFromString($dimension->getColumnIndex()));
        }

        // Loop through row dimensions
        foreach ($this->rowDimensions as $dimension) {
            $highestRow = max($highestRow, $dimension->getRowIndex());
        }

        // Cache values
        if ($highestColumn < 0) {
            $this->cachedHighestColumn = 'A';
        } else {
            $this->cachedHighestColumn = PHPExceller_Cell::stringFromColumnIndex(--$highestColumn);
        }
        $this->cachedHighestRow = $highestRow;

    }

    /**
     * Get hash code
     *
     * @return string    Hash code
     */
    public function getHashCode() {
        if ($this->dirty) {
            $this->hash = md5( $this->title .
                                $this->autoFilter .
                                ($this->protection->isProtectionEnabled() ? 't' : 'f') .
                                __CLASS__
                              );
            $this->dirty = false;
        }
        return $this->hash;
    }

    /**
     * Extract worksheet title from range.
     *
     * Example: extractSheetTitle("testSheet!A1") ==> 'A1'
     * Example: extractSheetTitle("'testSheet 1'!A1", true) ==> array('testSheet 1', 'A1');
     *
     * @param string $pRange    Range to extract title from
     * @param bool $returnRange    Return range? (see example)
     * @return mixed
     */
    public static function extractSheetTitle($pRange, $returnRange = false) {
        // Sheet title included?
        if (($sep = strpos($pRange, '!')) === false) {
            return '';
        }

        if ($returnRange) {
            return array(trim(substr($pRange, 0, $sep), "'"), substr($pRange, $sep + 1));
        }

        return substr($pRange, $sep + 1);
    }

    /**
     * Get hyperlink
     *
     * @param string $pCellCoordinate    Cell coordinate to get hyperlink for
     */
    public function getHyperlink($pCellCoordinate = 'A1')
    {
        // return hyperlink if we already have one
        if (isset($this->hyperlinkCollection[$pCellCoordinate])) {
            return $this->hyperlinkCollection[$pCellCoordinate];
        }

        // else create hyperlink
        $this->hyperlinkCollection[$pCellCoordinate] = new PHPExceller_Cell_Hyperlink();
        return $this->hyperlinkCollection[$pCellCoordinate];
    }

    /**
     * Set hyperlnk
     *
     * @param string $pCellCoordinate    Cell coordinate to insert hyperlink
     * @param    PHPExceller_Cell_Hyperlink    $pHyperlink
     * @return void
     */
    public function setHyperlink($pCellCoordinate = 'A1', PHPExceller_Cell_Hyperlink $pHyperlink = null)
    {
        if ($pHyperlink === null) {
            unset($this->hyperlinkCollection[$pCellCoordinate]);
        } else {
            $this->hyperlinkCollection[$pCellCoordinate] = $pHyperlink;
        }
    }

    /**
     * Hyperlink at a specific coordinate exists?
     *
     * @param string $pCoordinate
     * @return boolean
     */
    public function hyperlinkExists($pCoordinate = 'A1')
    {
        return isset($this->hyperlinkCollection[$pCoordinate]);
    }

    /**
     * Get collection of hyperlinks
     *
     * @return PHPExceller_Cell_Hyperlink[]
     */
    public function getHyperlinkCollection()
    {
        return $this->hyperlinkCollection;
    }

    /**
     * Get data validation
     *
     * @param string $pCellCoordinate Cell coordinate to get data validation for
     */
    public function getDataValidation($pCellCoordinate = 'A1')
    {
        // return data validation if we already have one
        if (isset($this->dataValidationCollection[$pCellCoordinate])) {
            return $this->dataValidationCollection[$pCellCoordinate];
        }

        // else create data validation
        $this->dataValidationCollection[$pCellCoordinate] = new PHPExceller_Cell_DataValidation();
        return $this->dataValidationCollection[$pCellCoordinate];
    }

    /**
     * Set data validation
     *
     * @param string $pCellCoordinate    Cell coordinate to insert data validation
     * @param    PHPExceller_Cell_DataValidation    $pDataValidation
     * @return void
     */
    public function setDataValidation($pCellCoordinate = 'A1', PHPExceller_Cell_DataValidation $pDataValidation = null)
    {
        if ($pDataValidation === null) {
            unset($this->dataValidationCollection[$pCellCoordinate]);
        } else {
            $this->dataValidationCollection[$pCellCoordinate] = $pDataValidation;
        }
    }

    /**
     * Data validation at a specific coordinate exists?
     *
     * @param string $pCoordinate
     * @return boolean
     */
    public function dataValidationExists($pCoordinate = 'A1')
    {
        return isset($this->dataValidationCollection[$pCoordinate]);
    }

    /**
     * Get collection of data validations
     *
     * @return PHPExceller_Cell_DataValidation[]
     */
    public function getDataValidationCollection()
    {
        return $this->dataValidationCollection;
    }

    /**
     * Accepts a range, returning it as a range that falls within the current highest row and column of the worksheet
     *
     * @param string $range
     * @return string Adjusted range value
     */
    public function shrinkRangeToFit($range) {
        $maxCol = $this->getHighestColumn();
        $maxRow = $this->getHighestRow();
        $maxCol = PHPExceller_Cell::columnIndexFromString($maxCol);

        $rangeBlocks = explode(' ', $range);
        foreach ($rangeBlocks as &$rangeSet) {
            $rangeBoundaries = PHPExceller_Cell::getRangeBoundaries($rangeSet);

            if (PHPExceller_Cell::columnIndexFromString($rangeBoundaries[0][0]) > $maxCol) {
                $rangeBoundaries[0][0] = PHPExceller_Cell::stringFromColumnIndex($maxCol);
            }
            if ($rangeBoundaries[0][1] > $maxRow) {
                $rangeBoundaries[0][1] = $maxRow;
            }
            if (PHPExceller_Cell::columnIndexFromString($rangeBoundaries[1][0]) > $maxCol) {
                $rangeBoundaries[1][0] = PHPExceller_Cell::stringFromColumnIndex($maxCol);
            }
            if ($rangeBoundaries[1][1] > $maxRow) {
                $rangeBoundaries[1][1] = $maxRow;
            }
            $rangeSet = $rangeBoundaries[0][0].$rangeBoundaries[0][1].':'.$rangeBoundaries[1][0].$rangeBoundaries[1][1];
        }
        unset($rangeSet);
        $stRange = implode(' ', $rangeBlocks);

        return $stRange;
    }

    /**
     * Get tab color
     *
     * @return PHPExceller_Style_Color
     */
    public function getTabColor()
    {
        if ($this->tabColor === null) {
            $this->tabColor = new PHPExceller_Style_Color();
        }
        return $this->tabColor;
    }

    /**
     * Reset tab color
     *
     * @return void
     */
    public function resetTabColor()
    {
        $this->tabColor = null;
        unset($this->tabColor);
    }

    /**
     * Tab color set?
     *
     * @return boolean
     */
    public function isTabColorSet()
    {
        return ($this->tabColor !== null);
    }

    /**
     * Copy worksheet (!= clone!)
     *
     * @return PHPExceller_Worksheet
     */
    public function copy() {
        $copied = clone $this;

        return $copied;
    }

    /**
     * Implement PHP __clone to create a deep clone, not just a shallow copy.
     */
    public function __clone() {
        foreach ($this as $key => $val) {
            if ($key == 'parent') {
                continue;
            }

            if (is_object($val) || (is_array($val))) {
                if ($key == 'cellCollection') {
                    $newCollection = clone $this->cellCollection;
                    $newCollection->copyCellCollection($this);
                    $this->cellCollection = $newCollection;
                } elseif ($key == 'drawingCollection') {
                    $newCollection = clone $this->drawingCollection;
                    $this->drawingCollection = $newCollection;
                } elseif (($key == 'autoFilter') && ($this->autoFilter instanceof PHPExceller_Worksheet_AutoFilter)) {
                    $newAutoFilter = clone $this->autoFilter;
                    $this->autoFilter = $newAutoFilter;
                    $this->autoFilter->setParent($this);
                } else {
                    $this->{$key} = unserialize(serialize($val));
                }
            }
        }
    }

    /**
     * Define the code name of the sheet
     *
     * @param null|string Same rule as Title minus space not allowed (but, like Excel, change silently space to underscore)
     * @return void
     * @throws PHPExceller_Exception
    */
    public function setCodeName($pValue=null){
        // Is this a 'rename' or not?
        if ($this->getCodeName() == $pValue)
        {
            // nothing to do here
            return;
        }
        $pValue = str_replace(' ', '_', $pValue);//Excel does this automatically without flinching, we are doing the same
        // Syntax check
        // throw an exception if not valid
        self::checkSheetCodeName($pValue);

        // We use the same code that setTitle to find a valid codeName else not using a space (Excel don't like) but a '_'
        if ($this->getParent()) {
            // Is there already such sheet name?
            if ($this->getParent()->sheetCodeNameExists($pValue)) {
                // Use name, but append with lowest possible integer

                if (PHPExceller_Shared_String::CountCharacters($pValue) > 29) {
                    $pValue = PHPExceller_Shared_String::Substring($pValue,0,29);
                }
                $i = 1;
                while ($this->getParent()->sheetCodeNameExists($pValue . '_' . $i)) {
                    ++$i;
                    if ($i == 10) {
                        if (PHPExceller_Shared_String::CountCharacters($pValue) > 28) {
                            $pValue = PHPExceller_Shared_String::Substring($pValue,0,28);
                        }
                    } elseif ($i == 100) {
                        if (PHPExceller_Shared_String::CountCharacters($pValue) > 27) {
                            $pValue = PHPExceller_Shared_String::Substring($pValue,0,27);
                        }
                    }
                }

                $pValue = $pValue . '_' . $i;// ok, we have a valid name
                //codeName is'nt used in formula : no need to call for an update
                //return $this->setTitle($altTitle,$updateFormulaCellReferences);
            }
        }
        // set name
        $this->codeName=$pValue;
    }
    /**
     * Return the code name of the sheet
     *
     * @return null|string
    */
    public function getCodeName(){
        return $this->codeName;
    }
    /**
     * Sheet has a code name ?
     * @return boolean
    */
    public function hasCodeName(){
        return !(is_null($this->codeName));
    }

    /**
     * Invalid characters in sheet title
     *
     * @var array
     */
    private static $invalidCharacters = array('*', ':', '/', '\\', '?', '[', ']');

}

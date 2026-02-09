<?php declare(strict_types=1);
/**
 * @author    Jacques Marneweck <jacques@siberia.co.za>
 * @copyright 2018-2026 Jacques Marneweck.  All rights strictly reserved.
 */

namespace Jacques\Reports;

use Jacques\Reports\Traits\{
    Borders, Colour, Margins
};
use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

/**
 * @method fromArray(array $source, $nullValue, $startCell, bool $strictNullComparison)
 * @method getActiveSheet()
 * @method getColumnDimension()
 * @method getActiveSheetIndex()
 * @method getHighestColumn($row)
 * @method getHighestDataColumn($row)
 * @method getHighestRow($column)
 * @method getHighestDataRow($column)
 * @method getHighestRowAndColumn()
 * @method getDefaultStyle()
 * @method getStyles()
 * @method getStyle($cellCoordinate)
 * @method getStyleByColumnAndRow($columnIndex1, $row1, $columnIndex2 = null, $row2 = null)
 * @method setCellValue(string $coordinate, $value)
 * @method setCellValueByColumnAndRow(int $columnIndex, int $row, $value)
 */
class Excel
{
    use Borders;
    use Colour;
    use Margins;

    /**
     * @var \PhpOffice\PhpSpreadsheet\Spreadsheet
     */
    private \PhpOffice\PhpSpreadsheet\Spreadsheet $spreadsheet;

    /**
     * @var \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
     */
    private \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet $sheet;

    /**
     * Index of the active worksheet that we are working on.
     *
     * @var int
     */
    private int $activesheet = 0;

    /**
     * Precalculate formulas.
     *
     * T313 - Need more speed with the PTA for report generation
     * The EXL Tracking Report takes 10+ minutes to generate when the default to
     * pre calculate formulas is on.  The Concentrix Tracking Report is taking 15
     * minutes to generate for The Helix site with the option disabled.
     */
    private bool $precaculateformulas = false;

    /**
     * Creates a new spreadsheet.
     *
     * @param string $title Title for the Worksheet
     */
    public function __construct(string $title)
    {
        \PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::setPaperSizeDefault(
            \PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::PAPERSIZE_A4
        );

        \PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::setOrientationDefault(
            \PhpOffice\PhpSpreadsheet\Worksheet\PageSetup::ORIENTATION_LANDSCAPE
        );

        $this->spreadsheet = new Spreadsheet();
        $this->sheet = $this->spreadsheet->getActiveSheet();

        $this->sheet->setShowGridlines(true);
        $this->sheet->setTitle($title);
    }

    public function loadTemplate(string $filename)
    {
        $reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
        $this->spreadsheet = $reader->load($filename);

        $this->sheet = $this->spreadsheet->getActiveSheet();
    }

    /**
     * Creates a new sheet and makes the new sheet the active sheet.
     *
     * @param string $title Title for the Worksheet
     *
     * @return void
     */
    public function createsheet(string $title): void
    {
        ++$this->activesheet;

        $this->spreadsheet->createsheet();
        $this->spreadsheet->setActiveSheetIndex($this->activesheet);

        $this->sheet = $this->spreadsheet->getActiveSheet();

        $this->sheet->setShowGridlines(true);
        $this->sheet->setTitle($title);
    }

    /**
     * Get properties for the spreadsheet.
     *
     * @return \PhpOffice\PhpSpreadsheet\Document\Properties
     */
    public function getProperties(): \PhpOffice\PhpSpreadsheet\Document\Properties
    {
        return $this->spreadsheet->getProperties();
    }

    /**
     * Set the active sheet index by the name of the sheet.
     *
     * @param string $name
     *
     * @return void
     */
    public function setActiveSheetIndex(int $index): void
    {
        $this->spreadsheet->setActiveSheetIndex($index);
        $this->sheet = $this->spreadsheet->getActiveSheet();
    }

    /**
     * Set the active sheet index by the name of the sheet.
     *
     * @param string $name
     *
     * @return void
     */
    public function setActiveSheetIndexByName(string $name): void
    {
        $this->spreadsheet->setActiveSheetIndexByName($name);
        $this->sheet = $this->spreadsheet->getActiveSheet();
    }

    /**
     * Enable / Disable precalculated formulas
     *
     * @param bool $precaculateformulas
     * @return void
     */
    public function setPreCalculateFormulas(bool $precaculateformulas): bool
    {
        $this->precaculateformulas = $precaculateformulas;
    }

    /**
     * Apply header styles over multiple rows making up the header.
     *
     * @param string $cells Columns for the first row (i.e. A1:Z1).
     *
     * @return void
     */
    public function applyHeaderStyleSingleRow(string $cells): void
    {
        $styleArray = [
            'font' => [
                'bold' => true,
            ],
            'borders' => [
                'top' => [
                    'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THICK,
                ],
                'bottom' => [
                    'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
                ],
            ],
            'fill' => [
                'fillType' => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID,
                'startColor' => [
                    'argb' => 'FFA0A0A0',
                ],
                'endColor' => [
                    'argb' => 'FFA0A0A0',
                ],
            ],
        ];

        $this->sheet->getStyle($cells)->applyFromArray($styleArray);
    }

    /**
     * Apply header styles over multiple rows making up the header.
     *
     * @param string $firstRowCells Columns for the first row (i.e. A1:Z1).
     * @param string $lastRowCells  Columns for the last row of the header (i.e. A2:Z2).
     *
     * @return void
     */
    public function applyHeaderStylesMultipleRows(string $firstRowCells, string $lastRowCells): void
    {
        $styleArray = [
            'font' => [
                'bold' => true,
            ],
            'borders' => [
                'top' => [
                    'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THICK,
                ],
            ],
            'fill' => [
                'fillType' => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID,
                'startColor' => [
                    'argb' => 'FFA0A0A0',
                ],
                'endColor' => [
                    'argb' => 'FFA0A0A0',
                ],
            ],
        ];
        $this->sheet->getStyle($firstRowCells)->applyFromArray($styleArray);

        $styleArray = [
            'font' => [
                'bold' => true,
            ],
            'borders' => [
                'bottom' => [
                    'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
                ],
            ],
        ];
        $this->sheet->getStyle($lastRowCells)->applyFromArray($styleArray);
    }

    /**
     * Set the auto size property on a column for the column width.
     *
     * @param string $firstCell First column (i.e. A)
     * @param string $lastcell  Last column (i.e. Z)
     *
     * @return void
     */
    public function applyAutoSize(string $firstCell, string $lastCell): void
    {
        /** @psalm-suppress StringIncrement */
        ++$lastCell;
        $col = $firstCell;

        while ($col !== $lastCell) {
            $this->sheet->getColumnDimension($col)->setAutoSize(true);
            /** @psalm-suppress StringIncrement */
            ++$col;
        }
    }

    /**
     * Set the same column width to specified column range.
     *
     * @param string $firstCell First column (i.e. A)
     * @param string $lastCell  Last column (i.e. Z)
     * @param float  $size      Array of columns in character units
     *
     * @return void
     *
     * @see   https://phpspreadsheet.readthedocs.io/en/latest/topics/recipes/#setting-a-columns-width
     */
    public function applySameSizePerColumn(string $firstCell, string $lastCell, float $size): void
    {
        /** @psalm-suppress StringIncrement */
        ++$lastCell;
        $col = $firstCell;

        while ($col !== $lastCell) {
            $this->sheet->getColumnDimension($col)->setWidth($size);
            /** @psalm-suppress StringIncrement */
            ++$col;
        }
    }

    /**
     * Set the the column width to specified column sizes
     *
     * @param string $firstCell First column (i.e. A)
     * @param string $lastCell  Last column (i.e. Z)
     * @param array  $sizes     Array of columns in character units
     *
     * @return void
     *
     * @see   https://phpspreadsheet.readthedocs.io/en/latest/topics/recipes/#setting-a-columns-width
     */
    public function applySizePerColumn(string $firstCell, string $lastCell, array $sizes): void
    {
        $col = $firstCell;
        /** @psalm-suppress StringIncrement */
        ++$lastCell;

        while ($col !== $lastCell) {
            $this->sheet->getColumnDimension($col)->setWidth($sizes[$col]);
            /** @psalm-suppress StringIncrement */
            ++$col;
        }
    }

    /**
     * Save the spreadsheet.
     *
     * @param string $filename
     *
     * @return void
     */
    public function save(string $filename): void
    {
        if (empty($filename)) {
            throw new \InvalidArgumentException('Please pass in a filename.');
        }

        $writer = IOFactory::createWriter($this->spreadsheet, 'Xlsx');

        $writer->setPreCalculateFormulas($this->precaculateformulas);
        $writer->save($filename);
    }

    public function getWorksheet()
    {
        return $this->spreadsheet;
    }

    /**
     * @param string $name
     * @param array $args
     */
    public function __call(string $name, array $args)
    {
        /**
         * Certain calls we need to send perform against the workbook and not a
         * worksheet.
         */
        if (\in_array($name, [
            'getActiveSheet',
            'getActiveSheetIndex',
            'getDefaultStyle',
            'getProperties',
            'getSheet',
            'getAllSheets',
            'getIndex',
            'getSheetByCodeName',
            'getSheetByName',
            'getSheetCount',
            'getSheetNames',
            'removeSheetByIndex',
            'sheetNameExists',
        ])) {
            return (\call_user_func_array(array($this->spreadsheet, $name), $args));
        }

        return (\call_user_func_array(array($this->sheet, $name), $args));
    }
}

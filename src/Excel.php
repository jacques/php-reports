<?php declare(strict_types=1);
/**
 * @author    Jacques Marneweck <jacques@siberia.co.za>
 * @copyright 2018-2020 Jacques Marneweck.  All rights strictly reserved.
 */

namespace Jacques\Reports;

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

class Excel
{
    /**
     * @var \PhpOffice\PhpSpreadsheet\Spreadsheet
     */
    private $spreadsheet;

    /**
     * @var \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
     */
    private $sheet;

    /**
     * Active sheet we are working on.
     * @var integer
     */
    private $activesheet = 0;

    /**
     * Creates a new spreadsheet.
     *
     * @param string $title Title for the Worksheet
     */
    public function __construct($title)
    {
        $this->spreadsheet = new Spreadsheet();
        $this->sheet = $this->spreadsheet->getActiveSheet();

        $this->sheet->setShowGridlines(true);
        $this->sheet->setTitle($title);
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
        $this->activesheet++;

        $this->spreadsheet->createsheet();
        $this->spreadsheet->setActiveSheetIndex($this->activesheet);
        $this->sheet = $this->spreadsheet->getActiveSheet();

        $this->sheet->setShowGridlines(true);
        $this->sheet->setTitle($title);
    }

    /**
     * Retreive the active sheet.
     *
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
     */
    public function getActiveSheet(): \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
    {
        return $this->spreadsheet->getActiveSheet();
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

    public function removeSheetByIndex(int $index): void
    {
        $this->spreadsheet->removeSheetByIndex($index);
    }

    /**
     * Get index for sheet.
     *
     * @param string $index
     *
     * @return int
     */
    public function getIndex(): int
    {
        return $this->spreadsheet->getIndex($this->sheet);
    }

    /**
     * Gets a sheet by name.
     *
     * @param string $name
     *
     * @return \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet|null
     */
    public function getSheetByName(string $name): ?\PhpOffice\PhpSpreadsheet\Worksheet\Worksheet
    {
        return $this->spreadsheet->getSheetByName($name);
    }

    /**
     * Set the active sheet index by the name of the sheet.
     *
     * @param string $name
     */
    public function setActiveSheetIndexByName(string $name): void
    {
        $this->spreadsheet->setActiveSheetIndexByName($name);
        $this->sheet = $this->spreadsheet->getActiveSheet();
    }

    /**
     * Apply header styles over multiple rows making up the header.
     *
     * @param string $cells Columns for the first row (i.e. A1:Z1).
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
        $col = $firstCell;
        /** @psalm-suppress StringIncrement */
        $lastCell++;

        while ($col !== $lastCell) {
            $this->sheet->getColumnDimension($col)->setAutoSize(true);
            /** @psalm-suppress StringIncrement */
            $col++;
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
        $writer = IOFactory::createWriter($this->spreadsheet, 'Xlsx');
        $writer->save($filename);
    }

    public function __call(string $name, array $args)
    {
        return (\call_user_func_array(array($this->sheet, $name), $args));
    }
}

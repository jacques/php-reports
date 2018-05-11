<?php
/**
 * @author    Jacques Marneweck <jacques@siberia.co.za>
 * @copyright 2018 Jacques Marneweck.  All rights strictly reserved.
 */

namespace Jacques\Reports;

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;

class Excel
{
    /**
     * @var \PhpOffice\PhpSpreadsheet\Spreadsheet|null
     */
    private $spreadsheet = null;

    /**
     * @var \PhpOffice\PhpSpreadsheet\Worksheet\Worksheet|null
     */
    private $sheet = null;

    /**
     * Active sheet we are working on.
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
     */
    public function createsheet($title)
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
     */
    public function getActiveSheet()
    {
        return $this->spreadsheet->getActiveSheet();
    }

    /**
     * Get properties for the spreadsheet.
     */
    public function getProperties()
    {
        return $this->spreadsheet->getProperties();
    }

    /**
     * Apply header styles over multiple rows making up the header.
     *
     * @param string $cells Columns for the first row (i.e. A1:Z1).
     */
    public function applyHeaderStyleSingleRow($cells)
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
     */
    public function applyHeaderStylesMultipleRows($firstRowCells, $lastRowCells)
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
     */
    public function applyAutoSize($firstCell, $lastCell)
    {
        foreach (range($firstCell, $lastCell) as $col) {
            $this->sheet->getColumnDimension($col)->setAutoSize(true);
        }
    }

    /**
     * Save the spreadsheet.
     *
     * @param string $filename
     */
    public function save($filename)
    {
        $writer = IOFactory::createWriter($this->spreadsheet, 'Xlsx');
        $writer->save($filename);
    }

    public function __call($name, $args)
    {
        return (call_user_func_array(array($this->sheet, $name), $args));
    }
}

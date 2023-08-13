<?php declare(strict_types=1);
/**
 * @author    Jacques Marneweck <jacques@siberia.co.za>
 * @copyright 2018-2023 Jacques Marneweck.  All rights strictly reserved.
 */

namespace Jacques\Reports\Tests\Unit;

use Brick\VarExporter\VarExporter;
use Jacques\Reports\Excel;

final class ExcelTest extends \PHPUnit\Framework\TestCase
{
    /**
     * Sets up the fixture, for example, opens a network connection.
     * This method is called before a test is executed.
     */
    protected function setUp(): void
    {
    }

    /**
     * Tears down the fixture, for example, closes a network connection.
     * This method is called after a test is executed.
     */
    protected function tearDown(): void
    {
    }

    public function testExample(): void
    {
        $spreadsheet = new \Jacques\Reports\Excel('TEST');
        self::assertInstanceOf('\Jacques\Reports\Excel', $spreadsheet);

        $properties = $spreadsheet->getProperties();
        self::assertEquals('Unknown Creator', $properties->getCreator());
        $properties->setCreator('Bilbo Baggins');
        $properties->setLastModifiedBy('Gandalf the Grey');
        $properties->setCompany('The Shire Inc');

        $properties = $spreadsheet->getProperties();
        self::assertEquals('The Shire Inc', $properties->getCompany());

        self::assertEquals('Calibri', $spreadsheet->getDefaultStyle()->getFont()->getName());
        self::assertEquals(11, $spreadsheet->getDefaultStyle()->getFont()->getSize());

        $spreadsheet->getDefaultStyle()->getFont()->setName('Gill Sans');
        $spreadsheet->getDefaultStyle()->getFont()->setSize(12);

        self::assertEquals('Gill Sans', $spreadsheet->getDefaultStyle()->getFont()->getName());
        self::assertEquals(12, $spreadsheet->getDefaultStyle()->getFont()->getSize());

        $spreadsheet->setCellValue('A1', 'First Name');
        $spreadsheet->setCellValue('B1', 'Last Name');
        $spreadsheet->setCellValue('C1', 'Email');

        $spreadsheet->setCellValue('A2', 'Joe');
        $spreadsheet->setCellValue('B2', 'Soap');
        $spreadsheet->setCellValue('C2', 'joe@example.com');

        $spreadsheet->applyAutoSize('A', 'C');
        $spreadsheet->applyHeaderStyleSingleRow('A1:C1');
        $spreadsheet->drawborders('A2:C2', 'outer');
        $spreadsheet->drawborders('A1:C1', 'top', 'thick');

        $spreadsheet->createSheet('Summary');
        self::assertEquals('Summary', $spreadsheet->getActiveSheet()->getTitle());

        $spreadsheet->setCellValue('A1', 'Personal Details');
        $spreadsheet->mergeCells('A1:C1');

        $spreadsheet->setCellValue('A2', 'First Name');
        $spreadsheet->setCellValue('B2', 'Last Name');
        $spreadsheet->setCellValue('C2', 'Email');

        $spreadsheet->setCellValue('A3', 'Joe');
        $spreadsheet->setCellValue('B3', 'Soap');
        $spreadsheet->setCellValue('C3', 'joe@example.com');

        $spreadsheet->applyHeaderStylesMultipleRows('A1:C1', 'A2:C2');

        $spreadsheet->applySizePerColumn('A', 'C', [
            'A' => 10,
            'B' => 20,
            'C' => 30,
        ]);
        self::assertEquals(10, $spreadsheet->getColumnDimension('A')->getWidth());
        self::assertEquals('20', $spreadsheet->getColumnDimension('B')->getWidth());
        self::assertEquals('30', $spreadsheet->getColumnDimension('C')->getWidth());

        $spreadsheet->applySameSizePerColumn('D', 'E', 50);
        self::assertEquals('50', $spreadsheet->getColumnDimension('D')->getWidth());
        self::assertEquals('50', $spreadsheet->getColumnDimension('E')->getWidth());

        $spreadsheet->setCellValue('D2', 'Location');
        $spreadsheet->setCellValue('E2', 'Continent');
        $spreadsheet->setCellValue('D3', 'Cape Town, South Africa');
        $spreadsheet->setCellValue('E3', 'Africa');

        $spreadsheet->applyHeaderStylesMultipleRows('A1:E1', 'A2:E2');
        $spreadsheet->drawborders('A3:E3', 'all');

        $spreadsheet->setActiveSheetIndexByName('TEST');
        self::assertEquals('TEST', $spreadsheet->getActiveSheet()->getTitle());

        $expected = [
            0 => 'TEST',
            1 => 'Summary',
        ];
        self::assertEquals($expected, $spreadsheet->getSheetNames());

        self::assertEquals(0, $spreadsheet->getIndex($spreadsheet->getActiveSheet()));

        $spreadsheet->setActiveSheetIndexByName('Summary');
        self::assertEquals('Summary', $spreadsheet->getActiveSheet()->getTitle());
        self::assertEquals(1, $spreadsheet->getIndex($spreadsheet->getActiveSheet()));
        $spreadsheet->save('foo.xlsx');

        $spreadsheet->setActiveSheetIndex(0);
        self::assertEquals('TEST', $spreadsheet->getActiveSheet()->getTitle());
        self::assertEquals(0, $spreadsheet->getIndex($spreadsheet->getActiveSheet()));
        self::assertEquals('joe@example.com', $spreadsheet->getActiveSheet()->getCell('C2')->getFormattedValue());

        $spreadsheet->setActiveSheetIndex(1);

        self::assertEquals('Summary', $spreadsheet->getActiveSheet()->getTitle());
        self::assertEquals(1, $spreadsheet->getIndex($spreadsheet->getActiveSheet()));
        self::assertEquals('joe@example.com', $spreadsheet->getActiveSheet()->getCell('C3')->getFormattedValue());

        $spreadsheet->removeSheetByIndex(1);
        self::assertCount(1, $spreadsheet->getSheetNames());
        self::assertEquals(['TEST'], $spreadsheet->getSheetNames());
        unset($spreadsheet);

        $reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
        $spreadsheet = $reader->load('foo.xlsx');

        $expected = [
            0 => 'TEST',
            1 => 'Summary',
        ];

        self::assertEquals($expected, $spreadsheet->getSheetNames());

        self::assertEquals('Summary', $spreadsheet->getActiveSheet()->getTitle());

        $spreadsheet->setActiveSheetIndex(0);
        self::assertEquals(0, $spreadsheet->getIndex($spreadsheet->getActiveSheet()));

        self::assertEquals('TEST', $spreadsheet->getActiveSheet()->getTitle());

        self::assertEquals('First Name', $spreadsheet->getActiveSheet()->getCell('A1'));
        self::assertEquals('Last Name', $spreadsheet->getActiveSheet()->getCell('B1'));
        self::assertEquals('Email', $spreadsheet->getActiveSheet()->getCell('C1'));

        self::assertEquals('Joe', $spreadsheet->getActiveSheet()->getCell('A2'));
        self::assertEquals('Soap', $spreadsheet->getActiveSheet()->getCell('B2'));
        self::assertEquals('joe@example.com', $spreadsheet->getActiveSheet()->getCell('C2'));

        $spreadsheet->setActiveSheetIndex(1);

        self::assertEquals('Summary', $spreadsheet->getActiveSheet()->getTitle());

        self::assertEquals('Personal Details', $spreadsheet->getActiveSheet()->getCell('A1'));

        self::assertEquals('First Name', $spreadsheet->getActiveSheet()->getCell('A2'));
        self::assertEquals('Last Name', $spreadsheet->getActiveSheet()->getCell('B2'));
        self::assertEquals('Email', $spreadsheet->getActiveSheet()->getCell('C2'));

        $spreadsheet->setActiveSheetIndexByName('TEST');

        self::assertEquals('TEST', $spreadsheet->getActiveSheet()->getTitle());

        $expected = [
            'alignment' => [
                'horizontal' => 'general',
                'indent' => 0,
                'readOrder' => 0,
                'shrinkToFit' => false,
                'textRotation' => 0,
                'vertical' => 'bottom',
                'wrapText' => false
            ],
            'borders' => [
                'bottom' => [
                    'borderStyle' => 'thin',
                    'color' => [
                        'argb' => 'FF000000'
                    ]
                ],
                'diagonal' => [
                    'borderStyle' => 'none',
                    'color' => [
                        'argb' => 'FF000000'
                    ]
                ],
                'diagonalDirection' => 0,
                'left' => [
                    'borderStyle' => 'none',
                    'color' => [
                        'argb' => 'FF000000'
                    ]
                ],
                'right' => [
                    'borderStyle' => 'none',
                    'color' => [
                        'argb' => 'FF000000'
                    ]
                ],
                'top' => [
                    'borderStyle' => 'thin',
                    'color' => [
                        'argb' => 'FF000000'
                    ]
                ]
            ],
            'fill' => [
                'endColor' => [
                    'argb' => 'FFA0A0A0'
                ],
                'fillType' => 'solid',
                'rotation' => 0,
                'startColor' => [
                    'argb' => 'FFA0A0A0'
                ]
            ],
            'font' => [
                'bold' => true,
                'color' => [
                    'argb' => 'FF000000'
                ],
                'italic' => false,
                'name' => 'Gill Sans',
                'size' => 12.0,
                'strikethrough' => false,
                'subscript' => false,
                'superscript' => false,
                'underline' => 'none',
                'baseLine' => 0,
                'chartColor' => null,
                'complexScript' => '',
                'eastAsian' => '',
                'latin' => '',
                'strikeType' => '',
                'underlineColor' => null,
                'scheme' => '',
            ],
            'numberFormat' => [
                'formatCode' => 'General'
            ],
            'protection' => [
                'locked' => 'inherit',
                'hidden' => 'inherit'
            ],
            'quotePrefx' => false
        ];
        self::assertEquals($expected, $spreadsheet->getActiveSheet()->getStyle('A1')->exportArray());

        $expected = [
            'alignment' => [
                'horizontal' => 'general',
                'indent' => 0,
                'readOrder' => 0,
                'shrinkToFit' => false,
                'textRotation' => 0,
                'vertical' => 'bottom',
                'wrapText' => false
            ],
            'borders' => [
                'bottom' => [
                    'borderStyle' => 'thin',
                    'color' => [
                        'argb' => 'FF000000'
                    ]
                ],
                'diagonal' => [
                    'borderStyle' => 'none',
                    'color' => [
                        'argb' => 'FF000000'
                    ]
                ],
                'diagonalDirection' => 0,
                'left' => [
                    'borderStyle' => 'thin',
                    'color' => [
                        'argb' => 'FF000000'
                    ]
                ],
                'right' => [
                    'borderStyle' => 'none',
                    'color' => [
                        'argb' => 'FF000000'
                    ]
                ],
                'top' => [
                    'borderStyle' => 'thin',
                    'color' => [
                        'argb' => 'FF000000'
                    ]
                ]
            ],
            'fill' => [
                'fillType' => 'none',
                'rotation' => 0.0,
            ],
            'font' => [
                'bold' => false,
                'color' => [
                    'argb' => 'FF000000'
                ],
                'italic' => false,
                'name' => 'Gill Sans',
                'size' => 12.0,
                'strikethrough' => false,
                'subscript' => false,
                'superscript' => false,
                'underline' => 'none',
                'baseLine' => 0,
                'chartColor' => null,
                'complexScript' => '',
                'eastAsian' => '',
                'latin' => '',
                'strikeType' => '',
                'underlineColor' => null,
                'scheme' => '',
            ],
            'numberFormat' => [
                'formatCode' => 'General'
            ],
            'protection' => [
                'locked' => 'inherit',
                'hidden' => 'inherit'
            ],
            'quotePrefx' => false
        ];
        self::assertEquals($expected, $spreadsheet->getActiveSheet()->getStyle('A2:C2')->exportArray());

        $expected = [
            'alignment' => [
                'horizontal' => 'general',
                'indent' => 0,
                'readOrder' => 0,
                'shrinkToFit' => false,
                'textRotation' => 0,
                'vertical' => 'bottom',
                'wrapText' => false
            ],
            'borders' => [
                'bottom' => [
                    'borderStyle' => 'none',
                    'color' => [
                        'argb' => 'FF000000'
                    ]
                ],
                'diagonal' => [
                    'borderStyle' => 'none',
                    'color' => [
                        'argb' => 'FF000000'
                    ]
                ],
                'diagonalDirection' => 0,
                'left' => [
                    'borderStyle' => 'none',
                    'color' => [
                        'argb' => 'FF000000'
                    ]
                ],
                'right' => [
                    'borderStyle' => 'none',
                    'color' => [
                        'argb' => 'FF000000'
                    ]
                ],
                'top' => [
                    'borderStyle' => 'none',
                    'color' => [
                        'argb' => 'FF000000'
                    ]
                ]
            ],
            'fill' => [
                'fillType' => 'none',
                'rotation' => 0.0,
            ],
            'font' => [
                'bold' => false,
                'color' => [
                    'argb' => 'FF000000'
                ],
                'italic' => false,
                'name' => 'Gill Sans',
                'size' => 12.0,
                'strikethrough' => false,
                'subscript' => false,
                'superscript' => false,
                'underline' => 'none',
                'baseLine' => 0,
                'chartColor' => null,
                'complexScript' => '',
                'eastAsian' => '',
                'latin' => '',
                'strikeType' => '',
                'underlineColor' => null,
                'scheme' => '',
            ],
            'numberFormat' => [
                'formatCode' => 'General'
            ],
            'protection' => [
                'locked' => 'inherit',
                'hidden' => 'inherit'
            ],
            'quotePrefx' => false
        ];
        self::assertEquals($expected, $spreadsheet->getActiveSheet()->getStyle('A3:C3')->exportArray());

        $spreadsheet->setActiveSheetIndexByName('Summary');
        self::assertEquals('Summary', $spreadsheet->getActiveSheet()->getTitle());

        $expected = [
            'alignment' => [
                'horizontal' => 'general',
                'indent' => 0,
                'readOrder' => 0,
                'shrinkToFit' => false,
                'textRotation' => 0,
                'vertical' => 'bottom',
                'wrapText' => false
            ],
            'borders' => [
                'bottom' => [
                    'borderStyle' => 'none',
                    'color' => [
                        'argb' => 'FF000000'
                    ]
                ],
                'diagonal' => [
                    'borderStyle' => 'none',
                    'color' => [
                        'argb' => 'FF000000'
                    ]
                ],
                'diagonalDirection' => 0,
                'left' => [
                    'borderStyle' => 'none',
                    'color' => [
                        'argb' => 'FF000000'
                    ]
                ],
                'right' => [
                    'borderStyle' => 'none',
                    'color' => [
                        'argb' => 'FF000000'
                    ]
                ],
                'top' => [
                    'borderStyle' => 'thick',
                    'color' => [
                        'argb' => 'FF000000'
                    ]
                ]
            ],
            'fill' => [
                'endColor' => [
                    'argb' => 'FFA0A0A0'
                ],
                'fillType' => 'solid',
                'rotation' => 0.0,
                'startColor' => [
                    'argb' => 'FFA0A0A0'
                ]
            ],
            'font' => [
                'bold' => true,
                'color' => [
                    'argb' => 'FF000000'
                ],
                'italic' => false,
                'name' => 'Gill Sans',
                'size' => 12.0,
                'strikethrough' => false,
                'subscript' => false,
                'superscript' => false,
                'underline' => 'none',
                'baseLine' => 0,
                'chartColor' => null,
                'complexScript' => '',
                'eastAsian' => '',
                'latin' => '',
                'strikeType' => '',
                'underlineColor' => null,
                'scheme' => '',
            ],
            'numberFormat' => [
                'formatCode' => 'General'
            ],
            'protection' => [
                'locked' => 'inherit',
                'hidden' => 'inherit'
            ],
            'quotePrefx' => false
        ];
        self::assertEquals($expected, $spreadsheet->getActiveSheet()->getStyle('A1')->exportArray());

        $sheet = $spreadsheet->getActiveSheet();

        $properties = $spreadsheet->getProperties();
        self::assertEquals('Bilbo Baggins', $properties->getCreator());
        self::assertEquals('Gandalf the Grey', $properties->getLastModifiedBy());
        self::assertEquals('The Shire Inc', $properties->getCompany());
    }
}

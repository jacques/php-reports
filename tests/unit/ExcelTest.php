<?php declare(strict_types=1);
/**
 * @author    Jacques Marneweck <jacques@siberia.co.za>
 * @copyright 2018-2020 Jacques Marneweck.  All rights strictly reserved.
 */

namespace Jacques\Reports\Tests\Unit;

use Jacques\Reports\Excel;

class ExcelTest extends \PHPUnit\Framework\TestCase
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

        self::assertEquals('Calibri', $spreadsheet->getDefaultStyle()->getFont()->getName());
        self::assertEquals(11, $spreadsheet->getDefaultStyle()->getFont()->getSize());

        $spreadsheet->getDefaultStyle()->getFont()->setName('Gill Sans');
        $spreadsheet->getDefaultStyle()->getFont()->setSize(12);

        $spreadsheet->setCellValue('A1', 'First Name');
        $spreadsheet->setCellValue('B1', 'Last Name');
        $spreadsheet->setCellValue('C1', 'Email');

        $spreadsheet->setCellValue('A2', 'Joe');
        $spreadsheet->setCellValue('B2', 'Soap');
        $spreadsheet->setCellValue('C2', 'joe@example.com');

        $spreadsheet->applyAutoSize('A', 'C');
        $spreadsheet->applyHeaderStyleSingleRow('A1:C1');

        $spreadsheet->createSheet('Summary');

        $spreadsheet->setCellValue('A1', 'Personal Details');
        $spreadsheet->mergeCells('A1:C1');

        $spreadsheet->setCellValue('A2', 'First Name');
        $spreadsheet->setCellValue('B2', 'Last Name');
        $spreadsheet->setCellValue('C2', 'Email');

        $spreadsheet->setCellValue('A3', 'Joe');
        $spreadsheet->setCellValue('B3', 'Soap');
        $spreadsheet->setCellValue('C3', 'joe@example.com');

        $spreadsheet->save('foo.xlsx');

        $reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
        $spreadsheet = $reader->load('foo.xlsx');

        $expected = [
            0 => 'TEST',
            1 => 'Summary',
        ];

        self::assertEquals($expected, $spreadsheet->getSheetNames());

        self::assertEquals('Summary', $spreadsheet->getActiveSheet()->getTitle());

        $spreadsheet->setActiveSheetIndex(0);

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

    }
}

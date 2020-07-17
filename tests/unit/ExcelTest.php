<?php declare(strict_types=1);
/**
 * @author    Jacques Marneweck <jacques@siberia.co.za>
 * @copyright 2018-2020 Jacques Marneweck.  All rights strictly reserved.
 */

namespace Jacques\Reports\Tests\Unit;

use Jacques\Reports\Excel;

class Test extends \PHPUnit\Framework\TestCase
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
        $spreadsheet->setCellValue('A1', 'First Name');
        $spreadsheet->setCellValue('B1', 'Last Name');
        $spreadsheet->setCellValue('C1', 'Email');

        $spreadsheet->setCellValue('A2', 'Joe');
        $spreadsheet->setCellValue('B2', 'Soap');
        $spreadsheet->setCellValue('C2', 'joe@example.com');

        $spreadsheet->applyAutoSize('A', 'C');
        $spreadsheet->applyHeaderStyleSingleRow('A1:C1');
        $spreadsheet->save('foo.xlsx');

        $reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
        $spreadsheet = $reader->load('foo.xlsx');

        self::assertEquals('TEST', $spreadsheet->getActiveSheet()->getTitle());

        self::assertEquals('First Name', $spreadsheet->getActiveSheet()->getCell('A1'));
        self::assertEquals('Last Name', $spreadsheet->getActiveSheet()->getCell('B1'));
        self::assertEquals('Email', $spreadsheet->getActiveSheet()->getCell('C1'));

        self::assertEquals('Joe', $spreadsheet->getActiveSheet()->getCell('A2'));
        self::assertEquals('Soap', $spreadsheet->getActiveSheet()->getCell('B2'));
        self::assertEquals('joe@example.com', $spreadsheet->getActiveSheet()->getCell('C2'));
    }
}

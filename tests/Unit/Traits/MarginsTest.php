<?php declare(strict_types=1);
/**
 * @author    Jacques Marneweck <jacques@siberia.co.za>
 * @copyright 2018-2022 Jacques Marneweck.  All rights strictly reserved.
 */

namespace Jacques\Reports\Tests\Unit\Traits;

use Brick\VarExporter\VarExporter;
use Jacques\Reports\Excel;

final class MarginsTest extends \PHPUnit\Framework\TestCase
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

    public function testInvalidPreset(): void
    {
        $spreadsheet = new \Jacques\Reports\Excel('TEST');
        self::assertInstanceOf('\Jacques\Reports\Excel', $spreadsheet);
        self::expectException(\InvalidArgumentException::class);
        self::expectExceptionMessage('Unknown margin preset \'superwide\' specified.');

        $spreadsheet->setMargins('superwide');
    }

    public function testSetNarrowMargins(): void
    {
        $spreadsheet = new \Jacques\Reports\Excel('TEST');
        self::assertInstanceOf('\Jacques\Reports\Excel', $spreadsheet);
        $spreadsheet->setMargins('narrow');

        $expected = (static function() {
            $class = new \ReflectionClass(\PhpOffice\PhpSpreadsheet\Worksheet\PageMargins::class);
            $object = $class->newInstanceWithoutConstructor();

            (function() {
                $this->left = 0.25;
                $this->right = 0.25;
                $this->top = 0.75;
                $this->bottom = 0.75;
                $this->header = 0.3;
                $this->footer = 0.3;
            })->bindTo($object, \PhpOffice\PhpSpreadsheet\Worksheet\PageMargins::class)();

            return $object;
        })();

        $margins = $spreadsheet->getPageMargins();
        self::assertEquals($expected, $margins);
    }
}

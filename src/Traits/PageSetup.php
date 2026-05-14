<?php declare(strict_types=1);
/**
 * @author    Jacques Marneweck <jacques@siberia.co.za>
 * @copyright 2018-2026 Jacques Marneweck.  All rights strictly reserved.
 */

namespace Jacques\Reports\Traits;

use \InvalidArgumentException;
use PhpOffice\PhpSpreadsheet\Worksheet\PageSetup;

trait PageSetup
{
    /**
     * Configure the active worksheet's page setup for portrait A4 printing at 100% scale with fit-to-width enabled.
     * Pass in the $scale to change the scale (zoom).
     *
     * @param \Jacques\Reports\Excel $report The report whose active worksheet will be configured.
     * @param int $scale Scale to zoom
     */
    public function printsetup(\Jacques\Reports\Excel $report, int $scale = 100): void
    {
        $report->getActiveSheet()->getPageSetup()
            ->setOrientation(PageSetup::ORIENTATION_PORTRAIT);
        $report->getActiveSheet()->getPageSetup()
            ->setPaperSize(PageSetup::PAPERSIZE_A4);
        $report->getActiveSheet()->getPageSetup()
            ->setScale($scale, true);
        $report->getActiveSheet()->getPageSetup()
            ->setFitToWidth(1);
        $report->getActiveSheet()->getPageSetup()
            ->setFitToHeight(0);
    }
}

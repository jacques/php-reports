<?php declare(strict_types=1);
/**
 * @author    Jacques Marneweck <jacques@siberia.co.za>
 * @copyright 2018-2021 Jacques Marneweck.  All rights strictly reserved.
 */

namespace Jacques\Reports\Traits;

trait Borders
{
    /**
     * Draws borders.
     *
     * @param \Jacques\Report\Excel $report
     * @param string $coords
     *
     * @void
     */
    public function drawborders(\Jacques\Reports\Excel $report, string $coords, ?string $type = 'outer'): void
    {
        switch ($type) {
            case 'all':
                $report->getStyle($coords)
                    ->applyFromArray([
                        'borders' => [
                            'allBorders' => [
                                'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
                            ],
                        ],
                    ]);
                break;
            case 'bottom':
                $report->getStyle($coords)
                    ->getBorders()->getBottom()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN);
                break;
            case 'outer':
            default:
                $report->getStyle($coords)
                    ->getBorders()->getTop()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN);
                $report->getStyle($coords)
                    ->getBorders()->getBottom()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN);
                $report->getStyle($coords)
                    ->getBorders()->getLeft()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN);
                $report->getStyle($coords)
                    ->getBorders()->getRight()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN);
            break;
        }
    }
}

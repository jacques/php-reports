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
     * @param string $coords
     * @param string $type (all | top | bottom | left | right | outer)
     *
     * @void
     */
    public function drawborders(string $coords, ?string $type = 'outer', ?string $style = 'thin'): void
    {
        switch ($style) {
            case 'dashdot':
                $style = \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_DASHDOT;
                break;
            case 'dashdotdot':
                $style = \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_DASHDOTDOT;
                break;
            case 'dashdotdot':
                $style = \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_DOUBLE;
                break;
            case 'hair':
                $style = \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_HAIR;
                break;
            case 'thin':
                $style = \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN;
                break;
            case 'thick':
                $style = \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THICK;
                break;
            default:
                throw \InvalidArgumentException(\sprintf('Unknown border style \'%s\' specified.', $style));
        }

        switch ($type) {
            case 'all':
                $this->sheet->getStyle($coords)
                    ->applyFromArray([
                        'borders' => [
                            'allBorders' => [
                                'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
                            ],
                        ],
                    ]);
                break;
            case 'top':
                $this->sheet->getStyle($coords)
                    ->getBorders()->getTop()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN);
                break;
            case 'bottom':
                $this->sheet->getStyle($coords)
                    ->getBorders()->getBottom()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN);
                break;
            case 'left':
                $this->sheet->getStyle($coords)
                    ->getBorders()->getLeft()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN);
                break;
            case 'right':
                $this->sheet->getStyle($coords)
                    ->getBorders()->getRight()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN);
                break;
            case 'outer':
            default:
                $this->sheet->getStyle($coords)
                    ->getBorders()->getTop()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN);
                $this->sheet->getStyle($coords)
                    ->getBorders()->getBottom()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN);
                $this->sheet->getStyle($coords)
                    ->getBorders()->getLeft()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN);
                $this->sheet->getStyle($coords)
                    ->getBorders()->getRight()->setBorderStyle(\PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN);
            break;
        }
    }
}

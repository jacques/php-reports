<?php declare(strict_types=1);
/**
 * @author    Jacques Marneweck <jacques@siberia.co.za>
 * @copyright 2018-2022 Jacques Marneweck.  All rights strictly reserved.
 */

namespace Jacques\Reports\Traits;

use \InvalidArgumentException;

trait Colour
{
    /**
     * Set the cell background colour.
     *
     * @param string $coords
     * @param string $colour
     *
     * @void
     */
    public function setCellColour(string $coords, string $colour): void
    {
        $this->sheet->getStyle($coords)
            ->getFill()
            ->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)
            ->getStartColor()
            ->setARGB($colour);
    }

    /**
     * Set the font colour.
     *
     * @param string $coords
     * @param string $colour
     *
     * @void
     */
    public function setFontColour(string $coords, string $colour): void
    {
        $this->sheet->getStyle($coords)
            ->getFont()
            ->getColor()
            ->setARGB($colour);
    }
}

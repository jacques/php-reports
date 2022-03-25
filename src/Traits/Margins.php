<?php declare(strict_types=1);
/**
 * @author    Jacques Marneweck <jacques@siberia.co.za>
 * @copyright 2018-2022 Jacques Marneweck.  All rights strictly reserved.
 */

namespace Jacques\Reports\Traits;

use \InvalidArgumentException;

trait Margins
{
    /**
     * Set the page margins.
     *
     * @param string $preset (narrow|normal|wide|custom)
     *
     * @void
     */
    public function setMargins(string $preset): void
    {
        $presets = [
            'narrow' => [
                'left' => 0.25,
                'right' => 0.25,
                'top' => 0.75,
                'bottom' => 0.75,
                'header' => 0.3,
                'footer' => 0.3,
            ],
            'normal' => [
                'left' => 0.7,
                'right' => 0.7,
                'top' => 0.75,
                'bottom' => 0.75,
                'header' => 0.3,
                'footer' => 0.3,
            ],
            'wide' => [
                'left' => 1,
                'right' => 1,
                'top' => 1,
                'bottom' => 1,
                'header' => 0.5,
                'footer' => 0.5,
            ],
        ];

        if (!\in_array($preset, \array_keys($presets))) {
            throw new \InvalidArgumentException(\sprintf("Unknown margin preset '%s' specified.", $preset));
        }

        $margins = $this->sheet->getPageMargins();
        $margins->setLeft($presets[$preset]['left']);
        $margins->setRight($presets[$preset]['right']);
        $margins->setTop($presets[$preset]['top']);
        $margins->setBottom($presets[$preset]['bottom']);
        $margins->setHeader($presets[$preset]['header']);
        $margins->setFooter($presets[$preset]['footer']);
    }
}

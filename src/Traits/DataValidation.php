<?php declare(strict_types=1);
/**
 * @author    Jacques Marneweck <jacques@siberia.co.za>
 * @copyright 2018-2021 Jacques Marneweck.  All rights strictly reserved.
 */

namespace Jacques\Reports\Traits;

trait DataValidation
{
    /**
     * Drop-down list for data validation.
     *
     * @param \Jacques\Report\Excel $report
     * @param string $coords
     * @param array|string the options for the drop-down list.
     * @param string $title
     * @param string $prompt
     *
     * @void
     */
    public function dropdown(\Jacques\Reports\Excel $report, string $coords, array|string $options, string $title = 'Pick from list', string $prompt = 'Please pick a value from the drop-down list.'): void
    {
        $validation = $report->getActiveSheet()->getCell($cell)->getDataValidation();
        $validation->setType(DataValidation::TYPE_LIST)
            ->setErrorStyle(DataValidation::STYLE_INFORMATION)
            ->setAllowBlank(false)
            ->setShowInputMessage(true)
            ->setShowErrorMessage(true)
            ->setShowDropDown(true)
            ->setErrorTitle('Input error')
            ->setError('Value is not in list.')
            ->setPromptTitle($title)
            ->setPrompt($prompt);

        if (is_array($options)) {
            $validation->setFormula1('"' . implode(',', $options) . '"');
        } else {
            $validation->setFormula1($options);
        }
    }
}

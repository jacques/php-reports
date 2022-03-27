<?php declare(strict_types=1);
/**
 * @author    Jacques Marneweck <jacques@siberia.co.za>
 * @copyright 2020-2021 Jacques Marneweck.  All rights strictly reserved.
 */

use Rector\Core\Configuration\Option;
use Rector\Core\ValueObject\PhpVersion;
use Rector\Php74\Rector\Property\TypedPropertyRector;
use Rector\PHPUnit\Set\PHPUnitSetList;
use Rector\Set\ValueObject\SetList;
use Symfony\Component\DependencyInjection\Loader\Configurator\ContainerConfigurator;

return static function (ContainerConfigurator $containerConfigurator): void {
    $parameters = $containerConfigurator->parameters();
    $parameters->set(Option::PHP_VERSION_FEATURES, PhpVersion::PHP_74);

    $services = $containerConfigurator->services();

    $parameters->set(Option::PATHS, ['src', 'tests']);

    foreach ([
        SetList::CODE_QUALITY,
        #SetList::CODE_QUALITY_STRICT,
        SetList::CODING_STYLE,
        #SetList::PERFORMANCE,
        PHPUnitSetList::PHPUNIT_60,
        PHPUnitSetList::PHPUNIT_70,
        PHPUnitSetList::PHPUNIT_75,
        PHPUnitSetList::PHPUNIT_80,
        PHPUnitSetList::PHPUNIT_90,
        PHPUnitSetList::PHPUNIT_91,
        SetList::PHP_70,
        SetList::PHP_71,
        SetList::PHP_72,
        SetList::PHP_73,
        SetList::PHP_74,
    ] as $set) {
        $containerConfigurator->import($set);
    }
};

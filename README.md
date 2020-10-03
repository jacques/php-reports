# Reports

[![Author](http://img.shields.io/badge/author-@jacques-blue.svg?style=flat-square)](https://twitter.com/jacques)
[![Software License](https://img.shields.io/badge/license-MIT-brightgreen.svg?style=flat-square)](LICENSE.md)
[![Scrutinizer Code Quality](https://scrutinizer-ci.com/g/jacques/php-reports/badges/quality-score.png?b=master)](https://scrutinizer-ci.com/g/jacques/php-reports/?branch=master)
[![Maintainability](https://api.codeclimate.com/v1/badges/750d2130ab688f3d8031/maintainability)](https://codeclimate.com/github/jacques/php-reports/maintainability)
[![Test Coverage](https://api.codeclimate.com/v1/badges/750d2130ab688f3d8031/test_coverage)](https://codeclimate.com/github/jacques/php-reports/test_coverage)

> Helpers for building reports with PHP (i.e. Excel Spreadsheets).  Use the
> [phpspreadsheet](https://github.com/phpoffice/phpspreadsheet) package to generate
> the Excel Spreadsheet.

---

## Installation

```
$ composer require jacques/php-reports:dev-master
```

```
{
    "require": {
        "jacques/php-reports": "dev-master"
    }
}
```

---

## Usage

```php
<?php declare(strict_types=1);

require_once __DIR__ . '/../vendor/autoload.php';

$report = new \Jacques\Reports\Excel('Cover Sheet');
$report->setCellValue('B2', 'Report Name');
$report->setCellValue('B4', 'Description');

$report->createSheet('User');

$report->save(__DIR__.'/tmp/file.xlsx');
```

---

## License

Report Helpers is licensed under the [MPL v.2.0](LICENSE).
This source code form is distributed under the License is distributed
on an "AS IS" BASIS, WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND,
either express or implied. See the License for the specific language
governing permissions and limitations under the License.

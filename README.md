# Builder

A wrapper for the PHPExcel library to help you quickly build Excel reports.

## Example Usage

```php
$builder = new Builder(
    new \Builder\Builders\PhpSpreadsheet(),
    '/var/cache',
);

$reportArray = [
    'headers' => [
        'Column 1',
        'Column B',
    ],
    'rows' => [
        [
            'Some Data',
            'Some Other Data',
        ],
        [
            'Some Data 2',
            'Some Other Data 2',
        ],
    ],
];

$builder->setSheets($reportArray);

$builder->setCreator('Workflow');
$builder->setTitle('Day Report');
$builder->setSheetTitles(['Data']);
$builder->setDescription('The Workflow Day Report');
$builder->setFilename('Workflow-Day_Report_' . $startDate->format('d_m_Y'));

// use generate() to output headers and force file download.
$builder->generate();

// use generateExcel() to create the file.
$builder->generateExcel();
```

## Development

### Todo

* Allow both caching when building a report as well as short term or perm-caching to a configured location.

## Testing

Minimal tests can be performed with PHPUnit.

### Unit Tests
`composer tests` or  `./vendor/bin/phpunit`

### Code Coverage
`composer coverage`

These will be available in `./builder_coverage`.

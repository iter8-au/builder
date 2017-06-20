# Builder

A wrapper for the PHPExcel and Spout libraries to help you quickly build Excel reports.

## Example Usage

```php
$app['builder.driver']    = 'spout'; // or 'phpexcel'
$app['builder.cache_dir'] = '/var/cache';

$app->register(new BuilderServiceProvider());
// --- OR ---
$app->register(
    new BuilderServiceProvider(),
    [
        'builder.driver'    => 'phpexcel',
        'builder.cache_dir' => '/var/cache',
    ]
);
```

```php
$builder = $app['builder'];

$reportArray = [
    [
        'Column 1' => 'Some Data',
        'Column B' => 'Some Other Data',
    ],
    [
        'Column 1' => 'Some Data 2',
        'Column B' => 'Some Other Data 2',
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

## Feature Parity
Feature | PHPExcel | Spout
------- | -------- | -----
Cell Alignment | Yes | No
Auto-sizing Columns | Yes | No
Custom Column Widths | Yes | No
Document Properties | Yes | No
Header Styling | Yes | Yes
Multiple Sheets | Yes | Yes


## Development

### Todo

* Allow both caching when building a report as well as short term or perm-caching to a configured location

## Testing

Minimal tests can be performed with PHPUnit.

`./vendor/bin/phpunit`

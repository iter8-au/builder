# Builder

A wrapper for the PHP Excel library to help you quickly build reports

## Example Usage

```php
$app['builder.cache_dir'] = '/var/cache';

$app->register(
    new BuilderServiceProvider(),
    [
        'builder.cache_dir' => $app['builder.cache_dir'],
    ]
);
```

```php
$builder = $app['builder'];

$reportArray = [
    [
        'Column 1' => 'Some Data',
        'Column B' => 'Some Other Data'
    ],
    [
            'Column 1' => 'Some Data 2',
            'Column B' => 'Some Other Data 2'
    ]
];

$builder->setSheets($reportArray);

$builder->setColumnWidths([
    0 => 30,
    1 => 15,
]);

$builder->setColumnStyles([
    0 => [
        'alignment' => ReportService::ALIGNMENT_CENTER
    ],
    1 => [
        'alignment' => ReportService::ALIGNMENT_CENTER
    ],
]);

$builder->setCreator('Workflow');
$builder->setTitle('Day Report');
$builder->setSheetTitles(['Data']);
$builder->setDescription("The Workflow Day Report");
$builder->setFilename('Workflow-Day_Report_' . $startDate->format('d_m_Y'));

$builder->generate();
```

## Development

### Todo

* Allow both caching when building a report as well as short term or perm-caching to a configured location

## Testing

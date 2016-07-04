# Builder

A wrapper for the PHP Excel library to help you quickly build reports

## Example Usage


```php
$app->register(
    new BuilderServiceProvider(),
    [
        'builder.cache_directory' => $app['builder.cache_dir'],
    ]
);
```

```php
$report = $app['service.report'];

// Set the Agent and Team Leader Qualifies columns to be wider
$report->setColumnWidths([
    1 => 30,
    2 => 15,
    3 => 15,
    4 => 25,
    5 => 15,
    6 => 20,
    7 => 15,
]);

$report->setColumnStyles([
    0 => [
        'alignment' => ReportService::ALIGNMENT_CENTER
    ],
    1 => [
        'alignment' => ReportService::ALIGNMENT_CENTER
    ],
    2 => [
        'alignment' => ReportService::ALIGNMENT_CENTER
    ],
    3 => [
        'alignment' => ReportService::ALIGNMENT_CENTER
    ],
    4 => [
        'alignment' => ReportService::ALIGNMENT_CENTER
    ],
    5 => [
        'alignment' => ReportService::ALIGNMENT_CENTER
    ],
    6 => [
        'alignment' => ReportService::ALIGNMENT_CENTER
    ],
    7 => [
        'alignment' => ReportService::ALIGNMENT_CENTER
    ],
]);

$report->setData($incentiveReportData);
$report->setCreator('X Sell');
$report->setTitle('Incentive Report');
$report->setDescription(
    sprintf(
        "This is the incentive report for the %s period",
        $reportPeriod
    )
);
$report->setFilename('incentive_report_' . $date . '_' . date('d_m_Y'));

$report->generate();
```


## Development

### Todo

* Allow both caching when building a report as well as short term or perm-caching to a configured location

## Testing

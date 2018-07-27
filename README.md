# Builder

A wrapper for the PhpSpreadsheet and Spout libraries to help you quickly build Excel reports.

## Requirements

- Silex ~2.x

## Example Usage

```php
$app['builder.default']   = 'spout'; // or 'phpspreadsheet'
$app['builder.cache_dir'] = '/var/cache';

$app->register(new BuilderServiceProvider());
// --- OR ---
$app->register(
    new BuilderServiceProvider(),
    [
        'builder.default'   => 'phpspreadsheet',
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

$builder->setCreator('App Name');
$builder->setTitle('My Spreadsheet');
$builder->setSheetTitles(['Sheet 1']);
$builder->setDescription('Spreadsheet that contains some data');
$builder->setFilename('App_Name_Spreadsheet_' . $startDate->format('d_m_Y'));

// use generate() to output headers and force file download.
$builder->generate();

// use generateExcel() to create the file.
$builder->generateExcel();
```

Both Builders are available under the `$app['builders']` key, but `$app['builder']` will be the default builder you specify.

### PhpSpreadsheet
Accessible via `$app['builders']['phpspreadsheet']`.

### Spout
Accessible via `$app['builders']['spout']`.

## Feature Parity
Feature | PhpSpreadsheet | Spout
------- | -------- | -----
Cell Alignment | Yes | No
Auto-sizing Columns | Yes | No
Custom Column Widths | Yes | No
Document Properties | Yes | No
Header Styling | Yes | Yes
Multiple Sheets | Yes | Yes


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

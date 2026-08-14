# SOP Reporter configuration reference

SOP Reporter creates the editable files `config/app_config.yaml` and `config/extraction_rules.yaml` on first start. Restart the tray app after editing either file.

Do not put the Gmail app password in YAML. It belongs only in Windows Credential Manager through the first-run prompt.

## `app_config.yaml`

### Email matching

```yaml
email:
  account: reports.inbox@gmail.com
  imap_host: imap.gmail.com
  imap_port: 993
  mailbox: INBOX
  search:
    from: sender@example.com
    subject_contains: Daily SOP
    since_days: 14
    unread_only: false
  attachments:
    filename_patterns:
      - "*.xlsx"
      - "*.xlsm"
```

- `account` is populated after first-run setup.
- `from` and `subject_contains` may be blank to omit that condition.
- `since_days` limits how far back the IMAP search goes. Deduplication still prevents already-handled messages from printing again.
- `unread_only` should normally remain `false`. Reading is performed with `BODY.PEEK`, so SOP Reporter does not intentionally change Gmail read/unread state.
- Attachment patterns use shell-style matching and are case-insensitive.
- Legacy `.xls` is not supported. Save it as `.xlsx` first.

### Schedule

```yaml
schedule:
  enabled: true
  days: [monday, tuesday, wednesday, thursday, friday]
  time: "07:00"
  polling_seconds: 20
```

Use 24-hour `HH:MM` time. The tray app must be running in the signed-in Windows session.

### Output and printing

```yaml
output:
  downloads_directory: downloads
  reports_directory: reports
  report_filename: "Olympia_SOP_%Y-%m-%d_%H%M%S.xlsx"

printer:
  enabled: true
  name: "Your Printer Name on Ne01:"
  paper_size: tabloid
  orientation: landscape
  copies: 1
  fit_to_pages_wide: 1
  fit_to_pages_tall: 0
```

Relative output folders are created in SOP Reporter's writable runtime folder. Absolute paths are also accepted.

The printer name must match the name Excel shows. Excel sometimes includes a port suffix such as `on Ne01:`. Leave it blank only if the Windows default printer is guaranteed to be correct.

`paper_size` is intentionally restricted to `tabloid`. Orientation may be `landscape` or `portrait`. A value of `0` for `fit_to_pages_tall` means automatic height.

### Software updates

```yaml
update:
  enabled: true
  repository: matijepekovic/SOP
  check_on_startup: true
  include_prereleases: false
```

SOP Reporter can update itself from the desktop. **Check for Updates** appears both in the tray menu and in the control window, and installs a newer build in place.

| Key | Meaning |
| --- | --- |
| `enabled` | Set to `false` to hide the update controls entirely. |
| `repository` | The `owner/name` GitHub repository holding the published releases. |
| `check_on_startup` | Check quietly a couple of seconds after launch. A newer version is reported in the control window; nothing is installed without a click. |
| `include_prereleases` | Offer prerelease builds as well as final ones. |

Updates come from the repository's GitHub Releases, so a release must exist with `SOPReporter.exe` attached. Pushing a `v*` tag builds and publishes one automatically; see `.github/workflows/windows-build.yml`. The tag and `sop_reporter.__version__` must match, and the workflow fails the build if they disagree, because the updater compares those two values to decide whether a build is newer.

Nothing is downloaded until you press **Check for Updates**, and nothing is installed until you confirm. An update is refused while a fetch/report/print run is in progress.

## `extraction_rules.yaml`

Rules are applied in this order:

1. Select a worksheet and header row.
2. Map and convert columns.
3. Apply row filters.
4. Split matching rows into separate reports when enabled.
5. Optionally group and aggregate within each report.
6. Sort final output rows.
7. Build, style, and print every report separately.

### Input layout

```yaml
input:
  sheet_name: Data
  header_row: 1
  data_start_row: 2
  stop_at_first_blank_row: false
  skip_blank_rows: true
```

Set `sheet_name` to `null` to use the first active worksheet. Header matching is case-insensitive and ignores surrounding spaces.

### Column mapping and conversion

```yaml
columns:
  - source: Sales Rep
    target: Representative
    type: text
    required: true
    include_in_report: true
    fill_down: false
    width: 24
  - source: Amount
    target: Amount
    type: currency
    required: true
    number_format: '$#,##0.00'
    width: 16
```

Supported types:

- `text`
- `integer`
- `number`
- `currency`
- `percent`
- `date`
- `datetime`
- `boolean`

`required: true` means the source header must exist. For an optional source, use `required: false` and optionally set `default`.

Set `include_in_report: false` for a mapped helper field used only for filtering or splitting. Set `fill_down: true` when a grouped export writes a value only on the first row of a section and leaves the following detail cells blank. Fill-down resets at each worksheet and ignores subtotal/total labels as new group values.

Currency strings such as `$1,250.00` and `(300.00)` are converted to numeric Excel values. Percent strings containing `%` are divided by 100. Supported text dates include `YYYY-MM-DD`, `MM/DD/YYYY`, `MM/DD/YY`, and `DD-Mon-YYYY`; true Excel date cells are preferred.

### Filters

```yaml
filters:
  mode: all
  rules:
    - column: Status
      operator: equals
      value: Approved
      case_sensitive: false
    - column: Amount
      operator: gte
      value: 1000
```

`mode: all` requires every rule to pass. `mode: any` requires at least one.

Supported operators:

| Operator | Configuration |
| --- | --- |
| `equals`, `not_equals` | `value` |
| `contains`, `not_contains` | `value` |
| `in`, `not_in` | `values: [A, B, C]` |
| `range` | `min` and/or `max`, inclusive |
| `gt`, `gte`, `lt`, `lte` | `value` |
| `is_blank`, `not_blank` | no value |

Filters may reference either a source header or a mapped target name. Text comparisons are case-insensitive unless `case_sensitive: true` is set.

The bundled rules require `Market` to equal `Olympia` and require a nonblank `Job Number`, which excludes the Salesforce subtotal and grand-total rows shown in the screenshot.

### Separate report per Sub Status

```yaml
split:
  enabled: true
  by: Sub Status
  title_template: 'OLYMPIA — {value}'
  filename_suffix: '_{value}'
  include_blank: true
  blank_label: No Sub Status
```

Splitting happens after filters and before optional grouping. Each distinct value becomes a separate workbook and a separate Excel print job. The status list is dynamic, so a new Sub Status does not require a code change. Unsafe filename characters are replaced automatically. Templates may use `{base_title}` and `{value}`.

### Grouping and aggregation

```yaml
grouping:
  enabled: true
  by:
    - Representative
    - Branch
  aggregations:
    - source: Job Number
      target: Approved Jobs
      operation: count
      number_format: '0'
    - source: Amount
      target: Total Sales
      operation: sum
      number_format: '$#,##0.00'
```

Grouping always references mapped target names. Supported operations are `sum`, `count`, `count_rows`, `avg`, `min`, and `max`. `count_rows` does not require `source`; every other operation does.

If grouping is disabled, every matching detail row is written to the report.

### Sorting

```yaml
sort:
  - column: Total Sales
    direction: desc
  - column: Representative
    direction: asc
```

Sorting references final output headers. Blank values always sort last.

### Report layout and styles

The `report` section controls the worksheet name, title, header row, filters, freeze pane, Tabloid page setup, fonts, colors, widths, formats, and alternating body fill. Colors are six-character RGB hex values without `#`.

The final Excel COM print step forces the configured Tabloid size and orientation again immediately before printing. This protects against printer-driver defaults overriding the workbook's saved page setup.

## Formula-source limitation

Extraction uses `openpyxl` with cached formula results. It does not calculate formulas. If a source attachment contains formulas, the sending system must save the workbook after Excel has calculated it so cached values are present.

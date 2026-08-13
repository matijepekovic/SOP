# SOP Reporter

SOP Reporter is a Windows tray application that downloads matching Excel attachments from Gmail, transforms them with editable YAML rules, builds a formatted workbook, and prints it through Microsoft Excel on forced 11 x 17 (Tabloid) paper.

The same pipeline runs automatically at 7:00 AM Monday through Friday and from the tray icon's **Run Now** command.

## What is implemented

- Direct Gmail IMAP access using a Gmail app password stored in Windows Credential Manager.
- Rule-driven column mapping, type conversion, filtering, grouping, aggregation, and sorting.
- An Olympia market filter with one separate workbook and Excel print job for every Sub Status found in the attachment.
- Styled `.xlsx` report generation without changing the source attachment.
- Excel COM printing from an isolated `EXCEL.EXE` instance with Tabloid size and configured orientation forced immediately before printing.
- Atomic state tracking that prevents an email from being printed twice.
- A tray menu with Run Now, Open Reports, Open Settings, Open Logs, and Exit.
- A weekday scheduler that runs inside the tray process.
- Frozen-executable-safe paths with a `%APPDATA%\SOPReporter` fallback when the executable folder is not writable.

## First run

1. Copy `SOPReporter.exe` into its permanent folder.
2. Start it once while signed into the Windows account that will run it.
3. Enter the Gmail address and its 16-character Gmail app password when prompted.
4. Right-click the tray icon and choose **Open Settings**.
5. Edit `app_config.yaml` to set the email sender/subject filters and printer name.
6. Review `extraction_rules.yaml`. Its defaults match the supplied Salesforce screenshot; update the sheet/header details after a real source workbook is available.
7. Use **Run Now** with **Microsoft Print to PDF** first.

The password is never written to YAML or a log. It is stored through `keyring` in Windows Credential Manager.

## Gmail setup

The Gmail account must have two-step verification enabled before an app password can be created. In the Google Account security settings, create an app password for SOP Reporter and enter that value in the first-run dialog. Regular Gmail passwords are not accepted by Gmail IMAP when app-password authentication is required.

## Configuration

Bundled files ending in `.default.yaml` are immutable templates. On first start, SOP Reporter copies them to writable files named:

- `config/app_config.yaml`
- `config/extraction_rules.yaml`

Edit the files without `.default` in their names. Restart SOP Reporter after a configuration change.

`app_config.yaml` controls Gmail matching, schedule, output folders, printer, and logging. `extraction_rules.yaml` controls the workbook sheet/header, source-to-output columns, filters, report splitting, grouping, aggregations, sorting, formats, widths, and report styling.

The included extraction rules follow the Salesforce layout visible in the supplied screenshot. They keep only `Market = Olympia`, remove subtotal rows, and dynamically create one report per distinct `Sub Status`. The exact sheet name, header row, and source headers still need confirmation against the first real emailed workbook.

## Run from source

Use 64-bit Python on Windows when the installed Microsoft Excel is 64-bit.

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
python -m sop_reporter
```

For a safe first test, set `printer.name` to the exact Excel printer name for Microsoft Print to PDF.

## Tests

The unit suite is cross-platform and does not contact Gmail or start Excel:

```powershell
python -m unittest discover -s tests -v
```

The tests cover extraction/filtering/grouping, Olympia filtering, per-status report splitting, report layout and styles, atomic state, MIME attachment handling with mocked IMAP, and pipeline deduplication.

## Build the executable

Run the build on Windows. PyInstaller does not cross-compile a Windows executable from Linux or macOS.

```powershell
.\scripts\build_windows.ps1
```

Outputs:

- `dist\SOPReporter-Debug.exe` — console visible for initial troubleshooting.
- `dist\SOPReporter.exe` — final tray-only release build.

The release file is a single executable. Default YAML and the tray icon are embedded and copied to the writable runtime folder on first launch.

If the project is hosted on GitHub, the included **Build Windows executables** workflow runs the tests and produces both files on a Windows runner. It can be started manually from the Actions tab.

## Start automatically at logon

Excel COM needs the user's interactive desktop. Configure a logon task, not a Windows service.

1. Open **Task Scheduler** and select **Create Task**.
2. On **General**, name it `SOP Reporter`, select the intended user, choose **Run only when user is logged on**, and leave **Run with highest privileges** unchecked.
3. On **Triggers**, add **At log on** for that user.
4. On **Actions**, add **Start a program**. Select `SOPReporter.exe`, and set **Start in** to the folder containing the executable.
5. On **Settings**, select **If the task is already running: Do not start a new instance**.
6. Save the task, sign out, sign back in, and confirm the tray icon appears.

The app remains resident. Its internal scheduler performs the 7:00 AM weekday run.

## Windows verification checklist

Before connecting the production printer:

1. Send a matching email with `tests/fixtures/salesforce_olympia_sample.xlsx` attached.
2. Select Microsoft Print to PDF in `app_config.yaml`.
3. Run **Run Now** and verify three separate reports and print jobs: Item Notification, Install Issue, and On Hold. Each should be landscape 11 x 17.
4. Run it again and verify the email is skipped.
5. Leave an unrelated workbook open in Excel and verify it remains open and unchanged.
6. Repeat several runs and confirm Task Manager has no orphaned `EXCEL.EXE` created by SOP Reporter.
7. Test the release executable on a clean Windows VM without Python installed.
8. Place the executable in a non-writable folder and confirm runtime data appears under `%APPDATA%\SOPReporter`.

Logs rotate in the runtime `logs` folder. Generated workbooks remain in `reports`, and downloaded source attachments remain in `downloads` for traceability.

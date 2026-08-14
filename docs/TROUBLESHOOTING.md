# SOP Reporter troubleshooting

## Start with the log

Right-click the tray icon and choose **Open Logs**. The current file is `sop_reporter.log`; older files have numeric suffixes.

Passwords are not written to logs.

## Gmail rejects login

- Confirm the address in `config/app_config.yaml` is correct.
- Use a Google app password, not the normal Gmail password.
- Remove spaces when pasting the app password; SOP Reporter also strips them automatically.
- Confirm IMAP access is permitted for that Google account or Workspace policy.
- To replace a saved password, open Windows **Credential Manager**, remove the entry named `SOPReporter Gmail`, and restart SOP Reporter.

## No matching email is found

- Temporarily reduce the filters: blank `from` or `subject_contains` one at a time.
- Increase `since_days`.
- Confirm the correct `mailbox` is configured.
- Keep `unread_only: false` unless only unread messages should ever qualify.
- Check that the attachment extension matches `filename_patterns`.

## Workbook extraction fails

- Confirm `sheet_name`, `header_row`, and source column names against the actual file.
- Header names are case-insensitive, but the words must otherwise match.
- Legacy `.xls` is unsupported; convert it to `.xlsx`.
- If formula cells extract as blank, the source workbook was probably saved without cached calculated values.

## A Sub Status is missing or mixed with another

- Confirm `split.enabled` is `true` and `split.by` exactly matches the mapped target `Sub Status`.
- Keep `fill_down: true` on both `Market` and `Sub Status` for grouped Salesforce exports that leave repeated cells blank.
- Confirm detail rows have a `Job Number`; the default nonblank filter intentionally removes subtotal and grand-total rows.
- If every row is missing, verify the source uses `Olympia` in the Market column and confirm the configured header row.

## Printing fails

1. Open the generated report manually in Excel to confirm Excel can read it.
2. In Excel, inspect the exact printer label and copy that label into `printer.name`.
3. Test first with Microsoft Print to PDF.
4. Confirm Excel and Python/PyInstaller use compatible 32-bit or 64-bit architecture.
5. Confirm the printer driver supports 11 x 17 Tabloid paper.

SOP Reporter never attaches to the user's already-open Excel window. It creates a separate hidden instance and closes only that instance.

## The printer asks for a Department Name and Password

Konica Minolta devices such as the C360i series can be configured with **User Authentication / Account Track**. When that is enabled, the driver raises a modal dialog asking for a Department Name and Password on every print job.

Entering the credentials in that popup is enough for a print you start by hand, but it will stall the scheduled 7:00 AM run: the dialog waits for a person who is not there, and the job never reaches the printer. The credentials have to be saved into the driver instead, so no dialog is raised at all.

Save them in **both** places, because they are stored separately:

1. **Printing Defaults** — this is the one that matters for scheduled runs, because it supplies the values to jobs that Excel spawns through COM automation.
   - Control Panel > Devices and Printers
   - Right-click the Konica Minolta printer > **Printer properties**
   - **Advanced** tab > **Printing Defaults...**
   - Open the **User Authentication/Account Track** settings
   - Leave **Department Name** blank, enter the account password (for example `0000`), and confirm with **Verify** if the driver offers it
   - Disable any "prompt for credentials" or "display dialog on each job" option
2. **Printing Preferences** — the per-user copy, used when someone prints interactively.
   - Same printer > **Printing preferences** > repeat the same entries

Then confirm it worked:

- Set `printer.enabled: true` and leave `printer.name` blank (or set it to the exact label Excel shows).
- Use the tray **Run Now** action, then step away from the keyboard. If the report reaches the printer with no dialog appearing, unattended printing is correctly configured.
- If a dialog still appears, the driver saved the credentials only to the interactive profile. Re-check the **Printing Defaults** path above, which is the copy the automated run reads.

## An uncertain print is not retried

SOP Reporter claims an email in `state.json` immediately before asking Excel to print. If Excel definitely fails before `PrintOut`, the claim is released. If Excel accepts `PrintOut` and then returns an error, or the PC shuts down during printing, SOP Reporter keeps the claim so the same email cannot print twice.

Check the printer queue and generated report before changing state. If a qualified operator confirms that nothing printed, close SOP Reporter, back up `state.json`, remove only that message's entry, and restart the app. Do not delete the entire state file unless every prior matching email is safely outside `since_days`; otherwise old reports can print again.

## Tray icon does not appear after logon

- Open Task Scheduler and run the SOP Reporter task manually.
- Confirm **Run only when user is logged on** is selected.
- Confirm **Run with highest privileges** is unchecked.
- Confirm **Start in** is the folder containing `SOPReporter.exe`.
- Check `sop_reporter.log` under the executable folder or `%APPDATA%\SOPReporter\logs`.

## Runtime files moved to AppData

If the executable folder is not writable, SOP Reporter stores editable configuration, logs, state, downloads, and reports under `%APPDATA%\SOPReporter`. This is expected for locations such as `Program Files`.

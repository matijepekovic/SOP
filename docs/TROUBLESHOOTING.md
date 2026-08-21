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

## Changing the Gmail sign-in

Use **Change Gmail Sign-in**, in the tray menu or beside the other buttons at the top of the control window. It asks for the address and app password again.

The new password is tested against Gmail before anything is saved. If Gmail refuses it the dialog reopens and says why, and the credential already stored is left untouched — so a failed attempt can never lock you out of a working setup. The same check runs during first-time setup, so a mistyped app password is refused on the spot rather than saved.

Nothing is written to the configuration file. The password lives only in Windows Credential Manager, under `SOPReporter Gmail`.

If the application will not start at all and you need to clear the credential by hand: **Credential Manager** → **Windows Credentials** → remove the `SOPReporter Gmail` entry, then start SOP Reporter and the setup dialog reappears.

## Updating from the desktop

Use **Check for Updates**, in the tray menu or at the bottom of the control window. If a newer build exists, **Install and Restart** downloads it, swaps it in, and relaunches SOP Reporter.

The swap works by renaming: Windows refuses to overwrite a running executable but does allow it to be renamed, so the outgoing build is moved aside to `SOPReporter.previous-<version>.exe` and deleted on the next start. If you ever need to go back, exit SOP Reporter, delete `SOPReporter.exe`, and rename that file back.

Things that stop an update:

- **"running from source"** — self-update applies only to the packaged `SOPReporter.exe`, not to `python -m sop_reporter.main`.
- **"the folder holding SOPReporter.exe is not writable"** — the application is somewhere locked down such as `Program Files`. Move it to a writable location such as the Desktop. Note that this is separate from where settings and logs live; those already fall back to `%APPDATA%\SOPReporter`.
- **"No published release was found"** — no GitHub release exists yet with `SOPReporter.exe` attached. Push a `v*` tag to build and publish one.
- **"GitHub is rate limiting update checks"** — anonymous GitHub API calls are capped per hour. Wait and try again.
- **A run is in progress** — updating is refused mid-run so a report is never interrupted halfway. Wait for the run to finish.

If the new build does not reappear after an update, start it from its shortcut. The relaunched process waits up to 45 seconds for the outgoing one to exit before claiming the single-instance lock, so a slow exit delays rather than blocks it.

## Runtime files moved to AppData

If the executable folder is not writable, SOP Reporter stores editable configuration, logs, state, downloads, and reports under `%APPDATA%\SOPReporter`. This is expected for locations such as `Program Files`.

# Backup and recovery

## Back up

1. In Google Sheets, choose **File → Make a copy** and include a date in the copy name.
2. For an offline snapshot, choose **File → Download → Microsoft Excel (.xlsx)**.
3. Keep the Apps Script deployment URL and the Git commit or release used by that installation with the backup record.
4. Back up before a code update, schema change, or bulk import.

## Recover

1. Stop entering new records in the affected sheet.
2. Make a copy of the affected sheet before attempting repairs.
3. Restore the newest known-good Google Sheet copy, or import the offline workbook into a new Sheet.
4. Confirm that the Apps Script project is bound to the intended Sheet and deploy the tested code version.
5. Verify Settings, user access, one client, one order, one packing slip, and PDF output before reopening the workflow to staff.

The application has no Bifro-hosted operating database to restore. The customer's Google Workspace backup and access policies are therefore the recovery boundary.

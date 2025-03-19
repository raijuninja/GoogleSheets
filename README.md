# GoogleSheets PowerShell Module

This module provides functions to interact with Google Sheets using a Google Service Account. Below are the available functions and their usage.

## Functions

### `Get-GoogleServiceAccountToken`

Retrieves an access token for a Google Service Account.

#### Get-GoogleServiceAccountToken Parameters

- `-CredentialsPath` (Mandatory): Path to the service account key JSON file.

#### Get-GoogleServiceAccountToken Example

```powershell
Get-GoogleServiceAccountToken -CredentialsPath "path/to/service-account-key.json"
```

---

### `New-ErrorTrackingSpreadsheet`

Creates a new Google Sheets spreadsheet for error tracking.

#### New-ErrorTrackingSpreadsheet Parameters

- `-CredentialsPath` (Mandatory): Path to the service account key JSON file.
- `-ErrorSheetName` (Optional): Name of the sheet for error tracking. Default is "Errors".

#### New-ErrorTrackingSpreadsheet Example

```powershell
New-ErrorTrackingSpreadsheet -CredentialsPath "path/to/service-account-key.json" -ErrorSheetName "CustomErrorSheet"
```

---

### `Write-ErrorToGoogleSheet`

Writes an error entry to a specified Google Sheets spreadsheet.

#### Parameters

- `-CredentialsPath` (Mandatory): Path to the service account key JSON file.
- `-SpreadsheetId` (Mandatory): ID of the target spreadsheet.
- `-ErrorSheetName` (Mandatory): Name of the sheet to write the error to.
- `-Integration` (Mandatory): Name of the integration where the error occurred.
- `-ErrorType` (Mandatory): Type of the error.
- `-ErrorMessage` (Mandatory): Description of the error.
- `-LogFilePath` (Mandatory): Path to the log file associated with the error.

#### Example

```powershell
Write-ErrorToGoogleSheet -CredentialsPath "path/to/service-account-key.json" `
    -SpreadsheetId "spreadsheet-id" `
    -ErrorSheetName "Errors" `
    -Integration "IntegrationName" `
    -ErrorType "Critical" `
    -ErrorMessage "An unexpected error occurred." `
    -LogFilePath "path/to/logfile.log"
```

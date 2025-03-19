# Implement your module commands in this script.


# Export only the functions using PowerShell standard verb-noun naming.
# Be sure to list each exported functions in the FunctionsToExport field of the module manifest file.
# This improves performance of command discovery in PowerShell.
Export-ModuleMember -Function *-*

# Function to get Google Service Account Token
function Get-GoogleServiceAccountToken {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$CredentialsPath
    )

    try {
        # Load the service account key file
        Write-Host "Loading service account key from: $CredentialsPath" -ForegroundColor Cyan
        $serviceAccountKey = Get-Content $CredentialsPath | ConvertFrom-Json
        Write-Host "Service account email: $($serviceAccountKey.client_email)" -ForegroundColor Cyan
        Write-Host "Project ID: $($serviceAccountKey.project_id)" -ForegroundColor Cyan

        # Create JWT header
        $header = @{
            alg = "RS256"
            typ = "JWT"
        } | ConvertTo-Json -Compress

        # Create JWT claim set
        $now = [DateTimeOffset]::Now.ToUnixTimeSeconds()
        $claimset = @{
            iss   = $serviceAccountKey.client_email
            scope = "https://www.googleapis.com/auth/spreadsheets"
            aud   = "https://oauth2.googleapis.com/token"
            exp   = $now + 3600
            iat   = $now
        } | ConvertTo-Json -Compress

        # Encode header and claim set
        $encodedHeader = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($header)).Replace('/', '_').Replace('+', '-').TrimEnd('=')
        $encodedClaimSet = [Convert]::ToBase64String([System.Text.Encoding]::UTF8.GetBytes($claimset)).Replace('/', '_').Replace('+', '-').TrimEnd('=')

        # Process the private key
        $privateKey = $serviceAccountKey.private_key.Replace("-----BEGIN PRIVATE KEY-----", "").Replace("-----END PRIVATE KEY-----", "").Replace("`n", "").Replace("`r", "")
        $keyBytes = [Convert]::FromBase64String($privateKey)
        $rsa = [System.Security.Cryptography.RSA]::Create()
        $rsa.ImportPkcs8PrivateKey($keyBytes, [ref]$null)

        # Sign the data
        $dataToSign = [System.Text.Encoding]::UTF8.GetBytes("$encodedHeader.$encodedClaimSet")
        $signature = $rsa.SignData($dataToSign, [System.Security.Cryptography.HashAlgorithmName]::SHA256, [System.Security.Cryptography.RSASignaturePadding]::Pkcs1)

        # Encode signature
        $encodedSignature = [Convert]::ToBase64String($signature).Replace('/', '_').Replace('+', '-').TrimEnd('=')

        # Create JWT token
        $jwt = "$encodedHeader.$encodedClaimSet.$encodedSignature"

        # Exchange JWT for access token
        $tokenParams = @{
            grant_type = "urn:ietf:params:oauth:grant-type:jwt-bearer"
            assertion  = $jwt
        }
        $response = Invoke-RestMethod -Uri "https://oauth2.googleapis.com/token" -Method Post -Body $tokenParams
        Write-Host "Successfully obtained access token" -ForegroundColor Green
        return $response.access_token
    }
    catch {
        Write-Host "Failed to get service account token: $($_.Exception.Message)" -ForegroundColor Red
        throw
    }
}

# Function to create a new error tracking spreadsheet
function New-ErrorTrackingSpreadsheet {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$CredentialsPath,

        [Parameter(Mandatory = $false)]
        [string]$ErrorSheetName = "Errors"
    )

    try {
        # Get service account token
        $accessToken = Get-GoogleServiceAccountToken -CredentialsPath $CredentialsPath

        # Create a new spreadsheet
        $createParams = @{
            Uri         = "https://sheets.googleapis.com/v4/spreadsheets"
            Method      = "POST"
            Headers     = @{
                Authorization = "Bearer $accessToken"
            }
            Body        = @{
                properties = @{
                    title = "PowerSchool Integration Error Tracking"
                }
                sheets     = @(
                    @{
                        properties = @{
                            title          = $ErrorSheetName
                            gridProperties = @{
                                frozenRowCount = 1
                            }
                        }
                    }
                )
            } | ConvertTo-Json -Depth 10
            ContentType = "application/json"
        }

        $createResponse = Invoke-RestMethod @createParams
        Write-Host "New error tracking spreadsheet created with ID: $($createResponse.spreadsheetId)" -ForegroundColor Green
        return $createResponse.spreadsheetId
    }
    catch {
        Write-Host "Failed to create error tracking spreadsheet: $($_.Exception.Message)" -ForegroundColor Red
        throw
    }
}

# Function to write an error to Google Sheets
function Write-ErrorToGoogleSheet {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$CredentialsPath,

        [Parameter(Mandatory = $true)]
        [string]$SpreadsheetId,

        [Parameter(Mandatory = $true)]
        [string]$ErrorSheetName,

        [Parameter(Mandatory = $true)]
        [string]$Integration,

        [Parameter(Mandatory = $true)]
        [string]$ErrorType,

        [Parameter(Mandatory = $true)]
        [string]$ErrorMessage,

        [Parameter(Mandatory = $true)]
        [string]$LogFilePath
    )

    try {
        # Get service account token
        $accessToken = Get-GoogleServiceAccountToken -CredentialsPath $CredentialsPath

        # Format timestamp
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

        # Append error data to the sheet
        $rowValues = , @(
            $timestamp,
            $Integration,
            $ErrorType,
            $ErrorMessage,
            "New",
            $LogFilePath,
            "",
            ""
        )

        $appendBody = @{
            range          = "$ErrorSheetName!A:H"
            majorDimension = "ROWS"
            values         = $rowValues
        }

        $appendParams = @{
            Uri         = "https://sheets.googleapis.com/v4/spreadsheets/$SpreadsheetId/values/$ErrorSheetName!A:H:append?valueInputOption=USER_ENTERED&insertDataOption=INSERT_ROWS"
            Method      = "POST"
            Headers     = @{
                Authorization = "Bearer $accessToken"
            }
            Body        = $appendBody | ConvertTo-Json -Depth 10
            ContentType = "application/json"
        }

        $appendResponse = Invoke-RestMethod @appendParams
        Write-Host "Error data appended successfully at range: $($appendResponse.updates.updatedRange)" -ForegroundColor Green
        return $true
    }
    catch {
        Write-Host "Failed to write error to Google Sheet: $($_.Exception.Message)" -ForegroundColor Red
        return $false
    }
}

# Export functions
Export-ModuleMember -Function Get-GoogleServiceAccountToken, New-ErrorTrackingSpreadsheet, Write-ErrorToGoogleSheet
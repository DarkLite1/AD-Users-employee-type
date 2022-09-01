#Requires -Version 5.1
#Requires -Modules Toolbox.HTML, Toolbox.EventLog, ImportExcel

<# 
    .SYNOPSIS   
        Report about all the user accounts in a specific OU that have a specific EmployeeType.

    .DESCRIPTION
        Report about all the users in AD that have the text 'TEMP' in the field 'EmployeeID'.
        This Excel report will is sent by mail to the end user and saved in the log folder.
        This script is specifically designed for GBR.

    .PARAMETER ImportFile
        A simple .JSON file that contains the parameters for the script.
        See Example.json
    
    .PARAMETER OU
        List of organizational units where to search for AD user accounts.

    .PARAMETER MailTo
        List of e-mail addresses that wil receive the report.

    .PARAMETER EmployeeType
        List of employee types to search for.

    .PARAMETER LogFolder
        Location for the log files.
#>

Param (
    [Parameter(Mandatory)]
    [String]$ScriptName,
    [Parameter(Mandatory)]
    [String]$ImportFile,
    [String]$LogFolder = "$env:POWERSHELL_LOG_FOLDER\AD Reports\AD Users\$ScriptName",
    [String[]]$ScriptAdmin = $env:POWERSHELL_SCRIPT_ADMIN
)

Begin {
    Try {
        Get-ScriptRuntimeHC -Start
        Import-EventLogParamsHC -Source $ScriptName
        Write-EventLog @EventStartParams

        #region Logging
        try {
            $logParams = @{
                LogFolder    = New-Item -Path $LogFolder -ItemType 'Directory' -Force -ErrorAction 'Stop'
                Name         = $ScriptName
                Date         = 'ScriptStartTime'
                NoFormatting = $true
            }
            $logFile = New-LogFileNameHC @LogParams
        }
        Catch {
            throw "Failed creating the log folder '$LogFolder': $_"
        }
        #endregion

        #region Import .json file
        $M = "Import .json file '$ImportFile'"
        Write-Verbose $M; Write-EventLog @EventOutParams -Message $M

        $file = Get-Content $ImportFile -Raw -EA Stop | ConvertFrom-Json
        #endregion

        #region Test .json file properties
        if (-not ($OU = $file.OU)) {
            throw "Input file '$ImportFile': Property 'OU' not found."
        }
        if (-not ($MailTo = $file.MailTo)) {
            throw "Input file '$ImportFile': Property 'MailTo' not found."
        }
        if (-not ($EmployeeType = $file.EmployeeType)) {
            throw "Input file '$ImportFile': Property 'EmployeeType' not found."
        }
        #endregion
    }
    Catch {
        Write-Warning $_
        Send-MailHC -To $ScriptAdmin -Subject 'FAILURE' -Priority 'High' -Message $_ -Header $ScriptName
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"
        Write-EventLog @EventEndParams; Exit 1
    }
}
Process {
    Try {
        #region Create search filter
        $filter = (
            $EmployeeType | ForEach-Object {
                "(employeeType -eq '{0}')" -f $_
            }
        ) -join ' -or '

        $M = "Get-ADUser filter '{0}'" -f $filter
        Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M
        #endregion

        #region Get users
        [Array]$users = foreach ($o in $OU) {
            $M = "Get user from OU '{0}'" -f $o
            Write-Verbose $M; Write-EventLog @EventVerboseParams -Message $M

            Get-ADUser -Filter $filter -SearchBase $o -Properties whenCreated, displayName, title, department, company, manager, EmployeeID, 
            extensionAttribute8, employeeType, CanonicalName, 
            description, co, Office, OfficePhone, 
            HomePhone, MobilePhone, ipPhone, Fax, pager, info, EmailAddress, 
            scriptPath, homeDirectory, AccountExpirationDate, LastLogonDate, 
            PasswordExpired, PasswordNeverExpires, LockedOut |
            Select-Object -Property @{
                Name       = 'CreationDate'
                Expression = { $_.whenCreated } 
            }, 
            DisplayName, Name, SamAccountName,
            @{
                Name       = 'LastName'
                Expression = { $_.Surname } 
            }, 
            @{
                Name       = 'FirstName'
                Expression = { $_.GivenName } 
            }, 
            Title, Department, Company,
            @{
                Name       = 'Manager'
                Expression = { 
                    if ($_.manager) { Get-ADDisplayNameHC $_.manager }
                }
            }, 
            EmployeeID,
            @{
                Name       = 'HeidelbergCementBillingID'
                Expression = { $_.extensionAttribute8 } 
            },
            employeeType,
            @{
                Name       = 'OU'
                Expression = {
                    ConvertTo-OuNameHC $_.CanonicalName
                }
            },
            Description,
            @{
                Name       = 'Country'
                Expression = { $_.co } 
            },
            Office, OfficePhone, HomePhone, MobilePhone, ipPhone, Fax, pager,
            @{
                Name       = 'Notes'
                Expression = { $_.info -replace "`n", ' ' } 
            },
            EmailAddress,
            @{
                Name       = 'LogonScript'
                Expression = { $_.scriptPath } 
            }, 
            @{
                Name       = 'TSUserProfile'
                Expression = {
                    Get-ADTsProfileHC $_.DistinguishedName 'UserProfile' 
                } 
            }, 
            @{
                Name       = 'TSHomeDirectory'
                Expression = { 
                    Get-ADTsProfileHC $_.DistinguishedName 'HomeDirectory' 
                }
            }, 
            @{
                Name       = 'TSHomeDrive'
                Expression = {
                    Get-ADTsProfileHC $_.DistinguishedName 'HomeDrive'
                }
            }, 
            @{
                Name       = 'TSAllowLogon'
                Expression = {
                    Get-ADTsProfileHC $_.DistinguishedName 'AllowLogon'
                }
            },
            HomeDirectory, AccountExpirationDate, LastLogonDate, PasswordExpired, 
            PasswordNeverExpires, LockedOut, Enabled
        }
        #endregion
    }
    Catch {
        Write-Warning $_
        Send-MailHC -To $ScriptAdmin -Subject 'FAILURE' -Priority 'High' -Message $_ -Header $ScriptName
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"
        Write-EventLog @EventEndParams; Exit 1
    }
}

End {
    Try {
        $mailParams = @{
            To        = $MailTo
            Bcc       = $ScriptAdmin
            Subject   = '{0} account{1}' -f
            $users.Count,
            $(if ($users.Count -ne 1) { 's' })
            LogFolder = $LogParams.LogFolder
            Header    = $ScriptName
            Save      = "$LogFile - Mail.html"
        }

        #region Export to Excel
        if ($users) {
            $excelParams = @{
                Path               = "$logFile - Log.xlsx"
                AutoSize           = $true
                FreezeTopRow       = $true
                TableName          = 'Users'
                WorkSheetName      = 'Users'
                NoNumberConversion = @(
                    'Employee ID', 'OfficePhone', 'HomePhone', 
                    'MobilePhone', 'ipPhone', 'Fax', 'Pager'
                )
            }
            $users | Export-Excel @excelParams

            $mailParams.Attachments = $excelParams.Path
        }
        #endregion

        #region Create HTML summary table
        $summaryTable = '<table>
            <tr><th>EmployeeType</th><th>Accounts</th></tr>'
        
        $summaryTable += foreach (
            $type in 
            $EmployeeType | Sort-Object
        ) {
            '<tr><td>{0}</td><td>{1}</td></tr>' -f $type,
            $(
                ($users | Where-Object { $_.employeeType -eq $type } |
                Measure-Object).Count
            )
        }

        $summaryTable += '</table>'
        #endregion
        
        #region Send e-mail
        $mailParams.Message = "Found a total of <b>{0}</b> for the following employee types in the active directory:
        $summaryTable
        {1}
        {2}" -f
        $(
            $mailParams.Subject
        ),
        $(
            if ($mailParams.Attachments) {
                '<p><i>* Check the attachments for details</i></p>'
            }
        ),
        $(
            $OU | ConvertTo-OuNameHC -OU | Sort-Object |
            ConvertTo-HtmlListHC -Header 'Organizational units:'
        )
        
        Get-ScriptRuntimeHC -Stop
        Send-MailHC @MailParams
        #endregion
    }
    Catch {
        Write-Warning $_
        Send-MailHC -To $ScriptAdmin -Subject 'FAILURE' -Priority 'High' -Message $_ -Header $ScriptName
        Write-EventLog @EventErrorParams -Message "FAILURE:`n`n- $_"; Exit 1
    }
    Finally {
        Write-EventLog @EventEndParams
    }
}
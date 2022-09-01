#Requires -Modules Pester
#Requires -Version 5.1

BeforeAll {
    $testOutParams = @{
        FilePath = (New-Item "TestDrive:/Test.json" -ItemType File).FullName
        Encoding = 'utf8'
    }

    $testScript = $PSCommandPath.Replace('.Tests.ps1', '.ps1')
    $testParams = @{
        ScriptName = 'Test (Brecht)'
        ImportFile = $testOutParams.FilePath
        LogFolder  = New-Item 'TestDrive:/log' -ItemType Directory
    }

    Mock Get-ADDisplayNameHC
    Mock Get-ADUser
    Mock Get-ADTSProfileHC
    Mock Send-MailHC
    Mock Write-EventLog
}
Describe 'the mandatory parameters are' {
    It '<_>' -ForEach 'ScriptName', 'ImportFile' {
        (Get-Command $testScript).Parameters[$_].Attributes.Mandatory | 
        Should -BeTrue
    }
}
Describe 'send an e-mail to the admin when' {
    BeforeAll {
        $MailAdminParams = {
            ($To -eq $ScriptAdmin) -and ($Priority -eq 'High') -and 
            ($Subject -eq 'FAILURE')
        }    
    }
    It 'the log folder cannot be created' {
        $testNewParams = $testParams.clone()
        $testNewParams.LogFolder = 'xxx::\notExistingLocation'

        .$testScript @testNewParams

        Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
            (&$MailAdminParams) -and 
            ($Message -like '*Failed creating the log folder*')
        }
    }
    Context 'the ImportFile' {
        It 'is not found' {
            $testNewParams = $testParams.clone()
            $testNewParams.ImportFile = 'nonExisting.json'
    
            .$testScript @testNewParams
    
            Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                (&$MailAdminParams) -and 
                ($Message -like "Cannot find path*nonExisting.json*")
            }
            Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                $EntryType -eq 'Error'
            }
        }
        Context 'property' {
            It 'OU is missing' {
                @{
                    # OU           = @('OU=BEL,OU=EU,DC=contoso,DC=com')
                    MailTo       = @('bob@contoso.com')
                    EmployeeType = @('Plant', 'Kiosk' )
                } | ConvertTo-Json -Depth 3 | Out-File @testOutParams

                .$testScript @testParams
                        
                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailAdminParams) -and 
                    ($Message -like "*$ImportFile*Property 'OU' not found*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
            It 'MailTo is missing' {
                @{
                    OU           = @('OU=BEL,OU=EU,DC=contoso,DC=com')
                    # MailTo       = @('bob@contoso.com')
                    EmployeeType = @('Plant', 'Kiosk' )
                } | ConvertTo-Json -Depth 3 | Out-File @testOutParams

                .$testScript @testParams
                        
                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailAdminParams) -and 
                    ($Message -like "*$ImportFile*Property 'MailTo' not found*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
            It 'EmployeeType is missing' {
                @{
                    OU     = @('OU=BEL,OU=EU,DC=contoso,DC=com')
                    MailTo = @('bob@contoso.com')
                    # EmployeeType = @('Plant', 'Kiosk' )
                } | ConvertTo-Json -Depth 3 | Out-File @testOutParams

                .$testScript @testParams
                        
                Should -Invoke Send-MailHC -Exactly 1 -ParameterFilter {
                    (&$MailAdminParams) -and 
                    ($Message -like "*$ImportFile*Property 'EmployeeType' not found*")
                }
                Should -Invoke Write-EventLog -Exactly 1 -ParameterFilter {
                    $EntryType -eq 'Error'
                }
            }
        }
    }
}
Describe 'when all tests pass' {
    BeforeAll {
        $testData = @{
            picard = @{
                AccountExpirationDate = (Get-Date).AddDays(+30)
                SamAccountName        = 'picard'
                displayName           = 'Jean-Luc Picard'
                title                 = 'Captain Enterprise'
                department            = 'Starfleet'
                co                    = 'EU'
                company               = 'Starfleet Inc.'
                description           = 'Starfleet captain'
                DistinguishedName     = "CN=Bxl,OU=EU,DC=contoso,DC=com"
                EmailAddress          = 'picard@starfleet.earth'
                EmployeeID            = 2
                EmployeeType          = 'Employee'
                extensionAttribute8   = 5
                Fax                   = '02 66 66 66'
                homeDirectory         = 'c:\home\picard'
                HomePhone             = '02 55 55 55'
                info                  = "smart`nThings"
                ipPhone               = '10.10.10.10'
                Manager               = 'Kirk'
                MobilePhone           = '0477 77 77 77'
                Name                  = 'Jean-Luc Picard'
                Office                = 'France'
                OfficePhone           = '02 22 22 22'
                pager                 = '02 77 77 77'
                CanonicalName         = 'contoso.com/EU/picard'
                ScriptPath            = 'c:\script.vbs'
                whenCreated           = (Get-Date).AddDays(-1)
                LastLogonDate         = (Get-Date)
                LockedOut             = $false
                PasswordExpired       = $false
                PasswordNeverExpires  = $true
            }
        }
        Mock Get-ADUser {
            [PSCustomObject]$testData.picard
        }
        Mock Get-ADTsProfileHC {
            'c:\tsUserProfile'
        } -ParameterFilter {
            ($DistinguishedName -eq $testData.picard.DistinguishedName) -and
            ($Property -eq 'UserProfile' )
        }
        Mock Get-ADTsProfileHC {
            'c:\tsHomeDir'
        } -ParameterFilter {
            ($DistinguishedName -eq $testData.picard.DistinguishedName) -and
            ($Property -eq 'HomeDirectory' )
        }
        Mock Get-ADTsProfileHC {
            'c:\tsHomeDrive'
        } -ParameterFilter {
            ($DistinguishedName -eq $testData.picard.DistinguishedName) -and
            ($Property -eq 'HomeDrive' )
        }
        Mock Get-ADTsProfileHC {
            'c:\tsAllowLogon'
        } -ParameterFilter {
            ($DistinguishedName -eq $testData.picard.DistinguishedName) -and
            ($Property -eq 'AllowLogon' )
        }
        Mock Get-ADDisplayNameHC {
            'Kirk'
        } -ParameterFilter {
            ($Name -eq $testData.picard.Manager)
        }
        Mock ConvertTo-OuNameHC {
            'the ou'
        } -ParameterFilter {
            ($Name -eq $testData.picard.CanonicalName)
        }

        @{
            OU           = @('OU=EU,DC=contoso,DC=com')
            MailTo       = @('bob@contoso.com')
            EmployeeType = @('Plant', 'Kiosk', 'Employee' )
        } | ConvertTo-Json -Depth 3 | Out-File @testOutParams

        . $testScript @testParams
    }
    Context 'Get-ADUser' {
        It 'is called' {
            Should -Invoke Get-ADUser -Times 1 -Exactly -Scope Describe
        }
        It 'with the correct arguments' {
            Should -Invoke Get-ADUser -Times 1 -Exactly -Scope Describe -ParameterFilter {
                ($SearchBase -eq 'OU=EU,DC=contoso,DC=com') -and
                ($Filter -eq "(employeeType -eq 'Plant') -or (employeeType -eq 'Kiosk') -or (employeeType -eq 'Employee')")
            }
        }
    }
    Context "export an Excel file with" {
        Context "worksheet 'Users'" {
            BeforeAll {
                $testExportedExcelRows = @(
                    [PSCustomObject]@{
                        AccountExpirationDate     = $testData.picard.AccountExpirationDate
                        Country                   = $testData.picard.co
                        Company                   = $testData.picard.company
                        CreationDate              = $testData.picard.whenCreated
                        Department                = $testData.picard.department
                        Description               = $testData.picard.description
                        DisplayName               = $testData.picard.displayName
                        FirstName                 = 'Jean-Luc'
                        HeidelbergCementBillingID = $testData.picard.extensionAttribute8
                        TSUserProfile             = 'c:\tsUserProfile'
                        TSHomeDirectory           = 'c:\tsHomeDir'
                        TSHomeDrive               = 'c:\tsHomeDrive'
                        TSAllowLogon              = 'c:\tsAllowLogon'
                        EmailAddress              = 'picard@starfleet.earth'
                        EmployeeID                = $testData.picard.EmployeeID
                        EmployeeType              = $testData.picard.employeeType
                        Fax                       = $testData.picard.Fax
                        HomeDirectory             = $testData.picard.homeDirectory
                        HomePhone                 = $testData.picard.HomePhone
                        ipPhone                   = $testData.picard.ipPhone
                        LastName                  = 'Picard'
                        LastLogonDate             = $testData.picard.LastLogonDate
                        LockedOut                 = $testData.picard.LockedOut
                        LogonScript               = $testData.picard.ScriptPath
                        Manager                   = 'Kirk'
                        MobilePhone               = $testData.picard.MobilePhone
                        Name                      = $testData.picard.Name
                        Notes                     = 'smart Things'
                        Office                    = $testData.picard.Office
                        OfficePhone               = $testData.picard.OfficePhone
                        OU                        = 'the ou'
                        Pager                     = $testData.picard.Pager
                        PasswordExpired           = $testData.picard.PasswordExpired
                        PasswordNeverExpires      = $testData.picard.PasswordNeverExpires
                        SamAccountName            = $testData.picard.SamAccountName
                        Title                     = $testData.picard.title
                    }
                )

                $testExcelLogFile = Get-ChildItem $testParams.LogFolder -File -Recurse -Filter '* - Log.xlsx'

                $actual = Import-Excel -Path $testExcelLogFile.FullName -WorksheetName 'Users'
            }
            It 'to the log folder' {
                $testExcelLogFile | Should -Not -BeNullOrEmpty
            }
            It 'with the correct total rows' {
                $actual | Should -HaveCount $testExportedExcelRows.Count
            }
            It 'with the correct data in the rows' {
                foreach ($testRow in $testExportedExcelRows) {
                    $actualRow = $actual | Where-Object {
                        $_.SamAccountName -eq $testRow.SamAccountName
                    }
                    $actualRow.AccountExpirationDate.ToString('yyyyMMdd HHmmss') |
                    Should -Be $testRow.AccountExpirationDate.ToString('yyyyMMdd HHmmss')
                    $actualRow.CreationDate.ToString('yyyyMMdd HHmmss') |
                    Should -Be $testRow.CreationDate.ToString('yyyyMMdd HHmmss')
                    $actualRow.Company | Should -Be $testRow.Company
                    $actualRow.Country | Should -Be $testRow.Country
                    $actualRow.Description | Should -Be $testRow.Description
                    $actualRow.Department | Should -Be $testRow.Department
                    $actualRow.DisplayName | Should -Be $testRow.DisplayName
                    $actualRow.EmailAddress | Should -Be $testRow.EmailAddress
                    $actualRow.EmployeeID | Should -Be $testRow.EmployeeID
                    $actualRow.EmployeeType | Should -Be $testRow.EmployeeType
                    $actualRow.Enabled | Should -Be $testRow.Enabled
                    $actualRow.Fax | Should -Be $testRow.Fax
                    $actualRow.HeidelbergCementBillingID | 
                    Should -Be $testRow.HeidelbergCementBillingID
                    $actualRow.HomeDirectory | Should -Be $testRow.HomeDirectory
                    $actualRow.LastLogonDate.ToString('yyyyMMdd HHmmss') |
                    Should -Be $testRow.LastLogonDate.ToString('yyyyMMdd HHmmss')
                    $actualRow.Title | Should -Be $testRow.Title
                    $actualRow.Manager | Should -Be $testRow.Manager
                    $actualRow.Office | Should -Be $testRow.Office
                    $actualRow.OfficePhone | Should -Be $testRow.OfficePhone
                    $actualRow.HomePhone | Should -Be $testRow.HomePhone
                    $actualRow.MobilePhone | Should -Be $testRow.MobilePhone
                    $actualRow.ipPhone | Should -Be $testRow.ipPhone
                    $actualRow.ipPhone | Should -Be $testRow.ipPhone
                    $actualRow.Pager | Should -Be $testRow.Pager
                    $actualRow.Notes | Should -Be $testRow.Notes
                    $actualRow.LockedOut | Should -Be $testRow.LockedOut
                    $actualRow.LogonScript | Should -Be $testRow.LogonScript
                    $actualRow.TSUserProfile | Should -Be $testRow.TSUserProfile
                    $actualRow.TSHomeDirectory | Should -Be $testRow.TSHomeDirectory
                    $actualRow.TSHomeDrive | Should -Be $testRow.TSHomeDrive
                    $actualRow.TSAllowLogon | Should -Be $testRow.TSAllowLogon
                    $actualRow.PasswordExpired | Should -Be $testRow.PasswordExpired
                    $actualRow.PasswordNeverExpires | Should -Be $testRow.PasswordNeverExpires
                }
            }
        }
    }
    Context 'send a mail to the user when SendMail.When is Always' {
        BeforeAll {
            $testMail = @{
                Header      = $testParams.ScriptName
                To          = 'bob@contoso.com'
                Bcc         = $ScriptAdmin
                Subject     = '1 account'
                Message     = "*Found a total of <b>1 account</b> for the following employee types in the active directory:*<table>*<tr><th>EmployeeType</th><th>Accounts</th></tr><tr><td>Employee</td><td>1</td></tr> <tr><td>Kiosk</td><td>0</td></tr> <tr><td>Plant</td><td>0</td></tr></table>*<p><i>* Check the attachments for details</i></p>*<h3>Organizational units:</h3><ul><li>CONTOSO.COM\EU</li></ul>"
                Attachments = '* - Log.xlsx'
            }
        }
        It 'Send-MailHC has the correct arguments' {
            $mailParams.Header | Should -Be $testMail.Header
            $mailParams.To | Should -Be $testMail.To
            $mailParams.Bcc | Should -Be $testMail.Bcc
            $mailParams.Subject | Should -Be $testMail.Subject
            $mailParams.Message | Should -BeLike $testMail.Message
            $mailParams.Attachments | Should -BeLike $testMail.Attachments
        }
        It 'Send-MailHC is called' {
            Should -Invoke Send-MailHC -Exactly 1 -Scope Describe -ParameterFilter {
                ($Header -eq $testMail.Header) -and
                ($To -eq $testMail.To) -and
                ($Bcc -eq $testMail.Bcc) -and
                ($Subject -eq $testMail.Subject) -and
                ($Attachments -like $testMail.Attachments) -and
                ($Message -like $testMail.Message)
            }
        }
    } 
}
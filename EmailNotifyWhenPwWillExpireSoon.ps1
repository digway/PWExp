<#
.Synopsis
    Checks AD for users whose PW is about to expire.
.Description
    Gets all AD Users, then checks their PW has been set recently.
    If not, sends a report to the user.
    
    Will also check for AD Users where their PW has already expired.
    It will send that report to pwadmin@etkg.com
.Example
    .\EmailNotifyWhenPwWillExpireSoon.ps1
    Takes the defaults, checks AD and sends emails as needed.
.Example
    .\EmailNotifyWhenPwWillExpireSoon.ps1 -DoNotSendEmail
    Will do everything but actually send emails to users and the help desk.
.Inputs
    Optional string parameters.
    Requires an HTML file exist in the same folder as the PS1 file.
.Outputs
    Creates log files and XML for emails that are sent.
.Notes
    Created by Djacobs@hbs.net
.Link
    http://www.hbs.net
#>
[CmdletBinding(DefaultParameterSetName = 'DefParamSet',
               SupportsShouldProcess = $true,
               PositionalBinding = $false,
               ConfirmImpact = 'Medium')]
Param (
    # This is the FQDN of the server that will relay our email messages out to users.
    [Parameter(Mandatory = $false,
               ParameterSetName = 'DefParamSet')]
    [ValidateNotNullOrEmpty()]
    $SmtpServer = 'smtp.etkg.com',
    
    # This is how far you want to go back to check for old passwords.
    [Parameter(Mandatory = $false,
               ParameterSetName = 'DefParamSet')]
    [int]$DaysToGoBack = 30,
    
    # If you only want to log results, but not actually send emails to recipients.
    [Parameter(Mandatory = $false,
               ParameterSetName = 'DefParamSet')]
    [switch]$DoNotSendEmail
)
Begin {
    Function Get-IgCurrentLineNumber {
        # Simply Displays the Line number within the script.
        [string]$line = $MyInvocation.ScriptLineNumber
        $line.PadLeft(4, '0')
    }
    
    Function Get-IgLocalDC {
        [CmdletBinding()]
        Param ()
        Write-Verbose -Message "Finding a Local Domain Controller."
        # http://www.onesimplescript.com/2012/03/using-powershell-to-find-local-domain.html
        
        # Set $ErrorActionPreference to continue so we don't see errors for the connectivity test
        $ErrorActionPreference = 'SilentlyContinue'
        
        # Get all the local domain controllers
        $allLocalDCs = ([System.DirectoryServices.ActiveDirectory.ActiveDirectorySite]::GetComputerSite()).Servers
        
        # Create an array for the potential DCs we could use
        $PotentialDCs = @()
        
        # Check connectivity to each DC
        ForEach ($LocalDC in $allLocalDCs) {
            Write-Verbose -Message "Testing '$($LocalDC.Name)'"
            
            # Create a new TcpClient object
            $TCPClient = New-Object System.Net.Sockets.TCPClient
            
            # Try connecting to port 389 on the DC
            $Connect = $TCPClient.BeginConnect($LocalDC.Name, 389, $null, $null)
            
            # Wait 250ms for the connection
            $Wait = $Connect.AsyncWaitHandle.WaitOne(250, $False)
            
            # If the connection was succesful add this DC to the array and close the connection
            If ($TCPClient.Connected) {
                Write-Verbose -Message "Could talk to '$($LocalDC.Name)' on port 389."
                
                # Add the FQDN of the DC to the array
                $PotentialDCs += $LocalDC.Name
                
                # Close the TcpClient connection
                $Null = $TCPClient.Close()
            } else {
                Write-Warning -Message "[Line: $(Get-IgCurrentLineNumber)] Error. Cannot talk to DC '$($LocalDC.Name)' on port 389."
            }
        }
        
        # Pick a random DC from the list of potentials
        if ($PotentialDCs.Count -eq 0) {
            Write-Error -Message "No Domain Controllers Found." -ErrorAction stop
        } else {
            $PotentialDCs | Get-Random
        }
    }
    
    Function Update-IgLog {
        <#
        .Synopsis
            Adds a Line into a log file and error log.
        .Description
            Adds a line of text into a log file or Error File.
            This program does not check for the existance of a log or error file, just assumes that one has already been created.
        .Example
            Update-IgLog -Message "File Created Successfully"
            This will add the time stamp and the line "File Created Successfully".
        .Example
            Update-IgLog "File Created Successfully"
            This will add the time stamp and the line "File Created Successfully".
        .Example
            Update-IgLog -Message "File Not Created" -IncludeErrorLog
            This will add the time stamp and the line "File Created Successfully". Will also add the same message to an Error File.
        .Notes
            Created by Donald Jacobs
        #>
        [CmdletBinding(DefaultParameterSetName = 'DefParamSet',
                       SupportsShouldProcess = $true,
                       PositionalBinding = $false,
                       ConfirmImpact = 'Medium')]
        Param (
            # This is the Message that will be logged into the Log File
            [Parameter(Mandatory = $False,
                       ValueFromPipeline = $true,
                       ValueFromPipelineByPropertyName = $true,
                       Position = 0)]
            $Message,
            
            # If you also want to send it to a different Error Log, include this switch.
            [Parameter(ParameterSetName = 'ParamSet02')]
            [switch]$IncludeErrorLog,
            
            # To include a section break in your log file, include this switch.
            # A section break is three blank lines in the log file.
            [Parameter(ParameterSetName = 'ParamSet03')]
            [switch]$SectionBreak,
            
            # To Add a minor break in your log file, include this switch.
            # A minor break is ************* in the log file.
            [Parameter(ParameterSetName = 'ParamSet04')]
            [switch]$MinorBreak
        )
        Begin {
        }
        Process {
            if (-not $PSBoundParameters.ContainsKey('Verbose')) {
                $VerbosePreference = $PSCmdlet.GetVariableValue('VerbosePreference')
            }
            Try {
                If ($SectionBreak) {
                    Add-Content -Value "`n`n`n" -Path $LogFile
                } ElseIf ($MinorBreak) {
                    If ($IncludeErrorLog) {
                        Add-Content -Value "$(Get-Date -Format o) *************" -Path $LogFile, $ErrorFile -ErrorAction Stop
                    } else {
                        Add-Content -Value "$(Get-Date -Format o) *************" -Path $LogFile -ErrorAction Stop
                    }
                } Else {
                    if (-not $PSBoundParameters.ContainsKey('Verbose')) {
                        $VerbosePreference = $PSCmdlet.GetVariableValue('VerbosePreference')
                    }
                    Write-Verbose -Message "$(Get-Date -Format o) $Message"
                    
                    If ($IncludeErrorLog) {
                        Add-Content -Value "$(Get-Date -Format o) $Message" -Path $LogFile, $ErrorFile -ErrorAction Stop
                    } else {
                        Add-Content -Value "$(Get-Date -Format o) $Message" -Path $LogFile -ErrorAction Stop
                    }
                }
            } Catch {
                Write-Error -Message "Cannot Update log because: $_" -ErrorAction Continue
            }
        }
        End {
        }
    }
    
    $igStyle = @"
<style>
body { color:#333333; font-family:Calibri,Tahoma; font-size: 11pt; }
th { font-weight:bold; color:#eeeeee; background-color:#000000; }
.odd  { background-color:#ffffff; }
.even { background-color:#dddddd; }
</style>
"@
    
    function Set-IgTableFormat {
        <#
        .Synopsis
            Allows CSS type formatting on Table Fragments
        .DESCRIPTION
            Classifies rows as odd or even to be used with a CSS Style Worksheet.
        .EXAMPLE
            Set-IgTableFormat -HTMLFragment $tableFragment
            Creates a formatted table based on CSS Style sheets.
        .EXAMPLE
            $tableFragment | Set-IgTableFormat
            Creates a formatted table based on CSS Style sheets.
        .Notes
            Must be used in conjunction with a CSS file or embedded style formatting.
        .Link
            http://powershell.org/wp/powershell-books/
        #>
        
        [CmdletBinding()]
        param (
            # This is the fragment of code you want to embellish.
            [Parameter(Mandatory = $True,
                       ValueFromPipeline = $True)]
            [string]$HTMLFragment,
            
            # This is the CSS code given to Even Rows. If you change this, change the style.
            [Parameter(Mandatory = $False)]
            [string]$EvenRow = 'even',
            
            # This is the CSS code given to Odd Rows. If you change this, change the style.
            [Parameter(Mandatory = $False)]
            [string]$OddRow = 'odd'
        )
        
        [xml]$xml = $HTMLFragment
        
        $table = $xml.SelectSingleNode('table')
        $classname = $OddRow
        
        foreach ($tr in $table.tr) {
            if ($classname -eq $EvenRow) {
                $classname = $OddRow
            } else {
                $classname = $EvenRow
            }
            
            $class = $xml.CreateAttribute('class')
            $class.value = $classname
            $tr.attributes.append($class) | Out-null
        }
        $xml.innerxml | out-string
    }
    
    #region to create a log based on where the PS1 file was run from and based on the name of the PS1 file, use the following
    $MyRootPath = Get-Item -Path $MyInvocation.MyCommand.Path
    $global:logFile = "$($MyRootPath.DirectoryName)\Log Files\Log_$($MyRootPath.BaseName)_$(get-date -Format yyyyMMddTHHmmss).txt"
    $global:errorFile = "$($MyRootPath.DirectoryName)\Log Files\Error_$($MyRootPath.BaseName)_$(get-date -Format yyyyMMddTHHmmss).txt"
    $global:EmailReports = "$($MyRootPath.DirectoryName)\Log Files"
    #endregion
    
    #region Check if 'Log Files' exists and if not, create it
    if (Test-Path -Path "$($MyRootPath.DirectoryName)\Log Files") {
        Write-Verbose -Message 'Log Files folder already existed.'
    } else {
        Write-Verbose -Message 'Need to create Log Files folder.'
        Try {
            $null = New-Item -Path "$($MyRootPath.DirectoryName)\Log Files" -ItemType Directory -ErrorAction Stop
            Write-Verbose -Message "Created Folder '$($MyRootPath.DirectoryName)\Log Files' because it did not exist."
        } Catch {
            # Program Terminating Event, send email to help desk to let them know program stopped.
            $Message = "[Line: $(Get-IgCurrentLineNumber)] Error. Could not create '$($MyRootPath.DirectoryName)\Log Files' because: $_"
            Write-Error -Message "$Message" -ErrorAction Stop
        }
    }
    #endregion
    
    #region Import the HTML file
    Try {
        $body = Get-Content "$($MyRootPath.DirectoryName)\PassExp.html" -ErrorAction Stop | Out-String
    } Catch {
        Write-Error -Message "[Line: $(Get-IgCurrentLineNumber)] Cannot get HTML file from '$($MyRootPath.DirectoryName)\PassExp.html' because $_" -ErrorAction Stop
    }
    #endregion
    
    #region define default email parameters
    $EmailDefaults = @{
        Subject         = "Password Expires Soon"
        SmtpServer      = $SmtpServer
        From            = "pwadmin@etkg.com"
        BodyAsHtml      = $true
        ErrorAction     = "Stop"
    }
    #endregion
    
    #region Use this DC
    $DC = Get-IgLocalDC
    #endregion
}
Process {
    #region Program Starting
    Update-IgLog -Message "[Line: $(Get-IgCurrentLineNumber)] Program Starting"
    Update-IgLog -Message "[Line: $(Get-IgCurrentLineNumber)] Going to use the Domain Controller '$DC' for the rest of the script."
    #endregion
    
    #region Get Users From Active Directory
    Try {
        $getaduserParams = @{
            Server         = $DC
            Filter         = "*"
            Properties     = @("PasswordNeverExpires", "PasswordExpired", "PasswordLastSet", "EmailAddress", "DisplayName", "msDS-UserPasswordExpiryTimeComputed")
            ErrorAction    = "Stop"
        }
        $allUsers = Get-ADUser @getaduserParams
        Update-IgLog -Message "[Line: $(Get-IgCurrentLineNumber)] Successfully queried '$($allUsers.Count)' users from the Domain Controller '$DC'."
    } Catch {
        Update-IgLog -Message "[Line: $(Get-IgCurrentLineNumber)] Error. Cannot query users from the Domain Controller '$DC' because $_"
        Write-Error -Message "[Line: $(Get-IgCurrentLineNumber)] Cannot get users from the DC '$DC' because $_" -ErrorAction Stop
    }
    #endregion
    
    #region Only check AD users that must change their password
    $UsersWhoHavePasswordsWeWantToCheck = $allUsers |
    Where-Object PasswordNeverExpires -EQ $false
    #endregion
    
    #region Loop through all users and check each one
    $usersWhosePwAlreadyExpired = @()
    foreach ($user in $UsersWhoHavePasswordsWeWantToCheck) {
        
        if ($user.PasswordExpired) {
            # Write-Host "The user '$($user.SamAccountName)' last set their password on $($user.PasswordLastSet) and has an expired password."
            $usersWhosePwAlreadyExpired += $user
        } else {
            $dateToCheck = Get-Date $user.PasswordLastSet
            
            if ($dateToCheck -lt (Get-Date).AddDays(- $daysToGoBack)) {
                # This means their PW is about to Expire.
                Update-IgLog -Message "[Line: $(Get-IgCurrentLineNumber)] This user '$($user.SamAccountName)' last set their password on '$dateToCheck', so warn this user."
                
                $willExpireOn = [datetime]::FromFileTime($user."msDS-UserPasswordExpiryTimeComputed")
                Update-IgLog -Message "[Line: $(Get-IgCurrentLineNumber)] The user '$($user.SamAccountName)' will have their PW expire on '$willExpireOn'."
                
                $UniqueBody = $body -replace 'displayName', $($user.DisplayName)
                $UniqueBody = $UniqueBody -replace 'pwExpiry', "$willExpireOn"
                $uniqueEmailParams = $EmailDefaults
                $uniqueEmailParams.To = $user.EmailAddress
                $uniqueEmailParams.Body = $UniqueBody
                
                #region Save a copy of the Email, Just in case
                $currentXmlEmailFile = "$($MyRootPath.DirectoryName)\Log Files\EmailTo_$($user.SamAccountName)_$(Get-Date -Format yyyyMMddTHHmmss).xml"
                Try {
                    $uniqueEmailParams | Export-Clixml -Path $currentXmlEmailFile -ErrorAction Stop
                    Update-IgLog -Message "[Line: $(Get-IgCurrentLineNumber)] Successfully saved a copy of the Email at '$currentXmlEmailFile'."
                } Catch {
                    Update-IgLog -Message "[Line: $(Get-IgCurrentLineNumber)] Error. Cannot save a copy of the Email at '$currentXmlEmailFile' because $_" -IncludeErrorLog
                }
                #endregion
                
                #region Send the email
                Try {
                    if ($DoNotSendEmail) {
                        Update-IgLog -Message "[Line: $(Get-IgCurrentLineNumber)] Not sending an Email because I was told NOT to. But I would have."
                    } else {
                        Send-MailMessage @uniqueEmailParams
                    }
                    Update-IgLog -Message "[Line: $(Get-IgCurrentLineNumber)] Successfully sent an email to '$($user.EmailAddress)' telling them their PW would expire soon."
                } Catch {
                    Update-IgLog -Message "[Line: $(Get-IgCurrentLineNumber)] Error. Cannot send an email to '$($user.EmailAddress)' telling them their PW would expire soon because $_" -IncludeErrorLog
                }
                #endregion
                
            } else {
                # This means they set their PW less than our value for $daysToGoBack, we will do nothing.
                # Write-Host "$dateToCheck $($user.SamAccountName) Newer than $daysToGoBack" -ForegroundColor Green
            }
        }
    }
    #endregion
    
    #region Create a report of users whose PW has expired.
    $BodyArrayTable = $usersWhosePwAlreadyExpired | Sort-Object SamAccountName | Select-Object SamAccountName, PasswordLastSet, PasswordExpired |
    ConvertTo-Html -Fragment |
    Out-String |
    Set-IgTableFormat
    $HtmlBody = "<H2>These Users have their PW already expired.</H2>$BodyArrayTable"
    $htmParams = @{
        'Head'   = "$igStyle"; 'Body' = "$HtmlBody"
    }
    
    $HelpDeskParams = @{
        To           = "pwadmin@etkg.com"
        From         = "pwadmin@etkg.com"
        Subject      = "$(Get-Date -Format d) Passwords Expired"
        Body         = "$(ConvertTo-HTML @htmParams)"
        BodyAsHtml   = $true
        SmtpServer   = $SmtpServer
        ErrorAction  = "Stop"
    }
    
    #region Save a copy, just i case.
    $xmlHdFile = "$($MyRootPath.DirectoryName)\Log Files\PwAlreadyExp_$(get-date -Format yyyyMMddTHHmmss).xml"
    Try {
        $HelpDeskParams | Export-Clixml -Path $xmlHdFile -ErrorAction Stop
        Update-IgLog -Message "[Line: $(Get-IgCurrentLineNumber)] Saved a copy of today's list of all users with expired passwords at '$($xmlHdFile)'."
    } Catch {
        Update-IgLog -Message "[Line: $(Get-IgCurrentLineNumber)] Error. Cannot save a copy of today's list of all users with expired passwords at '$($xmlHdFile)' because $_" -IncludeErrorLog
    }
    #endregion
    
    #region Send the Email to the HelpDesk
    Try {
        if ($DoNotSendEmail) {
            Update-IgLog -Message "[Line: $(Get-IgCurrentLineNumber)] Not sending an Email because I was told NOT to. But I would have."
        } else {
            Send-MailMessage HelpDeskParams
        }
        Update-IgLog -Message "[Line: $(Get-IgCurrentLineNumber)] Successfully send email to the help desk."
    } Catch {
        Update-IgLog -Message "[Line: $(Get-IgCurrentLineNumber)] Error. Cannot send email to help desk because $_" -IncludeErrorLog
    }
    #endregion
    
    #endregion
}
End {
    Update-IgLog -Message "[Line: $(Get-IgCurrentLineNumber)] Shakespearean Play Complete."
}

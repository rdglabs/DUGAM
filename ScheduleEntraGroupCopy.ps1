<#
.SYNOPSIS
Schedule groups to be copied from one user to another user in entra

.DESCRIPTION
Will schedule exchange and entra groups to be copied from one user to another user. 

.NOTES
  Version:        2.0.0
  Author:         Ryan Gillespie
  Creation Date:  2024/08/28
  ChangeLog: 
    2024/08/28 - Copied over existing tool to new file to be modified
    2024/08/31 - Removed mail functions from this script and moved them to a seprate one to be called from both parts
    2024/09/10 - Migrated all functions to NDU-Functions This will allow them to be called from both Manual ran and from schedule task scripts.
                 Cleaned up script some

#>

################################################################################################################################
#------------------------------------------------------Variables Declared------------------------------------------------------#
################################################################################################################################
Import-Module "$PSScriptRoot\NDU-Functions.psm1" -Force

#module check
#$modules = 'Microsoft.Graph','ExchangeOnlineManagement'
$config = Get-Content ".\Resources\config.json" -raw | ConvertFrom-Json

################################################################################################################################
#----------------------------------------------------------Functions-----------------------------------------------------------#
################################################################################################################################
<#
    View NDU-Functions.ps1 to see all functions supported by this tool
#>
################################################################################################################################
#---------------------------------------------------------Script Start---------------------------------------------------------#
################################################################################################################################

$logfile = "$PSScriptRoot\Logs\ScheduleTask-Entra-$(get-date -Format "yyyyddMM-hhmm").log"
Start-Transcript -Path $logfile

$logger = @{
    Name         = 'logfile'
    InstanceName = 'DUGMA-ST'
    FilePath     = 'C:\scripts\New-DomainUser\Logs\%date%-%username%.log'
    FileType     = 'csv'
    Headers      =  ‘Timestamp’, 'Level', 'Message','FunctionName',‘Username’,, ‘Type’, 'tags','InstanceName',‘Runspace’
    Enabled      = $true

}
Set-PSFLoggingProvider @logger
Set-PSFConfig -FullName PSFramework.Message.Style.FunctionName -Value $false
Set-PSFConfig -FullName PSFramework.Message.Style.Timestamp -Value $false
Write-PSFMessage -Level Verbose -Message "Program Schedule task start"

#------------------------------------------------Check if modules are installed------------------------------------------------#
<#
$installed = @((get-module $modules -ListAvailable).Name | Select-Object -Unique)
$notInstalled = Compare-Object $modules $installed -PassThru
if ($notInstalled) { # At least one module is missing.

  # Prompt for installing the missing ones.
  $promptText = @"
  The following modules aren't currently installed:
  
      $notInstalled
  
  Would you like to install them now?
"@
write-host $promptText

}
#>
connect-EntraExchange

#------------------------------------------------Connect to Entra via graph and Exchange-------------------------------------------------#


$users = import-csv .\Resources\EntraUsers.csv
$export = @()
foreach ($user in $users){
    if($user.Completed -eq $false){
        try {
            Get-MgUser -UserId $user.TargetUser -ErrorAction Stop
            Write-PSFMessage -Level Verbose -Message "Copying groups from $($user.SourceUser) to $($user.TargetUser)"
            $info = get-aduser -filter "UserPrincipalName -eq '$($user.TargetUser)'" -Properties UserPrincipalName, name, givenName, Manager | Select-Object UserPrincipalName, name, givenName, Manager
            $info.Manager= (get-aduser $info.Manager).UserPrincipalName
            Copy-EntraGroups -sourceEmail $user.SourceUser -targetEmail $user.TargetUser
            $emailBody = Get-EmailTemplate -user $info -Manager $info.Manager -Emailpassword $user.password
		    Send-MailtoManager -user $info -manager $info.Manager -EmailBody $emailBody -config $config
            $user.completed = $true
            $export +=$user
        }
        catch {
            Write-PSFMessage -Level error -Message "user not found in entra"
			$export +=$user
            try {
                Unregister-ScheduledTask -TaskName 'CopyEntraGroups' -Confirm:$False
                Write-PSFMessage -Level Verbose -Message "Schedule task removed"
            }
            catch {
                Write-PSFMessage -Level Error -Message "Error removing schedule task"
                Write-PSFMessage -Level Error -Message "$($Error[0])"
            }
            Write-PSFMessage -Level Verbose -Message "Attempting to rescheduling schedule task to run again in 15 minutes"
            set-EntraCopyGroups -timeDelay 15

            Send-MailMessage -SmtpServer $config.SMTP.address -To $config.SMTP.ErrorTo -From $config.SMTP.from -Subject "Error schedule task" -Body "could not find $($user.TargetUser). Trying again in 15 minutes" -BodyAsHtml
        }
        
        
    } else {
        $export +=$user
    }
}
    
$export | Select-Object -last 100 | export-csv .\Resources\EntraUsers.csv -Force


Disconnect-ExchangeOnline -Confirm:$false
Disconnect-Graph

try {
    Unregister-ScheduledTask -TaskName 'CopyEntraGroups' -Confirm:$False
    Write-PSFMessage -Level Verbose -Message "Schedule task removed"
}
catch {
    Write-PSFMessage -Level Error -Message "Error removing schedule task"
    Write-PSFMessage -Level Error -Message "$($Error[0])"
}

disable-PSFLoggingProvider -InstanceName 'DUGMA-ST' -name 'logfile'
Stop-Transcript
#TODO: test functionality

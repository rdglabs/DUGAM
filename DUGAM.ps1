<#
.SYNOPSIS
Domain User & Group Automator

.DESCRIPTION
This tool has two main functions. 

The first function is to create a new AD user account that is synced to Entra and Exchange. You will be prompted for information such as Name, Title, Department, Employee ID, Manager, Office Location, and Phone Extension. The tool will create the user in AD based on the information provided during the prompts, with office locations and OUs set up in the configuration file found under the resource folder. The user password will be automatically generated using word found in the password.dict file, along with a combination of number and special character. Once the user is created in AD, the tool will prompt you to decide if you would like to copy all groups from an existing user. If you choose to copy groups, it will prompt and search for that user. For the Entra/Exchange groups, it will create a scheduled task that will run once the user’s account is synced from AD. This scheduled task will copy over groups and then send a notification email to the manager with the login information. If you choose not to copy over groups, the tool will prompt you to notify the manager of the user’s login info immediately.

The second function is for mass group membership updates, which is intended for use on existing users. This function allows you to remove all existing groups from AD and Entra and copy over new groups from another user on both AD and Entra. It is also possible to not remove existing groups and only copy over new groups from another user. Not removing groups from the user could add the user to different site licensing groups if copying over groups on Entra. 

Authentication to Entra and Exchange is done via App Authentication with a certification. If the certification expires it will prompt to use your login. Your login must have all the correct permissions. If the certification expired the schedule task will not be able to run causing it error out and fail. A updated certification must deployed. 

.NOTES
  Version:        2.0.1
  Author:         Ryan Gillespie
  Creation Date:  2024/08/26
  ChangeLog: 
    V0.0.1 - 2024/08/26
        - Start work
        - functions started for group copies, and creating AD user from existing scripts
    V0.3.1 - 2024/08/28
        - New Password Function completed. Words are pulled from password.dict file under Resources folder
        - Completed functions and testing of basic functions. 
        - Starting work on interface
    V0.5.0 - 2024/08/30
        - UI and Function testing 
    V0.9.0 - 2024/08/30
        - Complete tool Testing
    V1.0.0 - 2024/08/30
        - V1.0.0 published
    V1.1.2 - 2024/09/01
        - migrated email functions to its own Ps1 file to be called from both CustomNew-ADUser and ScheduleEntraGroupCopy scripts
        - Created Email templates for manager notification using HTML vs inside function. Give great ability to customize the email layout
        - Added Email Template HTML files in Resource
    V1.2.1 - 2024/09/03
        - Updated static fields to pull from config file. All configuration of script should be done via the config.json file under Resources
    V1.3.5 - 2024/09/09
        - Request to mass update/remove/copy from user groups on existing users
        - Moved new user prompts to its own function. Added functions to remove all AD user's group. 
        - Created function for searching for AD user. Replaced AD searching with function call
    V2.0.1_Beta 
        - Added function to prompt for existing user group removal and group copy from another user
        - Added function to remove all AD user's group.
        - Created function for search for AD user
        - Updated Launch Title, Added Description under title.  
        - Migrated all functions to NDU-Functions.psm1 This will allow them to be called from both Manual ran and from schedule task scripts.
        - Updated Tool Launch section
        - Add comment_base Help to Functions in NDU-Function file. Parameters and Description field.
        - Tested Group Membership tool fully. Fixed errors with pulling groups and Administrative Units and ignoring dynamic groups from first where. 
        - Adjusted the UI Spacing with New lines/Foreground Colors
    V2.0.1 - 2024/09/11
        - Version Published
	V2.0.2 - 2024/09/12
		- update log naming to include username
    V2.0.3 - 2024/09/17
        - corrected typo's for ScheduleEntraGroupCopy.ps1 that caused it to fail
        - updated copy entra groups filtering.
    V2.0.4 - 2024/10/11
        - corrected typo's for DFW OU
        - changed "changed password on login" from always true to prompting
    V2.0.5 - 2024/10/14
        - Added check to new-domainuser to check if account with samaccountname exist
        - updated set-entracopygroup parameters to timeDelay from user source and target. Also migrated the call to export the users information outside of function
    V2.1.0 - 2024/10/15
        - Start added logging via PSFramework module
    V2.1.1 - 2024/11/07
        - Created Documentation, and added some more dynamic option in config file with respective prompts if they are needed. 



#>

#TODO: Error notification function. Currently have to manually view log files

################################################################################################################################
#------------------------------------------------------Variables Declared------------------------------------------------------#
###############################################################################################################################

Import-Module "$PSScriptRoot\NDU-Functions.psm1" -Force


#$attributes = @{}

#[string]$NewUser = $null
#[string]$Password = $null
#[string]$givenName = $null
#[string]$surName = $null



################################################################################################################################
#----------------------------------------------------------Functions-----------------------------------------------------------#
################################################################################################################################
<#
    View NDU-Functions.ps1 to see all functions supported by this tool
#>
################################################################################################################################
#---------------------------------------------------------Script Start---------------------------------------------------------#
################################################################################################################################




if (-Not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] 'Administrator')) {
    if ([int](Get-CimInstance -Class Win32_OperatingSystem | Select-Object -ExpandProperty BuildNumber) -ge 6000) {
        $CommandLine = "-File `"" + $MyInvocation.MyCommand.Path + "`" " + $MyInvocation.UnboundArguments
        Write-Host "You didn't run this script as an Administrator. This script will self elevate to run as an Administrator and continue." -ForegroundColor Red
        Start-Sleep .5
        Write-Host "                                                          3" -ForegroundColor Red
        Start-Sleep .5
        Write-Host "                                                          2" -ForegroundColor Red
        Start-Sleep .5
        Write-Host "                                                          1" -ForegroundColor Red
        Start-Sleep 1
        Start-Process -FilePath pwsh.exe -Verb Runas -ArgumentList $CommandLine
        exit
    }
}
$logFile = "C:\scripts\New-DomainUser\Logs\Domain_User&Group_Automator-$(($env:username).Replace('-adm',''))-.log"
Start-Transcript -Path $logFile

$logger = @{
    Name         = 'logfile'
    InstanceName = 'DUGMA'
    FilePath     = 'C:\scripts\New-DomainUser\Logs\%date%-%username%.log'
    FileType     = 'csv'
    Headers      =  ‘Timestamp’, 'Level', 'Message','FunctionName',‘Username’,, ‘Type’, 'tags','InstanceName',‘Runspace’
    Enabled      = $true

}
Set-PSFLoggingProvider @logger
Set-PSFConfig -FullName PSFramework.Message.Style.FunctionName -Value $false
Set-PSFConfig -FullName PSFramework.Message.Style.Timestamp -Value $false
Write-PSFMessage -Level Verbose -Message "Program Domain User & Group Membership Automator Start"
Clear-Host
do{
Clear-Host
Write-Host "--==================================================================================================================--" -ForegroundColor DarkCyan
Write-Host "||                                                                                                                  ||" -ForegroundColor DarkCyan
Write-Host "||                                     Domain User & Group Membership Automator                                     ||" -ForegroundColor DarkCyan
write-host "||                                                  Version $($version -replace '(?ms)Version:\s+([0-9.]+).+','$1')                                                   ||" -ForegroundColor DarkCyan
Write-Host "||                                                                                                                  ||" -ForegroundColor DarkCyan
Write-Host "--==================================================================================================================--" -ForegroundColor DarkCyan
Write-Host "`n"
write-host "This tool has two primary functions:"
write-host "1. Create a New AD/Entra User Account:" -ForegroundColor green
write-host "    Prompts for:" -ForegroundColor White -NoNewline
    write-host "Name, Title, Department, Employee ID, Manager, Office Location, and Phone Extension."  -ForegroundColor DarkCyan
write-host "    Group Copying:" -ForegroundColor White -NoNewline
    write-host "Copies group memberships from another user in AD, Entra, and Exchange. Entra and Exchange groups are" -ForegroundColor DarkCyan
	write-host "		  copied via a scheduled task to ensure the new account fully syncs." -ForegroundColor DarkCyan 
write-host "    Notification:" -ForegroundColor White -NoNewline
    write-host "Once the setup is complete, the tool notifies the user’s manager with the account details and password."  -ForegroundColor DarkCyan
write-host "2. Mass Group Membership:" -ForegroundColor green
write-host "    Use Case:" -NoNewline -ForegroundColor White
    write-host "Ideal for users migrating to different departments or sites." -ForegroundColor DarkCyan
write-host "    Remove Existing Memberships:" -ForegroundColor white -NoNewline
    write-Host "Clears all current group memberships." -ForegroundColor DarkCyan
write-host "    Copy New Memberships:" -ForegroundColor White -NoNewline
    write-host "Copies group memberships from another user." -ForegroundColor DarkCyan
write-host "`nAnything in magenta is required" -ForegroundColor Magenta 
Write-Host "`n`n"


$Q1title    = ''
$Q1question = 'Which Option would you like to launch? (Press ctrl+c to exit)'
$Q1Choices = @(
    [System.Management.Automation.Host.ChoiceDescription]::new("&New User", "Launch tool to start creating new user account in AD and entra")
    [System.Management.Automation.Host.ChoiceDescription]::new("&Group Mass Membership", "Launch tool to start mass group removal or copy additional groups")
    [System.Management.Automation.Host.ChoiceDescription]::new("&Tool Info", "Read the version notes and Description")
)
$Q1Answer = ($host.ui.PromptForChoice($Q1title,$Q1question,$Q1Choices,0))
Write-PSFMessage -Level Verbose -Message "$($Q1Answer) was selected"
switch ($Q1Answer) {
    0{
        Write-Host "Starting to Create a New AD User Account. Please answer the following prompts" -ForegroundColor DarkCyan
        Get-DomainUserData
    }
    1{
        Write-Host "Starting Mass Group Membership tool. Please answer the following prompts`n`n" -ForegroundColor DarkCyan
        Copy-DomainUserPrompt
    }
    2{
        $help = (get-help $PSCommandPath)
        $Q2Answer = ($host.ui.PromptForChoice("","View Detailed Description, Version Notes, or Both",@("&Description","&Version","&Manual"),0))
        Write-PSFMessage -Level Verbose -Message "$($Q1Answer) was selected"
        switch ($Q2Answer) {
            0{
                write-host "$($help.Description.Text)" -ForegroundColor DarkGreen
            }
            1{
                write-host "$($help.alertSet.alert.Text)" -ForegroundColor DarkGreen
            }
            2{
				
                start-process "msedge.exe" "$(Convert-Path .\Resources\ManualFiles\Manual.html)"
            }
            Default {}
        }
        pause
    }
    Default {}
}
}while(!($Q1Answer -ne 2))
Pause

disable-PSFLoggingProvider -InstanceName 'DUGMA' -name 'logfile'
Stop-Transcript
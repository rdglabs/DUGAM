<#
.SYNOPSIS
Function File for Custom Domain User creation and group management

.DESCRIPTION
Functions used by the  Custom Domain User creation and group management tool

.NOTES
  Author:         Ryan Gillespie
#>

$config = Get-Content "$PSScriptRoot\Resources\config.json" -raw | ConvertFrom-Json
$attributes = @{}

[string]$NewUser = $null
[string]$Password = $null
[string]$givenName = $null
[string]$surName = $null

Function Get-EmailTemplate
{
<#
.PARAMETER
    -user
        Specifies the user information to add to the email template. Must be an array with, UserPrincipalName, First name, and Full Name. Best option is to pull AD user info

        Required?       True

    -manager
        Specifies the user's manager and is used to address the email to. Expected format of first.last

        Required?       True

    -password
        The user's password in plaintext. 

        Required?       True

.DESCRIPTION
Load the email template from the HTML file and replaces the placeholders with user's name, password and the Manager's name 
#>
  PARAM(
    [parameter(Mandatory=$true)]
    $user,
    [parameter(Mandatory=$true)]
    $manager,
    [parameter(Mandatory=$true)]
    $Emailpassword
  )

  PROCESS
  {
    
	Write-PSFMessage -Level Host -Message "Attempting to send email to $($manager) for the user $($user) with password: $($Emailpassword)"
    #Get the mail template
    $mailTemplate = (Get-Content ("$PSScriptRoot\Resources\MailTemplateUserCreated.html")) | ForEach-Object {
      $_ 	-replace '{{manager.firstname}}', ($manager).Substring(0,($Manager.IndexOf('.'))) `
      -replace '{{user.UserPrincipalName}}', $user.UserPrincipalname `
      -replace '{{user.Password}}', $Emailpassword `
      -replace '{{user.fullname}}', $user.name `
      -replace '{{user.firstname}}', $user.givenName
    } | Out-String	
    
    Write-PSFMessage -Level InternalComment -Message "Email body has been setup to send to $($manager) for user $($user)"

    return $mailTemplate
  }
}

Function Send-MailtoManager
{
<#
.PARAMETER
    -user
        Specifies the user information to add to the email template. Must be an array with, UserPrincipalName, First name, and Full Name. Best option is to pull AD user info

        Required?       True

    -manager
        Specifies the user's manager and is used to address the email to. Expected format of first.last. used to get the email address to send to.

        Required?       True

    -config
        The config file information containing the SMTP setup

        Required?       True
        
.DESCRIPTION
Sends the email templated that is build by function "Get-EmailTemplate" and send it to the manager listed
#>
  PARAM(
    [parameter(Mandatory=$true)]
    $emailBody,
    [parameter(Mandatory=$true)]
    $user,
    [parameter(Mandatory=$true)]
    $manager,
    [parameter(Mandatory=$true)]
    $config
  )
  
  PROCESS	{
    
    #Create subject of the email
    $subject = $config.SMTP.subject -replace '{{user.fullname}}', $user.name
    
    #Set encoding
    $encoding = [System.Text.Encoding]::UTF8

	if($config.SMTP.SendManagerEmail){
        Try {
            Send-MailMessage -SmtpServer $config.SMTP.address -To $manager -bcc $config.SMTP.ITEmail  -From $config.SMTP.from -Subject $subject -Encoding $encoding -Body $emailBody -BodyAsHtml -Attachments $config.SMTP.Attachment
            Write-PSFMessage -Level host -Message "Email sent to maanger and BCC to IT Team"
        } Catch	{
        Write-Error "Failed to send email to manager, $_"
        Write-PSFMessage -Level Error -Message "Failed to send email to manager, $_"
        Write-PSFMessage -Level Error -Message "$($Error[0])"
        }
    }
    else {
        Try {
            Send-MailMessage -SmtpServer $config.SMTP.address -To $config.SMTP.ITEmail -From $config.SMTP.from -Subject $subject -Encoding $encoding -Body $emailBody -BodyAsHtml -Attachments $config.SMTP.Attachment
            Write-PSFMessage -Level host -Message "Email sent to It team"
        } Catch	{
        Write-Error "Failed to send email to It team, $_"
        Write-PSFMessage -Level Error -Message "Failed to send email to It team, $_"
        Write-PSFMessage -Level Error -Message "$($Error[0])"
        }
    }
  }
}

function Copy-DomainUserGroups {
<#
.PARAMETER
    -TargetUser
        User to copy group to. Expected format is: Admin.Smith

        Required?       True

    -SourceUser
        User to copy group from. Expected format is: Admin.Smith

        Required?       True
        
.DESCRIPTION
Will get all groups that the SourceUser is part of and add the TargetUser as a member to those group
#>
    param (
        [parameter(Mandatory=$true)]
        $TargetUser,
        [parameter(Mandatory=$true)]
        $SourceUser
    )
    Write-PSFMessage -level Verbose -Message "Copying groups from $($SourceUser) to $($TargetUser)"
    $groups = Get-ADUser -Identity $SourceUser -Properties memberOf | Select-Object -ExpandProperty memberOf
    Write-PSFMessage -Level Verbose -Message "groups to add $($TargetUser) to $($groups | Out-String)"
    foreach ($group in $groups) {
        Try	{
            Add-ADGroupMember -Identity $group -Members $TargetUser
            $groupName = ($group.Substring(0,$group.IndexOf(","))).replace("CN=","")
            Write-PSFMessage -Level Host -Message "$GroupName Membership copied"
        }	Catch	{
            Write-Error "Failed to add groups, $group"
            Write-PSFMessage -Level Error -Message "Failed to add groups, $($group) to $($TargetUser)"
            Write-PSFMessage -Level Error -Message "Error MSG: $($Error[0])"
        }
    }
}

function remove-DomainUserGroups {
<#
.PARAMETER
    -SamAccountName
        SameAccountName from AD for the target user

        Required?       True
        
.DESCRIPTION
Will remove all AD Group membership the of the AD user. The User running the tool must have the permissions to adjust the group
#>
    param (
        [parameter(Mandatory=$true)]
        $SamAccountName
    )
    

    if($host.ui.PromptForChoice("","Do you want to review each group as they are removed?",@('&Yes','&No'),1) -eq 0){
        $choice = $true} else {$choice=$false
    }
    
    Get-AdPrincipalGroupMembership -Identity $SamAccountName | Where-Object -Property Name -Ne -Value 'Domain Users' | ForEach-Object{
        Write-PSFMessage -Level Host -Message "Removing $($SamAccountName) from $($_.Name)" -ForegroundColor Yellow
        #Write-PSFMessage -Level Verbose -Message "Removing $($SamAccountName) from $($_.Name)"
        try {
            Remove-AdGroupMember -Identity $_ -Members $SamAccountName -Confirm:$choice
            Write-PSFMessage -Level Verbose -Message "Removed $($SamAccountName) from $($_)"
        }
        catch {
            Write-PSFMessage -Level Error -Message "Failed to removed $($SamAccountName) from $_"
            Write-PSFMessage -Level Error -Message "$($error[0])"
        } 
    }
    
}
function New-DomainUser {
<#
.PARAMETER
    -User
        Expecting an User Attributes ps custom object. Will use this to build the full Splatter to user creation

        Required?       True
        
.DESCRIPTION
Creates the AD user with the information in the parameter provide along with domain information from the config
#>
    param (
        [parameter(Mandatory=$true)]
        $user
    )
    
    

    $userAttributes = @{
        GivenName             = $user.givenName 
        Surname               = $user.surName 
        Name                  = $user.GivenName + " " + $user.Surname
        DisplayName           = $user.GivenName + " " + $user.Surname
        EmailAddress          = $user.GivenName + "." + $user.Surname + $config.Settings.maildomain
        StreetAddress         = $user.StreetAddress
        PostalCode            = $user.PostalCode
		state                 = $user.state
        city                  = $user.City
        Office                = $user.office
        OfficePhone           = $user.Extesion
        title                 = $user.title
		department			  = $user.department
        manager               = $user.manager
		Company               = $user.Company
        SamAccountName        = $user.GivenName + "." + $user.Surname
        UserPrincipalName     = $user.UserPrincipalName
        AccountPassword       = (ConvertTo-SecureString -AsPlainText $Global:Password -force)
        Enabled               = $true
        ChangePasswordAtLogon = if($host.ui.PromptForChoice("","Set Password to change on login",@('&Yes','&No'),0) -eq 0){$true}else{$false}
        PassThru              = $true
        Path                  = $user.OU   
    }

    Write-PSFMessage -Level Verbose -Message "User account info: $($userAttributes | out-string)"
    try{
        Get-ADUser $user.SamAccountName -ErrorAction stop
        write-host "User Account with same SamAccountName Exists" -ForegroundColor Red
        write-host "Please verify account information and rerun script" -ForegroundColor red
        Write-PSFMessage -Level Error -Message "Account found with matching SamAccountName $($user.SamAccountName)"
        pause
        return
    }
    catch{
        Write-PSFMessage -Level Host -Message "No matching account with SamAccountName found. Continuing...."
        
    }

    foreach ($Key in @($userAttributes.Keys)){
        if(-not $userAttributes[$key]){
            write-host "$($Key) is empty, will not be added to user creation"
            Start-Sleep -Milliseconds 50 
            $userAttributes.Remove($Key)
        }
    }
    Write-host "Will create User Account`n`n$($userAttributes | out-string)" -ForegroundColor Cyan
    pause
    New-ADUser @userAttributes
    for ($i = 1; $i -le 100; $i++ ) {
        Write-Progress -Activity "Creating Account" -Status "$i% Complete:" -PercentComplete $i
        Start-Sleep -Milliseconds 10
    }
    
	
	for ($i = 1; $i -le 100; $i++ ){
    try {
        get-ADuser $userAttributes.SamAccountName | Set-ADUser -Replace @{c="US";co="United States"}
        break
    }
    catch {
        Write-Host "Taking a little longer..."
        start-sleep 1
    }
}
try {
    Write-PSFMessage -Level Host -Message "User is created $(get-ADuser $userAttributes.SamAccountName)"
}
catch {
    Write-Error "Error Creating user account. "
    Write-PSFMessage -Level Error -Message "Error Creating user account."
    Write-PSFMessage -Level Error -Message "Error MSG: $($error[0])"
}
}

function new-Password {
<#
.PARAMETER
    None
        
.DESCRIPTION
Generates a passphase with two words from the password dict. Adjust the words to have uppercase letters, add special characters, and numbers. It will make sure the length is 14 character long.   
#>
    

    $Dict = Get-Content -Path "$PSScriptRoot\Resources\password.dict" | Sort-Object {Get-Random} | Select-Object -First $config.Settings.password.NumberofWords
    $WordList = @()
    $chars=@()

    Write-PSFMessage -Level Host -Message "Will use the following words to create password $($Dict)"

    foreach ($Word in $Dict) {
        $word = [char[]]$word
        for ($i = 0; $i -lt $word.Count; $i++) {
            if($i -in 0,1){
                if(get-random -InputObject 0,1){
                    $word[$i] = $word[$i].ToString().ToUpper()
                }
            }
            else{
                $word[$i] = $word[$i]
            }
        }
        $chars = $word -join ''
        $WordList += $chars -join ''
    }
    
    #$specialChar = ($config.Settings.password.SpecialChar).Split(",")
    $Delimiters = @()
    if($config.Settings.password.Numbers){
        $Delimiters += '1','2','3','4','5','6','7','8','9','0'   
    }
    if($config.Settings.password.SpecialChar.count -gt 0){
        $Delimiters += ($config.Settings.password.SpecialChar).Split(',')
    }
    for ($i = 0; $i -lt $config.Settings.password.NumberofWords; $i++) {
       $Global:password += $WordList[$i] + (Get-Random -InputObject $Delimiters)
    }

    for ($i = 1; $i -le 100; $i++ ) {
        Write-Progress -Activity "Creating Password" -Status "$i% Complete:" -PercentComplete $i
        Start-Sleep -Milliseconds 10
    }

    if($global:Password.length -lt $config.Settings.password.RequiredLength){
		write-host "Extending Length" -foreground red
		start-sleep -Milliseconds 100
        for ($i = 0; $i -lt 14-$global:Password.length; $i++) {
            $global:Password += Get-Random -InputObject $Delimiters
        }
    }
    if($global:Password -cmatch "[A-Z]"){
        $global:Password = (Get-Culture).TextInfo.ToTitleCase($global:Password)
    }
    if($config.Settings.password.SpecialChar -gt 0){
        if(!($password -cmatch "[$($config.Settings.password.SpecialChar)]")){
            $global:Password += (get-random -InputObject ($config.Settings.password.SpecialChar).split(','))
        }
    }   

    
    Write-PSFMessage -Level Host -Message "New Password generated: $($global:Password)`n`n" 
	pause
}

function set-EntraCopyGroups {
<#
.PARAMETER
    -TargetUser
        expecting the Target user's userprincipalname to schedule the task of copying over entra/exchange groups

        Required?       True
    
    -SourceUser
        expecting the Source user's userprincipalname to schedule the task of copying from entra/exchange groups

        Required?       True
        
.DESCRIPTION
Creates a schedule task that will executes the scheduleEntraGroupCopy.ps1 script
#>
    param (
        [parameter(Mandatory=$false)]
        [int16]$timeDelay = 45
    )

    

    write-host "`nPlease enter in your Admin Account login`n" -ForegroundColor Cyan
    $taskDate= (get-date).AddMinutes(($timeDelay))
    $cred = get-credential -UserName $([System.Security.Principal.WindowsIdentity]::GetCurrent().Name)
    $user = $cred.UserName
    $pass = $cred.GetNetworkCredential().Password
    $trigger = New-ScheduledTaskTrigger -Once -At $taskDate
    $action  = New-ScheduledTaskAction -Execute "C:\Program Files\PowerShell\7\pwsh.exe" -Argument '-ExecutionPolicy Bypass -file "$PSScriptRoot\ScheduleEntraGroupCopy.ps1"' -WorkingDirectory "$PSScriptRoot"

    try{
        Register-ScheduledTask -TaskName 'CopyEntraGroups' -TaskPath '\Truvant\' -Action $action -Trigger $trigger -RunLevel Highest -User $user -Password $pass -ErrorAction Stop
        Write-PSFMessage -Level Host -Message "Schedule Task is created to copy Entra/Exchange groups at $($taskDate)"
    } catch{
        Write-Host "Error Creating schedule task" -ForegroundColor Red
        Write-PSFMessage -Level Error -Message "Error creating schedule task. Error MSG: $($Error[0])"
    }   

}

function export-csvinfo {
<#
.PARAMETER
    -TargetUser
        expecting the Target user's userprincipalname to schedule the task of copying over entra/exchange groups

        Required?       True
    
    -SourceUser
        expecting the Source user's userprincipalname to schedule the task of copying from entra/exchange groups

        Required?       True
        
.DESCRIPTION
adds the targetuser info and sourceuser to a csv file that is used by ScheduleEntraGroupCopy.ps1 and acts a simiple log of user created with password. 
#>
    param (
        [parameter(Mandatory=$true)]
        $TargetUser,
        [parameter(Mandatory=$true)]
        $SourceUser,
        [parameter(Mandatory=$false)]
        $EntraCompleted = $false
    )

    

    $csv = [PSCustomObject]@{
        SourceUser = $SourceUser
        TargetUser = $TargetUser
        Password = $global:Password
        Completed = $EntraCompleted
    }
    try {
        $csv | export-csv $PSScriptRoot\Resources\EntraUsers.csv -Force -Append 
        Write-PSFMessage -Level Verbose -Message "$($csv | out-string) appended to csv file successfully"
    }
    catch {
        Write-PSFMessage -Level error -Message "failed to export csv info"
        Write-PSFMessage -Level error -Message "Error MSG: $($Error[0])"
    } 
}
function Find-DomainUser {
<#
.PARAMETER
    -Message
        The prompt message for asking for the user to search for

        Required?       True
    
    -Mando
        True/False field that forces an account to be found. 

        Required?       True
        
.DESCRIPTION
Creates a schedule task that will executes the scheduleEntraGroupCopy.ps1 script
#>
    param (
        [parameter(Mandatory=$true)]
        $message,
        [parameter(Mandatory=$false)]
        $Mando = $false

    )
  
    
    while ($true){
        if($LookupUser -like "exit*"){
           if($Mando){
                write-host "Field is mandatory. You must exit tool completely. Please enter in name"
           }
           Else{
                Write-host "Exiting Search"
                break
           }
        }
        
        write-host  $message -ForegroundColor DarkCyan -NoNewline
        write-host "`nYou can search by entering the first name, full name or User ID.`nEnter 'Exit' to exit user search:" -ForegroundColor DarkCyan -NoNewline
        $LookupUser = (Read-Host) + "*"
        $Users = Get-ADUser -Filter {(name -like $LookupUser) -or (SamAccountName -like $LookupUser)} -SearchBase $config.Settings.SearchOU -Properties * -ErrorAction Stop | Select-Object Name, SamAccountName, UserPrincipalName, Office, Title, Manager
        
        if(-not $users){
            Write-Host "User not found.  Please try again." -ForegroundColor Red
            Write-PSFMessage -Level Verbose -Message "$($users) not found"
        } else {
            if($Users.Count -gt 1){
                $User = $Users | Out-ConsoleGridView -Title "Select Target User" -OutputMode Single
            }Else{
                Write-host "Found user $($Users)"
                $User = $Users
            }
            write-host "`n`n"
            return $user 
        }
    }
    write-host "`n`n"
}

function Copy-EntraGroups{
<#
.PARAMETER
    -SourceEmail
        Email address of the user to copy groups from
        Required?       True
    
    -targetEmail
        Email address of the user to copy groups to
        Required?       True
        
.DESCRIPTION
Copies groups from sourceEmail to the TargetEmail. used on existing emails. 
#>
    param (
        [parameter(Mandatory=$true)]
        $sourceEmail,
        [parameter(Mandatory=$true)]
        $targetEmail
    )
    
    

    $data = @()

    try {
        $source = get-mguser -userid $sourceEmail -ErrorAction Stop
        write-host "Found user account $($source.DisplayName) to copy from"
        $target = get-mguser -userid $targetEmail -ErrorAction Stop
        write-host "Found user account $($target.DisplayName) to copy to"
        
    
        # Pull all groups of source user
        $groups = Get-MgUserMemberOf -UserId $source.Id -all | Select-Object *
    
        $groups |  Where-Object {(($_.AdditionalProperties).onPremisesSyncEnabled -ne $true) -and (($_.AdditionalProperties)."@odata.type" -ne "#microsoft.graph.administrativeUnit") -and (($_.AdditionalProperties).groupTypes -ne "DynamicMembership" -or ($_.AdditionalProperties).groupTypes.count -eq 0 )} | ForEach-Object {
            $groupsDetails = "" | Select-Object "id", "displayName", "groupType","mail","mailEnabled","securityEnabled"
            $groupProperties = $_.AdditionalProperties
            $groupsDetails.id = $_.Id
            $groupsDetails.DisplayName  =   $groupProperties.displayName
            $groupsDetails.groupType =  $groupProperties.groupTypes
            $groupsDetails.mail =   $groupProperties.mail
            $groupsDetails.mailEnabled =    $groupProperties.mailEnabled
            $groupsDetails.securityEnabled =    $groupProperties.securityEnabled
        
            $data += $groupsDetails
        }
    
        $Entragroups = $data | Where-Object {($_.groupType -eq 'Unified') -or $_.MailEnabled -ne 'true'}
        $maillist = $data | Where-Object {$_.MailEnabled -eq 'true' -and ($_.groupType -ne 'Unified' -or $null -eq ($_.groupType)[0])}
        # Add Target user to groups from example user
		
		write-host "Attempting to add $($target.DisplayName) the following Entra:`n$($Entragroups.DisplayName | out-string)`n`nFollowing Exchange/M365 Groups:`n$($maillist.DisplayName | out-string)"
		
        foreach ($group in $Entragroups) {
            try{
                new-MgGroupMember -GroupId $group.id -DirectoryObjectId $target.id -ErrorAction SilentlyContinue
                Write-PSFMessage -Level Host -Message "Adding $($target.DisplayName) to group $($group.DisplayName)"
            }
            catch{
                if($_.exception.message -like "*references already exist*"){
                    Write-PSFMessage -Level Host -Message "$($group.DisplayName) Has already been added to $($target.UserPrincipalName)"
                }
            }
        }
        foreach ($mail in $maillist) {
            try{
                Add-DistributionGroupMember -Identity $mail.DisplayName -Member $target.UserPrincipalName -ErrorAction Stop
                Write-PSFMessage -Level Host -Message "Adding $($target.DisplayName) to group $($mail.DisplayName)"
            }
            catch {
                if($Error[0].CategoryInfo.Reason -like "*MemberAlreadyExistsException*"){
                    Write-PSFMessage -Level Host -Message "$($mail.DisplayName) Has already been added to $($target.UserPrincipalName)"
                }
                elseif($Error[0].CategoryInfo.Reason -like "*RecipientTaskException*"){
                    try{
                        Add-UnifiedGroupLinks -Identity $mail.DisplayName -LinkType Members -Links $target.UserPrincipalName -ErrorAction Stop
                        Write-PSFMessage -Level Host -Message "Adding $($target.DisplayName) to group $($mail.DisplayName)"
                    }
                    catch {
                        $Error[0].CategoryInfo.Reason
                    }
                }
            }
    
        } 
    }
    catch {
        write-host "ERROR: unable to find user(s). Exiting" -ForegroundColor Red
        Write-PSFMessage -Level Error -Message "ERROR: unable to find user(s). Exiting"
        Write-PSFMessage -Level Error -Message "$($Error[0])"
        break
    }
    

}

function Remove-EntraExchangeGroups {
<#
.PARAMETER
    -user
        Email address of the user to remove groups from
        Required?       True
        
.DESCRIPTION
Removed all entra/exchange groups from User.
#>
    param (
        [parameter(Mandatory=$true)]
        $user
    )
    
    $data =@()
    try {
        $userInfo = get-mguser -userid $user -ErrorAction Stop

        Write-PSFMessage -Level Host -Message "Found user account $($userInfo.DisplayName) to remove groups from"
        start-sleep 3
        $groups = Get-MgUserMemberOf -UserId $userInfo.id -all | Select-Object *

        $groups |  Where-Object {(($_.AdditionalProperties).onPremisesSyncEnabled -ne $true) -and (($_.AdditionalProperties)."@odata.type" -ne "#microsoft.graph.administrativeUnit") -and (($_.AdditionalProperties).membershipType -ne "Dynamic")} | ForEach-Object {
            $groupsDetails = "" | Select-Object "id", "displayName", "groupType","mail","mailEnabled","securityEnabled"
            $groupProperties = $_.AdditionalProperties
            $groupsDetails.id = $_.Id
            $groupsDetails.DisplayName  =   $groupProperties.displayName
            $groupsDetails.groupType =  $groupProperties.groupTypes
            $groupsDetails.mail =   $groupProperties.mail
            $groupsDetails.mailEnabled =    $groupProperties.mailEnabled
            $groupsDetails.securityEnabled =    $groupProperties.securityEnabled
    
            $data += $groupsDetails
        }

    
        $EntraGroups = $data | Where-Object {($null -eq ($_.groupType)[0] -or $_.groupType -ne 'DynamicMembership') -and $_.MailEnabled -ne 'true'} -ErrorAction SilentlyContinue
        $mailList = $data | Where-Object {$_.MailEnabled -eq 'true' -and ($_.groupType -ne 'DynamicMembership' -or $null -eq ($_.groupType)[0])}

        Write-PSFMessage -Level Host -Message "removing $($userInfo.DisplayName) from:`n$($EntraGroups.DisplayName -join "`n")`n$($mailList.DisplayName -join "`n")"

        foreach ($group in $EntraGroups) {
            try{
                Remove-MgGroupMemberByRef -GroupId $group.id -DirectoryObjectId $userInfo.Id
                write-host "Removing $($userInfo.DisplayName) to group $($group.DisplayName)"
            }
            catch{
                Write-Error "Failed to removed $($group.displayName) due to: $($error[0])"
            }
        }

        foreach ($mail in $mailList) {
            try{
                Remove-DistributionGroupMember -Identity $mail.DisplayName -Member $userInfo.UserPrincipalName -Confirm:$false -ErrorAction Stop
                write-host "removing $($userInfo.DisplayName) from group $($mail.DisplayName)"
            }
            catch{
                try{
                    Remove-UnifiedGroupLinks -Identity $mail.DisplayName -LinkType Members -Links $userInfo.UserPrincipalName -Confirm:$false -ErrorAction Stop
                    write-host "Removing $($userInfo.DisplayName) to group $($mail.DisplayName)"
                }
                catch {
                       $Error[0].CategoryInfo.Reason
                }
            }
        }
    }
    catch {
        write-host "ERROR: unable to find user. Exiting" -ForegroundColor Red
        Write-PSFMessage -Level Error -Message "Unabled to find user account"
        Write-PSFMessage -Level Error -Message "$(error[0])"
        break
    }
}

function connect-EntraExchange {
<#
.PARAMETER
        
.DESCRIPTION
Connects to Entra/Exchange using app authentication. If App Authentications will prompt user for login. User must have permission on their account
#>
    
    try{
        $tenant_id = $config.settings.auth.tenant_id
        $certThumbprint = $config.settings.auth.certThumbprint
        $cert = Get-ChildItem Cert:\LocalMachine\My\$certThumbprint -ErrorAction Stop
        $client_id = $config.settings.auth.client_id
        Connect-Graph -ClientId $client_id -TenantId $tenant_id -Certificate $cert
        Write-Host "Connected to Graph via App Authentication" -ForegroundColor Yellow
        Write-PSFMessage -Level Verbose -Message "Connected to Graph via App Auth"
        }
    catch{
        Write-Host "Error connecting to MG Graph via App authentication. Trying to manually connect with user credentials`n" -ForegroundColor Red
        Write-PSFMessage -Level Error -Message "Error connecting to MG Graph via App authentication. Trying to manually connect with user credentials"
        try {
            connect-Graph -Scopes Directory.Read.All, Group.ReadWrite.All, User.ReadWrite.All -ErrorAction Stop
            Write-Host "Connected to Graph via Interactive browser" -ForegroundColor Yellow
            Write-PSFMessage -Level Verbose -Message "Connected graph via interactive browser"
        }
        catch {
            write-host "Error connecting. Exiting."
            Write-PSFMessage -Level Error -Message "Error connecting to graph"
            Write-PSFMessage -Level Error -Message "$($Error[0])"
            break
        }
        }
    try{
        Connect-ExchangeOnline -CertificateThumbPrint $config.settings.auth.certThumbprint -AppID $config.settings.auth.client_id -Organization $config.settings.auth.Organization -ErrorAction Stop -ShowBanner:$false
        Write-Host "Connected to Exchange via App Authentication" -ForegroundColor Yellow
        Write-PSFMessage -Level Verbose -Message "Connected to exchange via App Auth"
        }
    catch{
        Write-Host "Error connecting to Exchange via App authentication. Trying to manually connect with user credentials"
        Write-PSFMessage -Level Error -Message "Error connecting to Exchange via App authentication. Trying to manually connect with user credentials"
        try {
            Connect-ExchangeOnline -ErrorAction Stop -ShowBanner:$false
            Write-Host "Connected to Exchange via Interactive browser" -ForegroundColor Yellow
            Write-PSFMessage -Level Verbose -Message "Connected Exchange via interactive browser"
        }
        catch {
            write-host "Error connecting. Exiting."
            Write-PSFMessage -Level Error -Message "Error connecting to exchange"
            Write-PSFMessage -Level Error -Message "$($Error[0])"
            break
        }
        }
}

Function Copy-DomainUserPrompt{
<#
    .SYNOPSIS
    This function allows users to copy and remove groups between two accounts in Active Directory (AD) and Entra/Exchange.

    .DESCRIPTION
    This function will prompt the user for two accounts: the source user (from whom groups will be copied) and the target user (to whom groups will be copied).
    It provides the option to copy and remove groups from either or both Active Directory (AD) and Entra/Exchange.
#>
    
    #get target group 
    $TargetUserGroup =  Find-DomainUser -message "Enter the username to copy groups to." -Mando $true
    Write-PSFMessage -Level Verbose -Message "Target User is: $($TargetUserGroup | Out-String)"
    #get source group
    $SourceUserGroup =  Find-DomainUser -message "Enter the username to copy groups from." -Mando $true
    Write-PSFMessage -Level Verbose -Message "Source User is: $($SourceUserGroup | Out-String)"

    if($host.ui.PromptForChoice("","Do you want to adjust AD groups from $($TargetUserGroup.Name)",@('&Yes','&No'),1) -eq 0){
        if($host.ui.PromptForChoice("","Do you want to remove current AD groups from $($TargetUserGroup.Name)",@('&Yes','&No'),1) -eq 0){
            Write-PSFMessage -Level Verbose -Message "Removing all current groups from $($TargetUserGroup.name) and then adding all groups from $($SourceUserGroup.name)"
            Write-host "Will Remove the user's current AD groups first"
            remove-DomainUserGroups -SamAccountName $TargetUserGroup.SamAccountName
            Write-host "Will add groups to $($TargetUserGroup.name) from $($SourceUserGroup.name)"
            Copy-DomainUserGroups -TargetUser $TargetUserGroup.SamAccountName -SourceUser $SourceUserGroup.SamAccountName
        }
        else{
            Write-PSFMessage -Level Verbose -Message "Adding groups from $($SourceUserGroup.name) to $($TargetUserGroup.name)"
            Write-host "Will add groups to $($TargetUserGroup.name) from $($SourceUserGroup.name)"
            Copy-DomainUserGroups -TargetUser $TargetUserGroup.SamAccountName -SourceUser $SourceUserGroup.SamAccountName
        }
    }

    if($host.ui.PromptForChoice("","Do you want to Adjust Entra/Exchange groups from $($TargetUserGroup.Name)",@('&Yes','&No'),1) -eq 0){
        write-host "Connecting to Entra/Exchagne via app authentication"
        for ($i = 1; $i -le 100; $i++ ) {
            Write-Progress -Activity "Attempting to connect via App Authentication" -Status "$i% Complete:" -PercentComplete $i
            Start-Sleep -Milliseconds 10
        }
        connect-EntraExchange
        if($host.ui.PromptForChoice("","Do you want to remove current Entra/Exchange groups from $($TargetUserGroup.Name)",@('&Yes','&No'),1) -eq 0){
            Write-PSFMessage -Level Verbose -Message "Removing Entra groups first before copying over groups to $($TargetUserGroup.name) from $($SourceUserGroup.name)"
            Write-host "Will Remove the user's current Entra/Exchange groups first"
            Remove-EntraExchangeGroups -user $TargetUserGroup.UserPrincipalName
            Write-host "Will add groups to $($TargetUserGroup.name) from $($SourceUserGroup.name)"
            Copy-EntraGroups -targetEmail $TargetUserGroup.UserPrincipalName -sourceEmail $SourceUserGroup.UserPrincipalName
        }
        else{
            Write-PSFMessage -Level Verbose -Message "Adding groups to $($TargetUserGroup.name) from $($SourceUserGroup.name)"
            Write-host "Will add groups to $($TargetUserGroup.name) from $($SourceUserGroup.name)"
            Copy-EntraGroups -targetEmail $TargetUserGroup.UserPrincipalName -sourceEmail $SourceUserGroup.UserPrincipalName
        }
        write-host "Disconnecting.." -ForegroundColor DarkCyan
        Write-PSFMessage -Level Verbose -Message "Disconnecting from Graph and Exchange"
        Disconnect-ExchangeOnline -Confirm:$false
        Disconnect-Graph
    }
    
}

function Get-DomainUserData {
    <#
    .SYNOPSIS
    This function collects user details, creates the user, copies AD groups immediately and Entra/Exchange groups via a scheduled task from another user, and notifies the manager of the new user’s credentials

    .DESCRIPTION
    This function will prompt the user for the following information: First Name, Last Name, Job Title, Department, Employee ID, Manager, Office, and Phone Extension.
    Once the user is created, the tool provides the option to copy both Active Directory (AD) groups immediately and Entra/Exchange groups via a scheduled task from another user.
    After completion, it will notify the manager of the new user’s credentials.
    #>
    
    while(($givenName -eq '') -or ($NULL -eq $givenName)){
        $givenName = $(write-Host "Enter User's First Name: " -ForegroundColor Magenta -NoNewline; Read-Host)
    }
    while(($surName -eq '') -or ($NULL -eq $surName)){
        $surName = $(write-Host "Enter User's Last Name: " -ForegroundColor Magenta -NoNewline; Read-Host)
    }
    $attributes.add("givenName",$givenName)
    $attributes.add("SurName",$surName)
    $attributes.add("title",(Read-Host "Enter User's Title"))
    
    
    $attributes.add("department",(Read-Host "Enter User's department"))
    $attributes.add("Employee ID",(Read-Host "Enter User's Employee ID #"))
    
    $Manager =  Find-DomainUser -message "Enter the user's manager."
    Write-PSFMessage -Level host -Message "You selected $($Manager.Name)"
    $attributes.Add("Manager",$Manager.SamAccountName)	
    
    
    
    
    if($config.Settings.Company.count -gt 1) {
        $attributes.Add("Company",($config.Settings.Company | Out-ConsoleGridView -Title "Select the Company Name" -OutputMode Single))
    }
    else {
        $attributes.Add("Company",($config.Settings.Company))
    }
    
    write-host "`nPlease Select Office Location`n"
    Start-sleep -Milliseconds 800
    $office = $config.SiteCode | Out-ConsoleGridView -Title "select office location" -OutputMode Single
    Write-PSFMessage -Level host -Message "You selected $($Office.Name)"
    
    ##if Office fields are empty in config it will prompt for them. 
    ##City
    If($NULL -eq $config.site.$($office.Code).City){
        $attributes.add("City",(Read-host "What City is the user in"))
    } else {
        $attributes.add("City",$config.site.$($office.Code).City)
    }
    ##Office
    If($NULL -eq $config.site.$($office.Code).office){
        $attributes.add("office",(Read-host "What is the office name is the user in"))
    } else {
        $attributes.add("office",$config.site.$($office.Code).office)
    }
    #OU
    $attributes.add("OU",$config.site.$($office.Code).OU)
    <#try {
        if(Get-ADOrganizationalUnit -filter "DistringuishedName -eq '$($config.site.$($office.Code).OU)'") {
            $attributes.add("OU",$config.site.$($office.Code).OU)
        }
    }
    catch {
        Write-PSFMessage -level Error -message "Office OU is not found. Please correct the config file for $($office.Code). The current value is: $($config.site.$($office.Code).OU)"
        Pause
        exit
    }#>
    #State
    If($NULL -eq $config.site.$($office.Code).state){
        $attributes.add("State",(Read-host "What State is the user in"))
    } else {
        $attributes.add("State",$config.site.$($office.Code).State)
    }
    #PostalCode
    If($NULL -eq $config.site.$($office.Code).PostalCode){
        $attributes.add("postalCode",(Read-host "What PostalCode is the user in"))
    } else {
        $attributes.add("postalCode",$config.site.$($office.Code).PostalCode)
    }
    


    # $attributes.add("City",$config.site.$($office.Code).City)
    # $attributes.add("office",$config.site.$($office.Code).office) 
    # $attributes.add("postalCode",$config.site.$($office.Code).PostalCode)
    # $attributes.add("ou",$config.site.$($office.Code).OU)    
    
    
    if($host.ui.PromptForChoice("","Do you want Enter in user's phone extension",@('&Yes','&No'),1) -eq 0){
            $attributes.add("Extesion",(read-host "Enter User Extension Number"))
    }
    $attributes.add("UserPrincipalName",$attributes.GivenName + "." + $attributes.Surname + $config.Settings.maildomain)
    $attributes.add("SamAccountName",$attributes.GivenName + "." + $attributes.Surname)
    
    Write-PSFMessage -Level Verbose -Message "User entered in the following: $($attributes | Out-String)"
    
    write-host "`n`n"
    new-Password
    write-host "`n`n"
    New-DomainUser -user $attributes
    
    
    
    if(($host.ui.PromptForChoice("","Do you want to copy groups from an AD user",@('&Yes','&No'),1)) -eq 0){
        $SourceUser =  Find-DomainUser -message "Enter the user to copy groups from."

        Copy-DomainUserGroups -SourceUser $SourceUser.SamAccountName -TargetUser $attributes.SamAccountName
    
        if(($host.ui.PromptForChoice("","Do you want to copy groups from M365/Entra Also",@('&Yes','&No'),1)) -eq 0){
            Write-PSFMessage -Level Host -Message "Groups will be copied from $($SourceUser.Name) to $($attributes.givenName) $($attributes.Surname) on M365/Entra via scheduled task. `nOnce M365/Entra Groups are copied, notification email will be sent to manger with login info."
            set-EntraCopyGroups
            export-csvinfo -TargetUser $attributes.UserPrincipalName -SourceUser $SourceUser.UserPrincipalName

        }
    }
    else{
        write-host "No groups where copied. All groups must be added for AD" -ForegroundColor Red
        Write-PSFMessage -Level Verbose -Message "user selected to not copy any groups"
        export-csvinfo -TargetUser $attributes.UserPrincipalName -SourceUser $NewUser.UserPrincipalName -EntraCompleted $true
    
        if(($host.ui.PromptForChoice("","Do you want to notify the manager with user's login now?",@('&Yes','&No'),0)) -eq 0){
            $emailBody = Get-EmailTemplate -user $info -Manager $info.Manager -password $users.password
            Send-MailtoManager -user $info -manager $info.Manager -EmailBody $emailBody -config $config
            Write-host "Users login password is: $($Global:Password)"
            pause
        }else {
            Write-host "Users login password is: $($Global:Password)"
            pause
        }
    }
    
}



#--------------------------------------------------------Code Signature--------------------------------------------------------#
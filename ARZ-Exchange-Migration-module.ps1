#ARZ Exchange Online Management Shell
#Viktor Ahorner
#Date: 22.02.2023
#V1.4
# v1.2 - 12.01.2024 - [Peter Weghofer] added Reset-DelegatesAndRules


Write-Host 'Welcome to ARZ Exchange-Management-Module' -BackgroundColor Cyan -ForegroundColor DarkBlue
Write-Host 'Following functions have been integrated' -ForegroundColor  Cyan
Write-Host 'Export-MailboxFolderPermissions' -ForegroundColor Green
Write-Host 'Export-MailboxPermissions' -ForegroundColor Green
Write-Host 'Get-RoomDelegationReport' -ForegroundColor Green
Write-Host 'Enable-ExternalForwarding' -ForegroundColor Green
Write-Host 'Enable-InternalForwarding' -ForegroundColor Green
Write-Host 'Enable-ExternalForwardingForExternalDomain' -ForegroundColor Green
Write-Host 'Reset-DelegatesAndRules' -ForegroundColor Green
Write-Host 'New-CloudSharedMailbox' -ForegroundColor Green
Write-Host 'New-CloudRoomMailbox' -ForegroundColor Green
Write-Host 'New-CloudEquipmentMailbox' -ForegroundColor Green
Write-Host 'New-CloudDistributiongroup' -ForegroundColor Green
Write-Host 'Enable-MailboxDelegation' -ForegroundColor Green
Write-Host 'Complete-MailboxEnablement' -ForegroundColor Green



function Export-MailboxFolderPermissions {
    param($mailboxname)
    $exportlocation = 'C:\temp\'
    $mailboxfolders = Get-MailboxFolderStatistics -Identity $mailboxname 
    $mailboxpermissions = @()
    foreach ($mailboxfolder in $mailboxfolders) {
        Write-Host 'Trying to get permissions for '($mailboxfolder.identity) -ForegroundColor DarkGray
        try {
            $permission = Get-MailboxFolderPermission -Identity ($mailboxfolder.identity).replace($mailboxname, ($mailboxname + ':')) -ErrorAction stop
            Write-Host 'Successfull got permissions for '($mailboxfolder.identity) -ForegroundColor Green
            $mailboxpermissions += $permission
        }
        catch {
            Write-Host 'Unable to get permissions for '($mailboxfolder.identity) -ForegroundColor Magenta
        }
    }
    Write-Host 'Trying to export permissions to '($exportlocation + $mailboxname + 'permissions.csv') -ForegroundColor DarkGray
    try {
        $mailboxpermissions | ConvertTo-Csv | Out-File -FilePath ($exportlocation + $mailboxname + 'permissions.csv') -ErrorAction stop
        Write-Host 'Successfull exported permissions to '($exportlocation + $mailboxname + 'permissions.csv') -ForegroundColor Green
    }
    catch {
        Write-Host 'Unable to export permissions to '($exportlocation + $mailboxname + 'permissions.csv') -ForegroundColor Magenta
    }
}

function Export-MailboxPermissions {
    param($mailuser)
    $mailboxpermissions = @()
    foreach ($permission in (get-mailbox $mailuser | get-mailboxpermission)) {
        if (-not ($permission.User -like '*NT AUTHORITY*')) {
            write-host $permission.AccessRights' '$permission.User -ForegroundColor Green
            $mailboxpermissions += $permission
        }
    }

    return $mailboxpermissions
}
#Export-MailboxPermissions -mailuser mccloud.mailbox@onpremto.cloud

function Get-RoomDelegationReport {

    $resources = @()
    write-host 'Looking for roommailboxes' -ForegroundColor DarkGray
    $rooms = Get-Mailbox -RecipientTypeDetails roommailbox
    write-host $rooms.Count' roommailboxes found' -BackgroundColor DarkGray
    write-host 'Looking for euipmentmailboxes' -ForegroundColor DarkGray
    $equipment = Get-Mailbox -RecipientTypeDetails equipmentmailbox
    write-host $equipment.Count' equipmentmailboxes found' -BackgroundColor DarkGray
    write-host 'Looking for sharedmailbox' -ForegroundColor DarkGray
    $shared = Get-Mailbox -RecipientTypeDetails sharedmailbox
    write-host $shared.Count' sharedmailboxes found' -BackgroundColor DarkGray
    $resources += $rooms
    $resources += $equipment
    $resources += $shared
    $delegationlist = @()

    foreach ($resource in $resources) {
        $exportlocation = 'C:\temp\'
        Write-Host 'Looking for delegated  booking for '$resource.Alias
        $delegation = $resource | Get-CalendarProcessing

        if ($delegation.ResourceDelegates) {
            Write-Host 'Delegation found for '$resource.alias -ForegroundColor Green
            $delegationlist += $resource
        }

    }
    write-host 'Exporting report to '($exportlocation + 'resouce-calender-delegation-report.csv') -foregroundcolor DarkGray
    $delegationlist | select Displayname, UserPrincipalName, SamAccountName, ResourceType, ResourceCapacity | ConvertTo-Csv -Delimiter ';' | Out-File -FilePath ($exportlocation + 'resouce-calender-delegation-report.csv')
}

function Enable-ExternalForwarding {
    param($mailaddress, $externalrecipient)
    $usertounblock = $mailaddress
    Write-Host 'Looking for UserMailbox '$mailaddress -ForegroundColor DarkGray
    try {
        $mailbox = Get-Mailbox $usertounblock -ErrorAction stop
        Write-Host 'Successful' -ForegroundColor Green
        Write-Host 'Looking for Autoforward rule for External Recipients blocking rule' -ForegroundColor DarkGray
        try {
            $rule = Get-TransportRule | ? { $_.MessageTypeMatches -eq 'Autoforward' -and $_.SentToScope -eq 'NotInOrganization' } -ErrorAction stop
            Write-Host 'Successful' -ForegroundColor Green
            Write-Host 'Trying to configure Exception for '$rule.Name
            try {
                if (-not ($rule.ExceptIfFrom -like $mailbox.PrimarySmtpAddress)) {
                    Set-TransportRule $rule.Identity -ExceptIfFrom ($rule.ExceptIfFrom + $mailbox.PrimarySmtpAddress) -ErrorAction stop
                    Write-Host 'Successful' -ForegroundColor Green
                }
                else {
                    Write-Host 'User is already persisting in the Exception for '$rule.Name -ForegroundColor Yellow

                }
                if ($externalrecipient) {
                    try {
                        Write-Host 'Trying to create a new Inbound-Rule for Exernal-Recipient'
                        New-InboxRule -Name ('OnRequest-External-Forward ' + $externalrecipient) -Mailbox $mailbox -ForwardTo $externalrecipient
                        Write-Host 'Successful' -ForegroundColor Green

                    }
                    catch {
                        Write-Host 'Could not create a new inbox rule ' -ForegroundColor Magenta 
                        $mailusererror = $_.Exception.Message
                        Write-Host $mailusererror -BackgroundColor DarkYellow
                    }
                }
            }
            catch {
                Write-Host 'Setup exception!' -ForegroundColor Magenta 
                $mailusererror = $_.Exception.Message
                Write-Host $mailusererror -BackgroundColor DarkYellow
            }
        }
        catch {
            Write-Host 'Forwarding rule not found ' -ForegroundColor Magenta 
            $mailusererror = $_.Exception.Message
            Write-Host $mailusererror -BackgroundColor DarkYellow
        }
    }
    catch {
        Write-Host 'UserMailbox '$mailaddress' not found' -ForegroundColor Magenta
        $mailusererror = $_.Exception.Message
        Write-Host $mailusererror -BackgroundColor DarkYellow
    }

}

function Enable-InternalForwarding {
    param($mailaddress, $internalrecipient)
    $usertounblock = $mailaddress
    Write-Host 'Looking for UserMailbox '$mailaddress -ForegroundColor DarkGray
    try {
        $mailbox = Get-Mailbox $usertounblock -ErrorAction stop
        Write-Host 'Successful' -ForegroundColor Green


        if ($internalrecipient) {
            try {
                Write-Host 'Trying to create a new Inbound-Rule for Internal-Recipient'
                New-InboxRule -Name ('OnRequest-Internal-Forward ' + $internalrecipient) -Mailbox $mailbox -ForwardTo $internalrecipient
                Write-Host 'Successful' -ForegroundColor Green

            }
            catch {
                Write-Host 'Could not create a new inbox rule ' -ForegroundColor Magenta 
                $mailusererror = $_.Exception.Message
                Write-Host $mailusererror -BackgroundColor DarkYellow
            }
        }
    }
    catch {
        Write-Host 'UserMailbox '$mailaddress' not found' -ForegroundColor Magenta
        $mailusererror = $_.Exception.Message
        Write-Host $mailusererror -BackgroundColor DarkYellow
    }

}


function Enable-ExternalForwardingForExternalDomain {
    param([string]$externaldomain)
    try {
        $rule = Get-TransportRule | ? { $_.MessageTypeMatches -eq 'Autoforward' -and $_.SentToScope -eq 'NotInOrganization' } -ErrorAction stop
        Write-Host 'Transportrule-Found' -ForegroundColor DarkGray
        Write-Host 'Trying to configure Exception for '$externaldomain
        try {
            if (-not ($rule.ExceptIfRecipientAddressMatchesPatterns -like $externaldomain)) {
                Set-TransportRule $rule.Identity -ExceptIfRecipientAddressMatchesPatterns ($rule.ExceptIfRecipientAddressMatchesPatterns + $externaldomain) -ErrorAction stop
                Write-Host 'Successful' -ForegroundColor Green
            }
            else {
                Write-Host 'User is already persisting in the Exception for '$rule.Name -ForegroundColor Yellow

            }
            if ($externalrecipient) {
                try {
                    Write-Host 'Trying to create a new Inbound-Rule for Exernal-Recipient'
                    New-InboxRule -Name ('OnRequest-External-Forward ' + $externalrecipient) -Mailbox $mailbox -ForwardTo $externalrecipient
                    Write-Host 'Successful' -ForegroundColor Green

                }
                catch {
                    Write-Host 'Could not create a new inbox rule ' -ForegroundColor Magenta 
                    $mailusererror = $_.Exception.Message
                    Write-Host $mailusererror -BackgroundColor DarkYellow
                }
            }
        }
        catch {
            Write-Host 'Setup exception!' -ForegroundColor Magenta 
            $mailusererror = $_.Exception.Message
            Write-Host $mailusererror -BackgroundColor DarkYellow
        }
    }
    catch {
        Write-Host 'Forwarding rule not found ' -ForegroundColor Magenta 
        $mailusererror = $_.Exception.Message
        Write-Host $mailusererror -BackgroundColor DarkYellow
    }

}


#region Reset-DelegatesAndRules

function Reset-DelegatesAndRules {
    #static variables
    [string]$foldertype="Calendar"

    # Test if Exchange Online is connected and connect if not
    if (Get-ConnectionInformation|where {$_.state -eq "Connected" -and $_.Name -like "ExchangeOnline*"}) {
        write-host "Already connected to Exchange" -ForegroundColor DarkGray
    } else {
        Write-Host "Connecting to Exchange Online" -ForegroundColor DarkGray
        Connect-ExchangeOnline
    }

    $Tenant=(Get-ConnectionInformation|where {$_.state -eq "Connected" -and $_.Name -like "ExchangeOnline*"}).userprincipalname.split("@")[1]

    # prompt for mailaddress until valid mailaddress is entered and mailbox is found
    $SingleMailbox = $false
    while (!($SingleMailbox)) {
        $mailaddress = Read-Host "Enter Mailaddress"
        # check if mailaddress is valid
        if ($mailaddress -match "^[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+$") {
            # check if mailbox exists exactly once
            $Mailbox = @(Get-Mailbox -Identity $mailaddress -ErrorAction SilentlyContinue)
            if ($Mailbox) {
                if ($Mailbox.Count -eq 1) {
                    $SingleMailbox = $true
                    Write-host "Mailbox found: $($Mailbox[0].DisplayName)" -ForegroundColor Green
                } else {
                    Write-Host "Mailbox not found or not unique" -ForegroundColor Magenta
                }
            } else {
                Write-Host "Mailbox not found or not unique" -ForegroundColor Magenta
            }
        } else {
            Write-Host "Invalid Mailaddress Format. Correct Format is first.last@domain.com" -ForegroundColor Magenta
        }
    }

    if ($SingleMailbox){
        try {
            # MailboxFolder
            $CurrentMailbox = $($Mailbox[0])
            Write-host "Processing Mailbox $($CurrentMailbox.DisplayName) - $($CurrentMailbox.PrimarySmtpAddress)" -ForegroundColor DarkGray
            Write-Host "Getting Folderstatistics - this can take a couple of seconds..." -ForegroundColor DarkGray
            $folders=@(Get-MailboxFolderStatistics -Identity $CurrentMailbox | Where-Object {$_.foldertype -eq $foldertype})

            Write-Host "Resetting Calendar Permissions and DelegateUserCollection.." -ForegroundColor DarkGray
            foreach ($Folder in $folders){
                $FolderID = "{0}:\{1}" -f $CurrentMailbox.PrimarySmtpAddress,$Folder.Folderpath
                # normalize Slashes
                $CurrentFolderID = $FolderID.Replace("/","\").Replace("\\","\")
                Write-Host "Are you sure you want to remove all permissions from $($CurrentFolderID)? (y/n)" -ForegroundColor Yellow
                $confirmation = Read-Host
                if ($confirmation -eq "y") {
                    Write-Host "Removing permissions from $CurrentFolderID" -ForegroundColor DarkGray
                    Remove-Mailboxfolderpermission -Identity $CurrentFolderID -ResetDelegateUserCollection -Confirm:$false
                } else {
                    Write-Host "Aborting" -ForegroundColor Magenta
                }
            }
            Write-Host "Removing Inbox Rules..." -ForegroundColor DarkGray
            Get-InboxRule -Mailbox $CurrentMailbox -IncludeHidden | Remove-InboxRule -Confirm:$false
        } catch {
            Write-Host "An Error occurred !!" -ForegroundColor Magenta
            $DelegatePermissionsError = $_.Exception.Message
            Write-Host $DelegatePermissionsError -BackgroundColor DarkYellow
        }
    }
}
#endregion

function New-CloudSharedMailbox
{
param($identity)


do
{
Write-Host 'Would you like to create new shared mailbox it? (Y/N) : ' -ForegroundColor Cyan
$objectyn = Read-Host
}while(($objectyn -ne 'Y') -and ($objectyn -ne 'N'))

if($objectyn -eq 'Y')
{
Write-Host 'Starting with object creation' -ForegroundColor DarkGray
if($identity)
{
$remoteroutingaddress = ($identity+'.EXO@'+$global:routingdomain)
}
else
{
do
{
Write-Host 'Please enter Identity/SamAccountName of the AD-Object for which Shared-Mailbox should be enabled : ' -ForegroundColor Cyan
$identity = Read-Host
}while($identity -eq '')
do
{
Write-Host 'Please enter Displayname : ' -ForegroundColor Cyan
$displayname = Read-Host
}while($mandantenid = $null)

do
{
Write-Host 'Please enter primary Email-address : ' -ForegroundColor Cyan
$primarysmtp = Read-Host
}while($mandantenid = $null)


}

try
{
Write-Host 'Creating a new Mailbox '$identity -ForegroundColor DarkGray

New-Mailbox -Shared -PrimarySmtpAddress $primarysmtp -DisplayName $displayname -Name $identity 
Set-MailboxRegionalConfiguration -identity $identity -Language 'de-AT' -DateFormat 'dd.MM.yyyy' -TimeFormat 'HH:mm' -LocalizeDefaultFolderName:$true -ErrorAction stop
Write-Host 'Creation successfull' -ForegroundColor Green

}
catch
{

$mailusererror = $_.Exception.Message
if($mailusererror -like '*The proxy address*')
{
Write-Host $mailusererror -BackgroundColor DarkYellow
Write-Host 'Trying to resolve the proxy address conflict' -ForegroundColor DarkGray
$smtpaddress = $mailusererror.Split('"')[1]
$proxyauser = $mailusererror.Split('"')[3]
Try
{
$mailuser = Get-MailUser $smtpaddress.Split(':')[1] -ErrorAction stop

}
catch
{
$mailuser = Get-DistributionGroup $smtpaddress.Split(':')[1] -ErrorAction SilentlyContinue
}
Write-Host 'Proxy-Address conflict is persisting  '($smtpaddress.Split(':')[1])' from '$mailuser.DisplayName -ForegroundColor DarkGray
#Set-MailUser -Identity $mailuser.Alias -EmailAddresses @{remove=($smtpaddress.Split(':')[1])} -ErrorAction stop
Write-Host '------------- COPY following text to your ticket ---------------------'
Write-Host 'Sehr geehrter Kunde!'
Write-host 'Die von Ihnen gewählte Proxy-Adresse '($smtpaddress.Split(':')[1])' ist bereits vergeben.'
write-host 'Bitte bereinigen Sie die Proxy-Adresse für den User '$proxyauser', oder löschen Sie das AD-Objekt.'
write-host 'Wir erwarten auf baldige Rückmeldung wenn die Proxy-Adresse '($smtpaddress.Split(':')[1])' von Ihnen bereinigt wurde.'
write-host 'Wir wünschen Ihnen einen wunderschönen Tag!'
}


if($mailuser)
{
Write-Host 'A remote mailbox is already persisting for '($mailuser.UserPrincipalName) -ForegroundColor Green
}

}
$mailusererror = $_.Exception.Message
Write-Host $mailusererror -BackgroundColor DarkYellow
#Write-Host $mailusererror
}
}

function New-CloudroomMailbox
{
param($identity)


do
{
Write-Host 'Would you like to create new room mailbox it? (Y/N) : ' -ForegroundColor Cyan
$objectyn = Read-Host
}while(($objectyn -ne 'Y') -and ($objectyn -ne 'N'))

if($objectyn -eq 'Y')
{
Write-Host 'Starting with object creation' -ForegroundColor DarkGray
if($identity)
{
$remoteroutingaddress = ($identity+'.EXO@'+$global:routingdomain)
}
else
{
do
{
Write-Host 'Please enter Identity/SamAccountName of the AD-Object for which room-Mailbox should be enabled : ' -ForegroundColor Cyan
$identity = Read-Host
}while($identity -eq '')
do
{
Write-Host 'Please enter Displayname : ' -ForegroundColor Cyan
$displayname = Read-Host
}while($mandantenid -eq '')

do
{
Write-Host 'Please enter primary Email-address : ' -ForegroundColor Cyan
$primarysmtp = Read-Host
}while($primarysmtp -eq '')


}

try
{
Write-Host 'Creating a new Mailbox '$identity -ForegroundColor DarkGray

New-Mailbox -room -PrimarySmtpAddress $primarysmtp -DisplayName $displayname -Name $identity 
Write-Host 'Creation successfull' -ForegroundColor Green

}
catch
{

$mailusererror = $_.Exception.Message
if($mailusererror -like '*The proxy address*')
{
Write-Host $mailusererror -BackgroundColor DarkYellow
Write-Host 'Trying to resolve the proxy address conflict' -ForegroundColor DarkGray
$smtpaddress = $mailusererror.Split('"')[1]
$proxyauser = $mailusererror.Split('"')[3]
Try
{
$mailuser = Get-MailUser $smtpaddress.Split(':')[1] -ErrorAction stop

}
catch
{
$mailuser = Get-DistributionGroup $smtpaddress.Split(':')[1] -ErrorAction SilentlyContinue
}
Write-Host 'Proxy-Address conflict is persisting  '($smtpaddress.Split(':')[1])' from '$mailuser.DisplayName -ForegroundColor DarkGray
#Set-MailUser -Identity $mailuser.Alias -EmailAddresses @{remove=($smtpaddress.Split(':')[1])} -ErrorAction stop
Write-Host '------------- COPY following text to your ticket ---------------------'
Write-Host 'Sehr geehrter Kunde!'
Write-host 'Die von Ihnen gewählte Proxy-Adresse '($smtpaddress.Split(':')[1])' ist bereits vergeben.'
write-host 'Bitte bereinigen Sie die Proxy-Adresse für den User '$proxyauser', oder löschen Sie das AD-Objekt.'
write-host 'Wir erwarten auf baldige Rückmeldung wenn die Proxy-Adresse '($smtpaddress.Split(':')[1])' von Ihnen bereinigt wurde.'
write-host 'Wir wünschen Ihnen einen wunderschönen Tag!'
}


if($mailuser)
{
Write-Host 'A remote mailbox is already persisting for '($mailuser.UserPrincipalName) -ForegroundColor Green
}

}
$mailusererror = $_.Exception.Message
Write-Host $mailusererror -BackgroundColor DarkYellow
#Write-Host $mailusererror
}
}

function New-CloudEquipmentMailbox
{
param($identity)


do
{
Write-Host 'Would you like to create new Equipment mailbox it? (Y/N) : ' -ForegroundColor Cyan
$objectyn = Read-Host
}while(($objectyn -ne 'Y') -and ($objectyn -ne 'N'))

if($objectyn -eq 'Y')
{
Write-Host 'Starting with object creation' -ForegroundColor DarkGray
if($identity)
{
$remoteroutingaddress = ($identity+'.EXO@'+$global:routingdomain)
}
else
{
do
{
Write-Host 'Please enter Identity/SamAccountName of the AD-Object for which Equipment-Mailbox should be enabled : ' -ForegroundColor Cyan
$identity = Read-Host
}while($identity -eq '')
do
{
Write-Host 'Please enter Displayname : ' -ForegroundColor Cyan
$displayname = Read-Host
}while($displayname -eq '')

do
{
Write-Host 'Please enter primary Email-address : ' -ForegroundColor Cyan
$primarysmtp = Read-Host
}while($primarysmtp -eq '')


}

try
{
Write-Host 'Creating a new Mailbox '$identity -ForegroundColor DarkGray

New-Mailbox -Equipment -PrimarySmtpAddress $primarysmtp -DisplayName $displayname -Name $identity -ErrorAction stop
Write-Host 'Creation successfull' -ForegroundColor Green

}
catch
{

$mailusererror = $_.Exception.Message
if($mailusererror -like '*The proxy address*')
{
Write-Host $mailusererror -BackgroundColor DarkYellow
Write-Host 'Trying to resolve the proxy address conflict' -ForegroundColor DarkGray
$smtpaddress = $mailusererror.Split('"')[1]
$proxyauser = $mailusererror.Split('"')[3]
Try
{
$mailuser = Get-MailUser $smtpaddress.Split(':')[1] -ErrorAction stop

}
catch
{
$mailuser = Get-DistributionGroup $smtpaddress.Split(':')[1] -ErrorAction SilentlyContinue
}
Write-Host 'Proxy-Address conflict is persisting  '($smtpaddress.Split(':')[1])' from '$mailuser.DisplayName -ForegroundColor DarkGray
#Set-MailUser -Identity $mailuser.Alias -EmailAddresses @{remove=($smtpaddress.Split(':')[1])} -ErrorAction stop
Write-Host '------------- COPY following text to your ticket ---------------------'
Write-Host 'Sehr geehrter Kunde!'
Write-host 'Die von Ihnen gewählte Proxy-Adresse '($smtpaddress.Split(':')[1])' ist bereits vergeben.'
write-host 'Bitte bereinigen Sie die Proxy-Adresse für den User '$proxyauser', oder löschen Sie das AD-Objekt.'
write-host 'Wir erwarten auf baldige Rückmeldung wenn die Proxy-Adresse '($smtpaddress.Split(':')[1])' von Ihnen bereinigt wurde.'
write-host 'Wir wünschen Ihnen einen wunderschönen Tag!'
}


if($mailuser)
{
Write-Host 'A remote mailbox is already persisting for '($mailuser.UserPrincipalName) -ForegroundColor Green
}

}
$mailusererror = $_.Exception.Message
Write-Host $mailusererror -BackgroundColor DarkYellow
#Write-Host $mailusererror
}
}

function New-CloudDistributionGroup
{

do
{
Write-Host 'Please enter Identity/SamAccountName of the DistributionGroup which should be enabled : ' -ForegroundColor Cyan
$identity = Read-Host
}while($identity -eq '')

do
{
Write-Host 'Please enter Displayname of the DistributionGroup which should be enabled : ' -ForegroundColor Cyan
$displayname = Read-Host
}while($displayname -eq '')

do
{
Write-Host 'Please enter PrimarySMTPAddress of the DistributionGroup which should be enabled : ' -ForegroundColor Cyan
$primarysmtp = Read-Host
}while($displayname -eq '')

Write-Host 'Please enter Description of the DistributionGroup which should be enabled : ' -ForegroundColor Cyan
$description = Read-Host

#do
#{
#Write-Host 'Please enter Remote-RoutingAddress of the Shared-Mailbox which should be enabled : ' -ForegroundColor Cyan
#$remoteroutingaddress = Read-Host
#}while(-not ($remoteroutingaddress -like '*@*'))

try
{
Write-Host 'Trying to Enable a Remote DistributionGroup for '$identity -ForegroundColor DarkGray
New-DistributionGroup -Name $identity -Alias $identity -Description $description -DisplayName $displayname -PrimarySmtpAddress $primarysmtp
Write-Host 'Successful created a DistributionGroup '$identity' was successful' -ForegroundColor Green

}
catch
{

$mailusererror = $_.Exception.Message
if($mailusererror -like '*The proxy address*')
{
Write-Host $mailusererror -BackgroundColor DarkYellow
Write-Host 'Trying to resolve the proxy address conflict' -ForegroundColor DarkGray
$smtpaddress = $mailusererror.Split('"')[1]
$proxyauser = $mailusererror.Split('"')[3]
Try
{
$mailuser = Get-MailUser $smtpaddress.Split(':')[1] -ErrorAction stop

}
catch
{
$mailuser = Get-DistributionGroup $smtpaddress.Split(':')[1] -ErrorAction SilentlyContinue
}
Write-Host 'Proxy-Address conflict is persisting  '($smtpaddress.Split(':')[1])' from '$mailuser.DisplayName -ForegroundColor DarkGray
#Set-MailUser -Identity $mailuser.Alias -EmailAddresses @{remove=($smtpaddress.Split(':')[1])} -ErrorAction stop
Write-Host '------------- COPY following text to your ticket ---------------------'
Write-Host 'Sehr geehrter Kunde!'
Write-host 'Die von Ihnen gewählte Proxy-Adresse '($smtpaddress.Split(':')[1])' ist bereits vergeben.'
write-host 'Bitte bereinigen Sie die Proxy-Adresse für den User '$proxyauser', oder löschen Sie das AD-Objekt.'
write-host 'Wir erwarten auf baldige Rückmeldung wenn die Proxy-Adresse '($smtpaddress.Split(':')[1])' von Ihnen bereinigt wurde.'
write-host 'Wir wünschen Ihnen einen wunderschönen Tag!'

}

if($mailusererror -like '*is already mail-enabled*')
{
Write-Host $mailusererror -BackgroundColor DarkYellow
Write-Host 'Trying to resolve the issue' -ForegroundColor DarkGray
$smtpaddress = $mailusererror.Split('"')[1]
$mailuser = Get-DistributionGroup $identity
if($mailuser)
{
Write-Host 'A DistributionGroup is already persisting for '($mailuser.DisplayName) -ForegroundColor Green
}

}

#Write-Host $mailusererror
$mailusererror = $_.Exception.Message
Write-Host $mailusererror -BackgroundColor DarkYellow
}


}

function Enable-MailboxDelegation
{
param($identity)
#m365_smbx_m044_ALM_sendas@
$global:routingdomain = (Get-AcceptedDomain | ?{$_.domainname -like '*mail.onmicrosoft.com*'}).domainname

do
{
Write-Host 'Please enter Identity/SamAccountName of the AD-Object for which Email-Delegation should be enabled : ' -ForegroundColor Cyan
$identity = Read-Host
}while($identity -eq '')


$mbx = Get-Mailbox $identity
#$mbx.RecipientTypeDetails
if($mbx.RecipientTypeDetails -eq 'Sharedmailbox')
{
$fullaccess=('m365_smbx_EXO_'+$identity+'@'+$global:routingdomain)
$sendadaccess=('m365_smbx_EXO_'+$identity+'_sendas@'+$global:routingdomain)
$sendonbehalfaccess=('m365_smbx_EXO_'+$identity+'_onbehalf@'+$global:routingdomain)

try
{
Write-Host 'Mailbox is persisting '$identity -ForegroundColor DarkGray
#--- enable full access group start

    try
    {
     Write-Host 'Trying to enable fullaccess delegation group '$fullaccess -ForegroundColor DarkGray
     New-DistributionGroup -Name ('m365_smbx_EXO_'+$identity) -Type Security -PrimarySmtpAddress $fullaccess -ErrorAction stop
     Write-Host 'Waiting for 20 seconds to apply the next step' -BackgroundColor DarkGray
     Start-Sleep -Seconds 20
     Set-DistributionGroup -Identity ('m365_smbx_EXO_'+$identity) -HiddenFromAddressListsEnabled:$true -HiddenGroupMembershipEnabled:$true
     Write-Host 'Successfull enabled '$fullaccess -ForegroundColor green

    }
    catch
    {


$mailusererror = $_.Exception.Message
if($mailusererror -like '*The proxy address*')
{
Write-Host $mailusererror -BackgroundColor DarkYellow
Write-Host 'Trying to resolve the proxy address conflict' -ForegroundColor DarkGray
$smtpaddress = $mailusererror.Split('"')[1]
$proxyauser = $mailusererror.Split('"')[3]
Try
{
$mailuser = Get-MailUser $smtpaddress.Split(':')[1] -ErrorAction stop

}
catch
{
$mailuser = Get-DistributionGroup $smtpaddress.Split(':')[1] -ErrorAction SilentlyContinue
}
Write-Host 'Proxy-Address conflict is persisting  '($smtpaddress.Split(':')[1])' from '$mailuser.DisplayName -ForegroundColor DarkGray
#Set-MailUser -Identity $mailuser.Alias -EmailAddresses @{remove=($smtpaddress.Split(':')[1])} -ErrorAction stop
Write-Host '------------- COPY following text to your ticket ---------------------'
Write-Host 'Sehr geehrter Kunde!'
Write-host 'Die von Ihnen gewählte Proxy-Adresse '($smtpaddress.Split(':')[1])' ist bereits vergeben.'
write-host 'Bitte bereinigen Sie die Proxy-Adresse für den User '$proxyauser', oder löschen Sie das AD-Objekt.'
write-host 'Wir erwarten auf baldige Rückmeldung wenn die Proxy-Adresse '($smtpaddress.Split(':')[1])' von Ihnen bereinigt wurde.'
write-host 'Wir wünschen Ihnen einen wunderschönen Tag!'
}
if($mailusererror -like '*is already mail-enabled*' -or $mailusererror -like '*"SamAccountName" is used by another recipient object*')
{
Write-Host 'Distributiongroup is already enabled for '$sendadaccess -ForegroundColor Green
}

if($mailusererror -like '*already exists*')
{
#Write-Host $mailusererror -BackgroundColor DarkYellow
Write-Host 'Distributiongroup is already persisting for '$fullaccess -ForegroundColor Green
}

#Write-Host $mailusererror

     
    }
#--- enable full access group end
#--- enable sendas group start

    try
    {
      Write-Host 'Trying to enable sendas delegation group '$sendadaccess -ForegroundColor DarkGray
     New-DistributionGroup -Name ('m365_smbx_EXO_'+$identity+'_sendas') -Type Security -PrimarySmtpAddress $sendadaccess -ErrorAction stop
     Write-Host 'Waiting for 20 seconds to apply the next step' -BackgroundColor DarkGray
     Start-Sleep -Seconds 20
     Set-DistributionGroup -Identity ('m365_smbx_EXO_'+$identity+'_sendas') -HiddenFromAddressListsEnabled:$true -HiddenGroupMembershipEnabled:$true
     Write-Host 'Successfull enabled '$sendadaccess -ForegroundColor green

    }
    catch
    {


$mailusererror = $_.Exception.Message
if($mailusererror -like '*The proxy address*')
{
Write-Host $mailusererror -BackgroundColor DarkYellow
Write-Host 'Trying to resolve the proxy address conflict' -ForegroundColor DarkGray
$smtpaddress = $mailusererror.Split('"')[1]
$proxyauser = $mailusererror.Split('"')[3]
Try
{
$mailuser = Get-MailUser $smtpaddress.Split(':')[1] -ErrorAction stop

}
catch
{
$mailuser = Get-DistributionGroup $smtpaddress.Split(':')[1] -ErrorAction SilentlyContinue
}
Write-Host 'Proxy-Address conflict is persisting  '($smtpaddress.Split(':')[1])' from '$mailuser.DisplayName -ForegroundColor DarkGray
#Set-MailUser -Identity $mailuser.Alias -EmailAddresses @{remove=($smtpaddress.Split(':')[1])} -ErrorAction stop
Write-Host '------------- COPY following text to your ticket ---------------------'
Write-Host 'Sehr geehrter Kunde!'
Write-host 'Die von Ihnen gewählte Proxy-Adresse '($smtpaddress.Split(':')[1])' ist bereits vergeben.'
write-host 'Bitte bereinigen Sie die Proxy-Adresse für den User '$proxyauser', oder löschen Sie das AD-Objekt.'
write-host 'Wir erwarten auf baldige Rückmeldung wenn die Proxy-Adresse '($smtpaddress.Split(':')[1])' von Ihnen bereinigt wurde.'
write-host 'Wir wünschen Ihnen einen wunderschönen Tag!'
}

if($mailusererror -like '*is already mail-enabled*' -or $mailusererror -like '*"SamAccountName" is used by another recipient object*')
{
Write-Host 'Distributiongroup is already enabled for '$sendadaccess -ForegroundColor Green
}

if($mailusererror -like '*already exists*')
{
#Write-Host $mailusererror -BackgroundColor DarkYellow
Write-Host 'Distributiongroup is already persisting for '$sendadaccess -ForegroundColor Green
}

#Write-Host $mailusererror

     
    }

#--- enable sendas group end

#--- enable sendonbehalf group start

    try
    {
          Write-Host 'Trying to enable sendas delegation group '$sendonbehalfaccess -ForegroundColor DarkGray
     New-DistributionGroup -Name ('m365_smbx_EXO_'+$identity+'_onbehalf') -Type Security -PrimarySmtpAddress $sendonbehalfaccess -ErrorAction stop
     Write-Host 'Waiting for 20 seconds to apply the next step' -BackgroundColor DarkGray
     Start-Sleep -Seconds 20
     Set-DistributionGroup -Identity ('m365_smbx_EXO_'+$identity+'_onbehalf') -HiddenFromAddressListsEnabled:$true -HiddenGroupMembershipEnabled:$true
     Write-Host 'Successfull enabled '$sendonbehalfaccess -ForegroundColor green
    }
    catch
    {


$mailusererror = $_.Exception.Message
if($mailusererror -like '*The proxy address*')
{
Write-Host $mailusererror -BackgroundColor DarkYellow
Write-Host 'Trying to resolve the proxy address conflict' -ForegroundColor DarkGray
$smtpaddress = $mailusererror.Split('"')[1]
$proxyauser = $mailusererror.Split('"')[3]
Try
{
$mailuser = Get-MailUser $smtpaddress.Split(':')[1] -ErrorAction stop

}
catch
{
$mailuser = Get-DistributionGroup $smtpaddress.Split(':')[1] -ErrorAction SilentlyContinue
}
Write-Host 'Proxy-Address conflict is persisting  '($smtpaddress.Split(':')[1])' from '$mailuser.DisplayName -ForegroundColor DarkGray
#Set-MailUser -Identity $mailuser.Alias -EmailAddresses @{remove=($smtpaddress.Split(':')[1])} -ErrorAction stop
Write-Host '------------- COPY following text to your ticket ---------------------'
Write-Host 'Sehr geehrter Kunde!'
Write-host 'Die von Ihnen gewählte Proxy-Adresse '($smtpaddress.Split(':')[1])' ist bereits vergeben.'
write-host 'Bitte bereinigen Sie die Proxy-Adresse für den User '$proxyauser', oder löschen Sie das AD-Objekt.'
write-host 'Wir erwarten auf baldige Rückmeldung wenn die Proxy-Adresse '($smtpaddress.Split(':')[1])' von Ihnen bereinigt wurde.'
write-host 'Wir wünschen Ihnen einen wunderschönen Tag!'
}
if($mailusererror -like '*is already mail-enabled*' -or $mailusererror -like '*"SamAccountName" is used by another recipient object*')
{
Write-Host 'Distributiongroup is already enabled for '$sendadaccess -ForegroundColor Green
}
if($mailusererror -like '*already exists*')
{
#Write-Host $mailusererror -BackgroundColor DarkYellow
Write-Host 'Distributiongroup is already persisting for '$sendonbehalfaccess -ForegroundColor Green
}

#Write-Host $mailusererror

    }
#--- enable sendonbehalf group start

    }
    catch
    {

             $mailusererror = $_.Exception.Message
            Write-Host $mailusererror -BackgroundColor DarkYellow
    }


Write-Host $mailusererror -BackgroundColor DarkYellow
}

if($mbx.RecipientTypeDetails -eq 'RoomMailbox')
{
$rrmoderator = ('m365_rr_EXO_'+$identity+'_moderator@'+$global:routingdomain)
$rrwrite = ('m365_rr_EXO_'+$identity+'_write@'+$global:routingdomain)
$rrbook = ('m365_rr_EXO_'+$identity+'_book@'+$global:routingdomain)

if($identity)
{
}
else
{

}
try
{
Write-Host 'looking if mailbox is persisting '$identity -ErrorAction stop -ForegroundColor DarkGray
Get-Mailbox -Identity $identity -ErrorAction stop
Write-Host 'Mailbox is persisting '$identity -ForegroundColor DarkGray
    try
    {
     Write-Host 'Trying to enable Moderator delegation group '$rrmoderator -ForegroundColor DarkGray
     New-DistributionGroup -Name ('m365_rr_EXO_'+$identity+'_moderator') -Type Security -PrimarySmtpAddress $rrmoderator -ErrorAction stop
     Write-Host 'Waiting for 20 seconds to apply the next step' -BackgroundColor DarkGray
     Start-Sleep -Seconds 20
     Set-DistributionGroup -Identity ('m365_rr_EXO_'+$identity+'_moderator') -HiddenFromAddressListsEnabled:$true -HiddenGroupMembershipEnabled:$true
     Write-Host 'Successfull enabled '$rrmoderator -ForegroundColor green

    }
    catch
    {
     try
     {
     Write-Host ('m365_rr_EXO_'+$identity+'_moderator')' does not exists' -ForegroundColor DarkGray
     Write-Host 'Successfull created '$rrmoderator -ForegroundColor green
     }
     catch
     {

$mailusererror = $_.Exception.Message
if($mailusererror -like '*The proxy address*')
{
Write-Host $mailusererror -BackgroundColor DarkYellow
Write-Host 'Trying to resolve the proxy address conflict' -ForegroundColor DarkGray
$smtpaddress = $mailusererror.Split('"')[1]
$proxyauser = $mailusererror.Split('"')[3]
Try
{
$mailuser = Get-MailUser $smtpaddress.Split(':')[1] -ErrorAction stop

}
catch
{
$mailuser = Get-DistributionGroup $smtpaddress.Split(':')[1] -ErrorAction SilentlyContinue
}
Write-Host 'Proxy-Address conflict is persisting  '($smtpaddress.Split(':')[1])' from '$mailuser.DisplayName -ForegroundColor DarkGray
#Set-MailUser -Identity $mailuser.Alias -EmailAddresses @{remove=($smtpaddress.Split(':')[1])} -ErrorAction stop
Write-Host '------------- COPY following text to your ticket ---------------------'
Write-Host 'Sehr geehrter Kunde!'
Write-host 'Die von Ihnen gewählte Proxy-Adresse '($smtpaddress.Split(':')[1])' ist bereits vergeben.'
write-host 'Bitte bereinigen Sie die Proxy-Adresse für den User '$proxyauser', oder löschen Sie das AD-Objekt.'
write-host 'Wir erwarten auf baldige Rückmeldung wenn die Proxy-Adresse '($smtpaddress.Split(':')[1])' von Ihnen bereinigt wurde.'
write-host 'Wir wünschen Ihnen einen wunderschönen Tag!'
}
if($mailusererror -like '*is already mail-enabled*' -or $mailusererror -like '*"SamAccountName" is used by another recipient object*')
{
Write-Host 'Distributiongroup is already enabled for '$rrwrite -ForegroundColor Green
}

if($mailusererror -like '*already exists*')
{
#Write-Host $mailusererror -BackgroundColor DarkYellow
Write-Host 'Distributiongroup is already persisting for '$rrmoderator -ForegroundColor Green
}

#Write-Host $mailusererror

     }
    }
    try
    {
     Write-Host 'Trying to enable Write delegation group '$rrwrite -ForegroundColor DarkGray
     New-DistributionGroup -Name ('m365_rr_EXO_'+$identity+'_write') -Type Security -PrimarySmtpAddress $rrwrite -ErrorAction stop
     Write-Host 'Waiting for 20 seconds to apply the next step' -BackgroundColor DarkGray
     Start-Sleep -Seconds 20
     Set-DistributionGroup -Identity ('m365_rr_EXO_'+$identity+'_write') -HiddenFromAddressListsEnabled:$true -HiddenGroupMembershipEnabled:$true
     Write-Host 'Successfull enabled '$rrwrite -ForegroundColor green
    }
    catch
    {
     try
     {
     Write-Host ('m365_rr_EXO_'+$identity+'_write')' does not exists' -ForegroundColor DarkGray
     }
     catch
     {

$mailusererror = $_.Exception.Message
if($mailusererror -like '*The proxy address*')
{
Write-Host $mailusererror -BackgroundColor DarkYellow
Write-Host 'Trying to resolve the proxy address conflict' -ForegroundColor DarkGray
$smtpaddress = $mailusererror.Split('"')[1]
$proxyauser = $mailusererror.Split('"')[3]
Try
{
$mailuser = Get-MailUser $smtpaddress.Split(':')[1] -ErrorAction stop

}
catch
{
$mailuser = Get-DistributionGroup $smtpaddress.Split(':')[1] -ErrorAction SilentlyContinue
}
Write-Host 'Proxy-Address conflict is persisting  '($smtpaddress.Split(':')[1])' from '$mailuser.DisplayName -ForegroundColor DarkGray
#Set-MailUser -Identity $mailuser.Alias -EmailAddresses @{remove=($smtpaddress.Split(':')[1])} -ErrorAction stop
Write-Host '------------- COPY following text to your ticket ---------------------'
Write-Host 'Sehr geehrter Kunde!'
Write-host 'Die von Ihnen gewählte Proxy-Adresse '($smtpaddress.Split(':')[1])' ist bereits vergeben.'
write-host 'Bitte bereinigen Sie die Proxy-Adresse für den User '$proxyauser', oder löschen Sie das AD-Objekt.'
write-host 'Wir erwarten auf baldige Rückmeldung wenn die Proxy-Adresse '($smtpaddress.Split(':')[1])' von Ihnen bereinigt wurde.'
write-host 'Wir wünschen Ihnen einen wunderschönen Tag!'
}

if($mailusererror -like '*is already mail-enabled*' -or $mailusererror -like '*"SamAccountName" is used by another recipient object*')
{
Write-Host 'Distributiongroup is already enabled for '$rrwrite -ForegroundColor Green
}

if($mailusererror -like '*already exists*')
{
#Write-Host $mailusererror -BackgroundColor DarkYellow
Write-Host 'Distributiongroup is already persisting for '$rrwrite -ForegroundColor Green
}

#Write-Host $mailusererror

     }
    }
    try
    {
     Write-Host 'Trying to enable Book delegation group '$rrbook -ForegroundColor DarkGray
     New-DistributionGroup -Name ('m365_rr_EXO_'+$identity+'_book') -Type Security -PrimarySmtpAddress $rrbook -ErrorAction stop
     Write-Host 'Waiting for 20 seconds to apply the next step' -BackgroundColor DarkGray
     Start-Sleep -Seconds 20
     Set-DistributionGroup -Identity ('m365_rr_EXO_'+$identity+'_book') -HiddenFromAddressListsEnabled:$true -HiddenGroupMembershipEnabled:$true
     Write-Host 'Successfull enabled '$rrbook -ForegroundColor green
    }
    catch
    {
     try
     {
     Write-Host ('m365_rr_EXO_'+$identity+'_book')' does not exists' -ForegroundColor DarkGray
     }
     catch
     {

$mailusererror = $_.Exception.Message
if($mailusererror -like '*The proxy address*')
{
Write-Host $mailusererror -BackgroundColor DarkYellow
Write-Host 'Trying to resolve the proxy address conflict' -ForegroundColor DarkGray
$smtpaddress = $mailusererror.Split('"')[1]
$proxyauser = $mailusererror.Split('"')[3]
Try
{
$mailuser = Get-MailUser $smtpaddress.Split(':')[1] -ErrorAction stop

}
catch
{
$mailuser = Get-DistributionGroup $smtpaddress.Split(':')[1] -ErrorAction SilentlyContinue
}
Write-Host 'Proxy-Address conflict is persisting  '($smtpaddress.Split(':')[1])' from '$mailuser.DisplayName -ForegroundColor DarkGray
#Set-MailUser -Identity $mailuser.Alias -EmailAddresses @{remove=($smtpaddress.Split(':')[1])} -ErrorAction stop
Write-Host '------------- COPY following text to your ticket ---------------------'
Write-Host 'Sehr geehrter Kunde!'
Write-host 'Die von Ihnen gewählte Proxy-Adresse '($smtpaddress.Split(':')[1])' ist bereits vergeben.'
write-host 'Bitte bereinigen Sie die Proxy-Adresse für den User '$proxyauser', oder löschen Sie das AD-Objekt.'
write-host 'Wir erwarten auf baldige Rückmeldung wenn die Proxy-Adresse '($smtpaddress.Split(':')[1])' von Ihnen bereinigt wurde.'
write-host 'Wir wünschen Ihnen einen wunderschönen Tag!'
}
if($mailusererror -like '*is already mail-enabled*' -or $mailusererror -like '*"SamAccountName" is used by another recipient object*')
{
Write-Host 'Distributiongroup is already enabled for '$rrwrite -ForegroundColor Green
}
if($mailusererror -like '*already exists*')
{
#Write-Host $mailusererror -BackgroundColor DarkYellow
Write-Host 'Distributiongroup is already persisting for '$rrbook -ForegroundColor Green
}

#Write-Host $mailusererror

     }
    }
    }
    catch
    {

             $mailusererror = $_.Exception.Message
            Write-Host $mailusererror -BackgroundColor DarkYellow
    }
Write-Host $mailusererror -BackgroundColor DarkYellow
}

if($mbx.RecipientTypeDetails -eq 'EquipmentMailbox')
{
$eqmoderator = ('m365_eq_EXO_'+$identity+'_moderator@'+$global:routingdomain)
$eqwrite = ('m365_eq_EXO_'+$identity+'_write@'+$global:routingdomain)
$eqbook = ('m365_eq_EXO_'+$identity+'_book@'+$global:routingdomain)

if($identity)
{
}
else
{

}
try
{
Write-Host 'looking if mailbox is persisting '$identity -ErrorAction stop -ForegroundColor DarkGray
Get-Mailbox -Identity $identity -ErrorAction stop
Write-Host 'Mailbox is persisting '$identity -ForegroundColor DarkGray
    try
    {
     Write-Host 'Trying to enable Moderator delegation group '$eqmoderator -ForegroundColor DarkGray
     New-DistributionGroup -Name ('m365_eq_EXO_'+$identity+'_moderator') -Type Security -PrimarySmtpAddress $eqmoderator -ErrorAction stop
     Write-Host 'Waiting for 20 seconds to apply the next step' -BackgroundColor DarkGray
     Start-Sleep -Seconds 20
     Set-DistributionGroup -Identity ('m365_eq_EXO_'+$identity+'_moderator') -HiddenFromAddressListsEnabled:$true -HiddenGroupMembershipEnabled:$true
     Write-Host 'Successfull enabled '$eqmoderator -ForegroundColor green

    }
    catch
    {
     try
     {
     Write-Host ('m365_eq_EXO_'+$identity+'_moderator')' does not exists' -ForegroundColor DarkGray
     }
     catch
     {

$mailusererror = $_.Exception.Message
if($mailusererror -like '*The proxy address*')
{
Write-Host $mailusererror -BackgroundColor DarkYellow
Write-Host 'Trying to resolve the proxy address conflict' -ForegroundColor DarkGray
$smtpaddress = $mailusererror.Split('"')[1]
$proxyauser = $mailusererror.Split('"')[3]
Try
{
$mailuser = Get-MailUser $smtpaddress.Split(':')[1] -ErrorAction stop

}
catch
{
$mailuser = Get-DistributionGroup $smtpaddress.Split(':')[1] -ErrorAction SilentlyContinue
}
Write-Host 'Proxy-Address conflict is persisting  '($smtpaddress.Split(':')[1])' from '$mailuser.DisplayName -ForegroundColor DarkGray
#Set-MailUser -Identity $mailuser.Alias -EmailAddresses @{remove=($smtpaddress.Split(':')[1])} -ErrorAction stop
Write-Host '------------- COPY following text to your ticket ---------------------'
Write-Host 'Sehr geehrter Kunde!'
Write-host 'Die von Ihnen gewählte Proxy-Adresse '($smtpaddress.Split(':')[1])' ist bereits vergeben.'
write-host 'Bitte bereinigen Sie die Proxy-Adresse für den User '$proxyauser', oder löschen Sie das AD-Objekt.'
write-host 'Wir erwarten auf baldige Rückmeldung wenn die Proxy-Adresse '($smtpaddress.Split(':')[1])' von Ihnen bereinigt wurde.'
write-host 'Wir wünschen Ihnen einen wunderschönen Tag!'
}
if($mailusererror -like '*is already mail-enabled*' -or $mailusererror -like '*"SamAccountName" is used by another recipient object*')
{
Write-Host 'Distributiongroup is already enabled for '$eqwrite -ForegroundColor Green
}

if($mailusererror -like '*already exists*')
{
#Write-Host $mailusererror -BackgroundColor DarkYellow
Write-Host 'Distributiongroup is already persisting for '$eqmoderator -ForegroundColor Green
}

#Write-Host $mailusererror

     }
    }
    try
    {
     Write-Host 'Trying to enable Moderator delegation group '$eqwrite -ForegroundColor DarkGray
     New-DistributionGroup -Name ('m365_eq_EXO_'+$identity+'_write') -Type Security -PrimarySmtpAddress $eqwrite -ErrorAction stop
     Write-Host 'Waiting for 20 seconds to apply the next step' -BackgroundColor DarkGray
     Start-Sleep -Seconds 20
     Set-DistributionGroup -Identity ('m365_eq_EXO_'+$identity+'_write') -HiddenFromAddressListsEnabled:$true -HiddenGroupMembershipEnabled:$true
     Write-Host 'Successfull enabled '$eqwrite -ForegroundColor green

    }
    catch
    {
     try
     {
     Write-Host ('m365_eq_EXO_'+$identity+'_write')' does not exists' -ForegroundColor DarkGray
     }
     catch
     {

$mailusererror = $_.Exception.Message
if($mailusererror -like '*The proxy address*')
{
Write-Host $mailusererror -BackgroundColor DarkYellow
Write-Host 'Trying to resolve the proxy address conflict' -ForegroundColor DarkGray
$smtpaddress = $mailusererror.Split('"')[1]
$proxyauser = $mailusererror.Split('"')[3]
Try
{
$mailuser = Get-MailUser $smtpaddress.Split(':')[1] -ErrorAction stop

}
catch
{
$mailuser = Get-DistributionGroup $smtpaddress.Split(':')[1] -ErrorAction SilentlyContinue
}
Write-Host 'Proxy-Address conflict is persisting  '($smtpaddress.Split(':')[1])' from '$mailuser.DisplayName -ForegroundColor DarkGray
#Set-MailUser -Identity $mailuser.Alias -EmailAddresses @{remove=($smtpaddress.Split(':')[1])} -ErrorAction stop
Write-Host '------------- COPY following text to your ticket ---------------------'
Write-Host 'Sehr geehrter Kunde!'
Write-host 'Die von Ihnen gewählte Proxy-Adresse '($smtpaddress.Split(':')[1])' ist bereits vergeben.'
write-host 'Bitte bereinigen Sie die Proxy-Adresse für den User '$proxyauser', oder löschen Sie das AD-Objekt.'
write-host 'Wir erwarten auf baldige Rückmeldung wenn die Proxy-Adresse '($smtpaddress.Split(':')[1])' von Ihnen bereinigt wurde.'
write-host 'Wir wünschen Ihnen einen wunderschönen Tag!'
}

if($mailusererror -like '*is already mail-enabled*' -or $mailusererror -like '*"SamAccountName" is used by another recipient object*')
{
Write-Host 'Distributiongroup is already enabled for '$rrwrite -ForegroundColor Green
}

if($mailusererror -like '*already exists*')
{
#Write-Host $mailusererror -BackgroundColor DarkYellow
Write-Host 'Distributiongroup is already persisting for '$rrwrite -ForegroundColor Green
}

#Write-Host $mailusererror

     }
    }
    try
    {
         Write-Host 'Trying to enable Moderator delegation group '$eqbook -ForegroundColor DarkGray
     New-DistributionGroup -Name ('m365_eq_EXO_'+$identity+'_book') -Type Security -PrimarySmtpAddress $eqbook -ErrorAction stop
     Write-Host 'Waiting for 20 seconds to apply the next step' -BackgroundColor DarkGray
     Start-Sleep -Seconds 20
     Set-DistributionGroup -Identity ('m365_eq_EXO_'+$identity+'_book') -HiddenFromAddressListsEnabled:$true -HiddenGroupMembershipEnabled:$true
     Write-Host 'Successfull enabled '$eqbook -ForegroundColor green

    }
    catch
    {
     try
     {
     Write-Host ('m365_eq_EXO_'+$identity+'_book')' does not exists' -ForegroundColor DarkGray
     Write-Host 'Trying to create fullaccess delegation group '$eqbook -ForegroundColor DarkGray
     Write-Host 'Successfull created '$eqbook -ForegroundColor green
     }
     catch
     {

$mailusererror = $_.Exception.Message
if($mailusererror -like '*The proxy address*')
{
Write-Host $mailusererror -BackgroundColor DarkYellow
Write-Host 'Trying to resolve the proxy address conflict' -ForegroundColor DarkGray
$smtpaddress = $mailusererror.Split('"')[1]
$proxyauser = $mailusererror.Split('"')[3]
Try
{
$mailuser = Get-MailUser $smtpaddress.Split(':')[1] -ErrorAction stop

}
catch
{
$mailuser = Get-DistributionGroup $smtpaddress.Split(':')[1] -ErrorAction SilentlyContinue
}
Write-Host 'Proxy-Address conflict is persisting  '($smtpaddress.Split(':')[1])' from '$mailuser.DisplayName -ForegroundColor DarkGray
#Set-MailUser -Identity $mailuser.Alias -EmailAddresses @{remove=($smtpaddress.Split(':')[1])} -ErrorAction stop
Write-Host '------------- COPY following text to your ticket ---------------------'
Write-Host 'Sehr geehrter Kunde!'
Write-host 'Die von Ihnen gewählte Proxy-Adresse '($smtpaddress.Split(':')[1])' ist bereits vergeben.'
write-host 'Bitte bereinigen Sie die Proxy-Adresse für den User '$proxyauser', oder löschen Sie das AD-Objekt.'
write-host 'Wir erwarten auf baldige Rückmeldung wenn die Proxy-Adresse '($smtpaddress.Split(':')[1])' von Ihnen bereinigt wurde.'
write-host 'Wir wünschen Ihnen einen wunderschönen Tag!'
}
if($mailusererror -like '*is already mail-enabled*' -or $mailusererror -like '*"SamAccountName" is used by another recipient object*')
{
Write-Host 'Distributiongroup is already enabled for '$eqwrite -ForegroundColor Green
}
if($mailusererror -like '*already exists*')
{
#Write-Host $mailusererror -BackgroundColor DarkYellow
Write-Host 'Distributiongroup is already persisting for '$eqbook -ForegroundColor Green
}

#Write-Host $mailusererror

     }
    }
    }
    catch
    {

             $mailusererror = $_.Exception.Message
            Write-Host $mailusererror -BackgroundColor DarkYellow
    }
Write-Host $mailusererror -BackgroundColor DarkYellow
}

}

function Complete-MailboxEnablement
{
param($identity)

if($identity -eq $null)
{
do
{
Write-Host 'Please enter Identity/SamAccountName of the Mailobject which should be completed : ' -ForegroundColor Cyan
$identity = Read-Host
}while($identity -eq '')
}

$mailboxinfo = Get-Recipient $identity

switch ($mailboxinfo.RecipientType)
{
'UserMailbox' { Write-Host 'Setting up default language and region to de-at' -ForegroundColor DarkGray
Set-MailboxRegionalConfiguration -identity $identity -Language 'de-AT' -DateFormat "dd.MM.yyyy" -TimeFormat 'HH:mm' -LocalizeDefaultFolderName
Write-Host 'Language configuration completed' -ForegroundColor Green
do
{
Write-Host 'Please enter duration of litigation-hold (default 365) : ' -ForegroundColor Cyan
$duration = Read-Host
}while($duration -eq '')
try
{
Set-Mailbox -Identity $identity -LitigationHoldEnabled:$true -LitigationHoldDuration $duration -ErrorAction stop
Write-Host 'LitigationHold configured successful' -ForegroundColor Green
}
catch
{
$mailusererror = $_.Exception.Message
if($mailusererror -like '*Online license does*')
{
Write-Host 'The mailbox has no Exchange Plan2 license, please make sure that right license is assigned!!' -BackgroundColor Magenta
}
else
{
Write-Host $mailusererror -BackgroundColor DarkYellow
}
}
}

'MailUser' {Write-Host 'nothing to do' -ForegroundColor Green}
'MailContact' {Write-Host 'nothing to do' -ForegroundColor Green}
'MailUniversalDistributionGroup' {Write-Host 'nothing to do' -ForegroundColor Green}
'MailUniversalSecurityGroup' {Write-Host 'nothing to do' -ForegroundColor Green}

}


#Write-Host $mailusererror

}

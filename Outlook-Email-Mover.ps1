<#
Outlook Email Mover
.Author Daniel Keer
.Version 1.0.0
.author_uri https://theDXT.ca
.script_uri https://github.com/thedxt/Outlook-email-mover

.SYNOPSIS
Moves emails from a source folder to a destination folder in an Outlook mailbox.

.DESCRIPTION
The Outlook-email-mover function allows you to move emails from a source folder to a destination folder in an Outlook mailbox.
You can specify the user, source folder, destination root folder, destination sub-folder (optional), sender, and subject to filter the emails to be moved.
The function also supports a dry run mode to preview the emails that would be moved without actually moving them.

.PARAMETER user
This will be your UPN. This is required.

.PARAMETER source_folder
The name of the source folder from which you want to move emails. Typically, this is just Inbox. This is required.

.PARAMETER dest_root_folder
The name of the destination folder to which you want the emails moved. This is required.

.PARAMETER dest_sub_folder
The name of the destination sub-folder if you want the emails moved to a sub-folder. This parameter is optional, and if used, it must be a sub-folder in the root folder.

.PARAMETER sender
The email address that sent the email. This is used to help make sure only the emails you intend to move are moved. 

.PARAMETER subject
The subject line of the emails you want to target.

.PARAMETER dryrun
This is an optional setting and will output the emails that would be moved.

.EXAMPLE
Outlook-email-mover -user "name@email.com" -source_folder "Inbox" -dest_root_folder "Deleted Items" -sender "support@email.com" -subject "Hello" -dryrun

#>

function Outlook-email-mover {
    param(
    [Parameter (Mandatory = $true)] [String]$user,
    [Parameter (Mandatory = $true)] [String]$source_folder,
    [Parameter (Mandatory = $true)] [String]$dest_root_folder,
    [Parameter (Mandatory = $false)] [String]$dest_sub_folder,
    [Parameter (Mandatory = $true)] [String]$sender,
    [Parameter (Mandatory = $true)] [String]$subject,
    [Parameter (Mandatory = $false)] [Switch]$dryrun
    )

    #gather the folder ids for the source and destination folders
    Write-Host "Getting folder IDs"
    $find_Source_folder = Get-MgUserMailFolder -UserId $user -all | Where-Object { $_.DisplayName -eq $source_folder }
    $find_dest_root_folder = Get-MgUserMailFolder -UserId $user -all | Where-Object { $_.DisplayName -eq $dest_root_folder }
    
    # if the sub folder is set, then get the folder id for the sub folder using the root folder id to find it.
    if ($dest_sub_folder){
        $find_dest_sub_folder = Get-MgUserMailFolderChildFolder -UserId $user -MailFolderId $find_dest_root_folder.Id | Where-Object { $_.DisplayName -eq $dest_sub_folder }
    }

#build out the search filters using the inputs
$filter_build_from = "contains(from/emailAddress/address,'" + $sender + "')"
$filter_build_subject = "contains(subject,'" + $subject + "')"

#build out a single varriable with the filters combined with AND.
$combo_filter = $filter_build_from + " AND " + $filter_build_subject

# hunt for the messages that match the filter
write-host "I am Hunting for the Emails"
$find_emails = Get-MgUserMailFolderMessage -UserId $user -MailFolderId $find_Source_folder.Id -All -Filter $combo_filter

#if drryrun is set, then display the emails that would be moved.
if ($dryrun){
    write-host "Dry run is activated I will NOT move any Email(s)."
    Write-host "Here are the" $find_emails.count "Emails(s) that would be moved"
    $find_emails | Select-Object Subject, ReceivedDateTime | sort-object -Property ReceivedDateTime -Descending

}else {
    #if dryrun is not set, then move the emails.
    Write-Host "Dry run is not activated."
    Write-Host "I will move Email(s)"
    Write-Host "Moving" $find_emails.count "Emails(s)"
    #Loop through the messages and move them to the target folder
    foreach ($message in $find_emails) {
        #if the sub folder is set, then move the email to the sub folder.
        if ($dest_sub_folder){
            Write-host "Moving Email" $message.Subject "to" $dest_sub_folder
            Move-MgUserMailFolderMessage -UserId $user -MessageId $message.Id -MailFolderId $find_Source_folder.id -DestinationId $find_dest_sub_folder.id |  Out-Null
        }else{
            #if the sub folder is not set, then move the email to the root folder.
            Write-host "Moving Email" $message.Subject "to" $dest_root_folder
        Move-MgUserMailFolderMessage -UserId $user -MessageId $message.Id -MailFolderId $find_Source_folder.id -DestinationId $find_dest_root_folder.id |  Out-Null
    }
    }
    write-host "All done moving" $find_emails.count "Email(s)"
}

}

<#
Outlook Email Mover Connector
.Author Daniel Keer
.Version 1.0.0
.author_uri https://theDXT.ca
.script_uri https://github.com/thedxt/Outlook-email-mover

.SYNOPSIS
Connects to the Microsoft Graph API using either the Entra App or the CLI.

.DESCRIPTION
The Outlook-email-mover-connector function is used to authenticate with the Microsoft Graph API. It supports two connection methods:

1. Entra App: Connects to the Microsoft Graph API using the Entra App, which requires providing the AppID and TenantID.
2. CLI: Connects to the Microsoft Graph API using the CLI, which does not require any additional parameters.

.PARAMETER connect
Specifies the connection method to use. Valid values are 'EntraApp' and 'CLI'.

.PARAMETER appid
The AppID to use when connecting to the Microsoft Graph API using the Entra App.

.PARAMETER tennatid
The TenantID to use when connecting to the Microsoft Graph API using the Entra App.

.EXAMPLE
Outlook-email-mover-connector -connect EntraApp -appid "12345678-1234-1234-1234-123456789012" -tennatid "12345678-1234-1234-1234-123456789012"

#>
function Outlook-email-mover-connector {
    param(
        [Parameter (Mandatory = $true)] [ValidateSet('EntraApp','CLI')] [String]$connect,
        [Parameter (Mandatory = $false)] [String]$appid,
        [Parameter (Mandatory = $false)] [String]$tennatid
        )
    # Authenticate with Microsoft Graph
    if ($connect -eq "EntraApp"){
        # Connect to Microsoft Graph using the Entra App
        write-host "Connecting to Microsoft Graph using the Entra App"
        Connect-MgGraph -ClientId $appid -TenantId $tennatID -scopes "Mail.ReadWrite" -NoWelcome
    }
    
    if ($connect -eq "CLI"){
        # Connect to Microsoft Graph using CLI
        write-host "Connecting to Microsoft Graph using CLI"
        Connect-MgGraph -scopes "Mail.ReadWrite" -NoWelcome
    }
    
Write-host "You should now be connected to Microsoft Graph"

}

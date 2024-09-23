![image](https://github.com/user-attachments/assets/2bb351c8-d78b-4a5d-b864-708186c15137)

# Outlook Email Mover
A PowerShell script to move emails around in Outlook using Microsoft Graph.

> [!IMPORTANT]
>
> You must be connected to Microsoft Graph to run the script. [Outlook Email Mover Connector](#outlook-email-mover-connector) is a secondary function that can be used to connect to Microsoft Graph.
>
> You will need the Microsoft Graph cmdlets installed.
> 
> `user` must be defined as this is the mailbox the script will use.
> 
> `source_folder` must be defined as this is the location where the script looks for the emails. Typically this is just set to Inbox.
> 
> `dest_root_folder` must be defined as this is location the script will move the emails to.
> 
> `sender` must be defined as this is part of what the script uses to find the emails.
> 
> `subject` must be defined as this is the other part of what the script uses to find the emails.

`dest_sub_folder` is optional and if used it must be a sub-folder in the root folder.

`dryrun` is an optional setting that will output the emails that would be moved.


[More detailed documentation](https://thedxt.ca/)


> [!TIP]
> ### Example
> `Outlook-email-mover -user "my@email.com" -source_folder "Inbox" -dest_root_folder "Deleted Items" -sender "renewals@godaddy.com" -subject "Your GoDaddy Renewal Notice" -dryrun`
>

## Outlook Email Mover Connector
In the Outlook Email Mover a second function `Outlook-email-mover-connector` can be used to connect to Microsoft Graph.

The Outlook Email Mover Connector has two connect modes CLI and Entra App.
- `CLI` connect mode will call the `Connect-MgGraph` command to connect to Microsoft Graph with the required scope.
- `EntraApp` connect will connect to using an Entra App. You must provide the Entra App ID using `appid` and your Tenant ID using `tennatid `. It will connect as your account to the Entra Application with the required scope and will use delegated permissions when you run the script.

> [!TIP]
> ### Example
> `Outlook-email-mover-connector -connect EntraApp -appid "12345678-1234-1234-1234-123456789012" -tennatid "12345678-1234-1234-1234-123456789012"`

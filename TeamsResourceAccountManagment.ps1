<# Header
NAME: Teams Resource Account Management Script
Details: This script is designed to look at a SPO List that will check if a Resource Account needs to be added or removed from Teams.
 AUTHOR: Microsoft 365 Team , Pacific Gas and Electric Co
 DATE  : 05/31/2023
********************************************************************************** 
Do not make changes to this script without checking with SPoadminalerts@pge.com
**********************************************************************************
Versions:
 05/31/2023 1.00 Draft

#>
test 1


[CmdletBinding()]
PARAM (
)

Begin
{
    

    Function LogIt
    {
        [CmdletBinding()]
        PARAM (
            [Parameter(Mandatory = $true)][string]$LogFile,
            [Parameter(Mandatory = $true)][string]$LogLine
        )
        
        Process
        {
            #----------------[ TODO: Main Execution Code Here ]---------------
            $tmpLogLine = "[$((Get-Date).ToString('yyyy-MM-dd HH:mm:ss'))] $($LogLine)"
            Write-Host $tmpLogLine 
            $tmpLogLine | Out-File -FilePath $LogFile -Append
        }
    
    }
}
$Script:MyStartTime = Get-Date
    Write-Verbose "Start time = $($Script:MyStartTime)"
    
    #----------------[ TODO: Begin Code Here ]---------------

    Clear-Host

    Write-Host ""
    Write-Host ""
    Write-Host ""
    Write-Host ""
    Write-Host ""
    Write-Host ""
    Write-Host ""
    Write-Host ""
    
    
    $user = $($env:username).ToLower()
    if ($user -eq "s1e8admin")
    {
        #Running under s1e8admin
        $SPoAdminCredPath = "D:\Scripts\Credentials\s1e8_CredentialSPoSiteMangement.txt"    
    }
    elseif ($user -eq "jxgwadmin")
    {
        #Running under jxgwadmin
        $SPoAdminCredPath = "D:\Scripts\Credentials\jxgw_CredentialSPoSiteMangement.txt"        
    }

    elseif ($user -eq "svc-2262-prd-tasks")
    {
        #Running under pge\SVC-2262-Prd-Tasks
        $SPoAdminCredPath = "D:\Scripts\Credentials\svc-2262-prd-taskscloud_securepw.txt"     
    } 
    
    else
    {
        Exit
    }
    
    $SPoAdminAccount = "SVC-2262-Prd-Taskscloud@pge.onmicrosoft.com"
    $SPoAdmin = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $SPoAdminAccount, (Get-Content $SPoAdminCredPath | ConvertTo-SecureString) -ErrorVariable CredentialError
    
    # TODO: Import required modules
    #Import-Module SharePointPnPPowerShellOnline
    Import-Module MicrosoftTeams
    $ScriptName = "TeamsResourceAccountManagement"
    $ScriptOwner = "SPoAdminAlerts@pge.com"
    $ScriptStart = Get-Date
    $Test = $true
    $verbose = $true

    # TODO: Connect to Teams and SharePoint Online
    $siteUrl = "https://pge.sharepoint.com/sites/SPGovernance"
    $listName = "Teams Resource Account List"
    Connect-MicrosoftTeams -TenantId 44ae661a-ece6-41aa-bc96-7c2c85a08941 -Credential $SPoAdmin -ErrorAction SilentlyContinue -ErrorVariable TeamError
    Connect-PnPOnline -Url $siteUrl -Credential $SPoAdmin -ErrorAction SilentlyContinue -ErrorVariable TeamError

    
    
    #----------------[ TODO: Main Execution Code Here ]---------------



# Retrieve list items from SharePoint
$listItems = Get-PnPListItem -List $listName

foreach ($item in $listItems) {
    # Extract information from SharePoint List item
    $Created = $item["Created By"]
    $displayName = $item["Display Name"] 
    $emailAddress = $item["email address"] 
    $Accounttype = $item["Resource Account Type"] 
    $Account = $item["Add Account or Remove"] 

    

# Teams Resource Account Announcement Email Info
$TeamResourceAccountAnnouncementBody = [IO.File]::ReadAllText('D:\Scripts\Data\TeamResourceAccountAnnouncement.txt')
$TeamResourceAccountAnnouncementSubject = "INFO: Your Request to Add\Remove a Teams Resource Account has Been Processed [Resource Account Name]"
$TestResourceAccountAnnouncementSubject = "[TEST] INFO: Your Request to Add\Remove a Teams Resource Account has Been Processed [Resource Account Name]"

# Send email to Requestor
$From = "SharepointOnlineGovernance@PGE.onmicrosoft.com"
$EmailBody = $TeamResourceAccountAnnouncementBody.replace("[Resource Account Name]", $displayName)
$Subject = $TeamResourceAccountAnnouncementSubject.Replace("[Resource Account Name]", $displayName)
$SubjectTest = $TestResourceAccountAnnouncementSubject.replace("[Resource Account Name])", $displayName)
$sendTo = @()
$sendTo = $Created.user
$SendToTest = "s1e8@pge.com"
Send-MailMessage -to $sendTo -Subject $Subject -Body $emailbody -BodyAsHtml -SmtpServer "mailhost.utility.pge.com" -bcc $scriptowner -from $from 
if ($test -eq $true) {
    Send-MailMessage -to $SendToTest -Subject $SubjectTest -Body $emailbody -BodyAsHtml -SmtpServer "mailhost.utility.pge.com" -from $from 
}

  

# Perform action based on the specified value in the 'Account' column
    switch ($Account) {
        "Add Account" {
            # Check if the Resource Account already exists
            $existingAccount = Get-CsOnlineApplicationInstance -Filter { UserPrincipalName -eq $emailAddress }
            if ($existingAccount) {
                Write-Host "Resource Account for $emailAddress already exists. Skipping creation."
                continue
            }

            

     # Create a New Resource Account in Teams or Remove
    $resourceAccount = New-CsOnlineApplicationInstance -UserPrincipalName { UserPrincipalName -eq $emailAddress }
    Write-Host "Resource Account for $emailAddress has been added."
    Send-MailMessage -to $Created
    if ($test)
            {
                Write-log "test | $displayName | skip adding $($emailAddress)"
            }
                    }
        "Remove Account" {
            # Check if the Resource Account exists
            $existingAccount = Get-CsOnlineApplicationInstance -Filter { UserPrincipalName -eq $emailAddress }
            if ($existingAccount) {
                # Remove the Resource Account from Teams
                Remove-CsOnlineApplicationInstance -Identity $existingAccount.ObjectId
                Write-Host "Resource Account for $emailAddress has been removed."
            } else {
                Write-Host "Resource Account for $emailAddress does not exist. Skipping removal."
            }
        }
        default {
            Write-Host "No valid action specified for $emailAddress. Skipping."
        }
    }
}  
    

    # Add the Resource Account to a call queue or auto attendant based on the account type
    if ($accountType -eq "Call Queue") {
        # Add to call queue
        Add-CsOnlineApplicationInstanceToCallQueue -Identity $resourceAccount.ObjectId -CallQueueId $serviceId
    } elseif ($accountType -eq "Auto Attendant") {
        # Add to auto attendant
        Add-CsOnlineApplicationInstanceToAutoAttendant -Identity $resourceAccount.ObjectId -AutoAttendantId $serviceId
    } else {
        Write-Host "Invalid account type for $emailAddress. Skipping addition to service."
    }
{

    # Create Log File

    $exportCSV = "D:\Scripts\Logs\TeamsResourceAccountManagement\ProcessData_$((Get-Date -f yyyy-MM-dd-HHmmss).ToString()).csv"
   

    {
   $obj = New-Object -TypeName PSObject
        $obj | -Name "Display Name" -Value $($displayName.Displayname)
        $obj | -Name "email address" -Value $($emailAddress.emailAddress)
        $obj | -Name "Resource Account Type" -Value $($ccounttype.Accounttype)
        $obj | -Name "Add or Remove Account" -Value $($Account.Account)

        $obj | Export-Csv -Path $exportCSV -NoTypeInformation -Append
    }

     #----------------[ TODO: End Code Here ]---------------



    
    #----------------[ Cleanup ]---------------
    $Script:MyEndTime = Get-Date
    Write-Verbose "End time = $($Script:MyEndTime)"
    $Script:MyDuration = New-TimeSpan -Start $Script:MyStartTime -End $Script:MyEndTime
    Write-Verbose "Script run duration: $($Script:MyDuration.Days) days + $($Script:MyDuration.Hours) hours + $($Script:MyDuration.Minutes) minutes + $($Script:MyDuration.Seconds) seconds + $($Script:MyDuration.Milliseconds) milliseconds"
}

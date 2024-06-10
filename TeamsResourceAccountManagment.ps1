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
 06/10/2024 1.1 Update Draft

#>



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
    $Script:MyStartTime = Get-Date
    Write-Verbose "Start time = $($Script:MyStartTime)"

    # Create csv File

    $exportCSV = "D:\Scripts\Logs\TeamsResourceAccountManagement\ProcessData_$((Get-Date -f yyyy-MM-dd-HHmmss).ToString()).csv"
}

Process
{
    
    
    #----------------[ TODO: Begin Code Here ]---------------

    Clear-Host
    
    
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
    
    try
    {

        Connect-MicrosoftTeams -TenantId 44ae661a-ece6-41aa-bc96-7c2c85a08941 -Credential $SPoAdmin -ErrorAction Stop
        Connect-PnPOnline -Url $siteUrl -Credential $SPoAdmin -ErrorAction Stop
    }
    catch
    {
        $err = $_
        # Log error
        Exit
    }

    
    
    #----------------[ TODO: Main Execution Code Here ]---------------



    # Retrieve list items from SharePoint
    $listItems = Get-PnPListItem -List $listName

    foreach ($item in $listItems)
    {
        # Extract information from SharePoint List item
        $Created = $item["Author"]
        $displayName = $item["Title"] 
        $emailAddress = $item["emailaddress"] 
        $Accounttype = $item["ResourceAccountType"] 
        $Account = $item["AddAccountorRemove"] 

        $obj = New-Object -TypeName PSObject
        $obj | -Name "DisplayName" -Value $($displayName)
        $obj | -Name "emailaddress" -Value $($emailAddress)
        $obj | -Name "ResourceAccountType" -Value $($Accounttype)
        $obj | -Name "AddorRemoveAccount" -Value $($Account)
        $obj | -Name "Action1" -Value $($Null)
        $obj | -Name "Action2" -Value $($Null)


       
        # Teams Resource Account Announcement Email Info
        #$TeamResourceAccountAnnouncementBody = [IO.File]::ReadAllText('D:\Scripts\Data\TeamResourceAccountAnnouncement.txt')
        $TeamResourceAccountAnnouncementBody = Get-Content -Path 'D:\Scripts\Data\TeamResourceAccountAnnouncement.txt' -Raw
        $TeamResourceAccountAnnouncementSubject = "INFO: Your Request to Add\Remove a Teams Resource Account has Been Processed [Resource Account Name]"
        $TestResourceAccountAnnouncementSubject = "[TEST] INFO: Your Request to Add\Remove a Teams Resource Account has Been Processed [Resource Account Name]"

        # Load Emails
        $From = "SharepointOnlineGovernance@PGE.onmicrosoft.com"
        $EmailBody = $TeamResourceAccountAnnouncementBody.replace("[Resource Account Name]", $displayName)
        $Subject = $TeamResourceAccountAnnouncementSubject.Replace("[Resource Account Name]", $displayName)
        $SubjectTest = $TestResourceAccountAnnouncementSubject.replace("[Resource Account Name])", $displayName)
        $sendTo = @()
        $sendTo = $Created.email
        $SendToTest = "s1e8@pge.com"

        #if($script:Test){
        #$sendTo = "s1e8@pge.com"
        #$Subject = "[Test] $Subject"
        #Clear-Variable CC, BCC
 
        # Perform action based on the specified value in the 'Account' column
        switch ($Account)
        {
            "Add Account"
            {
                # Check if the Resource Account already exists
                $existingAccount = Get-CsOnlineApplicationInstance -Filter { UserPrincipalName -eq $emailAddress }
                if ($existingAccount)
                {
                    Write-log "Resource Account for $emailAddress already exists. Skipping creation."
                    $obj.Action1 = "Account alread exists"
                    continue
                }            

                # Create a New Resource Account in Teams 
              
                if ($test -eq $true)
                {
                    Send-MailMessage -to $SendToTest -Subject $SubjectTest -Body $emailbody -BodyAsHtml -SmtpServer "mailhost.utility.pge.com" -from $from 
                    $obj.Action1 = "Skip adding: Sent email"
                    Write-log "test | $displayName | skip adding $($emailAddress)"
                }
                else 
                {
                    $resourceAccount = New-CsOnlineApplicationInstance -UserPrincipalName { UserPrincipalName -eq $emailAddress }
                    Write-Host "Resource Account for $emailAddress has been added."
                    Send-MailMessage -to $sendTo -Subject $Subject -Body $emailbody -BodyAsHtml -SmtpServer "mailhost.utility.pge.com" -bcc $scriptowner -from $from 
                    $obj.Action1 = "Add new account"
                    # Add the Resource Account to a call queue or auto attendant based on the account type
                    if ($accountType -eq "Call Queue")
                    {
                        # Add to call queue
                        Add-CsOnlineApplicationInstanceToCallQueue -Identity $resourceAccount.ObjectId -CallQueueId $serviceId
                        $obj.Action2 = "Add call queue"
                    }
                    elseif ($accountType -eq "Auto Attendant")
                    {
                        # Add to auto attendant
                        Add-CsOnlineApplicationInstanceToAutoAttendant -Identity $resourceAccount.ObjectId -AutoAttendantId $serviceId
                        $obj.Action2 = "Add Auto Attendant"
                    }
                    else
                    {
                        Write-log "Invalid account type for $emailAddress. Skipping addition to service."
                        $obj.Action2 = "ERROR"
                    }
                }
            }

            # Remove Resource Account in Teams 
            "Remove Account"
            {
                # Check if the Resource Account exists
                $existingAccount = Get-CsOnlineApplicationInstance -Filter { UserPrincipalName -eq $emailAddress }
                if ($existingAccount)
                {
                    if ($test -eq $true)
                    {
                        Send-MailMessage -to $SendToTest -Subject $SubjectTest -Body $emailbody -BodyAsHtml -SmtpServer "mailhost.utility.pge.com" -from $from 
                                
                        Write-log "test | $displayName | skip adding $($emailAddress)"
                    } 
                    else
                    {
                   
                    

                        # Remove the Resource Account from Teams
                        Remove-CsOnlineApplicationInstance -Identity $existingAccount.ObjectId
                        Write-log "Resource Account for $emailAddress has been removed."
                    }
                }
                else
                {
                    Write-log "Resource Account for $emailAddress does not exist. Skipping removal."
                }
            }
            default
            {
                Write-log "No valid action specified for $emailAddress. Skipping."
            }
        }

        $obj | Export-Csv -Path $exportCSV -NoTypeInformation -Append
    }      

 

}

End
{

 
    

    #----------------[ TODO: End Code Here ]---------------

    # Disconenct from Teams

    # Disconnect from SharePoint PnP



    
    #----------------[ Cleanup ]---------------
    $Script:MyEndTime = Get-Date
    Write-Verbose "End time = $($Script:MyEndTime)"
    $Script:MyDuration = New-TimeSpan -Start $Script:MyStartTime -End $Script:MyEndTime
    Write-Verbose "Script run duration: $($Script:MyDuration.Days) days + $($Script:MyDuration.Hours) hours + $($Script:MyDuration.Minutes) minutes + $($Script:MyDuration.Seconds) seconds + $($Script:MyDuration.Milliseconds) milliseconds"
}
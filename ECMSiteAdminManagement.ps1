<#
NAME: ECM Admin Management Script
 AUTHOR: pge\pwr3 , Pacific Gas and Electric Co.
 DATE  : 09/20/2021
********************************************************************************** 
Do not make changes to this script without checking with Patrick Reeves (pge\pwr3)
**********************************************************************************
Versions:
 10/22/2021 1.0  Release
			     Manages adding/removing SCAs for Records365
                 https://pge.sharepoint.com/sites/SPoGov/Lists/ECMSiteManagement/
 02/12/2024 1.1  Replaced pwr3@pge.com with SPoadminalerts@pge.com
 03/26/2024 2.0  Updated to PNP.Powershell 2.x
#>
Function Write-Log{
	param ($Message, $color)
	$ErrMessage = $(Get-Date -f HH:mm:ss).ToString() + " | " + $Message
    if($Message.Contains("Error")){
        $Script:ErrorLog += $Message
    }
	if (!$color) { $color = "white" }
	If ($Script:Test -or $Script:Verbose) { Write-Host $Message -ForegroundColor $color }
	Write-Output $ErrMessage | Out-File $Script:LogFile -Append;
}
Function Get-ProperName{
	param ($Name)
    if($name){
        if($name.Contains("(")){
            $PrnS = $name.IndexOf("(")
            $Name = $Name.Substring(0,$($name.IndexOf("("))).Trim()
        }
	    $aryName = $Name.split(",")
        if($aryName.count -eq 2){
            Return $aryName[1].Trim() + " " + $aryName[0].Trim()
        }else{$Name}
    }else{$Name}
}
Function Fix-URL{
	param ($Url)
    if($Url -eq "https://pge.sharepoint.com"){
        $Url = "https://pge.sharepoint.com/"
    }elseif($Url -and $Url -ne "https://pge.sharepoint.com/"){
	    $u = $Url.Trim().ToLower()
	    if($u.EndsWith("/")){$Url = $u.Substring(0, $($u.Length - 1))}else{$url = $u}
    }
	$url
}
Function Send-ErrorEmail{
 param($Subject, $Body)
    if(!$Body){$Body = "No details passed"}
	Send-MailMessage -From "SPoadminalerts@pge.com" -To "SPoadminalerts@pge.com" -Subject $Subject -Body $Body -SmtpServer "mailhost.utility.pge.com"
    Write-Log "Error | Email | Sending Error Email to Admims" "Red"
	Write-Log "Error | Email | $Subject" "Yellow"
	Write-Log "Error | Email | $Body" "Yellow"
}
Function Check-User{
    param ($ChkID)
    Set-Variable CheckUser, CorpUser
    $UserReport = @{}    
    If($ChkID.Length -gt 0){
        $ChkID = $ChkID.ToLower()
        if($ChkID.Contains("@pge-corp.com")){
            Write-Log "Info | ID | [Check-User]PGE-Corp ID: $ChkID"
            try{
                $CheckUser = New-PnPUser -LoginName $ChkID -Connection $script:ctxSPoGovSite
                $CorpUser = Check-User $($CheckUser.LoginName).Replace("i:0#.f|membership|","")
                if($CorpUser){
                    $UserReport.Add("Email",$($CheckUser.Email))
                    $UserReport.Add("Name",$($CheckUser.Title))
                    $UserReport.Add("Title",$($CheckUser.Title))
                }else{
                    Write-Log "Warning | ID | [Check-User]Bad Corp ID: $ChkID"
                }
            }catch{
                Write-Log "Warning | ID | [Check-User]Disabled Corp ID: $ChkID"
            }
        }else{
            $ChkID = $ChkID.Replace("@pge.com","")
            if($ChkID.Length -eq 4){
                Try{
                    $CheckUser = Get-ADUser $ChkID -Properties * -ErrorAction SilentlyContinue
                    if(!$CheckUser.Enabled){
                        Clear-Variable CheckUser
                        Write-Log "Warning | ID | [Check-User]Disabled: $ChkID"
                    }else{
                        $UserReport.Add("Email",$CheckUser.UserPrincipalName)
                        $UN = $CheckUser.SurName + ", " + $CheckUser.GivenName + " [pge\" + $CheckUser.Name + "]"
                        $UserReport.Add("Name",$UN)
                        $UserReport.Add("Title",$($CheckUser.Title))
                    }
                }catch{
                    Write-Log "Warning | ID | [Check-User]Bad: $ChkID"
                }
            }else{
                Write-Log "Warning | ID | [Check-User]Skipping Admin/System account: $ChkID"
            }
        }
    }
    $UserReport
}
Function Get-Manager{
    param ($LanID)
    $CorpID = $LanID.replace("@pge.com","")
    $CorpID = $CorpID.replace("pge\","")
    $CorpID = $CorpID.replace("admin","")
    try{
        $mgr = get-aduser $CorpID -properties * | select manager
        $MgrID = ($mgr.manager).split("=")[1].split(",")[0]
        $MgrTitle = $(get-aduser $MgrID -properties * | select Title).Title
        If( ($MgrTitle.Contains("VP")) -or ($MgrTitle.Contains("CEO")) -or ($MgrTitle.Contains("CEO")) -or ($MgrTitle.Contains("Officer"))){
            $Result = $LanID
        }else{$Result = $MgrID}
    }catch{
        $Result = $LanID
    }
    $Result
}
Function Get-VP{
    param ($LanID)
    $CorpID = $LanID.replace("@pge.com","")
    $CorpID = $CorpID.replace("pge\","")
    $CorpID = $CorpID.replace("admin","")
    try{
        $usr = get-aduser $CorpID -Properties *
        $usrTitle = $Usr.Title
        If( ($usrTitle.Contains("VP")) -or ($usrTitle.Contains("CEO")) -or ($usrTitle.Contains("CEO")) -or ($usrTitle.Contains("Officer"))){
            $Result = $LanID
        }else{
            $mgr = get-aduser $CorpID -properties * 
            $MgrID = ($mgr.manager).split("=")[1].split(",")[0]
            $result = Get-VP $mgrID
        }
    }catch{
        $Result = ""
    }
    $Result
}

Function Send-Email{
 Param(
[Parameter(Mandatory)] $Subject,
[Parameter(Mandatory)] $Body,
[Parameter(Mandatory)] $From,
[Parameter(Mandatory=$false)] $To,
[Parameter(Mandatory=$false)] $CC,
[Parameter(Mandatory=$false)] $BCC,
[Parameter(Mandatory=$false)] $Attachments
)
    [Int32]   $maxAttempts       = 5;
    [Int32]   $failureDelay      = 2;
    [Int32]   $numAttempts       = 0;  
    [Boolean] $messageSent       = $false;
    if($script:ItemUrl){
        $LogUrl = $script:ItemUrl
    }elseif($script:SiteUrl){
        $LogUrl = $script:SiteUrl
    }else{
        $LogUrl = "unknown"
    }
    
    if($script:Test){
        $To = $script:ScriptOwner
        $Subject = "[Test] $Subject"
        Clear-Variable CC, BCC
    }
    if($script:SuppressAllEmails){
        $To = $script:ScriptOwner
        $Subject = "[Supress] $Subject"
        Clear-Variable CC, BCC
    }
    if($To.Count -eq 0){
        $To = "SPoadminalerts@pge.com"
        if($script:VPRepress){
            $Subject = "[VPRepressed] $Subject $($script:SPSite.StorageUsageCurrent)"
        }else{
            $Subject = "[NoToList] $Subject $($script:SPSite.StorageUsageCurrent)"
        }
        Clear-Variable CC, BCC
    }
    $T = @()
    Foreach($E in $To){
        if(($e.Length -gt 0) -and !$t.Contains($E)){
            $t += $E
        }
    }
    $To = $T
    $T = @()
    Foreach($E in $CC){
        if(($e.Length -gt 0) -and !$t.Contains($E)){
            $t += $E
        }
    }
    $CC = $T
    while (($numAttempts -le $maxAttempts) -and (!$messageSent)) {
        try {
            if($Attachments){
                if(($cc.count -gt 0) -and ($bcc.count -gt 0)){
                    Send-MailMessage -Body $Body -Subject $Subject -From $From -To $To -Cc $CC -Bcc $BCC -BodyAsHtml -SmtpServer "mailhost.utility.pge.com" -Attachments $Attachments -ErrorAction Stop
                }elseif($cc.count -gt 0){
                    Send-MailMessage -Body $Body -Subject $Subject -From $From -To $To -Cc $CC -BodyAsHtml -SmtpServer "mailhost.utility.pge.com" -Attachments $Attachments -ErrorAction Stop
                }elseif($bcc.count -gt 0){
                    Send-MailMessage -Body $Body -Subject $Subject -From $From -To $To -Bcc $BCC -BodyAsHtml -SmtpServer "mailhost.utility.pge.com" -Attachments $Attachments -ErrorAction Stop
                }else{
                    Send-MailMessage -Body $Body -Subject $Subject -From $From -To $To -BodyAsHtml -SmtpServer "mailhost.utility.pge.com" -Attachments $Attachments -ErrorAction Stop
                }
            }else{
                if(($cc.count -gt 0) -and ($bcc.count -gt 0)){
                    Send-MailMessage -Body $Body -Subject $Subject -From $From -To $To -Cc $CC -Bcc $BCC -BodyAsHtml -SmtpServer "mailhost.utility.pge.com" -ErrorAction Stop
                }elseif($cc.count -gt 0){
                    Send-MailMessage -Body $Body -Subject $Subject -From $From -To $To -Cc $CC -BodyAsHtml -SmtpServer "mailhost.utility.pge.com" -ErrorAction Stop
                }elseif($bcc.count -gt 0){
                    Send-MailMessage -Body $Body -Subject $Subject -From $From -To $To -Bcc $BCC -BodyAsHtml -SmtpServer "mailhost.utility.pge.com" -ErrorAction Stop
                }else{
                    Send-MailMessage -Body $Body -Subject $Subject -From $From -To $To -BodyAsHtml -SmtpServer "mailhost.utility.pge.com" -ErrorAction Stop
                }
            }
            $messageSent = $true
            Write-Log "Info | $LogUrl | [Send-Email]Email Successful" Green
        }catch{
            Write-Log "Info | $LogUrl | [Send-Email]Email Retrying" Yellow
            $numAttempts++
            sleep -Seconds $failureDelay
        }
        if($numAttempts -eq $maxAttempts){
            Write-Log "Error | $LogUrl | [Send-Email]Email Failed" Red
            Send-ErrorEmail -Subject "Send email failed: $Subject" -Body "$LogUrl"
        }
    }
}


#Update Credential File
#$CredPath = "D:\Scripts\Credentials\pwr3_CloudCredentialSPoSiteMangement.txt"
#'PasswordHere' |  ConvertTo-SecureString -AsPlainText -Force | ConvertFrom-SecureString | Out-File $CredPath

#Import-Module D:\Scripts\Test\Pnp.PowerShell\PnP.PowerShell\1.10.0\PnP.PowerShell.psd1
Import-Module "D:\Scripts\PowerShellModules\PnPPowerShell\pnp.powershell.2.2.117\PnP.PowerShell.psd1"

$ScriptName = "ECMSiteAdminManagement"
$ScriptOwner = "SPoadminalerts@pge.com"
$ScriptStart = Get-Date
$Test = $false
$verbose = $true
$EmailBody = "Script Start $ScriptStart<br> Modes:<br>Test- $Test<br>Verbose- $verbose"
Send-MailMessage -From $ScriptOwner -To $ScriptOwner -Subject "$ScriptName Script start" -Body $EmailBody -BodyAsHtml -SmtpServer "mailhost.utility.pge.com"


$SPOAdminUrl = "https://pge-admin.sharepoint.com"
$SPoGovSite = "https://pge.sharepoint.com/sites/SPGovernance"
$SPoGovListUrl = "Lists/SiteCollectionManagement"
$SPoGovListName = "Site Collection Management"


$ECMListUrl = "Lists/ECMSiteManagement"
$ECMListName = "ECM Site Management"

$FromOveride = "SPoadminalerts@pge.com"
$flgFromOverride = $false
$ERIMEmail = @("Microsoft365ERIMSupport@pge.com", "SPoadminalerts@pge.com")

$user = $($env:username).ToLower()
if($user -eq "pwr3admin"){
    #Running under pwr3admin
    $SPoAdminCredPath = "D:\Scripts\Credentials\pge_CredentialSPoSiteMangement.txt"
}Elseif($user -eq "svc-2262-prd-tasks"){
    #Running under pge\SVC-2262-Prd-Tasks
    $SPoAdminCredPath = "D:\Scripts\Credentials\svc-2262-prd-taskscloud_securepw.txt"
}else{
    Send-MailMessage -Body "script run under wrong user $user" -From "SPoadminalerts@pge.com" -To "SPoadminalerts@pge.com" -Subject "Error: Monthly Stat Update" -SmtpServer  "mailhost.utility.pge.com"
    Exit
}

$SPoAdminAccount = "SVC-2262-Prd-Taskscloud@pge.onmicrosoft.com"
$SPoAdmin = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $SPoAdminAccount, (Get-Content $SPoAdminCredPath | ConvertTo-SecureString) -ErrorVariable CredentialError

if($CredentialError){
    Send-MailMessage -Body $CredentialError[0] -From "SPoadminalerts@pge.com" -To "SPoadminalerts@pge.com" -Subject "Error: Monthly Stat Update" -SmtpServer  "mailhost.utility.pge.com"
    Exit
}


$ECMAddAdminBody = [IO.File]::ReadAllText('D:\Scripts\Data\ECMAddAdminBody.txt')
$ECMRemoveAdminBody = [IO.File]::ReadAllText('D:\Scripts\Data\ECMRemoveAdminBody.txt')
$LargeListQuery = "<View Scope='RecursiveAll'><RowLimit>5000</RowLimit></View>" 

#Clear Log
$LogFile = "D:\Scripts\Logs\$ScriptName\"
$RunDate = Get-Date
$DT = $RunDate.AddDays(-15);
Get-ChildItem -Path $LogFile -Force -File | Where-Object { $_.CreationTime -lt $DT } | Remove-Item -Force;
#New Log file
$LogFile = $LogFile + $Scriptname + "_" + $(Get-Date -f yyyy-MM-dd-HHmmss).ToString() + ".txt";
Write-Output "Time | Status | Site | Message" | Out-File $LogFile -Append;
Write-Log "Info | pge.sharepoint.com | Start ECMSiteAdminManagement"


$ctxSPoGovSite = Connect-PnPOnline -Url $SPoGovSite -Credentials $SPoAdmin -ReturnConnection
$SPGovList = Get-PnPList -Identity $SPoGovListUrl -Connection $ctxSPoGovSite
$GovListID = $SPGovList.ID.Guid
$SPGovItems = Get-PnPListItem -List $GovListID -PageSize 5000 -Connection $ctxSPoGovSite
$ECMList = Get-PnPList -Identity $ECMListUrl -Connection $ctxSPoGovSite
$ECMListID = $ECMList.ID.Guid
$SPItems = Get-PnPListItem -List $ECMListID -PageSize 5000 -Connection $ctxSPoGovSite
#$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid?SerializationLevel=Full -Credential $SPoAdmin -Authentication Basic -AllowRedirection
#Import-PSSession $Session -AllowClobber
Connect-ExchangeOnline -Credential $SPoAdmin -ShowBanner:$False
$ctxSPoAdmin = Connect-PnPOnline -Url $SPOAdminUrl -Credentials $SPoAdmin -ReturnConnection
$ctxStart = Get-Date
$SPSites = Get-PnPTenantSite -Detailed -Connection $ctxSPoAdmin

Write-Log "Info | pge.sharepoint.com | Create Site quick lookup"
$AllSites = "***************"
$i = 0
foreach($SPSite in $SPSites){
    $Add = "[" + $I.ToString().PadLeft(7,"0") + "]" + $(fix-url $SPSite.Url)
    $AllSites += $Add
    $i++
}
$AllSites += "[]"

Write-Log "Info | pge.sharepoint.com | Create SPoGov list quick lookup"
$AllSPoGovList = "***************"
$i = 0
foreach($SPItem in $SPGovItems){
    $Add = "[" + $I.ToString().PadLeft(7,"0") + "]" + $(fix-url $SPItem["Title"])
    $AllSPoGovList += $Add
    $i++
}
$AllSPoGovList += "[]"

$iItem = -1
Write-Log "Info | pge.sharepoint.com | Begin ECM Admin Add/Remove Process"

Set-Variable ItemUrl, PSCA, SSCA, SPSite, Find, AdminActionError, ECMAdmin, Email, OwnerList,o365Group,ctx
$ErrorLog = @()

$SPItem = $SPItems[0]

$Pending = 0
$aryPending = @()
$Added = 0
$aryAdded = @()
$Active = 0
$aryActive = @()
$Removed = 0
$AryRemoved = @()
$CCList = @("Microsoft365ERIMSupport@pge.com","SPoadminalerts@pge.com")
$From = "SPoadminalerts@pge.com"

ForEach ($SPItem in $SPItems){
	Clear-Variable ItemUrl, PSCA, SSCA, SPSite, Find, ECMAdmin, Email, OwnerList, o365Group, ctx
    $iItem++
    $SPItemID = $SPItem.ID
    $ItemUrl = fix-url $SPItem["Title"]
    Write-Log "Info | $ItemUrl | Processing Item $iItem/$SPItemID"
    $UpdateRecord = @{}
    $ToList = @()
    $CcList = @()
    $OwnerList = @()      
    try{
        if($ItemUrl -eq "https://pge.sharepoint.com/"){
            $process = $true
        }elseif(($ItemUrl.Contains("https://pge.sharepoint.com/sites/"))){
            $process = $true
        }else{
            Write-Log "Error | $ItemUrl | Bad Url - Item $SPItemID"
            $process = $false
        }
    }catch{$process = $false}
    if(($SPItem["Status"] -eq "Done") -or ($SPItem["Status"] -eq "Deleted")){
        Write-Log "Info | $ItemUrl | Item marked as Done or Deleted"
        $Process = $false
    }




    if($process){
        try{
            $User = Check-User $SPitem["ECMAdmin"].Email
            $ECMAdmin = $User.EMail
            $ECMName = $($User.Name)
            if($ECMAdmin){
                $Find = $true
                Write-log "Info | $ItemUrl | validated ECMAdmin $ECMAdmin"
            }else{
                $Find = $false
                Write-log "Error | $ItemUrl | Could not validate ECMAdmin $($SPitem["ECMAdmin"].Email)"
            }
        }catch{
            $find = $false
            Write-log "Error | $ItemUrl | Could not validate ECMAdmin $($SPitem["ECMAdmin"].Email)"
        }
	    $StartDate = $SPItem["StartDate"]
        $EndDate = $SPItem["EndDate"]
        $Status = $SPItem["Status"]
	    $wfStatus = $SPItem["wfStatus"]
#CheckIfSiteExists
        $SearchUrl = "]" + $ItemUrl + "["
        $Index = $AllSites.IndexOf($SearchUrl)
#Get Site Collection info
        if($Index -gt -1){
            $strID = $AllSites.Substring($index-7,7)
            Try{
                [int]$ID = [convert]::ToInt32($strID, 10)
                $SPSite = $SPSites[$ID]
            }Catch{
                Write-log "Error | $ItemUrl | Quick lookup failed in SPSites at index $strID"
                $UpdateItem.Add("Status","Deleted")
                $Find = $false
            }
            
        }else{
            Write-log "Error | $ItemUrl | Quick lookup failed in SPSites at index $strID"
            $Find = $false
        }
#Get SPGov list info
        $Index = $AllSPoGovList.IndexOf($SearchUrl)
        if($Index -gt -1){
            $strID = $AllSPoGovList.Substring($index-7,7)
            Try{
                [int]$ID = [convert]::ToInt32($strID, 10)
                $SPGovItem = $SPGovItems[$ID]
            }Catch{
                Write-log "Error | $ItemUrl | Quick lookup failed in SPGovItems at index $strID"
                $Find = $false
            }
            
        }else{
            Write-log "Error | $ItemUrl | Quick lookup failed SPGovItems at index $strID"
            $Find = $false
        }        

#########################################################
	    if($Find){
            $GroupID = $SPSite.GroupID.ToString()
            if($GroupID -ne "00000000-0000-0000-0000-000000000000"){
                Write-log "Info | $ItemUrl | Office 365 Group site" Cyan
                $i = 0
                $StillLooking = $true
                do{
                    Try{
                        $o365Group = Get-UnifiedGroup -Identity $GroupID -ErrorAction Stop
                        Write-Log "Info | $SiteUrl | o365 group found"
                    }catch{
                        Write-Log "Warning | $SiteUrl | Refreshing Exchange connection"
                        try{
                            Connect-ExchangeOnline -Credential $SPoAdmin -ShowBanner:$False
                        }catch{
                            End-ConnectionFail "MS Exchange" $($SPoAdmin.UserName)
                            Exit
                        }
                        Write-Log "Warning | $SiteUrl | Sleeping for 5 seconds"
                        sleep -Seconds 5
                        $o365Group = Get-UnifiedGroup -Identity $GroupID -ErrorAction SilentlyContinue
                    }
                    $SMTP = $o365Group.PrimarySmtpAddress
                    if($SMTP){   
                        $StillLooking = $false
                    }else{
                        $i++
                    }
                }while($StillLooking -and ($i -lt 6))
                if($StillLooking){
                    Write-Log "Error | $SiteUrl | Unable to find o365 group"
                }
                
            }else{
                $ctx = Connect-PnPOnline -Url $ItemUrl -Credentials $SPoAdmin -ReturnConnection
            }

            Write-Log "Info | $ItemUrl | Site Exists" Green
            if(($StartDate -lt $RunDate) -and ($Status -eq "New")){
                if($o365Group){
                    try{
                        Add-UnifiedGroupLinks -Identity $SMTP -LinkType "Members" -Links $ECMAdmin
                        Add-UnifiedGroupLinks -Identity $SMTP -LinkType "Owners" -Links $ECMAdmin
                        $UpdateRecord.Add("Status","Active")
                        $UpdateRecord.Add("wfStatus","Active")
                        Write-Log "Success | $ItemUrl | Added $ECMAdmin as group owner"
                        $Email = "Added"
                    }catch{
                        Write-Log "Error | $ItemUrl | Unable to add $ECMAdmin as group owner"
                    }
                }else{
                    Clear-Variable AdminActionError
                    Set-PnPTenantSite -Url $ItemUrl -Owners $ECMAdmin -ErrorAction SilentlyContinue -ErrorVariable $AdminActionError -Connection $ctxSPoAdmin
                    if($AdminActionError){
                        Write-Log "Error | $ItemUrl | Unable to add $ECMAdmin as site Admin"
                    }else{
                        $UpdateRecord.Add("Status","Active")
                        $UpdateRecord.Add("wfStatus","Active")
                        Write-Log "Success | $ItemUrl | Added $ECMAdmin as site Admin"
                        $Email = "Added"
                    }
                }
            }elseif(($EndDate -lt $RunDate) -and ($Status -eq "Active")){
                if($o365Group){
                    try{
                        Remove-UnifiedGroupLinks -Identity $SMTP -LinkType "Owners" -Links $ECMAdmin -Confirm:$false
                        Remove-UnifiedGroupLinks -Identity $SMTP -LinkType "Members" -Links $ECMAdmin -Confirm:$false
                        $UpdateRecord.Add("Status","Done")
                        $UpdateRecord.Add("wfStatus","Done")
                        Write-Log "Success | $ItemUrl | Removed $ECMAdmin as group owner"
                        $Email = "Removed"
                    }catch{
                        Write-Log "Error | $ItemUrl | Unable to remove $ECMAdmin as group owner"
                    }
                }else{
                    Try{
                        $Remove = $ECMAdmin
                        Remove-PnPSiteCollectionAdmin -Owners $ECMAdmin -Connection $ctx
                        $UpdateRecord.Add("Status","Done")
                        $UpdateRecord.Add("wfStatus","Done")
                        Write-Log "Success | $ItemUrl | Removed $ECMAdmin as site Admin"
                        $Email = "Removed"
                    }catch{
                        Write-Log "Error | $ItemUrl | Unable to remove $ECMAdmin as site Admin"
                    }
                }

            }elseif($Status -eq "Active"){
                $Active++
                $aryActive += "Active - $($SPSite.Url) - Ending: $($EndDate.ToString('MM/dd/yyyy'))"
            }elseif($Status -eq "New"){
                $Pending++
                $aryPending += "Pending - $($SPSite.Url) - Start: $($StartDate.ToString('MM/dd/yyyy'))"
            }
            If($Email){
                $OwnerWL = @()
                $ToList = @()
                $IAcount = 0
                $OwnerWL += $SPGovItem["SiteAdmins"]
                if(!$OwnerWL){
                    $OwnerWL = @()
                    $OwnerWL += $((Get-PnPListItem -List $SPoGovListID -Id $SPGovItem.ID  -Fields "SiteAdmins" -Connection $ctxSPoGovSite).FieldValues).SiteAdmins
                }
                if($OwnerWL.Count -gt 0){
	    	        foreach($Owner in $OwnerWL){
                        Clear-Variable User
                        #Check if user email
                        if($Owner){
                            if(((($($Owner.Email).Length -eq 12)) -and ($Owner.Email.Contains("@pge.com"))) -or ($($Owner.Email).Contains("@pge-corp.com"))){
                                $IAcount++
                            }
                            $User = Check-User $Owner.Email
                            if($User.count -gt 0){
                                $ToList += $User.Email
                            }
                        }
	                }
                }
                if($Email -eq "Added"){
                    $Added++
                    $aryAdded += "Added - $($SPSite.Url) - Ending: $($EndDate.ToString('MM/dd/yyyy'))"
                    $Subject = "ECMAdmin added to " + $SPSite.Url
                    $EmailBody = $ECMAddAdminBody.Replace("[SiteUrl]", $SPSite.Url)
                    $EmailBody = $EmailBody.Replace("[EndDate]", $($EndDate.ToString('MM/dd/yyyy')))

                }elseif($Email -eq "Removed"){
                    $Removed++
                    $aryAdded += "Removed - $($SPSite.Url)"
                    $Subject = "ECMAdmin removed from " + $SPSite.Url
                    $EmailBody = $ECMRemoveAdminBody.Replace("[SiteUrl]", $SPSite.Url)
                }
                $EmailBody = $EmailBody.Replace("[ECMAdmin]", $ECMName)
                $CCList = $ECMAdmin
                Send-Email -Subject $Subject -Body $EmailBody -To $ToList -CC $CCList -From "SPoadminalerts@pge.com"
            }
        }else{
            Write-Log "Error | $ItemUrl | Unable to find site"
        }
    }
    if($UpdateRecord.Count -gt 0){           
	    $xxx = Set-PnPListItem -List $ECMListID -Identity $SPItemID -Values $UpdateRecord -UpdateType SystemUpdate -Connection $ctxSPoGovSite
        Write-Log "Info | $ItemUrl | List Item updated $($xxx.ID)"
    }
}

$EmailBody = "ECM Admin Management Report<br><br>"
$ScriptEnd = Get-Date
$EmailBody += "Script Start: $ScriptStart<br>"
$EmailBody += "Script End: $ScriptEnd<br>"
$RT = $ScriptEnd - $ScriptStart
$Runtime = [Math]::Round((($RT.days * 24) + $RT.Hours) * 60 + $RT.Minutes + $RT.Seconds/60, 1)
$EmailBody += "Run Time: $Runtime minutes"
If($aryPending.count -gt 0){
    $EmailBody += "Pending- $Pending <br>"
    foreach($Item in $aryPending){
        $EmailBody += "$Item<br>"
    }
}else{
    $EmailBody += "No Pending items<br>"
}
$EmailBody += "<br>"
If($aryAdded.count -gt 0){
    $EmailBody += "Added- $Added <br>"
    foreach($Item in $aryAdded){
        $EmailBody += "$Item<br>"
    }
}else{
    $EmailBody += "No Added items<br>"
}
$EmailBody += "<br>"
If($aryActive.count -gt 0){
    $EmailBody += "Active- $Active <br>"
    foreach($Item in $aryActive){
        $EmailBody += "$Item<br>"
    }
}else{
    $EmailBody += "No Active items<br>"
}
$EmailBody += "<br>"
If($aryRemoved.count -gt 0){
    $EmailBody += "Removed- $Removed <br>"
    foreach($Item in $aryRemoved){
        $EmailBody += "$Item<br>"
    }
}else{
    $EmailBody += "No Removed items<br>"
}
$EmailBody += "<br>"
If($ErrorLog.count -gt 0){
    $EmailBody += "Errors- $($ErrorLog.Count) <br>"
    foreach($Item in $ErrorLog){
        $EmailBody += "$Item<br>"
    }
}else{
    $EmailBody += "No Errors<br>"
}
$EmailBody += "<br>"

Send-MailMessage -From $ScriptOwner -To $ScriptOwner -Subject "$ScriptName Script done" -Body $EmailBody -BodyAsHtml -SmtpServer "mailhost.utility.pge.com"
Write-Log "End | Script Complete"

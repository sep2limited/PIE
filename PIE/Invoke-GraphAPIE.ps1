﻿using namespace System.Collections.Generic

<#
  #===================================================#
  # GraphAPIE - MS Graph Phishing Intelligence Engine #
  # v4.2.G  --  January 2022                          #
  #===================================================#

# Copyright 2021 LogRhythm Inc.   
# Licensed under the MIT License. See LICENSE file in the project root for full license information.
# Graph API migration performed by Jon Cumiskey SEP2 Ltd (jon.cumiskey@sep2.co.uk)

INSTALL:
    Review lines 41 through 76
        Add credentials and mail service provider details under this section

    Review Lines 77 through 120
        For each setting that you would like to enable, change the value from $false to $true

USAGE:
    Configure as a scheduled task to run every 15-minutes:
        C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -command "& 'E:\Invoke-PIE.ps1'"
#>
$banner = @"
______ _     _     _     _               _____      _       _ _ _                             _____            _                   ___  _____ 
| ___ \ |   (_)   | |   (_)             |_   _|    | |     | | (_)                           |  ___|          (_)                /   ||  _  |
| |_/ / |__  _ ___| |__  _ _ __   __ _    | | _ __ | |_ ___| | |_  __ _  ___ _ __   ___ ___  | |__ _ __   __ _ _ _ __   ___     / /| || |/' |
|  __/| '_ \| / __| '_ \| | '_ \ / _  |   | || '_ \| __/ _ \ | | |/ _  |/ _ \ '_ \ / __/ _ \ |  __| '_ \ / _  | | '_ \ / _ \   / /_| ||  /| |
| |   | | | | \__ \ | | | | | | | (_| |  _| || | | | ||  __/ | | | (_| |  __/ | | | (_|  __/ | |__| | | | (_| | | | | |  __/   \___  |\ |_/ /
\_|   |_| |_|_|___/_| |_|_|_| |_|\__, |  \___/_| |_|\__\___|_|_|_|\__, |\___|_| |_|\___\___| \____/_| |_|\__, |_|_| |_|\___|       |_(_)___/ 
                                  __/ |                            __/ |                                  __/ |                            
                                 |___/                            |___/                                  |___/                            
"@

# Mask errors
$ErrorActionPreference= 'continue'

# ================================================================================
# DEFINE GLOBAL PARAMETERS AND CAPTURE CREDENTIALS
#
# ****************************** EDIT THIS SECTION ******************************
# ================================================================================



# Choose how to handle credentials - set the desired flag to $true
#     Be sure to set credentials or xml file location below
$EncodedXMLCredentials = $false

# XML Configuration - store credentials in an encoded XML (best option)
if ( $EncodedXMLCredentials ) {
    # ================================================================================
    #              Create PowerShell Stored Credential - Recommended
    #      PS E:\PIE\> Get-Credential | Export-Clixml PIE_cred.xml
    # ================================================================================
    $CredentialsFile = "E:\PIE\PIE_cred.xml"
    $PSCredential = Import-CliXml -Path $CredentialsFile 
} else {
    # ================================================================================
    #              Plain Text Credentials - Not Recommended
    #      Set line 43 to $false to leverage Username and Password variables.
    # ================================================================================
##Initial Setup parameters for creds
$TenantDomainName = ""
$ApplicationID = ""
$AccessSecret = '' #| ConvertTo-SecureString -AsPlainText -Force
}

# E-mail address for SOC Mailbox.  Typically same as Username value.
$SocMailbox = ""


# ================================================================================
#   Microsoft Graph API Setup and Token Request using Client Creds
#  
#  
# ================================================================================


#GraphAPI Auth Request
$Body = @{    
Grant_Type    = "client_credentials"
Scope         = "https://graph.microsoft.com/.default"
client_Id     = $ApplicationID
Client_Secret = $AccessSecret
} 

$ConnectGraph = Invoke-RestMethod -Uri "https://login.microsoftonline.com/$TenantDomainName/oauth2/v2.0/token" `
-Method POST -Body $Body

#GraphAPI Token
$token = $ConnectGraph.access_token

#Initial GraphAPI Test Request
$GraphTestURL = "https://graph.microsoft.com/v1.0/users/$SocMailbox/messages"
$testoutput = (Invoke-RestMethod -Headers @{Authorization = "Bearer $($token)"} -Uri $GraphTestURL -Method Get)



# ================================================================================
# LogRhythm SIEM Configuration
# ================================================================================

# LogRhythm Web Console Hostname
$LogRhythmHost = ""

# Case Tagging and User Assignment
# Default value - modify to match your case tagging schema. Note "PIE" tag is used with the Case Management Dashboard.
$LrCaseTags = ("PIE")

# Add as many users as you would like, separate them like so: "user1", "user2"...
$LrCaseCollaborators = ("")

# Playbook Assignment based on Playbook Name
$LrCasePlaybook = ("Phishing")

# Enable LogRhythm log search
$LrLogSearch = $true

# Enable LogRhythm Case output
$LrCaseOutput = $true

# Enable LogRhythm TrueIdentity Lookup
$LrTrueId = $true

# ================================================================================
# Third Party Analytics
# ================================================================================

# For each supported module, set the flag to $true.
# Note these modules must be appropriately setup and configured as part of LogRhythm.Tools.
# For additional details on LogRhyhtm.Tools setup and configuration, visit: 
# https://github.com/LogRhythm-Tools/LogRhythm.Tools

# VirusTotal
$virusTotal = $true

# URL Scan
$urlscan = $true

# Shodan.io
$shodan = $true

# ================================================================================
# END GLOBAL PARAMETERS
# ************************* DO NOT EDIT BELOW THIS LINE *************************
# ================================================================================

#Load dependencies required to parse emails removed by ZIP
#Ref https://pscustomobject.github.io/powershell/howto/PowerShell-Parse-Eml-File/

function Convert-EmlFile
{
<#
    .SYNOPSIS
        Function will parse an eml files.

    .DESCRIPTION
        Function will parse eml file and return a normalized object that can be used to extract infromation from the encoded file.

    .PARAMETER EmlFileName
        A string representing the eml file to parse.

    .EXAMPLE
        PS C:\> Convert-EmlFile -EmlFileName 'C:\Test\test.eml'

    .OUTPUTS
        System.Object
#>
    [CmdletBinding()]
    [OutputType([object])]
    param
    (
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]
        $EmlFileName
    )

    # Instantiate new ADODB Stream object
    $adoStream = New-Object -ComObject 'ADODB.Stream'

    # Open stream
    $adoStream.Open()

    # Load file
    $adoStream.LoadFromFile($EmlFileName)

    # Instantiate new CDO Message Object
    $cdoMessageObject = New-Object -ComObject 'CDO.Message'

    # Open object and pass stream
    $cdoMessageObject.DataSource.OpenObject($adoStream, '_Stream')

    return $cdoMessageObject

    }
# ================================================================================
# Date, File, and Global Email Parsing
# ================================================================================
# To support and facilitate accessing and creating resources
$pieFolder = $PSScriptRoot

# Folder Structure
$phishLog = "$pieFolder\logs\ongoing-phish-log.csv"
$caseFolder = "$pieFolder\cases\"
$tmpFolder = "$pieFolder\tmp\"
$runLog = "$pieFolder\logs\pierun.txt"
$phishingKeywords = Get-Content "$pieFolder\regex-swiftonsecurity.txt"

$LogsOutputFolder = (Join-Path $pieFolder -ChildPath "logs")
if (!(Test-Path $LogsOutputFolder)) {
    Try {
        New-Item -Path $LogsOutputFolder -ItemType Directory -Force | Out-Null
    } Catch {
        Write-Host "Unable to create folder: $LogsOutputFolder"
    } 
}

# PIE Version
$PIEVersion = 4.0

$PIEModules = @("logrhythm.tools")
ForEach ($ReqModule in $PIEModules){
    If ($null -eq (Get-Module $ReqModule -ListAvailable -ErrorAction SilentlyContinue)) {
        if ($ReqModule -like "logrhythm.tools") {
            Write-Host "$ReqModule is not installed.  Please install $ReqModule to continue."
            Write-Host "Please visit https://github.com/LogRhythm-Tools/LogRhythm.Tools"
            Return 0
        } else {
            Write-Verbose "Installing $ReqModule from default repository"
            Install-Module -Name $ReqModule -Force
            Write-Verbose "Importing $ReqModule"
            Import-Module -Name $ReqModule
        }
    } ElseIf ($null -eq (Get-Module $ReqModule -ErrorAction SilentlyContinue)){
        Try {
            Write-Host "Importing Module: $ReqModule"
            Import-Module -Name $ReqModule
        } Catch {
            Write-Host "Unable to import module: $ReqModule"
            Return 0
        }

    }
}

New-PIELogger -logSev "s" -Message "BEGIN - PIE Process" -LogFile $runLog -PassThru

# LogRhythm Tools Version
$LRTVersion = $(Get-Module -name logrhythm.tools | Select-Object -ExpandProperty Version) -join ","
New-PIELogger -logSev "i" -Message "PIE Version: $PIEVersion" -LogFile $runLog -PassThru
New-PIELogger -logSev "i" -Message "LogRhythm Tools Version: $LRTVersion" -LogFile $runLog -PassThru

s
$CaseOutputFolder = (Join-Path $pieFolder -ChildPath "cases")
if (!(Test-Path $CaseOutputFolder)) {
    Try {
        New-PIELogger -logSev "i" -Message "Creating folder: $CaseOutputFolder" -LogFile $runLog -PassThru
        New-Item -Path $CaseOutputFolder -ItemType Directory -Force | Out-Null
    } Catch {
        New-PIELogger -logSev "e" -Message "Unable to create folder: $CaseOutputFolder" -LogFile $runLog -PassThru
    } 
}

$CaseTmpFolder = (Join-Path $pieFolder -ChildPath "tmp")
if (!(Test-Path $CaseTmpFolder)) {
    Try {
        New-PIELogger -logSev "i" -Message "Creating folder: $CaseTmpFolder" -LogFile $runLog -PassThru
        New-Item -Path $CaseTmpFolder -ItemType Directory -Force | Out-Null
    } Catch {
        New-PIELogger -logSev "e" -Message "Unable to create folder: $CaseTmpFolder" -LogFile $runLog -PassThru
    }
    
}

# Email Parsing Varibles
$interestingFiles = @('pdf', 'exe', 'zip', 'doc', 'docx', 'docm', 'xls', 'xlsx', 'xlsm', 'ppt', 'pptx', 'arj', 'jar', '7zip', 'tar', 'gz', 'html', 'htm', 'js', 'rpm', 'bat', 'cmd')
$interestingFilesRegex = [string]::Join('|', $interestingFiles)
$specialPattern = "[^\u0000-\u007F]"
$imageext = ".png",".jpg",".gif"

# ================================================================================
# MEAT OF THE PIE
# ================================================================================
#Create phishLog if file does not exist.
if ( $(Test-Path $phishLog -PathType Leaf) -eq $false ) {
    New-PIELogger -logSev "a" -Message "No phishlog detected.  Created new $phishLog" -LogFile $runLog -PassThru
    Try {
        Set-Content $phishLog -Value "Guid,Timestamp,MessageId,SenderAddress,RecipientAddress,Subject"
    } Catch {
        New-PIELogger -logSev "e" -Message "Unable to create file: $phishLog" -LogFile $runLog -PassThru
    }    
}

#Open Inbox
Try {
    $testoutput
} Catch {
    New-PIELogger -logSev "i" -Message "Unable to open Mail Inbox." -LogFile $runLog -PassThru
}


#Gain List of mailfolders from Graph

    $GraphFoldersRequest = "https://graph.microsoft.com/v1.0/users/$SocMailbox/mailfolders"
    $GraphFolders = (Invoke-RestMethod -Headers @{Authorization = "Bearer $($token)"} -Uri $GraphFoldersRequest -Method Get)

#Get ID of Inbox folder

    $inboxid = $GraphFolders.value | Where-Object {$_.displayName -eq 'Inbox'} | select id
    $inboxidid = $inboxid.id

#Get list of child folders of Inbox

    $GraphInboxFolderListRequest = "https://graph.microsoft.com/v1.0/users/$SocMailbox/mailfolders/$inboxidid/childfolders"

    $GraphInboxFolderList = (Invoke-RestMethod -Headers @{Authorization = "Bearer $($token)"} -Uri $GraphInboxFolderListRequest -Method Get)


#Get the count of messages in Inbox, this count will be used to determine if processing of the mailbox is required.

    $GraphInboxFolderContentRequest = "https://graph.microsoft.com/v1.0/users/$SocMailbox/mailfolders/$inboxidid/messages"

    $GraphInboxFolderContentList = (Invoke-RestMethod -Headers @{Authorization = "Bearer $($token)"} -Uri $GraphInboxFolderContentRequest -Method Get)

    $InboxNewMail = $GraphInboxFolderContentList.value.Count

#Get the list of messages IDs from inbox to be processed

    $InboxMailIDs = @($GraphInboxFolderContentList.value.id)

#Validate Inbox/COMPLETED Folder

    $InboxCompleted = $GraphInboxFolderList.value | Where-Object {$_.DisplayName -eq 'COMPLETED'}
    $InboxCompleted = $InboxCompleted.id


# If the folder does not exist, create it.
if (!$InboxCompleted) {
    # Setup to create folders:
      $GraphInboxSubfolderCreateRequest = "https://graph.microsoft.com/v1.0/users/$SocMailbox/mailFolders/$inboxidid/childFolders"

    $completedbody = @{
        "displayName" = "COMPLETED"
        "isHidden"= $false } | ConvertTo-Json

    $GraphPusherCompleted = (Invoke-RestMethod -Headers @{Authorization = "Bearer $($token)"; "Content-Type" = "application/json"} -Uri $GraphInboxSubfolderCreateRequest -Method Post -body $completedbody)

    $InboxCompleted = $GraphPusherCompleted.id
}


#Validate Inbox/SKIPPED Folder

$InboxSkipped = $GraphInboxFolderList.value | Where-Object {$_.DisplayName -eq 'SKIPPED'}
$InboxSkipped = $InboxSkipped.id

#TEMP Remove this
New-PIELogger -logSev "s" -Message "Mail Count $InboxNewMail" -LogFile $runLog -PassThru
New-PIELogger -logSev "s" -Message "$GraphInboxFolderContentList" -LogFile $runLog -PassThru


# If the folder does not exist, create it.
if (!$InboxSkipped) {
    # Setup to create folders:
      $GraphInboxSubfolderCreateRequest = "https://graph.microsoft.com/v1.0/users/$SocMailbox/mailFolders/$inboxidid/childFolders"

    $skippedbody = @{
        "displayName" = "SKIPPED"
        "isHidden"= $false } | ConvertTo-Json

    $GraphPusherSkipped = (Invoke-RestMethod -Headers @{Authorization = "Bearer $($token)"; "Content-Type" = "application/json"} -Uri $GraphInboxSubfolderCreateRequest -Method Post -body $skippedbody)

    $InboxSkipped = $GraphPusherSkipped.id
}


################################
###Begin main detection loop###
###############################

if ($InboxNewMail -eq 0) {
    New-PIELogger -logSev "i" -Message "No new reports detected" -LogFile $runLog -PassThru
} else {
    New-PIELogger -logSev "i" -Message "New inbox items detected.  Proceeding to evaluate for PhishReports." -LogFile $runLog -PassThru

    #Initially the email is only a potential phish (i.e. boolean 0)
    $maliciousEmail = "false"
    
    ######################################################
    # Loop through each inbox item to identify standard msg PhishReports
    foreach ($i in $InboxMailIDs) {
        $ValidSubmissionAttachment = $false

        New-PIELogger -logSev "s" -Message "Begin processing newReports" -LogFile $runLog -PassThru
        $StartTime = (get-date).ToUniversalTime()
        New-PIELogger -logSev "i" -Message "Processing Start Time: $($StartTime)" -LogFile $runLog -PassThru

            $GraphAttachmentRequest = "https://graph.microsoft.com/v1.0/users/$SocMailbox/messages/$i/attachments"
            $GraphAttachments = (Invoke-RestMethod -Headers @{Authorization = "Bearer $($token)"} -Uri $GraphAttachmentRequest  -Method Get)

            #Just grab attachments to emails that are attached email messages, i.e. the raw phish, using two ways - either file extension, or rfc822.
            $attachmentid = $GraphAttachments.value | Where-Object {$_.name -match "\.msg$"}
            if (!$attachmentid) {
            $attachmentid = $GraphAttachments.value | Where-Object {$_.contentType -eq "message/rfc822"}

            }

            ################################################
            #Second attempt, if we can't get an .msg then try and grab .eml file

            if (!$attachmentid) {
            $attachmentid = $GraphAttachments.value | Where-Object {$_.name -match "\.eml$"}
            $emlid = $GraphAttachments.value | Where-Object {$_.name -match "\.eml$"}
            $emlid = $emlid.id

            #Unzip logic
            foreach ($eml in $emlid){
                $EmlAttachmentScrapeRequest = "https://graph.microsoft.com/v1.0/users/$socMailbox/messages/$i/attachments/$eml/?`$expand=microsoft.graph.itemattachment/item"
                $EmlAttachmentScrape = (Invoke-RestMethod -Headers @{Authorization = "Bearer $($token)"} -Uri $EmlAttachmentScrapeRequest  -Method Get)
              
              #Here is where we write the ZIP file and expand archive

              $emlname = $EmlAttachmentScrape.name
              $emlname = $emlname -replace '[\[\]\/\*\:\<\>\|"]'
              $emlname = $tmpFolder + $emlname
              $emlfile = [Convert]::FromBase64String($EmlAttachmentScrape.contentBytes)
                [IO.File]::WriteAllBytes($emlname, $emlfile)

              
               try { 
               
               #Extract Basic Mail Metadata
               $ReloadedUnzipped = Convert-EmlFile -EmlFileName $emlname

               #Manually extract Email
               $MimeAnalysis = Get-Content -LiteralPath $emlname

               
                }
               catch { 
               New-PIELogger -logSev "i" -Message "Failed to load $emlname - Terminating" -LogFile $runLog -PassThru 
               $attachmentid = $null
               }

               #Here we start loading in the good stuff

               #Set it as a zip file for further analysis
               #JC this was reused code from ZIP analysis, need to review if we broke anything
                             
                
                $from = $ReloadedUnzipped.From | Select-String -pattern '\<(?<email>\S+@[^\>]+)\>$'
                $from = $from.Matches.groups.value[1]
                #$ReloadedUnzipped.Subject
                $fromdisplay = $ReloadedUnzipped.From | Select-String -pattern '"(?<email>[^"]+)"'
                $fromdisplay = $fromdisplay.Matches.groups.value[1]


            }
            }

            ##########################################
            #Third attempt - try to use the ZIP file

            if (!$attachmentid) {
            $attachmentid = $GraphAttachments.value | Where-Object {$_.name -match "\.zip$"}
            $zipid = $GraphAttachments.value | Where-Object {$_.name -match "\.zip$"}
            $zipid = $zipid.id

            #Unzip logic
            foreach ($iz in $zipid){
                $ZipAttachmentScrapeRequest = "https://graph.microsoft.com/v1.0/users/$socMailbox/messages/$i/attachments/$iz/?`$expand=microsoft.graph.itemattachment/item"
                $ZipAttachmentScrape = (Invoke-RestMethod -Headers @{Authorization = "Bearer $($token)"} -Uri $ZipAttachmentScrapeRequest  -Method Get)
              
              #Here is where we write the ZIP file and expand archive

              $zipname = $ZipAttachmentScrape.name
              $zipname = $zipname -replace '[\[\]\/\*\:\<\>\|"]'
              $zipname = $tmpFolder + $zipname
              $zipfile = [Convert]::FromBase64String($ZipAttachmentScrape.contentBytes)
                [IO.File]::WriteAllBytes($zipname, $zipfile)

                Expand-Archive -Path $zipname -DestinationPath $tmpFolder -Force

                #Now we start working with the unzipped file and load it in

               $FileNames = Get-ChildItem $tmpFolder
               $unzipname = $FileNames.name -match '\.eml$'
               $unzipname = $tmpFolder + $unzipname

               #$unzipname = $zipname -replace 'zip$','eml'
               try { 
               
               #Extract Basic Mail Metadata
               $ReloadedUnzipped = Convert-EmlFile -EmlFileName $unzipname
               Start-Sleep -s 5

               #Manually extract Email
               $MimeAnalysis = Get-Content -LiteralPath $unzipname

               
                }
               catch { 
               New-PIELogger -logSev "i" -Message "Failed to load $UnzipName - Terminating" -LogFile $runLog -PassThru 
               $attachmentid = $null
               }

               #Here we start loading in the good stuff

               #Set it as a zip file for further analysis
               $isazip = "yes"


                $from = $ReloadedUnzipped.From | Select-String -pattern '\<(?<email>\S+@[^\>]+)\>$'
                $from = $from.Matches.groups.value[1]
                #$ReloadedUnzipped.Subject
                $fromdisplay = $ReloadedUnzipped.From | Select-String -pattern '"(?<email>[^"]+)"'
                $fromdisplay = $fromdisplay.Matches.groups.value[1]



               }
              }
            $attachmentidid = @($attachmentid.id)

            #Bomb out and move on if nothing attached.
            if (!$attachmentid) {

            $FolderDestSkipped += @($i)
            New-PIELogger -logSev "s" -Message "No phishing email was attached. Moving this message to skipped. " -LogFile $runLog -PassThru


            }

            
            #############################
            #Now we have a good file loaded in from our phish report, lets take a look at it.
            #Continue if good
            else {


                foreach ($i2 in $attachmentidid){
                $GraphAttachmentScrapeRequest = "https://graph.microsoft.com/v1.0/users/$socMailbox/messages/$i/attachments/$i2/?`$expand=microsoft.graph.itemattachment/item"
                $GraphAttachmentScrape = (Invoke-RestMethod -Headers @{Authorization = "Bearer $($token)"} -Uri $GraphAttachmentScrapeRequest  -Method Get)

                $ValidSubmissionAttachment = $true



        
        # Load NewReport
        $NewReport = $GraphInboxFolderContentList.value | where-object {$_.id -eq $i}

        # Establish Submission PSObject
        $Attachments = [list[object]]::new()

        # Report GUID
        $ReportGuid = $(New-Guid | Select-Object -ExpandProperty Guid)
        New-PIELogger -logSev "i" -Message "Evaluation GUID: $($ReportGuid)" -LogFile $runLog -PassThru

        # Add data for evaluated email
        $ReportEvidence = [PSCustomObject]@{
            Meta = [PSCustomObject]@{ 
                GUID = $null
                Timestamp = $null
                Metrics = [PSCustomObject]@{ 
                    Begin = $null
                    End = $null
                    Duration = $null
                }
                Version = [PSCustomObject]@{ 
                    PIE = $null
                    LRTools = $null
                }
            }
            ReportSubmission = [PSCustomObject]@{
                Sender = $null
                SenderDisplayName = $null
                Recipient = $null
                Subject = [PSCustomObject]@{
                    Original = $null
                    Modified = $null
                }
                UtcDate = $null
                MessageId = $null
                Attachment = [PSCustomObject]@{
                    Name = $null
                    Type = $null
                    Hash = $null
                }
            }
            EvaluationResults = [PSCustomObject]@{
                Sender = $null
                SenderDisplayName = $null
                Recipient = [PSCustomObject]@{
                    To = $null
                    CC = $null
                }
                UtcDate = $null
                Subject = [PSCustomObject]@{
                    Original = $null
                    Modified = $null
                }
                Body = [PSCustomObject]@{
                    Original = $null
                    Modified = $null
                }
                HTMLBody = [PSCustomObject]@{
                    Original = $null
                    Modified = $null
                }
                Headers = [PSCustomObject]@{
                    Source = $null
                    Details = $null
                }
                Attachments = $null
                Links = [PSCustomObject]@{
                    Source = $null
                    Value = $null
                    Details = $null
                }
                LogRhythmTrueId = [PSCustomObject]@{
                    Sender = $null
                    Recipient = $null
                }
            }
            LogRhythmCase = [PSCustomObject]@{
                Number = $null
                Url = $null
                Details = $null
            }
            LogRhythmSearch = [PSCustomObject]@{
                TaskID = $null
                Status = $null
                Summary = [PSCustomObject]@{
                    Quantity = $null
                    Recipient = $null
                    Sender = $null
                    Subject = $null
                }
                Details = [PSCustomObject]@{
                    SendAndSubject = [PSCustomObject]@{ 
                        Sender = $null
                        Subject = $null
                        Recipients = $null
                        Quantity = $null
                    }
                    Sender = [PSCustomObject]@{ 
                        Sender = $null
                        Recipients = $null
                        Subjects = $null
                        Quantity = $null
                    }
                    Subject = [PSCustomObject]@{ 
                        Senders = $null
                        Recipients = $null
                        Subject = $null
                        Quantity = $null
                    }
                }
            }
        }

        # Set initial time data
        $ReportEvidence.Meta.Timestamp = $StartTime.ToString("yyyy-MM-ddTHHmmssffZ")
        $ReportEvidence.Meta.Metrics.Begin = $StartTime.ToString("yyyy-MM-ddTHHmmssffZ")

        # Set PIE Metadata versions
        $ReportEvidence.Meta.Version.PIE = $PIEVersion
        $ReportEvidence.Meta.Version.LRTools = $LRTVersion

        # Set PIE Meta GUID
        $ReportEvidence.Meta.Guid = $ReportGuid

        # Set ReportSubmission data

      

        $ReportEvidence.ReportSubmission.Sender = $($NewReport.sender.emailAddress.address)
        $ReportEvidence.ReportSubmission.SenderDisplayName = $($NewReport.sender.emailAddress.name)
        $ReportEvidence.ReportSubmission.Recipient = $($NewReport.toRecipients.emailaddress.address)
        $ReportEvidence.ReportSubmission.Subject.Original = $($NewReport.Subject)
        $ReportEvidence.ReportSubmission.UtcDate = $($NewReport.sentDateTime)#.ToString("yyyy-MM-ddTHHmmssffZ")
        $ReportEvidence.ReportSubmission.MessageId = $($NewReport.internetMessageId) -replace '(\<|\>)',''

   

        }


        # Track the user who reported the message
        New-PIELogger -logSev "i" -Message "Sent By: $($ReportEvidence.ReportSubmission.Sender)  Reported Subject: $($ReportEvidence.ReportSubmission.Subject.Original)" -LogFile $runLog -PassThru
        

        #This is the phisher's email address
        #$GraphAttachmentScrape.item.sender.emailAddress.address
        
 
#############################


        # Extract and load attached e-mail attachment

        foreach ($ii in $($attachmentidid)) {

           
            $GraphAttachmentMimeScrapeRequest = "https://graph.microsoft.com/v1.0/users/$SocMailbox/messages/$i/attachments/$ii/`$value" 

            $GraphAttachmentMimeScrape = (Invoke-RestMethod -Headers @{Authorization = "Bearer $($token)"} -Uri $GraphAttachmentMimeScrapeRequest  -Method Get)
            $AttachmentSubDirectory = $ReportEvidence.ReportSubmission.Subject.Original -replace '[\[\]\/\*\:\<\>\|"]'
            $AttachmentDirectory = Join-Path -Path $tmpFolder -ChildPath $AttachmentSubDirectory
            $AttachmentOutputFile = $ReportEvidence.ReportSubmission.MessageID + "_mime.txt"
            $AttachmentOutputFile = $AttachmentOutputFile -replace '[\[\]\/\*\:\<\>\|"]'
            $AttachmentOutputFile = $AttachmentDirectory + "-" + $AttachmentOutputFile
            $GraphAttachmentMimeScrape | Out-File $AttachmentOutputFile

            if (!$isazip)
            {
            $MimeAnalysis = Get-Content $AttachmentOutputFile | Out-String
            

            #Extract Content from MimeMesage
            $EncodedAttachment = $MimeAnalysis | Select-String -pattern '(?msi)Content-Transfer-Encoding: base64..(?<msg>.*?).--(Apple-Mail=)?_' -AllMatches

            #Get Filenames
            $attachmentSelect = $GraphAttachmentMimeScrape | select-string -pattern '(?msi)filename=\"?(?<msg>[^\"\n]+)"?;?.' -AllMatches

            if ($EncodedAttachment) {


               $AttachmentsRaw = @()


                 for ($ixx = 0; $ixx -lt $EncodedAttachment.Matches.Count; $ixx++) {
 
                        #$atb64 = $EncodedAttachment.Matches[$ixx].groups[1].value
                        $atb64 = $EncodedAttachment.Matches[$ixx].Captures.Groups.Value[2]

                        $atname = $attachmentSelect.Matches[$ixx].Groups[1].value 

                
                $atname = $atname -replace "\r\n","" #Remove carriage returns from filenames.
                $ReportEvidence.ReportSubmission.Attachment.Name = $atname 


                $attachmentsRaw += New-Object -TypeName psobject -Property @{Name=$atname; b64=$atb64}



                  } 

                
                $AttachCount = $attachmentsRaw.count

                }
 
              }


            if ($isazip) {
            

            $pattern = '(?msi)Content-Type: (?<type>\S+)\r\nContent-Transfer-Encoding: base64\r\nContent-Disposition: attachment;\sfilename="(?<filename>[^"]+)"\r\n(?<msg>.*?)\r\n\r\n--='
            
            $MimeAnalysis = $MimeAnalysis | Out-String
            $ZipEmlattachmentSelect = $MimeAnalysis | select-string -pattern $pattern -AllMatches

           
            if ($ZipEmlattachmentSelect) {


               $AttachmentsRaw = @()


                 for ($izz = 0; $izz -lt $ZipEmlattachmentSelect.Matches.Count; $izz++) {
 
                        $atb64 = $ZipEmlattachmentSelect.Matches[$izz].Groups["msg"].value

                        $atname = $ZipEmlattachmentSelect.Matches[$izz].Groups["filename"].value

                
                $atname = $atname -replace "\r\n","" #Remove carriage returns from filenames.
                $ReportEvidence.ReportSubmission.Attachment.Name = $atname 


                $AttachmentsRaw += New-Object -TypeName psobject -Property @{Name=$atname; b64=$atb64}



                  } 

                
                $AttachCount = $AttachmentsRaw.count
 
              }


                }
            } 




        # Load e-mail from file
        $Eml = $GraphAttachmentMimeScrape

        $ReportEvidence.EvaluationResults.Subject.Original = $ReloadedUnzipped.Subject
        if ($($ReportEvidence.EvaluationResults.Subject.Original) -Match "$specialPattern") {
            New-PIELogger -logSev "i" -Message "Creating Message Subject without Special Characters to support Case Note." -LogFile $runLog -PassThru
            $ReportEvidence.EvaluationResults.Subject.Modified = $ReportEvidence.EvaluationResults.Subject.Original -Replace "$specialPattern","?"
        }
        New-PIELogger -logSev "d" -Message "Message subject: $($ReportEvidence.EvaluationResults.Subject)" -LogFile $runLog -PassThru
        
 
 #Plain text message processing         
 if (!$isazip) {

 New-PIELogger -logSev "s" -Message "Begin MSG format message processing logic" -LogFile $runLog -PassThru         

                #Plain text Message Body
                ##This regex reads the Mime body for content-type=text/plain or text/html and captures everything into capture group until the --_ delimter.
                # Set ReportEvidence HTMLBody Content
                    $MimeBody = $GraphAttachmentMimeScrape | Select-String -pattern "(?msi)Content-Type: text\/(html)(?<bodyoutput>.*?)--_"
                    $MimeBodyPlain = $GraphAttachmentMimeScrape | Select-String -pattern "(?msi)Content-Type: text\/(plain)(?<bodyoutput>.*?)--_"

                    $ReportEvidence.EvaluationResults.Body.Original = $GraphAttachmentScrape.item.body.content

                #Parse the Body to make it readable
                    #This removes the '=' EOL message and new lines from 	URLs, which would otherwise cause URL parsing issues.
                    $ReportEvidence.EvaluationResults.Body.Original = $ReportEvidence.EvaluationResults.Body.Original -replace "(?msi)=\r\n",""

                    #This transformation removes an extrenaus "3D" character from URL's causing formatting issues
                    $ReportEvidence.EvaluationResults.Body.Original = $ReportEvidence.EvaluationResults.Body.Original -replace "(?msi)=3Dhttp","=http"

                if ($($ReportEvidence.EvaluationResults.Body.Original) -Match "$specialPattern") {
                    New-PIELogger -logSev "i" -Message "Creating Message Body without Special Characters to support Case Note." -LogFile $runLog -PassThru
                    $ReportEvidence.EvaluationResults.Body.Modified = $ReportEvidence.EvaluationResults.Body.Original -Replace "$specialPattern","?"
                }                     

                #Headers
                New-PIELogger -logSev "d" -Message "Processing Headers" -LogFile $runLog -PassThru
                #This line gets all data before content-type to get message headers
                $ReportEvidence.EvaluationResults.Headers.Source = $GraphAttachmentScrape.item.internetMessageHeaders
                #Rename 'name' array header to 'field', for presentation to invoke-pieheaderinspect
                $ReportEvidence.EvaluationResults.Headers.Source = $ReportEvidence.EvaluationResults.Headers.Source | select @{n='field';e={$_.name}},value
                Try {
                    New-PIELogger -logSev "i" -Message "Parsing Header Details" -LogFile $runLog -PassThru
                    $ReportEvidence.EvaluationResults.Headers.Details = Invoke-PieHeaderInspect -Headers $ReportEvidence.EvaluationResults.Headers.Source
                } Catch {
                    New-PIELogger -logSev "e" -Message "Parsing Header Details" -LogFile $runLog -PassThru
                }

                New-PIELogger -logSev "s" -Message "Begin Parsing URLs" -LogFile $runLog -PassThru                  
                #Check if HTML Body exists else populate links from Text Body
                New-PIELogger -logSev "i" -Message "Identifying URLs" -LogFile $runLog -PassThru
                    
                if ( $($GraphAttachmentScrape.item.body.content -like "*<html>*") ) { 

                $ReportEvidence.EvaluationResults.HTMLBody.Original = $GraphAttachmentScrape.item.body.content

                    New-PIELogger -logSev "d" -Message "Processing MIME to raw HTML" -LogFile $runLog -PassThru

                     # Pull URL data from HTMLBody Content
                    New-PIELogger -logSev "d" -Message "Processing URLs from message HTML body" -LogFile $runLog -PassThru
                    #This is where we scrape links
                    $HTML = New-Object -ComObject "HTMLFile"
                    $HTML.IHTMLDocument2_write($ReportEvidence.EvaluationResults.Body.Original)
                    $HTMLLinks = $HTML.links | % href
            }


        }


  #ZIPPED message processing    

  if ($isazip) {

            New-PIELogger -logSev "s" -Message "Beginning ZIP file message processing logic" -LogFile $runLog -PassThru         


                    $ReportEvidence.EvaluationResults.Body.Original = $ReloadedUnzipped.HTMLBody
                    $ReportEvidence.EvaluationResults.Recipient 
                    $ReportEvidence.EvaluationResults.UtcDate = $ReloadedUnzipped.ReceivedTime
                    $ReportEvidence.EvaluationResults.HTMLBody


             

                if ($($ReportEvidence.EvaluationResults.Body.Original) -Match "$specialPattern") {
                    New-PIELogger -logSev "i" -Message "Creating Message Body without Special Characters to support Case Note." -LogFile $runLog -PassThru
                    $ReportEvidence.EvaluationResults.Body.Modified = $ReportEvidence.EvaluationResults.Body.Original -Replace "$specialPattern","?"
                }                     

                #Headers
                New-PIELogger -logSev "d" -Message "Processing Headers" -LogFile $runLog -PassThru
                #This line gets all data before content-type to get message headers, load it back in to turn it into a single value
                $MimeAnalysis = Out-String -InputObject $MimeAnalysis
                $ZIPHeaders = $MimeAnalysis | Select-String -Pattern '(?msi)^(?<capture>.*?)Mime-Version: \d'

                

                $ReportEvidence.EvaluationResults.Headers.Source.value = $ZIPHeaders.Matches[0].Value

                ##JC this doesn't work yet #Rename 'name' array header to 'field', for presentation to invoke-pieheaderinspect
                $ReportEvidence.EvaluationResults.Headers.Source = $ReportEvidence.EvaluationResults.Headers.Source | select @{n='field';e={$_.name}},value
                Try {
                    New-PIELogger -logSev "i" -Message "Parsing Header Details" -LogFile $runLog -PassThru
                    $ReportEvidence.EvaluationResults.Headers.Details = Invoke-PieHeaderInspect -Headers $ReportEvidence.EvaluationResults.Headers.Source
                } Catch {
                    New-PIELogger -logSev "e" -Message "Unable to parse Header Details" -LogFile $runLog -PassThru
                }

                New-PIELogger -logSev "s" -Message "Begin Parsing URLs" -LogFile $runLog -PassThru                  
                #Check if HTML Body exists else populate links from Text Body
                New-PIELogger -logSev "i" -Message "Identifying URLs" -LogFile $runLog -PassThru
                    
                if ( $($GraphAttachmentScrape.item.body.content -like "*<html>*") ) { 

                $ReportEvidence.EvaluationResults.HTMLBody.Original = $GraphAttachmentScrape.item.body.content

                    New-PIELogger -logSev "d" -Message "Processing MIME to raw HTML" -LogFile $runLog -PassThru

                     # Pull URL data from HTMLBody Content
                    New-PIELogger -logSev "d" -Message "Processing URLs from message HTML body" -LogFile $runLog -PassThru
                    #This is where we scrape links
                    $HTML = New-Object -ComObject "HTMLFile"
                    $HTML.IHTMLDocument2_write($ReportEvidence.EvaluationResults.Body.Original)
                    $HTMLLinks = $HTML.links | % href
} 

         }


            #Here we sanitise links we have scraped

            #Declare array
            $formatlink = @()

            foreach ($iiii in $HTMLLinks){
          
            if ( $iiii -like "*safelinks.protection.outlook.com*" ) {
                    
                    $striplink = $iiii.Substring($iiii.IndexOf('=') + 1)
                    $formatlink += @([System.Web.HttpUtility]::UrlDecode($striplink))
                    
                    }

            elseif ( $iiii -notlike "*safelinks.protection.outlook.com*" ) {

                    $formatlink += @($iiii)

            
                }
                     }

            
            #Only include links that start with HTTP and remove duplicates.
            $formatlink  = $formatlink  -match "^http"
            $formatlink  = $formatlink  | Select-Object -Unique


 
            $ReportEvidence.EvaluationResults.Links.Source = "HTML" #}
            #$ReportEvidence.EvaluationResults.Links.Value = $formatlink
 


            # Create copy of HTMLBody with special characters removed.
            if ($($ReportEvidence.EvaluationResults.Body.Original) -Match "$specialPattern") {
                New-PIELogger -logSev "i" -Message "Creating HTMLBody without Special Characters to support Case Note." -LogFile $runLog -PassThru
                $ReportEvidence.EvaluationResults.HTMLBody.Original = $ReportEvidence.EvaluationResults.HTMLBody.Original -Replace "$specialPattern","?"
            
        } #else {
            New-PIELogger -logSev "a" -Message "Processing URLs from Message body" -LogFile $runLog -PassThru
            $ReportEvidence.EvaluationResults.Links.Source = "HTML"

            #Extract unwanted file extensions and deduplicate
            $ReportEvidence.EvaluationResults.Links.Value = $(Get-PIEURLsFromHtml -HtmlSource $($ReportEvidence.EvaluationResults.Body.Original)) | Where-Object {$_.extension -notin $imageext}
            $ReportEvidence.EvaluationResults.Links.Value = $ReportEvidence.EvaluationResults.Links.Value | Sort-Object -Property url -Unique
   
        #} 
        New-PIELogger -logSev "s" -Message "End Parsing URLs - Total Count $ReportEvidence.EvaluationResults.Links.Value.Count" -LogFile $runLog -PassThru

        New-PIELogger -logSev "s" -Message "Begin Attachment block" -LogFile $runLog -PassThru
        New-PIELogger -logSev "i" -Message "Attachment Count: $($AttachCount)" -LogFile $runLog -PassThru
        if ( $AttachCount -gt 0 ) {
            # Validate path tmpFolder\attachments exists
            if (Test-Path "$tmpFolder\attachments" -PathType Container) {
                New-PIELogger -logSev "i" -Message "Folder $tmpFolder\attachments\ exists" -LogFile $runLog -PassThru
            } else {
                New-PIELogger -logSev "i" -Message "Creating folder: $tmpFolder\attachments\" -LogFile $runLog -PassThru
                Try {
                    New-Item -Path "$tmpFolder\attachments" -type Directory -Force | Out-null
                } Catch {
                    New-PIELogger -logSev "e" -Message "Unable to create folder: $tmpFolder\attachments\" -LogFile $runLog -PassThru
                }
            }                
            # Get the filename and location 

            
                  

            ForEach ($ixy in $attachmentsraw.Name) {
                $attachmentFull = $tmpFolder+"attachments\"+$ixy
                New-PIELogger -logSev "d" -Message "Attachment Name: $ixy" -LogFile $runLog -PassThru
                New-PIELogger -logSev "i" -Message "Checking attachment against interestingFilesRegex" -LogFile $runLog -PassThru
                If ($ixy -match $interestingFilesRegex) {
                    New-PIELogger -logSev "d" -Message "Saving Attachment to destination: $tmpFolder\attachments\$ixy" -LogFile $runLog -PassThru

                      # Establish FileStream to faciliating extracting e-mail attachment
                    $TmpSavePath = $attachmentFull
                    $SaveFileMode = [System.IO.FileMode]::Create
                    $SaveFileAccess = [System.IO.FileAccess]::Write
                    $SaveFileShare = [System.IO.FileShare]::Read

          

                    #Refactored file write 

                    #$bytes = @()
                   # foreach ($iz in $AttachmentsRaw)
                   # {
                   $b64output = $AttachmentsRaw | where name -eq $ixy | select b64
                   $bytes = [Convert]::FromBase64String($b64output.b64)
                   #) }


                   #Save it

                    Try {
                        [IO.File]::WriteAllBytes($TmpSavePath, $bytes) 
                        Start-Sleep 5
                    } Catch {
                        New-PIELogger -logSev "e" -Message "Unable to save file $TmpSavePath." -LogFile $runLog -PassThru
                    }
                    # Release FileStream
                     #A little pause needed so we don't write an empty file by closing too early. 

                    #Create file hash
                
                    $ReportEvidence.ReportSubmission.Attachment.Hash = @(Get-FileHash $TmpSavePath -Algorithm SHA256)
                    $HashPrep =  $GraphAttachmentScrape.item.attachments | where name -eq $ixy | select contenttype

                    $Attachment = [PSCustomObject]@{
                        Name = $ixy
                        Type = $HashPrep.contentType
                        Hash = $ReportEvidence.ReportSubmission.Attachment.Hash
                        Plugins = [pscustomobject]@{
                            VirusTotal = $null
                        }
                        Links = [PSCustomObject]@{
                            Source = $null
                            Value = $null
                            Details = $null
                        }
                    }
                    # Add Attachment object to Attachments list
                    if ($Attachments -notcontains $attachment) {
                        $Attachments.Add($Attachment)
                    }
                }
            }
            $ReportEvidence.EvaluationResults.Attachments = $Attachments
        }

        #Load in sender metadata - if the value does not exist using standard msg graphattachmentscrape then get it using the unzipped technique
        
        $ReportEvidence.EvaluationResults.Sender = $GraphAttachmentScrape.item.sender.emailAddress.address
        if (!$ReportEvidence.EvaluationResults.Sender)
         { $ReportEvidence.EvaluationResults.Sender = $from }

        $ReportEvidence.EvaluationResults.SenderDisplayName = $GraphAttachmentScrape.item.sender.emailAddress.name
        if (!$ReportEvidence.EvaluationResults.SenderDisplayName)
         { $ReportEvidence.EvaluationResults.SenderDisplayName = $fromdisplay }

        New-PIELogger -logSev "i" -Message "Origin sender set to: $($ReportEvidence.EvaluationResults.Sender )" -LogFile $runLog -PassThru
        if ($LrTrueId) {
            New-PIELogger -logSev "s" -Message "LogRhythm API - Begin TrueIdentity Lookup" -LogFile $runLog -PassThru
            New-PIELogger -logSev "i" -Message "LogRhythm API - TrueID Sender: $($ReportEvidence.EvaluationResults.Sender)" -LogFile $runLog -PassThru
            $LrTrueIdSender = Get-LrIdentities -Identifier $ReportEvidence.EvaluationResults.Sender
            if ($LrTrueIdSender) {
                New-PIELogger -logSev "i" -Message "LogRhythm API - Sender Identitity Id: $($LrTrueIdSender.identityId)" -LogFile $runLog -PassThru
                $ReportEvidence.EvaluationResults.LogRhythmTrueId.Sender = $LrTrueIdSender
            }
            Start-Sleep 0.2
            New-PIELogger -logSev "i" -Message "LogRhythm API - TrueID Recipient: $($ReportEvidence.ReportSubmission.Sender)" -LogFile $runLog -PassThru
            $LrTrueIdRecipient = Get-LrIdentities -Identifier $ReportEvidence.ReportSubmission.Sender
            if ($LrTrueIdRecipient) {
                New-PIELogger -logSev "i" -Message "LogRhythm API - Recipient Identitity Id: $($LrTrueIdRecipient.identityId)" -LogFile $runLog -PassThru
                $ReportEvidence.EvaluationResults.LogRhythmTrueId.Recipient = $LrTrueIdRecipient
            }
            New-PIELogger -logSev "s" -Message "LogRhythm API - End TrueIdentity Lookup" -LogFile $runLog -PassThru
        }
        

        if ($Eml.To.count -ge 1 -or $Eml.To.Length -ge 1) {
            $ReportEvidence.EvaluationResults.Recipient.To = $GraphAttachmentScrape.item.toRecipients.emailAddress.address
        } else {
            $ReportEvidence.EvaluationResults.Recipient.To = $ReportEvidence.ReportSubmission.Sender
        }
        $ReportEvidence.EvaluationResults.Recipient.CC = $GraphAttachmentScrape.item.ccRecipients.emailAddress.address
        $ReportEvidence.EvaluationResults.UtcDate = $GraphAttachmentScrape.item.sentDateTime
        New-PIELogger -logSev "i" -Message "Origin Sender Display Name set to: $($ReportEvidence.EvaluationResults.SenderDisplayName)" -LogFile $runLog -PassThru
    
        # Begin Section - Search
        if ($LrLogSearch) {
            New-PIELogger -logSev "s" -Message "LogRhythm API - Begin Log Search" -LogFile $runLog -PassThru
            $LrSearchTask = Invoke-PIELrMsgSearch -EmailSender $($ReportEvidence.EvaluationResults.Sender) -Subject $($ReportEvidence.EvaluationResults.Subject.Original) -SocMailbox $SocMailbox
            New-PIELogger -logSev "i" -Message "LogRhythm Search API - TaskId: $($LrSearchTask.TaskId) Status: Starting" -LogFile $runLog -PassThru
        }

        New-PIELogger -logSev "s" -Message "Begin Attachment Processing" -LogFile $runLog -PassThru
        ForEach ($Attachment in $ReportEvidence.EvaluationResults.Attachments) {
            New-PIELogger -logSev "i" -Message "Attachment: $($Attachment.Name)" -LogFile $runLog -PassThru
            if ($LrtConfig.VirusTotal.ApiKey) {
                New-PIELogger -logSev "i" -Message "VirusTotal - Submitting Hash: $($Attachment.Hash.Hash)" -LogFile $runLog -PassThru
                $VTResults = Get-VtHashReport -Hash $Attachment.Hash.Hash
                # Response Code 0 = Result not in dataset
                if ($VTResults.response_code -eq 0) {
                    New-PIELogger -logSev "i" -Message "VirusTotal - Result not in dataset." -LogFile $runLog -PassThru
                    $VTResponse = [PSCustomObject]@{
                        Status = $true
                        Note = $VTResults.verbose_msg
                        Results = $VTResults
                    }
                    $Attachment.Plugins.VirusTotal = $VTResponse
                } elseif ($VTResults.response_code -eq 1) {
                    # Response Code 1 = Result in dataset
                    New-PIELogger -logSev "i" -Message "VirusTotal - Result in dataset." -LogFile $runLog -PassThru
                    $maliciousEmail = "true"

                    $VTResponse = [PSCustomObject]@{
                        Status = $true
                        Note = $VTResults.verbose_msg
                        Results = $VTResults
                    }
                    $Attachment.Plugins.VirusTotal = $VTResults
                } else {
                    New-PIELogger -logSev "e" -Message "VirusTotal - Request failed." -LogFile $runLog -PassThru
                    $VTResponse = [PSCustomObject]@{
                        Status = $false
                        Note = "Requested failed."
                        Results = $VTResults
                    }
                    $Attachment.Plugins.VirusTotal = $VTResponse
                }
                # Inspect Attachment for URLs
                $AttachmentTypes_UrlScrape = @("text/html")
                if ($AttachmentTypes_UrlScrape -contains $Attachment.Type ) {
                    $Attachment.Links.Source = $(Get-Content -Path $Attachment.Hash.Path -Raw).ToString()
                    New-PIELogger -logSev "i" -Message "Attachment - URLs processing from Text Source" -LogFile $runLog -PassThru
                    $Attachment.Links.Value = Get-PIEURLsFromText -Text $Attachment.Links.Source
                    New-PIELogger -logSev "i" -Message "Attachment - URL Count: $($Attachment.Links.Value.count)" -LogFile $runLog -PassThru
                }

                # Process URL Details from Attachments
                If ($Attachment.Links.Value) {
                    $AttachmentUrlDetails = [list[pscustomobject]]::new()
                    ForEach ($AttachmentURL in $Attachment.Links.Value) {
                        
                        $DetailResults = Get-PIEUrlDetails -Url $AttachmentURL -EnablePlugins
                        if ($AttachmentUrlDetails -NotContains $DetailResults) {
                            $AttachmentUrlDetails.Add($DetailResults)
                        }
                    }
                    $Attachment.Links.Details = $AttachmentUrlDetails
                }
            }
        }
        New-PIELogger -logSev "s" -Message "End Attachment Processing" -LogFile $runLog -PassThru


        New-PIELogger -logSev "s" -Message "Begin Link Processing" -LogFile $runLog -PassThru
        $EmailUrls = [list[string]]::new()
        if ($ReportEvidence.EvaluationResults.Links.Value) {
            $UrlDetails = [list[pscustomobject]]::new()
            if ($ReportEvidence.EvaluationResults.Links.Source -like "HTML") { 
                New-PIELogger -logSev "i" -Message "Link processing from HTML Source" -LogFile $runLog -PassThru
                $EmailUrls = $ReportEvidence.EvaluationResults.Links.Value

                    $repl = $ReportEvidence.EvaluationResults.Links.Value -replace 'https?:\/\/',''
                    $DomainGroups = $repl -replace '\/.*$',''
                $UniqueDomainValues = $DomainGroups | Sort-Object | Get-Unique
                #Hygiene for ZIP files
                $UniqueDomainValues -replace '^.*=',''

                New-PIELogger -logSev "i" -Message "Links: $($EmailUrls.count) Domains: $($UniqueDomainValues.Count)" -LogFile $runLog -PassThru
                ForEach ($UniqueDomains in $UniqueDomainValues) {  
                    #New-PIELogger -logSev "i" -Message "Domain: $($UniqueDomains.Name) URLs: $($UniqueDomains.Count)" -LogFile $runLog -PassThru
                    if ($UniqueDomainValues.count -ge 2) {
                        # Collect details for initial
                        #$ScanTarget = $EmailUrls | Where-Object -Property hostname -like $UniqueDomains.Name | Select-Object -ExpandProperty Url -First 1
                        $ScanTarget = $EmailUrls |  Select-Object -ExpandProperty Url -First 1
                        New-PIELogger -logSev "i" -Message "Retrieve Domain Details - Url: $ScanTarget" -LogFile $runLog -PassThru
                        $DetailResults = Get-PIEUrlDetails -Url $ScanTarget -EnablePlugins -VTDomainScan

                       if ($DetailResults.Plugins.VirusTotal.detected_urls.count > 0)

                             { $maliciousEmail = "true" }

                        if ($UrlDetails -NotContains $DetailResults) {
                            $UrlDetails.Add($DetailResults)
                        }

                        # Provide summary but skip plugin output for remainder URLs - Here is where we submit to VirusTotal etc.
                        $SummaryLinks = $EmailUrls
                        ForEach ($SummaryLink in $SummaryLinks) {
                            $DetailResults = Get-PIEUrlDetails -Url $SummaryLink#.URL
                            if ($UrlDetails -NotContains $DetailResults) {
                                $UrlDetails.Add($DetailResults)

                            }
                        }
                    } else {
                        $ScanTargets = $EmailUrls | Where-Object {$_.Type -like "URL"}
                        #Remove anomalous links
                        $EmailUrls = $EmailUrls | Where-Object {$_.url -match "http"}
                        ForEach ($ScanTarget in $ScanTargets) {

                            $ScanTarget = $ScanTarget.url
                          

                          ##Note: To get the VT Submission plugin I had to change the Get-PIEURLDetails module file, changing from VT lines from "scantarget.url" to "scantarget"

                            New-PIELogger -logSev "i" -Message "Retrieve URL Details - Url: $ScanTarget" -LogFile $runLog -PassThru
                            $DetailResults = Get-PIEUrlDetails -Url $ScanTarget -EnablePlugins
                            if ($UrlDetails -NotContains $DetailResults) {
                                $UrlDetails.Add($DetailResults)


                            }
                        }
                    }
                }
            }
            # URLs pulled from e-mail body as Text
            if ($ReportEvidence.EvaluationResults.Links.Source -like "Text") {
                $EmailUrls = $ReportEvidence.EvaluationResults.Links.Value
                ForEach ($EmailURL in $EmailUrls) {
                    $DetailResults  = Get-PIEUrlDetails -Url $EmailURL
                    if ($UrlDetails -NotContains $DetailResults) {
                        $UrlDetails.Add($DetailResults)
                    }
                }
            }
            # Add the UrlDetails results to the ReportEvidence object.
            if ($UrlDetails) {
                $ReportEvidence.EvaluationResults.Links.Details = $UrlDetails
            }
        }
        New-PIELogger -logSev "s" -Message "End - Link Processing" -LogFile $runLog -PassThru
        
        New-PIELogger -logSev "s" -Message "Begin - Link Summarization" -LogFile $runLog -PassThru
        if ($ReportEvidence.EvaluationResults.Links.Details) {       
            New-PIELogger -logSev "s" -Message "Begin - Link Summary from Message Body" -LogFile $runLog -PassThru           
            New-PIELogger -logSev "d" -Message "Writing list of unique domains to $tmpFolder`domains.txt" -LogFile $runLog -PassThru
            Try {
                $($ReportEvidence.EvaluationResults.Links.Details.ScanTarget | Select-Object -ExpandProperty Domain -Unique) | Set-Content -Path "$tmpFolder`domains.txt"
            } Catch {
                New-PIELogger -logSev "e" -Message "Unable to write to file $tmpFolder`domains.txt" -LogFile $runLog -PassThru
            }
            
            New-PIELogger -logSev "d" -Message "Writing list of unique urls to $tmpFolder`links.txt" -LogFile $runLog -PassThru
            Try {
                $($ReportEvidence.EvaluationResults.Links.Details.ScanTarget | Select-Object -ExpandProperty Url -Unique) | Set-Content -Path "$tmpFolder`links.txt"
            } Catch {
                New-PIELogger -logSev "e" -Message "Unable to write to file $tmpFolder`links.txt" -LogFile $runLog -PassThru
            }
            
            $CountLinks = $($ReportEvidence.EvaluationResults.Links.Details.ScanTarget | Select-Object -ExpandProperty Url -Unique | Measure-Object | Select-Object -ExpandProperty Count)
            New-PIELogger -logSev "i" -Message "Total Unique Links: $countLinks" -LogFile $runLog -PassThru

            $CountDomains = $($ReportEvidence.EvaluationResults.Links.Details.ScanTarget | Select-Object -ExpandProperty Domain -Unique | Measure-Object | Select-Object -ExpandProperty Count)
            New-PIELogger -logSev "i" -Message "Total Unique Domains: $countDomains" -LogFile $runLog -PassThru
            New-PIELogger -logSev "s" -Message "End - Link Summary from Message Body" -LogFile $runLog -PassThru
        }
        
        
        
        if ($ReportEvidence.EvaluationResults.Attachments.Links.Details) { 
            New-PIELogger -logSev "s" -Message "Begin - Link Summary from Attachments" -LogFile $runLog -PassThru
            if (Test-Path "$tmpFolder`domains.txt" -PathType Leaf) {
                Try {
                    $($ReportEvidence.EvaluationResults.Attachments.Links.Details.ScanTarget | Select-Object -ExpandProperty Domain -Unique) | Set-Content -Path "$tmpFolder`domains.txt"
                } Catch {
                    New-PIELogger -logSev "e" -Message "Unable to write to file $tmpFolder`domains.txt" -LogFile $runLog -PassThru
                }
            } else {
                New-PIELogger -logSev "d" -Message "Appending list of unique domains to $tmpFolder`domains.txt" -LogFile $runLog -PassThru
                Try {
                    $($ReportEvidence.EvaluationResults.Attachments.Links.Details.ScanTarget | Select-Object -ExpandProperty Domain -Unique) | Add-Content -Path "$tmpFolder`domains.txt"
                } Catch {
                    New-PIELogger -logSev "e" -Message "Unable to append to file $tmpFolder`domains.txt" -LogFile $runLog -PassThru
                }
            }
            
            New-PIELogger -logSev "d" -Message "Writing list of unique urls to $tmpFolder`links.txt" -LogFile $runLog -PassThru
            if (Test-Path "$tmpFolder`links.txt" -PathType Leaf) {
                Try {
                    $($ReportEvidence.EvaluationResults.Attachments.Links.Details.ScanTarget | Select-Object -ExpandProperty Url -Unique) | Set-Content -Path "$tmpFolder`links.txt"
                } Catch {
                    New-PIELogger -logSev "e" -Message "Unable to write to file $tmpFolder`links.txt" -LogFile $runLog -PassThru
                }
            } else {
                Try {
                    $($ReportEvidence.EvaluationResults.Attachments.Links.Details.ScanTarget | Select-Object -ExpandProperty Url -Unique) | Add-Content -Path "$tmpFolder`links.txt"
                } Catch {
                    New-PIELogger -logSev "e" -Message "Unable to append to file $tmpFolder`links.txt" -LogFile $runLog -PassThru
                }
            }
            
            $CountLinks = $($ReportEvidence.EvaluationResults.Attachments.Links.Details.ScanTarget | Select-Object -ExpandProperty Url -Unique | Measure-Object | Select-Object -ExpandProperty Count)
            New-PIELogger -logSev "i" -Message "Total Unique Links: $countLinks" -LogFile $runLog -PassThru

            $CountDomains = $($ReportEvidence.EvaluationResults.Attachments.Links.Details.ScanTarget | Select-Object -ExpandProperty Domain -Unique | Measure-Object | Select-Object -ExpandProperty Count)
            New-PIELogger -logSev "i" -Message "Total Unique Domains: $countDomains" -LogFile $runLog -PassThru
            New-PIELogger -logSev "s" -Message "End - Link Summary from Attachments" -LogFile $runLog -PassThru
        }
        New-PIELogger -logSev "s" -Message "End - Link Summarization" -LogFile $runLog -PassThru

        #Perform decisioning based upon SwiftonSecurity keywords


            foreach($KeywordRegEx in $phishingKeywords)
            {
                if($ReportEvidence.EvaluationResults.Subject.Original -match $KeywordRegEx) 
                {
                
                $maliciousEmail = "true"
                New-PIELogger -logSev "i" -Message "Subject:  $phishsubject | Matches Keyword: $KeywordRegEx"  -LogFile $runLog -PassThru
                
                }
            }


        # Create a case folder
        New-PIELogger -logSev "s" -Message "Creating Evidence Folder" -LogFile $runLog -PassThru
        $caseID = Get-Date -Format M-d-yyyy_h-m-s
        if ( $ReportEvidence.EvaluationResults.Sender.Contains("@") -eq $true) {
            $spammerName = $ReportEvidence.EvaluationResults.Sender.Split("@")[0]
            $spammerDomain = $ReportEvidence.EvaluationResults.Sender.Split("@")[1]
            New-PIELogger -logSev "d" -Message "Spammer Name: $spammerName Spammer Domain: $spammerDomain" -LogFile $runLog -PassThru
            $caseID = $caseID+"_Sender_"+$spammerName+".at."+$spammerDomain           
        } else {
            New-PIELogger -logSev "d" -Message "Case created as Fwd Message source" -LogFile $runLog -PassThru
            $caseID = $caseID+"_"+$ReportEvidence.EvaluationResults.SenderDisplayName
        }
        try {
            New-PIELogger -logSev "i" -Message "Creating Directory: $caseFolder$caseID" -LogFile $runLog -PassThru
            mkdir $caseFolder$caseID | out-null
        } Catch {
            New-PIELogger -logSev "e" -Message "Unable to create directory: $caseFolder$caseID" -LogFile $runLog -PassThru
        }
        # Support adding Network Share Location to the Case
        $hostname = hostname

        # Copy evidence files to case folder
        New-PIELogger -logSev "i" -Message "Moving interesting files to case folder from $tmpFolder" -LogFile $runLog -PassThru
        Try {
            Copy-Item -Force -Recurse "$tmpFolder*" -Destination $caseFolder$caseID | Out-Null
        } Catch {
            New-PIELogger -logSev "e" -Message "Unable to copy contents from $tmpFolder to $CaseFolder$CaseId" -LogFile $runLog -PassThru
        }
        
        # Cleanup Temporary Folder
        New-PIELogger -logSev "i" -Message "Purging contents from $tmpFolder" -LogFile $runLog -PassThru
        Try {
            Remove-Item "$tmpFolder*" -Force -Recurse | Out-Null
        } Catch {
            New-PIELogger -logSev "e" -Message "Unable to purge contents from $tmpFolder" -LogFile $runLog -PassThru
        }  
        
        # Resume Section - Search
        if ($LrSearchTask) {
            New-PIELogger -logSev "s" -Message "LogRhythm API - Poll Search Status" -LogFile $runLog -PassThru
            New-PIELogger -logSev "i" -Message "LogRhythm Search API - TaskId: $($LrSearchTask.TaskId)" -LogFile $runLog -PassThru
            if ($LrSearchTask.StatusCode -eq 200) {
                do {
                    $SearchStatus = Get-LrSearchResults -TaskId $LrSearchTask.TaskId -PageSize 1000 -PageOrigin 0
                    Start-Sleep 10
                    New-PIELogger -logSev "i" -Message "LogRhythm Search API - TaskId: $($LrSearchTask.TaskId) Status: $($SearchStatus.TaskStatus)" -LogFile $runLog -PassThru
                } until ($SearchStatus.TaskStatus -like "Completed*")
                New-PIELogger -logSev "i" -Message "LogRhythm Search API - TaskId: $($LrSearchTask.TaskId) Status: $($SearchStatus.TaskStatus)" -LogFile $runLog -PassThru
                $ReportEvidence.LogRhythmSearch.TaskId = $LrSearchTask.TaskId
                $LrSearchResults = $SearchStatus
                $ReportEvidence.LogRhythmSearch.Status = $SearchStatus.TaskStatus
            } else {
                New-PIELogger -logSev "s" -Message "LogRhythm Search API - Unable to successfully initiate " -LogFile $runLog -PassThru
                $ReportEvidence.LogRhythmSearch.Status = "Error"
            }

            if ($($ReportEvidence.LogRhythmSearch.Status) -like "Completed*" -and ($($ReportEvidence.LogRhythmSearch.Status) -notlike "Completed: No Results")) {
                $LrSearchResultLogs = $LrSearchResults.Items
                
                # Build summary:
                $ReportEvidence.LogRhythmSearch.Summary.Quantity = $LrSearchResultLogs.count
                $ReportEvidence.LogRhythmSearch.Summary.Recipient = $LrSearchResultLogs | Select-Object -ExpandProperty recipient -Unique
                $ReportEvidence.LogRhythmSearch.Summary.Sender = $LrSearchResultLogs | Select-Object -ExpandProperty sender -Unique
                $ReportEvidence.LogRhythmSearch.Summary.Subject = $LrSearchResultLogs | Select-Object -ExpandProperty subject -Unique
                # Establish Unique Sender & Subject log messages
                $LrSendAndSubject = $LrSearchResultLogs | Where-Object {$_.sender -like $($ReportEvidence.EvaluationResults.Sender) -and $_.subject -like $($ReportEvidence.EvaluationResults.Subject.Original)}
                $ReportEvidence.LogRhythmSearch.Details.SendAndSubject.Quantity = $LrSendAndSubject.count
                $ReportEvidence.LogRhythmSearch.Details.SendAndSubject.Recipients = $LrSendAndSubject | Select-Object -ExpandProperty recipient -Unique
                $ReportEvidence.LogRhythmSearch.Details.SendAndSubject.Subject = $LrSendAndSubject | Select-Object -ExpandProperty subject -Unique
                $ReportEvidence.LogRhythmSearch.Details.SendAndSubject.Sender = $LrSendAndSubject | Select-Object -ExpandProperty sender -Unique
                # Establish Unique Sender log messages
                $LrSender = $LrSearchResultLogs | Where-Object {$_.sender -like $($ReportEvidence.EvaluationResults.Sender) -and $_ -notcontains $LrSendAndSubject}
                $ReportEvidence.LogRhythmSearch.Details.Sender.Quantity = $LrSender.count
                $ReportEvidence.LogRhythmSearch.Details.Sender.Recipients = $LrSender | Select-Object -ExpandProperty recipient -Unique
                $ReportEvidence.LogRhythmSearch.Details.Sender.Subjects = $LrSender | Select-Object -ExpandProperty subject -Unique
                $ReportEvidence.LogRhythmSearch.Details.Sender.Sender = $LrSender | Select-Object -ExpandProperty sender -Unique
                # Establish Unique Subject log messages
                $LrSubject = $LrSearchResultLogs | Where-Object {$_.subject -like $($ReportEvidence.EvaluationResults.Subject.Original) -and $_ -notcontains $LrSender -and $_ -notcontains $LrSendAndSubject}
                $ReportEvidence.LogRhythmSearch.Details.Subject.Quantity = $LrSubject.count
                $ReportEvidence.LogRhythmSearch.Details.Subject.Recipients = $LrSubject | Select-Object -ExpandProperty recipient -Unique
                $ReportEvidence.LogRhythmSearch.Details.Subject.Subject = $LrSubject | Select-Object -ExpandProperty subject -Unique
                $ReportEvidence.LogRhythmSearch.Details.Subject.Senders = $LrSubject | Select-Object -ExpandProperty sender -Unique
            }
            New-PIELogger -logSev "s" -Message "LogRhythm API - End Log Search" -LogFile $runLog -PassThru
        }
        # End Section - Search


        # Conclude runtime metrics
        $EndTime = (get-date).ToUniversalTime()
        New-PIELogger -logSev "i" -Message "Processing End Time: $($EndTime)" -LogFile $runLog -PassThru
        $ReportEvidence.Meta.Metrics.End = $EndTime.ToString("yyyy-MM-ddTHHmmssffZ")
        $Duration = New-Timespan -Start $StartTime -End $EndTime
        $ReportEvidence.Meta.Metrics.Duration = $Duration.ToString("%m\.%s\.%f")
        
        # Create Summary Notes for Case Output
        $CaseSummaryNote = Format-PIECaseSummary -ReportEvidence $ReportEvidence
        $CaseEvidenceSummaryNote = Format-PIEEvidenceSummary -EvaluationResults $ReportEvidence.EvaluationResults

        if ($ReportEvidence.EvaluationResults.Headers.Details) {
            $CaseEvidenceHeaderSummary = Format-PieHeaderSummary -ReportEvidence $ReportEvidence
        }

        if ($LrCaseOutput) {
            New-PIELogger -logSev "s" -Message "LogRhythm API - Create Case" -LogFile $runLog -PassThru
            if ( $ReportEvidence.EvaluationResults.Sender.Contains("@") -eq $true) {

            if ($maliciousEmail -eq "true" ) {

                New-PIELogger -logSev "i" -Message "DECISION - Verified Phish - LogRhythm API - Create Case with Sender Info" -LogFile $runLog -PassThru
                $caseSummary = "Verified phishing email from $($ReportEvidence.EvaluationResults.Sender) was reported on $($ReportEvidence.ReportSubmission.UtcDate) UTC by $($ReportEvidence.ReportSubmission.Sender). The subject of the email is ($($ReportEvidence.EvaluationResults.Subject.Original))."
                $CaseDetails = New-LrCase -Name "Verified Phish : $spammerName [at] $spammerDomain" -Priority 2 -Summary $caseSummary -PassThru

                }

            if ($maliciousEmail -eq "false" ) {

                New-PIELogger -logSev "i" -Message "DECISION - Possible Phish - LogRhythm API - Create Case with Sender Info" -LogFile $runLog -PassThru
                $caseSummary = "Potential phishing email from $($ReportEvidence.EvaluationResults.Sender) was reported on $($ReportEvidence.ReportSubmission.UtcDate) UTC by $($ReportEvidence.ReportSubmission.Sender). The subject of the email is ($($ReportEvidence.EvaluationResults.Subject.Original))."
                $CaseDetails = New-LrCase -Name "Potential Phish : $spammerName [at] $spammerDomain" -Priority 3 -Summary $caseSummary -PassThru

                }

            } else {
                New-PIELogger -logSev "i" -Message "NO DECISION - LogRhythm API - Create Case without Sender Info" -LogFile $runLog -PassThru
                $caseSummary = "Phishing email was reported on $($ReportEvidence.ReportSubmission.UtcDate) UTC by $($ReportEvidence.ReportSubmission.Sender). The subject of the email is ($($ReportEvidence.EvaluationResults.Subject.Original))."
                $CaseDetails = New-LrCase -Name "Phishing Message Reported" -Priority 3 -Summary $caseSummary -PassThru
                
            }
            Start-Sleep .2

            # Set ReportEvidence CaseNumber
            $ReportEvidence.LogRhythmCase.Number = $CaseDetails.number

            Try {
                $ReportEvidence.LogRhythmCase.Number | Out-File "$caseFolder$caseID\lr_casenumber.txt"
            } Catch {
                New-PIELogger -logSev "e" -Message "Unable to move $pieFolder\plugins\lr_casenumber.txt to $caseFolder$caseID\" -LogFile $runLog -PassThru
            }
            
            # Establish Case URL to ReportEvidence Object
            $ReportEvidence.LogRhythmCase.Url = "https://$LogRhythmHost/cases/$($ReportEvidence.LogRhythmCase.Number)"
            New-PIELogger -logSev "i" -Message "Case URL: $($ReportEvidence.LogRhythmCase.Url)" -LogFile $runLog -PassThru

            # Update case Earliest Evidence
            if ($ReportEvidence.EvaluationResults.UtcDate) {
                # Based on recipient's e-mail message recieve timestamp from origin sender
                #[datetime] $EvidenceTimestamp = [datetime]::parseexact($ReportEvidence.EvaluationResults.UtcDate, "yyyy-MM-ddTHHmmssffZ", $null)
                [datetime] $EvidenceTimestamp = $ReportEvidence.EvaluationResults.UtcDate
                Update-LrCaseEarliestEvidence -Id $($ReportEvidence.LogRhythmCase.Number) -Timestamp $EvidenceTimestamp
            } else {
                # Based on report submission for evaluation
                try { [datetime] $EvidenceTimestamp = [datetime]::parseexact($ReportEvidence.ReportSubmission.UtcDate, "yyyy-MM-ddTHHmmssffZ", $null) }
                catch { $EvidenceTimestamp = [datetime]::parseexact($ReportEvidence.ReportSubmission.UtcDate, "yyyy-MM-ddTHH:mm:ssZ", $null) }
                Update-LrCaseEarliestEvidence -Id $($ReportEvidence.LogRhythmCase.Number) -Timestamp $EvidenceTimestamp
            }


            # Tag the case
            if ( $LrCaseTags ) {
                New-PIELogger -logSev "i" -Message "LogRhythm API - Applying case tags" -LogFile $runLog -PassThru
                ForEach ($LrTag in $LrCaseTags) {
                    $TagStatus = Get-LrTags -Name $LrTag -Exact
                    Start-Sleep 0.2
                    if (!$TagStatus) {
                            $TagStatus = New-LrTag -Tag $LrTag -PassThru
                            Start-Sleep 0.2
                    }
                    if ($TagStatus) {
                            Add-LrCaseTags -Id $ReportEvidence.LogRhythmCase.Number -Tags $TagStatus.Number
                            New-PIELogger -logSev "i" -Message "LogRhythm API - Adding tag $LrTag Tag Number $($TagStatus.number)" -LogFile $runLog -PassThru
                            Start-Sleep 0.2
                    }
                }
                New-PIELogger -logSev "i" -Message "LogRhythm API - End applying case tags" -LogFile $runLog -PassThru
            }

            # Adding and assigning other users
            
            if ( $LrCaseCollaborators ) {
                New-PIELogger -logSev "s" -Message "Begin - LogRhythm Case Collaborators Block" -LogFile $runLog -PassThru
                ForEach ($LrCaseCollaborator in $LrCaseCollaborators) {
                    $LrCollabortorStatus = Get-LrUsers -Name $LrPlaybook -Exact
                    if ($LrCollabortorStatus) {
                        New-PIELogger -logSev "i" -Message "LogRhythm API - Adding Collaborator:$LrCaseCollaborator to Case:$($ReportEvidence.LogRhythmCase.Number)" -LogFile $runLog -PassThru
                        Add-LrCaseCollaborators -Id $ReportEvidence.LogRhythmCase.Number -Names $LrCaseCollaborator
                    } else {
                        New-PIELogger -logSev "e" -Message "LogRhythm API - Collaborator:$LrCaseCollaborator not found or not accessible due to permissions." -LogFile $runLog -PassThru
                    }
                }        
                New-PIELogger -logSev "s" -Message "End - LogRhythm Case Collaborators Block" -LogFile $runLog -PassThru      
            } else {
                New-PIELogger -logSev "d" -Message "LogRhythm API - Collaborators Omision - Collaborators not defined" -LogFile $runLog -PassThru
            }

            # Add case playbook if playbook has been defined.
            if ($LrCasePlaybook) {
                New-PIELogger -logSev "s" -Message "Begin - LogRhythm Playbook Block" -LogFile $runLog -PassThru
                ForEach ($LrPlaybook in $LrCasePlaybook) {
                    $LrPlayBookStatus = Get-LrPlaybooks -Name $LrPlaybook -Exact
                    if ($LrPlayBookStatus.Code -eq 404) {
                        New-PIELogger -logSev "e" -Message "LogRhythm API - Playbook:$LrPlaybook not found or not accessible due to permissions." -LogFile $runLog -PassThru
                    } else {
                        New-PIELogger -logSev "i" -Message "LogRhythm API - Adding Playbook:$LrPlaybook to Case:$($ReportEvidence.LogRhythmCase.Number)" -LogFile $runLog -PassThru
                        $AddLrPlaybook = Add-LrCasePlaybook -Id $ReportEvidence.LogRhythmCase.Number -Playbook $LrPlaybook
                        if ($AddLrPlaybook) {
                            New-PIELogger -logSev "e" -Message "LogRhythm API - Playbook:$LrPlaybook Error:$($AddLrPlaybook.Note)" -LogFile $runLog -PassThru
                        }
                    }
                }
                New-PIELogger -logSev "s" -Message "End - LogRhythm Playbook Block" -LogFile $runLog -PassThru
            } else {
                New-PIELogger -logSev "d" -Message "LogRhythm API - Playbook Omision - Playbooks not defined" -LogFile $runLog -PassThru
            }


            # Add Link plugin output to Case
            ForEach ($UrlDetails in $ReportEvidence.EvaluationResults.Links.Details) {
                if ($shodan) {
                    if ($UrlDetails.Plugins.Shodan) {
                        $CasePluginShodanNote = $UrlDetails.Plugins.Shodan | Format-ShodanTextOutput
                        Add-LrNoteToCase -id $ReportEvidence.LogRhythmCase.Number -Text $($CasePluginShodanNote).subString(0, [System.Math]::Min(20000, $CasePluginShodanNote.Length))
                    }
                }
                if ($urlscan) {
                    if ($UrlDetails.Plugins.urlscan) {
                        $CasePluginUrlScanNote = $UrlDetails.Plugins.urlscan | Format-UrlscanTextOutput -Type "summary"
                        Add-LrNoteToCase -id $ReportEvidence.LogRhythmCase.Number -Text $($CasePluginUrlScanNote).subString(0, [System.Math]::Min(20000, $CasePluginUrlScanNote.Length))
                    }
                }
                if ($virusTotal) {
                    if ($UrlDetails.Plugins.VirusTotal) {
                        $CasePluginVTNote = $UrlDetails.Plugins.VirusTotal  | Format-VTTextOutput 
                        Add-LrNoteToCase -id $ReportEvidence.LogRhythmCase.Number -Text $($CasePluginVTNote).subString(0, [System.Math]::Min(20000, $CasePluginVTNote.Length))
                    }
                }
            }

            # Add Attachment plugin output to Case
            ForEach ($AttachmentDetails in $ReportEvidence.EvaluationResults.Attachments) {
                if ($virusTotal) {
                    if ($AttachmentDetails.Plugins.VirusTotal.Status) {
                        $CasePluginVTNote = $AttachmentDetails.Plugins.VirusTotal.Results | Format-VTTextOutput 
                        Add-LrNoteToCase -id $ReportEvidence.LogRhythmCase.Number -Text $($CasePluginVTNote).subString(0, [System.Math]::Min(20000, $CasePluginVTNote.Length))
                    }
                }
                # Add Link plugin output to Case
                ForEach ($AttachmentUrlDetails in $AttachmentDetails.Links.Details) {
                    if ($shodan) {
                        if ($AttachmentUrlDetails.Plugins.Shodan) {
                            $CasePluginShodanNote = $AttachmentUrlDetails.Plugins.Shodan | Format-ShodanTextOutput
                            Add-LrNoteToCase -id $ReportEvidence.LogRhythmCase.Number -Text $($CasePluginShodanNote).subString(0, [System.Math]::Min(20000, $CasePluginShodanNote.Length))
                        }
                    }
                    if ($urlscan) {
                        if ($AttachmentUrlDetails.Plugins.urlscan) {
                            $CasePluginUrlScanNote = $AttachmentUrlDetails.Plugins.urlscan | Format-UrlscanTextOutput -Type "summary"
                            Add-LrNoteToCase -id $ReportEvidence.LogRhythmCase.Number -Text $($CasePluginUrlScanNote).subString(0, [System.Math]::Min(20000, $CasePluginUrlScanNote.Length))
                        }
                    }
                    if ($virusTotal) {
                        if ($AttachmentUrlDetails.Plugins.VirusTotal) {
                            $CasePluginVTNote = $AttachmentUrlDetails.Plugins.VirusTotal  | Format-VTTextOutput 
                            Add-LrNoteToCase -id $ReportEvidence.LogRhythmCase.Number -Text $($CasePluginVTNote).subString(0, [System.Math]::Min(20000, $CasePluginVTNote.Length))
                        }
                    }
                }
            }

            # Copy E-mail Message text body to case
            New-PIELogger -logSev "i" -Message "LogRhythm API - Copying e-mail body text to case" -LogFile $runLog -PassThru
            if ( $ReportEvidence.EvaluationResults.Body.Original ) {
                $DefangBody = $ReportEvidence.EvaluationResults.Body.Original.subString(0, [System.Math]::Min(19900, $ReportEvidence.EvaluationResults.Body.Original.Length)).Replace('<http','<hxxp')
                $NoteStatus = Add-LrNoteToCase -Id $ReportEvidence.LogRhythmCase.Number -Text "=== Reported Message Body ===`r`n--- BEGIN ---`r`n$DefangBody`r`n--- END ---" -PassThru
                if ($NoteStatus.Error) {
                    New-PIELogger -logSev "e" -Message "LogRhythm API - Unable to add ReportEvidence.EvaluationResults.Body to LogRhythm Case." -LogFile $runLog -PassThru
                    New-PIELogger -logSev "d" -Message "LogRhythm API - Code: $($NoteStatus.Error.Code) Note: $($NoteStatus.Error.Note)" -LogFile $runLog -PassThru
                }
            }

            if ($CaseEvidenceHeaderSummary) {
                New-PIELogger -logSev "i" -Message "LogRhythm API - Copying e-mail header details summary to case" -LogFile $runLog -PassThru
                $NoteStatus = Add-LrNoteToCase -Id $ReportEvidence.LogRhythmCase.Number -Text $CaseEvidenceHeaderSummary.Substring(0,[System.Math]::Min(20000, $CaseEvidenceHeaderSummary.Length)) -PassThru
                if ($NoteStatus.Error) {
                    New-PIELogger -logSev "e" -Message "LogRhythm API - Unable to add CaseEvidenceHeaderSummary to LogRhythm Case." -LogFile $runLog -PassThru
                    New-PIELogger -logSev "d" -Message "LogRhythm API - Code: $($NoteStatus.Error.Code) Note: $($NoteStatus.Error.Note)" -LogFile $runLog -PassThru
                }
            }

            # Search note
            if ($ReportEvidence.LogRhythmSearch.Summary.Quantity -ge 1) {
                New-PIELogger -logSev "i" -Message "LogRhythm API - Add LogRhythm Search summary note to case" -LogFile $runLog -PassThru
                $LrSearchSummary = $(Format-PIESearchSummary -ReportEvidence $ReportEvidence)
                $NoteStatus = Add-LrNoteToCase -Id $ReportEvidence.LogRhythmCase.Number -Text $LrSearchSummary.Substring(0,[System.Math]::Min(20000, $LrSearchSummary.Length)) -PassThru
                if ($NoteStatus.Error) {
                    New-PIELogger -logSev "e" -Message "LogRhythm API - Unable to add LogRhythm Search summary note to case." -LogFile $runLog -PassThru
                    New-PIELogger -logSev "d" -Message "LogRhythm API - Code: $($NoteStatus.Error.Code) Note: $($NoteStatus.Error.Note)" -LogFile $runLog -PassThru
                }
            }

            # Add Link/Attachment Summary as second Case note
            if ($CaseEvidenceSummaryNote) {
                $NoteStatus = Add-LrNoteToCase -Id $ReportEvidence.LogRhythmCase.Number -Text $CaseEvidenceSummaryNote.Substring(0,[System.Math]::Min(20000, $CaseEvidenceSummaryNote.Length)) -PassThru
                if ($NoteStatus.Error) {
                    New-PIELogger -logSev "e" -Message "LogRhythm API - Unable to add CaseEvidenceSummaryNote to LogRhythm Case." -LogFile $runLog -PassThru
                    New-PIELogger -logSev "d" -Message "LogRhythm API - Code: $($NoteStatus.Error.Code) Note: $($NoteStatus.Error.Note)" -LogFile $runLog -PassThru
                }
            }


            # Add overall summary as last, top most case note.
            Add-LrNoteToCase -Id $ReportEvidence.LogRhythmCase.Number -Text $CaseSummaryNote.Substring(0,[System.Math]::Min(20000, $CaseSummaryNote.Length))

            # If we have Log Results, add these to the case.
            if (($($ReportEvidence.LogRhythmSearch.Status) -like "Completed*") -and ($($ReportEvidence.LogRhythmSearch.Status) -notlike "Completed: No Results")) {
                Add-LrLogsToCase -Id $ReportEvidence.LogRhythmCase.Number -Note "Message trace matching sender or subject of the submitted e-mail message." -IndexId $($ReportEvidence.LogRhythmSearch.TaskId)
            }

            $ReportEvidence.LogRhythmCase.Details = Get-LrCaseById -Id $ReportEvidence.LogRhythmCase.Number
            # End Section - LogRhythm Case Output
        }



# ================================================================================
# Case Closeout
# ================================================================================

        # Write PIE Report Json object out to Case as Evidence
        New-PIELogger -logSev "i" -Message "Case Json - Writing ReportEvidence to $($caseFolder)$($caseId)\Case_Report.json" -LogFile $runLog -PassThru
        Try {
            $ReportEvidence | ConvertTo-Json -Depth 50 | Out-File -FilePath "$caseFolder$caseID\Case_Report.json"
        } Catch {
            New-PIELogger -logSev "e" -Message "Case Json - Unable to write ReportEvidence to $($caseFolder)$($caseId)\Case_Report.json" -LogFile $runLog -PassThru
        }

        if ($LrSearchResultLogs) {
            New-PIELogger -logSev "i" -Message "Case Logs - Writing Search Result logs to $($caseFolder)$($caseId)\Case_Logs.csv" -LogFile $runLog -PassThru
            Try {
                $LrSearchResultLogs | Export-Csv -Path "$caseFolder$caseID\Case_Logs.csv" -Force -NoTypeInformation
            } Catch {
                New-PIELogger -logSev "e" -Message "Case Logs - Unable to write Search Result logs to $($caseFolder)$($caseId)\Case_Logs.csv" -LogFile $runLog -PassThru
            }
        }
         
        # Write TXT Report as Evidence
        New-PIELogger -logSev "s" -Message "Case File - Begin - Writing details to Case File." -LogFile $runLog -PassThru
        New-PIELogger -logSev "i" -Message "Case File - Writing to $($caseFolder)$($caseId)\Case_Report.txt" -LogFile $runLog -PassThru
        $CaseFile = "\Case_Report.txt"
        Try {
            $CaseSummaryNote | Out-File -FilePath $caseFolder$caseID$CaseFile
        } Catch {
            New-PIELogger -logSev "e" -Message "Case File - Unable to write to $($caseFolder)$($caseId)\Case_Report.txt" -LogFile $runLog -PassThru
        }
        $CaseEvidenceSummaryNote | Out-File -FilePath $caseFolder$caseID$CaseFile -Append

        if ($CaseEvidenceHeaderSummary) {
            New-PIELogger -logSev "i" -Message "Case File - Copying e-mail header details summary to Case File" -LogFile $runLog -PassThru
            $CaseEvidenceHeaderSummary | Out-File -FilePath $caseFolder$caseID$CaseFile -Append
        }

        New-PIELogger -logSev "i" -Message "Case File - Appending URL Details to Case File." -LogFile $runLog -PassThru
        $EvidenceSeperator = "-----------------------------------------------`r`n"
        # Add Link plugin output to TXT Case
        ForEach ($UrlDetails in $ReportEvidence.EvaluationResults.Links.Details) {
            if ($UrlDetails.Plugins.Shodan) {
                $EvidenceSeperator  | Out-File -FilePath $caseFolder$caseID$CaseFile -Append
                $($UrlDetails.Plugins.Shodan | Format-ShodanTextOutput) | Out-File -FilePath $caseFolder$caseID$CaseFile -Append
            }
            if ($UrlDetails.Plugins.urlscan) {
                $EvidenceSeperator  | Out-File -FilePath $caseFolder$caseID$CaseFile -Append
                $($UrlDetails.Plugins.urlscan | Format-UrlscanTextOutput)  | Out-File -FilePath $caseFolder$caseID$CaseFile -Append
            }
            if ($UrlDetails.Plugins.VirusTotal) {
                $EvidenceSeperator  | Out-File -FilePath $caseFolder$caseID$CaseFile -Append
                $($UrlDetails.Plugins.VirusTotal | Format-VTTextOutput)  | Out-File -FilePath $caseFolder$caseID$CaseFile -Append
            }
        }

        # Add Attachment Link plugin output to TXT Case
        if ($ReportEvidence.EvaluationResults.Attachments) {
            # Add Attachment plugin output to Case
            New-PIELogger -logSev "i" -Message "Case File - Appending Attachment Details to Case File." -LogFile $runLog -PassThru
            ForEach ($AttachmentDetails in $ReportEvidence.EvaluationResults.Attachments) {
                    $EvidenceSeperator  | Out-File -FilePath $caseFolder$caseID$CaseFile -Append
                    if ($AttachmentDetails.Plugins.VirusTotal.Status) {
                        $($AttachmentDetails.Plugins.VirusTotal.Results | Format-VTTextOutput)  | Out-File -FilePath $caseFolder$caseID$CaseFile -Append
                    }
            }


            ForEach ($AttachmentUrlDetails in $ReportEvidence.EvaluationResults.Attachments.Links.Details) {
                New-PIELogger -logSev "i" -Message "Case File - Appending Embedded URL details from Attachment to Case File." -LogFile $runLog -PassThru
                if ($AttachmentUrlDetails.Plugins.Shodan) {
                    $EvidenceSeperator  | Out-File -FilePath $caseFolder$caseID$CaseFile -Append
                    $($AttachmentUrlDetails.Plugins.Shodan | Format-ShodanTextOutput) | Out-File -FilePath $caseFolder$caseID$CaseFile -Append
                }
                if ($AttachmentUrlDetails.Plugins.urlscan) {
                    $EvidenceSeperator  | Out-File -FilePath $caseFolder$caseID$CaseFile -Append
                    $($AttachmentUrlDetails.Plugins.urlscan | Format-UrlscanTextOutput)  | Out-File -FilePath $caseFolder$caseID$CaseFile -Append
                }
                if ($AttachmentUrlDetails.Plugins.VirusTotal) {
                    $EvidenceSeperator  | Out-File -FilePath $caseFolder$caseID$CaseFile -Append
                    $($AttachmentUrlDetails.Plugins.VirusTotal | Format-VTTextOutput)  | Out-File -FilePath $caseFolder$caseID$CaseFile -Append
                }
            }
        }
        New-PIELogger -logSev "s" -Message "Case File - End - Writing details to Case File." -LogFile $runLog -PassThru

        New-PIELogger -logSev "i" -Message "Phish Log - Writing details to Phish Log." -LogFile $runLog -PassThru
        Try {
            $PhishLogContent = "$($ReportEvidence.Meta.GUID),$($ReportEvidence.Meta.Timestamp),$($ReportEvidence.ReportSubmission.MessageId),$($ReportEvidence.EvaluationResults.Sender),$($ReportEvidence.ReportSubmission.Sender),$($ReportEvidence.ReportSubmission.Subject.Original)"
            $PhishLogContent | Out-File -FilePath $PhishLog -Append -Encoding ascii
        } Catch {
            New-PIELogger -logSev "e" -Message "Phish Log - Unable to write to $PhishLog" -LogFile $runLog -PassThru
            echo $PhishLogContent
        #}
    #}

    #Cleanup Variables prior to next evaluation
    New-PIELogger -logSev "i" -Message "Resetting analysis varaiables" -LogFile $runLog -PassThru

    $attachmentFull = $null
    $attachment = $null
    $attachments = $null
    $caseID = $null
    $maliciousEmail = $null
    $isazip = $null
    Clear-Variable Attach*
    New-PIELogger -logSev "i" -Message "End - Processing for GUID: $($ReportEvidence.Meta.Guid)" -LogFile $runLog -PassThru
}
}
}
}
##############################
##End of main detection loop##
##############################

# Move items from inbox to target folders
New-PIELogger -logSev "s" -Message "Begin - Mailbox Cleanup" -LogFile $runLog -PassThru


if ($InboxMailIDs){
#First do the skipped...
 foreach ($SkippedItem in $FolderDestSkipped) {
           

            $GraphSkippedBody = @{ 
                "destinationId" = "$InboxSkipped" } | ConvertTo-Json

            $GraphSkippedMoveRequest = "https://graph.microsoft.com/v1.0/users/$SocMailbox/messages/$SkippedItem/move" 

            $GraphSkippedMove = (Invoke-RestMethod -Headers @{Authorization = "Bearer $($token)"} -ContentType "application/json" -Uri $GraphSkippedMoveRequest -Body $GraphSkippedBody -Method Post)
            }


#Then do the rest

$CompletedMessages = @($InboxMailIDs | Where-Object { $FolderDestSkipped -notcontains $_ })

#Add CompletedMessages
 foreach ($CompMsg in $CompletedMessages) {
           

            $GraphCompletedBody = @{ 
                "destinationId" = "$InboxCompleted" } | ConvertTo-Json

            $GraphCompletedMoveRequest = "https://graph.microsoft.com/v1.0/users/$SocMailbox/messages/$CompMsg/move" 

            $GraphCompletedMove = (Invoke-RestMethod -Headers @{Authorization = "Bearer $($token)"} -ContentType "application/json" -Uri $GraphCompletedMoveRequest -Body $GraphCompletedBody -Method Post)
            }

}
else
{ 
New-PIELogger -logSev "s" -Message "Nothing to cleanup" -LogFile $runLog -PassThru
}

New-PIELogger -logSev "s" -Message "End - Mailbox Cleanup" -LogFile $runLog -PassThru

# ================================================================================
# LOG ROTATION
# ================================================================================

# Log rotation script stolen from:
#      https://gallery.technet.microsoft.com/scriptcenter/PowerShell-Script-to-Roll-a96ec7d4

function Reset-Log { 
    #function checks to see if file in question is larger than the paramater specified if it is it will roll a log and delete the oldes log if there are more than x logs. 
    param(
        [string]$fileName, 
        [int64]$filesize = 1mb, 
        [int] $logcount = 5) 
     
    $logRollStatus = $true 
    if(test-path $filename) { 
        $file = Get-ChildItem $filename 
        #this starts the log roll 
        if((($file).length) -ige $filesize) { 
            $fileDir = $file.Directory 
            $fn = $file.name #this gets the name of the file we started with 
            $files = Get-ChildItem $filedir | ?{$_.name -like "$fn*"} | Sort-Object lastwritetime 
            $filefullname = $file.fullname #this gets the fullname of the file we started with 
            #$logcount +=1 #add one to the count as the base file is one more than the count 
            for ($i5 = ($files.count); $i5 -gt 0; $i5--) {  
                #[int]$fileNumber = ($f).name.Trim($file.name) #gets the current number of the file we are on 
                $files = Get-ChildItem $filedir | Where-Object {$_.name -like "$fn*"} | Sort-Object lastwritetime 
                $operatingFile = $files | Where-Object {($_.name).trim($fn) -eq $i5} 
                if ($operatingfile) {
                    $operatingFilenumber = ($files | Where-Object {($_.name).trim($fn) -eq $i}).name.trim($fn)
                } else {
                    $operatingFilenumber = $null
                } 
                
                if (($null -eq $operatingFilenumber) -and ($i5 -ne 1) -and ($i5 -lt $logcount)) { 
                    $operatingFilenumber = $i5 
                    $newfilename = "$filefullname.$operatingFilenumber" 
                    $operatingFile = $files | ?{($_.name).trim($fn) -eq ($i5-1)} 
                    write-host "moving to $newfilename" 
                    move-item ($operatingFile.FullName) -Destination $newfilename -Force 
                } elseif($i5 -ge $logcount) { 
                    if($null -eq $operatingFilenumber) {  
                        $operatingFilenumber = $i5 - 1 
                        $operatingFile = $files | Where-Object {($_.name).trim($fn) -eq $operatingFilenumber} 
                    } 
                    write-host "deleting " ($operatingFile.FullName) 
                    remove-item ($operatingFile.FullName) -Force 
                } elseif($i5 -eq 1) { 
                    $operatingFilenumber = 1 
                    $newfilename = "$filefullname.$operatingFilenumber" 
                    write-host "moving to $newfilename" 
                    move-item $filefullname -Destination $newfilename -Force 
                } else { 
                    $operatingFilenumber = $i5 +1  
                    $newfilename = "$filefullname.$operatingFilenumber" 
                    $operatingFile = $files | ?{($_.name).trim($fn) -eq ($i5-1)} 
                    write-host "moving to $newfilename" 
                    move-item ($operatingFile.FullName) -Destination $newfilename -Force    
                } 
            }    
        } else { 
            $logRollStatus = $false
        }
    } else { 
        $logrollStatus = $false 
    }
    $LogRollStatus 
}

New-PIELogger -logSev "s" -Message "Begin Reset-Log block" -LogFile $runLog -PassThru
Reset-Log -fileName $phishLog -filesize 25mb -logcount 10 | Out-Null
Reset-Log -fileName $runLog -filesize 50mb -logcount 10 | Out-Null
New-PIELogger -logSev "s" -Message "End Reset-Log block" -LogFile $runLog -PassThru
New-PIELogger -logSev "i" -Message "Close mailbox connection" -LogFile $runLog -PassThru
# Kill Office365 Session and Clear Variables
#$MailClient.Disconnect($true)
New-PIELogger -logSev "s" -Message "PIE Execution Completed" -LogFile $runLog -PassThru
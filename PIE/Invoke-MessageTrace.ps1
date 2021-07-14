# Date Variables
$date = Get-Date
$oldAF = (Get-Date).AddDays(-10)
$96Hours = (Get-Date).AddHours(-96)
$48Hours = (Get-Date).AddHours(-48)
$24Hours = (Get-Date).AddHours(-24)
$inceptionDate = (Get-Date).AddMinutes(-16)
$phishDate = (Get-Date).AddMinutes(-31)
$day = Get-Date -Format MM-dd-yyyy

# Folder Structure
$pieFolder = "C:\PIE-main\PIE"
$traceLog = "$pieFolder\logs\ongoing-trace-log.csv"
$phishLog = "$pieFolder\logs\ongoing-phish-log.csv"
$spamTraceLog = "$pieFolder\logs\ongoing-outgoing-spam-log.csv"
$analysisLog = "$pieFolder\logs\analysis.csv"
$lastLogDateFile = "$pieFolder\logs\last-log-date.txt"
$tmpLog = "$pieFolder\logs\tmp.csv"
$caseFolder = "$pieFolder\cases\"
$tmpFolder = "$pieFolder\tmp\"
$confFolder = "$pieFolder\conf\"
$runLog = "$pieFolder\logs\pierun.txt"
$log = $true
try {
    $lastLogDate = [DateTime]::SpecifyKind((Get-Content -Path $lastLogDateFile),'Utc')
}
catch {
    $lastLogDate = $inceptionDate
}

 
 # Timestamp Function
function Get-TimeStamp {
    return "[{0:MM/dd/yy} {0:HH:mm:ss}]" -f (Get-Date)
}


 ================================================================================
# Office 365 API Authentication
# ================================================================================

if ( $EncodedXMLCredentials ) {
    try {
        $cred = Import-Clixml -Path $CredentialsFile
        $Username = $cred.Username
        $Password = $cred.GetNetworkCredential().Password
    } catch {
        Write-Error ("Could not find credentials file: " + $CredentialsFile)
        Write-Output "$(Get-TimeStamp) ERROR - Could not find credentials file: $CredentialsFile" | Out-File $runLog -Append
        Break;
    }
}

try {
    if (-Not ($password)) {
        $cred = Get-Credential
    } Else {
        $securePass = ConvertTo-SecureString -string $password -AsPlainText -Force
        $cred = New-Object -typename System.Management.Automation.PSCredential -argumentlist $username, $securePass
    }

    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $cred -Authentication Basic -AllowRedirection
    Import-PSSession $Session -AllowClobber
    Write-Output "$(Get-TimeStamp) INFO - Open Office 365 connection" | Out-File $runLog -Append
} Catch {
    Write-Error "Access Denied..."
    Write-Output "$(Get-TimeStamp) ERROR - Office 365 connection Access Denied" | Out-File $runLog -Append
    Break;
}




    
     # scrape all mail - ongiong log generation
    # new scrape mail - by sslawter - LR Community
    Write-Output "$(Get-TimeStamp) STATUS - Begin processing messageTrace" | Out-File $runLog -Append
    foreach ($page in 1..1000) {
        $messageTrace = Get-MessageTrace -StartDate $lastlogDate -EndDate $date -Page $page | Select MessageTraceID,Received,*Address,*IP,Subject,Status,Size,MessageID
        if ($messageTrace.Count -ne 0) {
            $messageTraces += $messageTrace
            Write-Verbose "Page #: $page"
            Write-Output "$(Get-TimeStamp) INFO - Processing page $page" | Out-File $runLog -Append
        }
        else {
            break
        }
    }
    
    $messageTracesSorted = $messageTraces | Sort-Object Received
    $messageTracesSorted | Export-Csv $traceLog -NoTypeInformation -Append
    ($messageTracesSorted | Select-Object -Last 1).Received.GetDateTimeFormats("O") | Out-File -FilePath $lastLogDateFile -Force -NoNewline
    Write-Output "$(Get-TimeStamp) STATUS - Completed messageTrace" | Out-File $runLog -Append

    $traceSize = Get-Item $traceLog
if ($traceSize.Length -gt 49MB ) {
    Start-Sleep -Seconds 30
    Reset-Log -fileName $traceLog -filesize 50mb -logcount 10
}


# Kill Office365 Session and Clear Variables
Remove-PSSession $Session
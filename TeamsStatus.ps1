# Change these values
$KasaBulbIP = '192.168.12.34'
$SlackKey = 'xoxb-Numbers-Numbers-MoreNumbers-LongHexadecimal'
$SlackUserID = 'U03QXXXXX'

$Time = Get-Date

# Local Teams log FTW
$StatusLog = Get-Content -Path "$env:APPDATA\Microsoft\Teams\logs.txt" -Tail 1000 |
    Select-String -Pattern 'StatusIndicatorStateService: Added' |
    Select-Object -Last 1

# RegEx all the things
[void]($StatusLog -match '(.*?)\sGMT.*StatusIndicatorStateService:\sAdded\s(.*)\s\(.*:\s(.*)\s-\>\s(.*)\)')
$LogTime = Get-Date $Matches[1]
$TeamsStatus = $Matches[2]
$StatusOld = $Matches[3]
$StatusNew = $Matches[4]

# Quit if there is no new log entry
# "NewActivity" overwrites OnThePhone/Presenting, but could happen in or out of a call
  # Ignore NewActivity and continue using the prior status instead
If ($Time.AddMinutes(-1) -gt $LogTime -or $TeamsStatus -eq 'NewActivity') {
    exit
}

If ($TeamsStatus -match 'OnThePhone|Presenting') {
    # call python externally to use python-kasa to change the light
    kasa --bulb --host $KasaBulbIP hsv 0 70 100

    $SlackStatus = @{
        # slack status and message to set
        status_text  = 'Teams call'
        status_emoji = ':no_entry:'
    }
} Else {
    # set bulb back to normal
    kasa --bulb --host $KasaBulbIP hsv 0 0 100
    
    $SlackStatus = @{
        # slack status will be cleared
        status_text  = ''
        status_emoji = ''
    }
}

# Log the intended change
"$($Time.ToString('s'));$($LogTime.ToString('s'));$StatusOld;$StatusNew" | Out-File $PSScriptRoot\log.txt -Append

# Build Slack API call details
$Splat = @{
    Headers = @{
        Authorization  = "Bearer $SlackKey"
        'Content-type' = 'application/json'
        charset        = 'utf-8'
    }
    Method  = 'Post'
    Uri     = 'https://slack.com/api/users.profile.set'
    Body    = @{
        user    = $SlackUserID
        profile = $SlackStatus
    } | ConvertTo-Json
    ErrorAction = 'Stop'
}

$response = Invoke-RestMethod @Splat

If ($response.ok -eq $true) {
    # Log the success
    "$TeamsStatus -- $($response.profile.status_emoji) -- $($response.profile.status_text)" | Out-File $PSScriptRoot\log.txt -Append
} Else {
    # Log the failure
    "ERROR -- $($response.error) -- $($response.warning)" | Out-File $PSScriptRoot\log.txt -Append
}

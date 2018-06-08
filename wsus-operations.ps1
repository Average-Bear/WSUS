<#
.SYNOPSIS
    Insert general purpose here.

.DESCRIPTION
    Insert script description here.
    
.NOTES
    Author: happysysadm.com
    Edited: JBear
#>
param(

    $Action = $false,
    $Today = (Get-Date),
    $Comments = (New-Object System.Collections.ArrayList),
    $d0 = (Get-Date -Day 1 -Month $($Today.Month) -Year $Today.Year)
)

$Comments.Add( "Today is $($Today.ToLongDateString())`n" )

switch ($d0.DayOfWeek) {

    "Sunday"    { $patchTuesday0 = $d0.AddDays(9)  }
    "Monday"    { $patchTuesday0 = $d0.AddDays(8)  } 
    "Tuesday"   { $patchTuesday0 = $d0.AddDays(7)  }
    "Wednesday" { $patchTuesday0 = $d0.AddDays(13) }
    "Thursday"  { $patchTuesday0 = $d0.AddDays(12) }
    "Friday"    { $patchTuesday0 = $d0.AddDays(11) }
    "Saturday"  { $patchTuesday0 = $d0.AddDays(10) }
}

$d1 = Get-Date -Day 1 -Month $($Today.Month + 1) -Year $Today.Year

switch ($d1.DayOfWeek) {

    "Sunday"    { $patchTuesday1 = $d1.AddDays(9)  }
    "Monday"    { $patchTuesday1 = $d1.AddDays(8)  }
    "Tuesday"   { $patchTuesday1 = $d1.AddDays(7)  }
    "Wednesday" { $patchTuesday1 = $d1.AddDays(13) }
    "Thursday"  { $patchTuesday1 = $d1.AddDays(12) }
    "Friday"    { $patchTuesday1 = $d1.AddDays(11) }
    "Saturday"  { $patchTuesday1 = $d1.AddDays(10) }

}

if($Today.Date -le $patchTuesday0.Date) {

    $patchTuesday = $patchTuesday0
}

else { 

    $patchTuesday = $patchTuesday1 
}

$d0 = Get-Date -Day 1 -Month $($Today.Month) -Year $Today.Year

switch ($d0.DayOfWeek) {

    "Sunday"    { $FourthMonday0 = $d0.AddDays(22) }
    "Monday"    { $FourthMonday0 = $d0.AddDays(21) }
    "Tuesday"   { $FourthMonday0 = $d0.AddDays(20) }
    "Wednesday" { $FourthMonday0 = $d0.AddDays(26) }
    "Thursday"  { $FourthMonday0 = $d0.AddDays(25) }
    "Friday"    { $FourthMonday0 = $d0.AddDays(24) }
    "Saturday"  { $FourthMonday0 = $d0.AddDays(23) }
}
    
$d1 = Get-Date -Day 1 -Month $($Today.Month + 1) -Year $Today.Year

switch ($d1.DayOfWeek) {

    "Sunday"    { $FourthMonday1 = $d1.AddDays(22) }
    "Monday"    { $FourthMonday1 = $d1.AddDays(21) }
    "Tuesday"   { $FourthMonday1 = $d1.AddDays(20) }
    "Wednesday" { $FourthMonday1 = $d1.AddDays(26) }
    "Thursday"  { $FourthMonday1 = $d1.AddDays(25) }
    "Friday"    { $FourthMonday1 = $d1.AddDays(24) }
    "Saturday"  { $FourthMonday1 = $d1.AddDays(23) }
}

if($Today.Date -le $FourthMonday0.Date) {

    $FourthMonday = $FourthMonday0 
}

else {

    $FourthMonday = $FourthMonday1
}

if($Today.Date -le $FourthMonday0.AddDays(1).Date) {

    $StandardApprovalDay = $FourthMonday0.AddDays(1)
}

else{

    $StandardApprovalDay= $FourthMonday1.AddDays(1)  
}

if($Today.Date -le $FourthMonday0.AddDays(1).Date) {

    $CriticalApprovalDay = $FourthMonday0.AddDays(7)
}

else {

    $CriticalApprovalDay= $FourthMonday1.AddDays(7) 
}

if($Today.Date -eq $PatchTuesday.Date) {

    $Comments.Add( "==> It's patch Tuesday!`n" )
    $Action = $true
}

else {

    $Comments.Add( "Next Patch Tuesday is in $((New-TimeSpan -Start $Today.Date -End $patchTuesday.Date).days) days on $($patchTuesday.ToLongDateString())`n" )
}

if($Today.Date -eq $FourthMonday.Date) {

    $Comments.Add( "==> It's fourth monday of the month - synching WSUS with Microsoft!`n" )
    $Action = $true
    $startTime = (Get-Date -f dd-MM-yyyy)
    (Get-WsusServer).GetSubscription().StartSynchronization()
}

else {

    $Comments.Add( "Next Sync will happen in $((New-TimeSpan -Start $Today.Date -End $FourthMonday.Date).days) days on $($FourthMonday.ToLongDateString())`n" )
}

if($Today.Date -eq $StandardApprovalDay.Date) {

    $Comments.Add( "==> It's the day after fourth monday of the month - approving for Standard servers`n" )
    $Action = $true
    $WSUS = Get-WsusServer
    $allupDates = $WSUS.GetUpDates() 
    $alltargetgroups = $WSUS.GetComputerTargetGroups()
    $computergroups = ($alltargetgroups | ? name -match 'Standard').name
    $computergroups | % {

        Get-WsusUpDate -Approval Unapproved -Status FailedOrNeeded | Approve-WsusUpDate -Action Install -TargetGroupName $_ –Verbose

    }

    $startTime = (Get-Date -f dd-MM-yyyy)
}

else {

    $Comments.Add( "Next approval for Standard servers will happen in $((New-TimeSpan -Start $Today.Date -End $StandardApprovalDay.Date).days) days on $($StandardApprovalDay.ToLongDateString())`n" )
}

if($Today.Date -eq $CriticalApprovalDay.Date) { 

    $Comments.Add( "==> It's the 7th day after fourth monday of the month - approving for User-Touchy and Mission-Critical servers`n" )
    $Action = $true
    $WSUS = Get-WsusServer
    $allupDates = $WSUS.GetUpDates() 
    $alltargetgroups = $WSUS.GetComputerTargetGroups()
    $computergroups = ($alltargetgroups | ? name -match 'touchy|critical').Name

    $computergroups | % {

        Get-WsusUpDate -Approval Unapproved -Status FailedOrNeeded | Approve-WsusUpDate -Action Install -TargetGroupName $_ –Verbose
    }

    $startTime = (get-Date -f dd-MM-yyyy)

}

else {

    $Comments.Add( "Next approval for User-Touchy and Mission-Critical servers will happen in $((New-TimeSpan -Start $Today.Date -End $CriticalApprovalDay.Date).days) days on $($CriticalApprovalDay.ToLongDateString())`n" )
}

if($Today.day -eq 7) {

    $Comments.Add( "==> Today is WSUS monthly clean up day`n" )
    $Action = $true
}

else {

    $Comments.Add( "Next WSUS monthly clean up will happen in $((New-TimeSpan -Start $Today.Date -End $(Get-Date -Day 7 -Month $($Today.Month + 1) -Year $Today.Year -OutVariable Datenextcleanup).Date).Days) days on $($Datenextcleanup.ToLongDateString())`n" )
}

$Comments

if(!$Action) {

    $Comments.Add( "<i style='color:red'>No actions to be done today</i>`n" )
}

$CommentsHTML = "<p style='color:blue'>" + $Comments.Replace("`n",'<br>') + "</p>"
$WSUS = Get-WsusServer
$AllTargetGroups = $WSUS.GetComputerTargetGroups()
$PatchReport = $AllTargetGroups | ForEach {

    $Group = $_.Name

    $_.GetTotalSummary() | ForEach {

        [PSCustomObject]@{

            TargetGroup = $Group
            Needed = ($_.NotInstalledCount + $_.DownloadedCount)
            "Installed/NotApplicable" = ($_.NotApplicableCount + $_.InstalledCount)
            NoStatus = $_.UnknownCount
            PendingReboot = $_.InstalledPendingRebootCount
        }
    }
}

$params = @{
    
    'encoding'=[System.Text.Encoding]::UTF8	
    'To' = 'recipient@domain.com'
    'From' = 'sender@domain.com'
    'SmtpServer' = "smtphost"
    'BodyAsHtml' = $true
    'Subject' = "WSUS - Patch Report"        
    'Body' = (($Commentshtml) + "<br>" + ($PatchReport | ConvertTo-Html | Out-String))   
}

Send-MailMessage @paramsd

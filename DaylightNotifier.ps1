[datetime]$currentDate=[DateTime]::Now.Date
$currentDateFormatted = 
$lattitude=50.4433
$longitude=30.5247
$url="https://api.sunrise-sunset.org/json?lat=$lattitude&lng=$longitude&date=$($currentDate.ToString("yyyy-MM-dd"))"
$outputDir = "Output"
$outputFileName = "Daylight.json"
$outputFilePath = Join-Path $outputDir $outputFileName
$outlookRootFolderName = "Gennadii_Saltyshchak@epam.com"
$outlookCalendarName = "Daylight"
$outlookEventLocation="$city($country)"
$outlookEventSensitivity = 2 # olPrivate

Function Create-Notifications {
    $output = Invoke-RestMethod -Uri $url
    $timeZoneOffset = [TimeZoneInfo]::Local.GetUtcOffset($currentDate)
    $sunRiseTime = [DateTime]::Parse($output.results.sunrise).TimeOfDay + $timeZoneOffset
    $sunSetTime = [DateTime]::Parse($output.results.sunset).TimeOfDay + $timeZoneOffset
    $sunRise = $currentDate + $sunRiseTime
    $sunSet = $currentDate + $sunSetTime

    #TODO: Refactoring: reduce code duplication
    New-OutlookCalendarMeeting -CalendarName $outlookCalendarName `
        -Subject "Sunrise at $sunRiseTime" `
        -Body $sunRiseTime `
        -Location $outlookEventLocation `
        -MeetingStart $sunRise `
        -MeetingDuration 0 `
        -Sensitivity $outlookEventSensitivity `
        -Categories "Sunrise" `
        -CheckDuplicates `
        -RootFolderName $outlookRootFolderName
    New-OutlookCalendarMeeting -CalendarName $outlookCalendarName `
        -Subject "Sunset at $sunSetTime" `
        -Body $sunSetTime `
        -Location $outlookEventLocation `
        -MeetingStart $sunSet `
        -MeetingDuration 0 `
        -Sensitivity $outlookEventSensitivity `
        -Categories "Sunset" `
        -CheckDuplicates `
        -RootFolderName $outlookRootFolderName
}

Push-Location (Split-Path $MyInvocation.MyCommand.Path -Parent)
Import-Module .\Modules\OutlookTools\OutlookTools.psm1
If (!(Test-Path $outputDir)) {
    New-Item  -Type Directory $outputDir | Out-Null
}
Create-Notifications
Pop-Location

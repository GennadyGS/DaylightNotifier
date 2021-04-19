$lattitude=50.4433
$longitude=30.5247
$outputDir = "Output"
$outlookRootFolderName = "Gennadii_Saltyshchak@epam.com"
$outlookCalendarName = "Daylight"
$outlookEventLocation="$city($country)"
$outlookEventSensitivity = 2 # olPrivate

Function CreateNotifications {
    param (
        [DateTime] $date
    )
    $url="https://api.sunrise-sunset.org/json?lat=$lattitude&lng=$longitude&date=$($date.ToString("yyyy-MM-dd"))"
    $output = Invoke-RestMethod -Uri $url
    $timeZoneOffset = [TimeZoneInfo]::Local.GetUtcOffset($date)
    $sunRiseTime = [DateTime]::Parse($output.results.sunrise).TimeOfDay + $timeZoneOffset
    $sunSetTime = [DateTime]::Parse($output.results.sunset).TimeOfDay + $timeZoneOffset
    $sunRise = $date + $sunRiseTime
    $sunSet = $date + $sunSetTime

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
CreateNotifications ([DateTime]::UtcNow.Date.AddDays(1))
Pop-Location

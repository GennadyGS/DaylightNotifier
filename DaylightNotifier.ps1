[datetime]$currentDate=Get-Date
[int]$day=$currentDate.Day
[int]$month=$currentDate.Month
[int]$timeZone=[TimeZoneInfo]::Local.BaseUtcOffset.Hours
[int]$dst=[int]$currentDate.IsDaylightSavingTime()
[double]$latitude=50.5
[double]$longitude=30.5
[string]$url="http://www.earthtools.org/sun/$latitude/$longitude/$day/$month/$timeZone/$dst"
[string]$outputDir = "Output"
[string]$outputFileName = "Daylight.xml"
[string]$outputFilePath = Join-Path $outputDir $outputFileName
[string]$outlookCalendarName = "Daylight"
[string]$outlookEventLocation="Киев"
[int]$outlookEventSensitivity = 2 # olPrivate

Function Get-Data {
    [Microsoft.PowerShell.Commands.WebResponseObject]$webResponse = Invoke-WebRequest $url
    [IO.Stream]$contentStream = $webResponse.RawContentStream
    $contentStream.Position = 0
    [IO.StreamReader]$streamReader = new-object IO.StreamReader($contentStream)
    $streamReader.ReadToEnd() | Out-File $outputFilePath -Encoding utf8
}

Function Create-Notifications {
    [xml] $data = Get-Content $outputFilePath
    #TODO: Refactoring: reduce code duplication
    #TODO: Twilights
    New-OutlookCalendarMeeting -CalendarName $outlookCalendarName `
        -Subject "Восход солнца в $($data.sun.morning.sunrise)" `
        -Body $data.sun.morning.sunrise `
        -Location $outlookEventLocation `
        -MeetingStart ($currentDate.Date + $data.sun.morning.sunrise) `
        -MeetingDuration 0 `
        -Sensitivity $outlookEventSensitivity `
        -Categories "Sunrise" `
        -CheckDuplicates
    New-OutlookCalendarMeeting -CalendarName $outlookCalendarName `
        -Subject "Закат солнца в $($data.sun.evening.sunset)" `
        -Body $data.sun.evening.sunset `
        -Location $outlookEventLocation `
        -MeetingStart ($currentDate.Date + $data.sun.evening.sunset) `
        -MeetingDuration 0 `
        -Sensitivity $outlookEventSensitivity `
        -Categories "Sunset" `
        -CheckDuplicates
}

Push-Location (Split-Path $MyInvocation.MyCommand.Path -Parent)
Import-Module .\Modules\OutlookTools\OutlookTools.psm1
If (!(Test-Path $outputDir)) {
    New-Item  -Type Directory $outputDir | Out-Null
}
Get-Data
Create-Notifications
Pop-Location

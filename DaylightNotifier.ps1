[datetime]$currentDate=Get-Date
[string]$apiKey="7fff5ef13308f08a"
[string]$country="Ukraine"
[string]$city="Kyiv"
[string]$url="http://api.wunderground.com/api/$apiKey/astronomy/q/$country/$city.xml"
[string]$outputDir = "Output"
[string]$outputFileName = "Daylight.xml"
[string]$outputFilePath = Join-Path $outputDir $outputFileName
[string]$outlookRootFolderName = "Gennadii_Saltyshchak@epam.com"
[string]$outlookCalendarName = "Daylight"
[string]$outlookEventLocation="$city($country)"
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
    [Xml.XmlElement]$sunriseNode=$data.response.sun_phase.sunrise
    [System.TimeSpan]$sunrise=New-Object System.TimeSpan -ArgumentList $sunriseNode.hour, $sunriseNode.minute, 0
    New-OutlookCalendarMeeting -CalendarName $outlookCalendarName `
        -Subject "Sunrise at $sunrise" `
        -Body $sunrise `
        -Location $outlookEventLocation `
        -MeetingStart ($currentDate.Date + $sunrise) `
        -MeetingDuration 0 `
        -Sensitivity $outlookEventSensitivity `
        -Categories "Sunrise" `
        -CheckDuplicates `
        -RootFolderName $outlookRootFolderName
    [Xml.XmlElement]$sunsetNode=$data.response.sun_phase.sunset
    [System.TimeSpan]$sunset=New-Object System.TimeSpan -ArgumentList $sunsetNode.hour, $sunsetNode.minute, 0
    New-OutlookCalendarMeeting -CalendarName $outlookCalendarName `
        -Subject "Sunset at $sunset" `
        -Body $sunset `
        -Location $outlookEventLocation `
        -MeetingStart ($currentDate.Date + $sunset) `
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
Get-Data
Create-Notifications
Pop-Location

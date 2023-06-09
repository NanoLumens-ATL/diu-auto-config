add-type -AssemblyName Microsoft.VisualBasic
add-type -AssemblyName System.Windows.Forms

function Find-Record {
    Write-Host "Searching for record with project name: $ProjectName" -ForegroundColor "Magenta"

    $Url = "$Domain/records/query"
    If ($Debug) {Write-Host "Url: $Url" -BackgroundColor "Black" -ForegroundColor "Cyan"}
    $RequestBody = @{
        from="$ProjectsTableId"
        where="{6.EX.'$ProjectName'}"
        select=@('3')
    } | ConvertTo-Json
    If ($Debug) {Write-Host "RequestBody: $RequestBody" -BackgroundColor "Black" -ForegroundColor "Cyan"}

    try {
        $Response = Invoke-WebRequest -Uri $Url -Method 'Post' -Body $RequestBody -Headers $Headers -ContentType 'application/json' | ConvertFrom-Json
        $RecordId = $Response.data.'3'.value
        if ($RecordId -ne $null) {Write-Host "Record found with ID: $RecordId" -ForegroundColor "Green"} else {throw [Exception] "Couldn't find record! Please make sure spelling matches what is in Quickbase"}
        return $RecordId
    } catch {
        Write-Host $PSItem.Exception.Message -ForegroundColor red
        Exit
    }
}

function Find-Files {
    Write-Host "Searching for files with RecordID: $RecordId" -ForegroundColor "Magenta"

    $Url = "$Domain/records/query"
    if ($Debug) {Write-Host "Url: $Url" -BackgroundColor "Black" -ForegroundColor "Cyan"}
    $RequestBody = @{
        from="$ProjectsTableId"
        where="{3.EX.'$RecordId'}"
        select=@('25', '30', '386', '608')
    } | ConvertTo-Json
    if ($Debug) {Write-Host "RequestBody: $RequestBody" -BackgroundColor "Black" -ForegroundColor "Cyan"}

    try {
        $Response = Invoke-WebRequest -Uri $Url -Method 'POST' -Body $RequestBody -Headers $Headers -ContentType 'application/json' | ConvertFrom-Json
        $ElectricalDrawingFile = $Response.data.'25'.value
        if ($ElectricalDrawingFile -ne $null) {Write-Host "Electrical Drawing File found:" $ElectricalDrawingFile.versions[$ElectricalDrawingFile.versions.count - 1].fileName "With latest version:"$ElectricalDrawingFile.versions.count -ForegroundColor "Green"}
        if ($Debug) {ConvertTo-Json $SubmittalFile | Write-Host -BackgroundColor "Black" -ForegroundColor "Cyan"}
        $MappingFile = $Response.data.'30'.value
        if ($MappingFile -ne $null) {Write-Host "Mapping File found:" $MappingFile.versions[$MappingFile.versions.count - 1].fileName "With latest version:"$MappingFile.versions.count -ForegroundColor "Green"}
        if ($Debug) {ConvertTo-Json $MappingFile | Write-Host -BackgroundColor "Black" -ForegroundColor "Cyan"}
        $ProductionDrawingFile = $Response.data.'386'.value
        if ($ProductionDrawingFile -ne $null) {Write-Host "Production Drawing File found:" $ProductionDrawingFile.versions[$ProductionDrawingFile.versions.count - 1].fileName "With latest version:"$ProductionDrawingFile.versions.count -ForegroundColor "Green"}
        if ($Debug) {ConvertTo-Json $SubmittalFile | Write-Host -BackgroundColor "Black" -ForegroundColor "Cyan"}
        $SubmittalFile = $Response.data.'608'.value
        if ($SubmittalFile -ne $null) {Write-Host "Submittal File found:" $SubmittalFile.versions[$SubmittalFile.versions.count - 1].fileName "With latest version:"$SubmittalFile.versions.count -ForegroundColor "Green"}
        if ($Debug) {ConvertTo-Json $SubmittalFile | Write-Host -BackgroundColor "Black" -ForegroundColor "Cyan"}

        $Files = @(
            @{
                Uri = $Domain + $ElectricalDrawingFile.url
                OutFile = $ElectricalDrawingFile.versions[$ElectricalDrawingFile.versions.count - 1].fileName
                Name = "Electrical Drawing"
            },
            @{
                Uri = $Domain + $MappingFile.url
                OutFile = $MappingFile.versions[$MappingFile.versions.count - 1].fileName
                Name = "Mapping File"
            },
            @{
                Uri = $Domain + $ProductionDrawingFile.url
                OutFile = $ProductionDrawingFile.versions[$ProductionDrawingFile.versions.count - 1].fileName
                Name = "Production Drawing"
            },
            @{
                Uri = $Domain + $SubmittalFile.url
                OutFile = $SubmittalFile.versions[$SubmittalFile.versions.count - 1].fileName
                Name = "Submittal File"
            }
        )

        return $Files
    } catch {
        Write-Error -Message $PSItem.Exception.Message
    }
}

function Download-Files {
    Param($Files)
    Write-Host "Downloading Files for project: $ProjectName" -ForegroundColor "Magenta"

    if ($Debug) {ConvertTo-Json $Files | Write-Host -BackgroundColor "Black" -ForegroundColor "Cyan"}
    $DesktopDir = [Environment]::GetFolderPath("Desktop")
    $Dir = New-Item -ItemType Directory -Force -Path $DesktopDir\$ProjectName

    foreach ($File in $Files) {
	    try {
		    $FilePath = [string]::Format("{0}\{1}", $Dir, $File.OutFile)
            if ($Debug) {Write-Host "FilePath: "$FilePath -ForegroundColor cyan}
		    $Response = Invoke-WebRequest -Uri $File.Uri -Method GET -Headers $Headers
            # if ($Debug) {Write-Host $Response -ForegroundColor cyan}
		    $String64 = [System.Text.Encoding]::ASCII.GetString($Response.content)
            # if ($Debug) {Write-Host $String64 -ForegroundColor cyan}
		    $Bytes = [Convert]::FromBase64String($String64)
            # if ($Debug) {Write-Host $Bytes -ForegroundColor cyan}
		    [IO.File]::WriteAllBytes($FilePath, $Bytes)
		    Write-Host "Downloaded"$File.OutFile -ForegroundColor green
	    } catch {
		    Write-Host "Failed to download"$File.Name -ForegroundColor red -BackgroundColor black
	    }
    }
}

$Debug = $false
if ($Debug) {Write-Host debug is on} else {Write-Host debug is off}

$Domain = "https://api.quickbase.com/v1"
$ProjectsTableId = "bmb347gyc"

$UserToken = "b76ve3_e7y7_0_byr7cmvbu435a3d8mpfykbvcu4rn"
$ProjectName = Read-Host -Prompt "Please type in project name"

$Headers = New-Object 'System.Collections.Generic.Dictionary[String,String]'
$Headers.Add("QB-Realm-Hostname", "nanolumens.quickbase.com")
$Headers.Add("Authorization", "QB-USER-TOKEN $UserToken")

$RecordId = Find-Record
$Files = Find-Files
Download-Files $Files
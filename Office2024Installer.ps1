$base = "C:\Office2024"
$odtUrl = "https://download.microsoft.com/download/6c1eeb25-cf8b-41d9-8d0d-cc1dbc032140/officedeploymenttool_19929-20062.exe"

$odtExe = "$base\officedeploymenttool.exe"
$setup = "$base\setup.exe"
$config = "$base\config.xml"
$officeFolder = "$base\Office"
$officeDataFolder = "$base\Office\Data"
$estimatedGB = 4.5

New-Item -ItemType Directory -Path $base -Force | Out-Null
Set-Location $base

function Test-OfficeCache {
    if (!(Test-Path $officeDataFolder)) {
        return $false
    }

    $files = Get-ChildItem $officeDataFolder -Recurse -File -ErrorAction SilentlyContinue
    $totalSize = ($files | Measure-Object Length -Sum).Sum

    if ($files.Count -lt 10) {
        return $false
    }

    if ($totalSize -lt 1GB) {
        return $false
    }

    return $true
}

if (!(Test-Path $setup)) {
    Write-Host "Downloading Office Deployment Tool..."
    Invoke-WebRequest $odtUrl -OutFile $odtExe

    Write-Host "Extracting Office Deployment Tool..."
    Start-Process `
        -FilePath $odtExe `
        -ArgumentList "/quiet", "/extract:$base" `
        -Wait
}

if (!(Test-Path $setup)) {
    Write-Host "setup.exe missing."
    Read-Host "Press Enter to exit"
    exit
}

function Ask-App {
    param([string]$Name)

    while ($true) {
        Write-Host ""
        Write-Host "Install $Name ? [Y/N] " -NoNewline

        $key = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
        $choice = $key.Character.ToString().ToUpper()

        if ($choice -eq "Y") {
            Write-Host "Y"
            return $true
        }

        if ($choice -eq "N") {
            Write-Host "N"
            return $false
        }
    }
}

$apps = @(
    "Word",
    "Excel",
    "PowerPoint",
    "Outlook",
    "OneNote",
    "Access",
    "Publisher"
)

$excluded = @()

foreach ($app in $apps) {
    if (-not (Ask-App $app)) {
        $excluded += $app
    }
}

$excluded += "Teams"
$excluded += "Lync"
$excluded += "OneDrive"
$excluded = $excluded | Sort-Object -Unique

$excludeXml = ""

foreach ($app in $excluded) {
    $excludeXml += "      <ExcludeApp ID=`"$app`"/>`r`n"
}

$xml = @"
<Configuration>
  <Add OfficeClientEdition="64" Channel="PerpetualVL2024">
    <Product ID="ProPlus2024Volume">
      <Language ID="en-us"/>
$excludeXml    </Product>
  </Add>
  <Display Level="Full" AcceptEULA="TRUE"/>
</Configuration>
"@

Set-Content `
    -Path $config `
    -Value $xml `
    -Encoding UTF8

Write-Host ""
Write-Host "Generated config.xml:"
Write-Host ""
Get-Content $config
Write-Host ""

if (!(Test-OfficeCache)) {
    Write-Host ""
    Write-Host "No valid Office cache found. Downloading Office files..."

    $download = Start-Process `
        -FilePath $setup `
        -ArgumentList "/download", $config `
        -WorkingDirectory $base `
        -PassThru

    $lastSize = 0
    $lastTime = Get-Date

    while (-not $download.HasExited) {
        Start-Sleep -Milliseconds 700

        if (Test-Path $officeDataFolder) {
            $size = (
                Get-ChildItem `
                    $officeDataFolder `
                    -Recurse `
                    -ErrorAction SilentlyContinue |
                Measure-Object Length -Sum
            ).Sum
        } else {
            $size = 0
        }

        $now = Get-Date
        $seconds = ($now - $lastTime).TotalSeconds
        $delta = $size - $lastSize

        if ($seconds -gt 0 -and $delta -gt 0) {
            $speed = ($delta / 1MB) / $seconds
        } else {
            $speed = 0
        }

        $gb = [math]::Round($size / 1GB, 2)

        $percent = [math]::Min(
            100,
            [math]::Round(($gb / $estimatedGB) * 100, 1)
        )

        Write-Progress `
            -Activity "Downloading Office 2024" `
            -Status "$gb GB | $([math]::Round($speed,2)) MB/s | $percent%" `
            -PercentComplete $percent

        Write-Host (
            "Downloaded: {0} GB | Speed: {1:N2} MB/s | {2}%" `
            -f $gb, $speed, $percent
        )

        $lastSize = $size
        $lastTime = $now
    }

    Write-Progress -Activity "Downloading Office 2024" -Completed

    if ($download.ExitCode -ne 0) {
        Write-Host ""
        Write-Host "Download failed."
        Write-Host "Exit code: $($download.ExitCode)"
        Read-Host "Press Enter to exit"
        exit
    }

} else {
    Write-Host ""
    Write-Host "Valid Office cache found."
    Write-Host "Skipping download."
}

Write-Host ""
Write-Host "Installing Office..."

$install = Start-Process `
    -FilePath $setup `
    -ArgumentList "/configure", $config `
    -WorkingDirectory $base `
    -Wait `
    -PassThru

if ($install.ExitCode -eq 0) {
    Write-Host ""
    Write-Host "Office installed successfully."
} else {
    Write-Host ""
    Write-Host "Install failed."
    Write-Host "Exit code: $($install.ExitCode)"
}

Read-Host "Press Enter to exit"

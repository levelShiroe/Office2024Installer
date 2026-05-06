Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$base = "C:\Office2024"
$odtUrl = "https://download.microsoft.com/download/6c1eeb25-cf8b-41d9-8d0d-cc1dbc032140/officedeploymenttool_19929-20062.exe"
$odtExe = "$base\officedeploymenttool.exe"
$setup = "$base\setup.exe"
$config = "$base\config.xml"
$officeFolder = "$base\Office"
$officeDataFolder = "$base\Office\Data"
$estimatedGB = 4.5

function Log($msg) {
    $statusBox.AppendText("[$(Get-Date -Format HH:mm:ss)] $msg`r`n")
    $statusBox.ScrollToCaret()
    [System.Windows.Forms.Application]::DoEvents()
}

function Get-OfficeDownloadSize {
    if (Test-Path $officeDataFolder) {
        return (Get-ChildItem $officeDataFolder -Recurse -ErrorAction SilentlyContinue | Measure-Object Length -Sum).Sum
    }

    if (Test-Path $officeFolder) {
        return (Get-ChildItem $officeFolder -Recurse -ErrorAction SilentlyContinue | Measure-Object Length -Sum).Sum
    }

    return 0
}

function Create-Config {
    $excluded = @()

    foreach ($cb in $checkboxes) {
        if (-not $cb.Checked) {
            $excluded += $cb.Text
        }
    }

    # Always exclude junk/extras
    $excluded += "Teams"
    $excluded += "Lync"      # Skype for Business
    $excluded += "OneDrive"
    $excluded = $excluded | Sort-Object -Unique

    $excludeXml = ""
    foreach ($app in $excluded) {
        $excludeXml += "      <ExcludeApp ID=`"$app`"/>`r`n"
    }

    $xml = @"
<Configuration>
  <Add OfficeClientEdition="64" Channel="PerpetualVL2024" SourcePath="C:\Office2024">
    <Product ID="ProPlus2024Volume">
      <Language ID="en-us"/>
$excludeXml    </Product>
  </Add>
  <Display Level="Full" AcceptEULA="TRUE"/>
  <Property Name="FORCEAPPSHUTDOWN" Value="TRUE"/>
</Configuration>
"@

    $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
    [System.IO.File]::WriteAllText($config, $xml, $utf8NoBom)
}

function Update-DownloadStats {
    param(
        [double]$GB,
        [double]$Speed,
        [double]$Percent,
        [string]$ETA
    )

    $progress.Value = [math]::Min(100, [int]$Percent)
    $speedLabel.Text = "Speed: $([math]::Round($Speed, 2)) MB/s"
    $downloadLabel.Text = "Downloaded: $GB GB / $estimatedGB GB"
    $etaLabel.Text = "ETA: $ETA"

    [System.Windows.Forms.Application]::DoEvents()
}

function Run-Installer {
    $installButton.Enabled = $false
    $progress.Value = 0
    $speedLabel.Text = "Speed: 0 MB/s"
    $downloadLabel.Text = "Downloaded: 0 GB / $estimatedGB GB"
    $etaLabel.Text = "ETA: --:--"

    try {
        New-Item -ItemType Directory -Path $base -Force | Out-Null
        Set-Location $base

        Log "Starting Office 2024 installer..."

        if (!(Test-Path $setup)) {
            Log "Downloading Office Deployment Tool..."
            Invoke-WebRequest $odtUrl -OutFile $odtExe
            $progress.Value = 10

            Log "Extracting Office Deployment Tool..."
            Start-Process -FilePath $odtExe -ArgumentList "/quiet", "/extract:$base" -Wait
            $progress.Value = 20
        } else {
            Log "Office Deployment Tool already exists."
            $progress.Value = 20
        }

        if (!(Test-Path $setup)) {
            Log "ERROR: setup.exe missing."
            return
        }

        Log "Creating config.xml..."
        Create-Config
        $progress.Value = 25

        if (!(Test-Path $officeFolder)) {
            Log "Downloading Office files..."

            $download = Start-Process `
                -FilePath $setup `
                -ArgumentList "/download", "config.xml" `
                -WorkingDirectory $base `
                -PassThru

            $lastSize = 0
            $lastTime = Get-Date

            while (-not $download.HasExited) {
                Start-Sleep -Milliseconds 700

                $size = Get-OfficeDownloadSize
                $now = Get-Date
                $seconds = ($now - $lastTime).TotalSeconds
                $delta = $size - $lastSize

                if ($seconds -gt 0 -and $delta -gt 0) {
                    $speed = ($delta / 1MB) / $seconds
                } else {
                    $speed = 0
                }

                $gb = [math]::Round($size / 1GB, 2)
                $percent = [math]::Min(100, [math]::Round(($gb / $estimatedGB) * 100, 1))

                if ($speed -gt 0) {
                    $remainingGB = [math]::Max(0, $estimatedGB - $gb)
                    $etaSeconds = ($remainingGB * 1024) / $speed
                    $etaText = "{0:mm\:ss}" -f ([TimeSpan]::FromSeconds($etaSeconds))
                } else {
                    $etaText = "--:--"
                }

                Update-DownloadStats -GB $gb -Speed $speed -Percent $percent -ETA $etaText

                $lastSize = $size
                $lastTime = $now
            }

            if ($download.ExitCode -ne 0) {
                Log "ERROR: Download failed. Exit code: $($download.ExitCode)"
                return
            }

            Update-DownloadStats -GB $estimatedGB -Speed 0 -Percent 100 -ETA "00:00"
            Log "Office files downloaded."
        } else {
            Log "Office files already exist. Skipping download."
            $progress.Value = 70
            $downloadLabel.Text = "Downloaded: already cached"
            $speedLabel.Text = "Speed: 0 MB/s"
            $etaLabel.Text = "ETA: 00:00"
        }

        Log "Installing Office..."
        $progress.Value = 80

        Set-Location $base

        # Raid boss fix: run ODT exactly like manual CMD from C:\Office2024
        $install = Start-Process `
            -FilePath "cmd.exe" `
            -ArgumentList "/c", "cd /d C:\Office2024 && setup.exe /configure config.xml" `
            -Wait `
            -PassThru

        if ($install.ExitCode -eq 0) {
            $progress.Value = 100
            Log "Office installed successfully."
        } else {
            Log "ERROR: Install failed. Exit code: $($install.ExitCode)"
            Log "Manual fallback: open CMD as admin, run: cd /d C:\Office2024 && setup.exe /configure config.xml"
        }
    }
    catch {
        Log "ERROR: $($_.Exception.Message)"
    }
    finally {
        $installButton.Enabled = $true
    }
}

$form = New-Object System.Windows.Forms.Form
$form.Text = "Office 2024 Installer Alpha"
$form.Size = New-Object System.Drawing.Size(540, 560)
$form.StartPosition = "CenterScreen"

$title = New-Object System.Windows.Forms.Label
$title.Text = "Office 2024 LTSC Installer"
$title.Font = New-Object System.Drawing.Font("Segoe UI", 16, [System.Drawing.FontStyle]::Bold)
$title.Location = New-Object System.Drawing.Point(20, 20)
$title.Size = New-Object System.Drawing.Size(480, 35)
$form.Controls.Add($title)

$subtitle = New-Object System.Windows.Forms.Label
$subtitle.Text = "Select apps to install. Teams, OneDrive, and Skype/Lync are always excluded."
$subtitle.Location = New-Object System.Drawing.Point(22, 60)
$subtitle.Size = New-Object System.Drawing.Size(480, 25)
$form.Controls.Add($subtitle)

$appNames = @("Word", "Excel", "PowerPoint", "Outlook", "OneNote", "Access", "Publisher")
$checkboxes = @()
$y = 100

foreach ($app in $appNames) {
    $cb = New-Object System.Windows.Forms.CheckBox
    $cb.Text = $app
    $cb.Location = New-Object System.Drawing.Point(35, $y)
    $cb.Size = New-Object System.Drawing.Size(180, 25)

    if ($app -eq "Word" -or $app -eq "Excel") {
        $cb.Checked = $true
    }

    $checkboxes += $cb
    $form.Controls.Add($cb)
    $y += 30
}

$installButton = New-Object System.Windows.Forms.Button
$installButton.Text = "Install"
$installButton.Location = New-Object System.Drawing.Point(280, 100)
$installButton.Size = New-Object System.Drawing.Size(180, 40)
$installButton.Add_Click({ Run-Installer })
$form.Controls.Add($installButton)

$exitButton = New-Object System.Windows.Forms.Button
$exitButton.Text = "Exit"
$exitButton.Location = New-Object System.Drawing.Point(280, 150)
$exitButton.Size = New-Object System.Drawing.Size(180, 40)
$exitButton.Add_Click({ $form.Close() })
$form.Controls.Add($exitButton)

$speedLabel = New-Object System.Windows.Forms.Label
$speedLabel.Text = "Speed: 0 MB/s"
$speedLabel.Location = New-Object System.Drawing.Point(25, 285)
$speedLabel.Size = New-Object System.Drawing.Size(220, 20)
$form.Controls.Add($speedLabel)

$downloadLabel = New-Object System.Windows.Forms.Label
$downloadLabel.Text = "Downloaded: 0 GB / $estimatedGB GB"
$downloadLabel.Location = New-Object System.Drawing.Point(25, 310)
$downloadLabel.Size = New-Object System.Drawing.Size(260, 20)
$form.Controls.Add($downloadLabel)

$etaLabel = New-Object System.Windows.Forms.Label
$etaLabel.Text = "ETA: --:--"
$etaLabel.Location = New-Object System.Drawing.Point(300, 285)
$etaLabel.Size = New-Object System.Drawing.Size(160, 20)
$form.Controls.Add($etaLabel)

$progress = New-Object System.Windows.Forms.ProgressBar
$progress.Location = New-Object System.Drawing.Point(25, 340)
$progress.Size = New-Object System.Drawing.Size(475, 25)
$form.Controls.Add($progress)

$statusBox = New-Object System.Windows.Forms.TextBox
$statusBox.Location = New-Object System.Drawing.Point(25, 380)
$statusBox.Size = New-Object System.Drawing.Size(475, 120)
$statusBox.Multiline = $true
$statusBox.ScrollBars = "Vertical"
$statusBox.ReadOnly = $true
$form.Controls.Add($statusBox)

[void]$form.ShowDialog()

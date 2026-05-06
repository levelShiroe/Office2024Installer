Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$base = "C:\Office2024"
$odtUrl = "https://download.microsoft.com/download/6c1eeb25-cf8b-41d9-8d0d-cc1dbc032140/officedeploymenttool_19929-20062.exe"
$odtExe = "$base\officedeploymenttool.exe"
$setup = "$base\setup.exe"
$config = "$base\config.xml"
$officeFolder = "$base\Office"

function Log($msg) {
    $statusBox.AppendText("[$(Get-Date -Format HH:mm:ss)] $msg`r`n")
    $statusBox.ScrollToCaret()
    [System.Windows.Forms.Application]::DoEvents()
}

function Create-Config {
    $excluded = @()

    foreach ($cb in $checkboxes) {
        if (-not $cb.Checked) {
            $excluded += $cb.Text
        }
    }

    # Always exclude unwanted extras
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

    Set-Content -Path $config -Value $xml -Encoding UTF8
}

function Run-Installer {
    $installButton.Enabled = $false
    $progress.Value = 0

    try {
        New-Item -ItemType Directory -Path $base -Force | Out-Null
        Set-Location $base

        Log "Starting Office 2024 installer..."

        if (!(Test-Path $setup)) {
            Log "Downloading Office Deployment Tool..."
            Invoke-WebRequest $odtUrl -OutFile $odtExe
            $progress.Value = 15

            Log "Extracting Office Deployment Tool..."
            Start-Process -FilePath $odtExe -ArgumentList "/quiet", "/extract:$base" -Wait
            $progress.Value = 25
        } else {
            Log "Office Deployment Tool already exists."
            $progress.Value = 25
        }

        if (!(Test-Path $setup)) {
            Log "ERROR: setup.exe missing."
            return
        }

        Log "Creating config.xml..."
        Create-Config
        $progress.Value = 35

        if (!(Test-Path $officeFolder)) {
            Log "Downloading Office files..."
            $download = Start-Process -FilePath $setup -ArgumentList "/download", $config -WorkingDirectory $base -Wait -PassThru

            if ($download.ExitCode -ne 0) {
                Log "ERROR: Download failed. Exit code: $($download.ExitCode)"
                return
            }

            $progress.Value = 70
            Log "Office files downloaded."
        } else {
            Log "Office files already exist. Skipping download."
            $progress.Value = 70
        }

        Log "Installing Office..."
        $install = Start-Process -FilePath $setup -ArgumentList "/configure", $config -WorkingDirectory $base -Wait -PassThru

        if ($install.ExitCode -eq 0) {
            $progress.Value = 100
            Log "Office installed successfully."
        } else {
            Log "ERROR: Install failed. Exit code: $($install.ExitCode)"
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
$form.Size = New-Object System.Drawing.Size(520, 520)
$form.StartPosition = "CenterScreen"

$title = New-Object System.Windows.Forms.Label
$title.Text = "Office 2024 LTSC Installer"
$title.Font = New-Object System.Drawing.Font("Segoe UI", 16, [System.Drawing.FontStyle]::Bold)
$title.Location = New-Object System.Drawing.Point(20, 20)
$title.Size = New-Object System.Drawing.Size(460, 35)
$form.Controls.Add($title)

$subtitle = New-Object System.Windows.Forms.Label
$subtitle.Text = "Select apps to install. Teams, OneDrive, and Skype/Lync are always excluded."
$subtitle.Location = New-Object System.Drawing.Point(22, 60)
$subtitle.Size = New-Object System.Drawing.Size(460, 25)
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
$installButton.Location = New-Object System.Drawing.Point(260, 100)
$installButton.Size = New-Object System.Drawing.Size(180, 40)
$installButton.Add_Click({ Run-Installer })
$form.Controls.Add($installButton)

$exitButton = New-Object System.Windows.Forms.Button
$exitButton.Text = "Exit"
$exitButton.Location = New-Object System.Drawing.Point(260, 150)
$exitButton.Size = New-Object System.Drawing.Size(180, 40)
$exitButton.Add_Click({ $form.Close() })
$form.Controls.Add($exitButton)

$progress = New-Object System.Windows.Forms.ProgressBar
$progress.Location = New-Object System.Drawing.Point(25, 330)
$progress.Size = New-Object System.Drawing.Size(455, 25)
$form.Controls.Add($progress)

$statusBox = New-Object System.Windows.Forms.TextBox
$statusBox.Location = New-Object System.Drawing.Point(25, 370)
$statusBox.Size = New-Object System.Drawing.Size(455, 90)
$statusBox.Multiline = $true
$statusBox.ScrollBars = "Vertical"
$statusBox.ReadOnly = $true
$form.Controls.Add($statusBox)

[void]$form.ShowDialog()

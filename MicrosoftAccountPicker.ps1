Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

function Load-Accounts {
    $accounts = @()
    $user = $env:USERPROFILE
    $files = Get-ChildItem -Path "$user\AppData\Local\Microsoft\OneAuth\accounts"

    foreach ($file in $files) {
        $content = Get-Content $file.FullName -Raw

        # Extract email
        $email = $null
        if ($content -match '"account_hints"\s*:\s*"\[\\"([^\\]+)') {
            $email = $matches[1]
        }

        # Determine status using whole word check
        if ($content -match '\bdisassociated\b') {
            $status = 'Disassociated'
        } elseif ($content -match '\bassociated\b') {
            $status = 'Associated'
        } else {
            $status = 'Unknown'
        }

        $accounts += [PSCustomObject]@{
            FileName = $file.Name
            Email    = $email
            Status   = $status
        }
    }

    return $accounts
}

function Refresh-UI {
    $form.Controls.Clear()
    $form.Controls.Add($panel)
    $panel.Controls.Clear()

    $accounts = Load-Accounts
    $y = 10

    foreach ($account in $accounts) {
        # Label
        $label = New-Object System.Windows.Forms.Label
        $label.Text = "$($account.Email) - $($account.Status)"
        $label.Location = New-Object System.Drawing.Point(10, $y)
        $label.Size = New-Object System.Drawing.Size(300, 20)

        # Set color
        switch ($account.Status) {
            'Associated'    { $label.ForeColor = [System.Drawing.Color]::Green }
            'Disassociated' { $label.ForeColor = [System.Drawing.Color]::Red }
            default         { $label.ForeColor = [System.Drawing.Color]::Black }
        }

        # Associate button
        $associateBtn = New-Object System.Windows.Forms.Button
        $associateBtn.Text = "Associate"
        $associateBtn.Location = New-Object System.Drawing.Point(320, $y)
        $associateBtn.Size = New-Object System.Drawing.Size(80, 25)
        $associateBtn.Tag = $account.FileName
        $associateBtn.Add_Click({
            $filePath = "$env:USERPROFILE\AppData\Local\Microsoft\OneAuth\accounts\$($this.Tag)"
            $content = Get-Content $filePath -Raw

            # Replace only whole words to avoid 'disdisassociated'
            $newContent = $content -replace '\bdisassociated\b', 'associated'

            Set-Content -Path $filePath -Value $newContent
            Refresh-UI
        })

        # Disassociate button
        $disassociateBtn = New-Object System.Windows.Forms.Button
        $disassociateBtn.Text = "Disassociate"
        $disassociateBtn.Location = New-Object System.Drawing.Point(420, $y)
        $disassociateBtn.Size = New-Object System.Drawing.Size(100, 25)
        $disassociateBtn.Tag = $account.FileName
        $disassociateBtn.Add_Click({
            $filePath = "$env:USERPROFILE\AppData\Local\Microsoft\OneAuth\accounts\$($this.Tag)"
            $content = Get-Content $filePath -Raw

            # Replace only whole words
            $newContent = $content -replace '\bassociated\b', 'disassociated'

            Set-Content -Path $filePath -Value $newContent
            Refresh-UI
        })

        $panel.Controls.AddRange(@($label, $associateBtn, $disassociateBtn))
        $y += 35
    }
}

# Create form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Microsoft Account Picker"
$form.Size = New-Object System.Drawing.Size(650, 500)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false
$form.MinimizeBox = $false

# Add scrollable panel
$panel = New-Object System.Windows.Forms.Panel
$panel.Dock = "Fill"
$panel.AutoScroll = $true
$form.Controls.Add($panel)

# Load and display UI
Refresh-UI

# Show form
[void]$form.ShowDialog()

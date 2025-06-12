Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# CSV File Path
$csvPath = "$PSScriptRoot\inventory.csv"

# Create Form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Inventory Entry"
$form.Size = New-Object System.Drawing.Size(600, 400)
$form.StartPosition = "CenterScreen"

# Dropdown Label and Menu for Floor
$floorLabel = New-Object System.Windows.Forms.Label
$floorLabel.Text = "Floor:"
$floorLabel.Location = New-Object System.Drawing.Point(10, 20)
$form.Controls.Add($floorLabel)

$floorDropdown = New-Object System.Windows.Forms.ComboBox
$floorDropdown.Location = New-Object System.Drawing.Point(60, 18)
$floorDropdown.Width = 60
$floorDropdown.Items.AddRange(17..21)
$floorDropdown.DropDownStyle = 'DropDownList'
$form.Controls.Add($floorDropdown)

# Desk
$deskLabel = New-Object System.Windows.Forms.Label
$deskLabel.Text = "Desk:"
$deskLabel.Location = New-Object System.Drawing.Point(140, 20)
$form.Controls.Add($deskLabel)

$deskBox = New-Object System.Windows.Forms.TextBox
$deskBox.Location = New-Object System.Drawing.Point(190, 18)
$form.Controls.Add($deskBox)

# Header labels
$headers = "Device","Model","Serial Number","Asset Tag"
$yOffset = 60
for ($i = 0; $i -lt $headers.Count; $i++) {
    $label = New-Object System.Windows.Forms.Label
    $label.Text = $headers[$i]
    $label.Location = New-Object System.Drawing.Point(10 + $i*140, $yOffset)
    $form.Controls.Add($label)
}

# Device Entry Fields
$devices = @("Monitor 1", "Monitor 2", "Computer", "Dock")
$entries = @{}

$yOffset += 30
foreach ($device in $devices) {
    $entries[$device] = @{}
    $label = New-Object System.Windows.Forms.Label
    $label.Text = $device
    $label.Location = New-Object System.Drawing.Point(10, $yOffset)
    $form.Controls.Add($label)

    # Model
    $model = New-Object System.Windows.Forms.TextBox
    $model.Location = New-Object System.Drawing.Point(150, $yOffset)
    $form.Controls.Add($model)
    $entries[$device]["Model"] = $model

    # Serial
    $serial = New-Object System.Windows.Forms.TextBox
    $serial.Location = New-Object System.Drawing.Point(290, $yOffset)
    $form.Controls.Add($serial)
    $entries[$device]["Serial"] = $serial

    # Asset Tag
    $tag = New-Object System.Windows.Forms.TextBox
    $tag.Location = New-Object System.Drawing.Point(430, $yOffset)
    $form.Controls.Add($tag)
    $entries[$device]["Tag"] = $tag

    $yOffset += 30
}

# Error MessageBox
function Show-Error {
    param([string]$msg)
    [System.Windows.Forms.MessageBox]::Show($msg, "Error", 'OK', 'Error')
}

# Submit Button
$submitButton = New-Object System.Windows.Forms.Button
$submitButton.Text = "Submit"
$submitButton.Location = New-Object System.Drawing.Point(250, $yOffset + 20)
$form.Controls.Add($submitButton)

# Button Click Handler
$submitButton.Add_Click({
    $floor = $floorDropdown.SelectedItem
    $desk = $deskBox.Text.Trim()

    if (-not $floor -or -not $desk) {
        Show-Error "Floor and Desk fields must be filled."
        return
    }

    $newRows = @()
    $existing = @()
    if (Test-Path $csvPath) {
        $existing = Import-Csv $csvPath
    }

    $serialsSeen = @{}
    $tagsSeen = @{}

    foreach ($device in $devices) {
        $model = $entries[$device]["Model"].Text.Trim()
        $serial = $entries[$device]["Serial"].Text.Trim()
        $tag = $entries[$device]["Tag"].Text.Trim()

        $isDock = $device -eq "Dock"
        if (-not $isDock -and ($model -eq "" -or $serial -eq "" -or $tag -eq "")) {
            Show-Error "All fields for $device must be filled (except Dock)."
            return
        }

        # Check for same serial or tag in same device category
        if ($serial -ne "" -and $serialsSeen[$device] -contains $serial) {
            Show-Error "$device serial number '$serial' is entered more than once."
            return
        }
        if ($tag -ne "" -and $tagsSeen[$device] -contains $tag) {
            Show-Error "$device asset tag '$tag' is entered more than once."
            return
        }
        $serialsSeen[$device] += $serial
        $tagsSeen[$device] += $tag

        # Check in existing CSV
        foreach ($row in $existing) {
            if ($row.Device -eq $device -and $row."Serial Number" -eq $serial -and $serial -ne "") {
                Show-Error "$device serial number '$serial' already exists at floor $($row.Floor), desk $($row.Desk)"
                return
            }
            if ($row.Device -eq $device -and $row."Asset Tag" -eq $tag -and $tag -ne "") {
                Show-Error "$device asset tag '$tag' already exists at floor $($row.Floor), desk $($row.Desk)"
                return
            }
        }

        $newRows += [PSCustomObject]@{
            Floor = $floor
            Desk = $desk
            Device = $device
            Model = $model
            "Serial Number" = $serial
            "Asset Tag" = $tag
        }
    }

    # Save to CSV
    if (-not (Test-Path $csvPath)) {
        $newRows | Export-Csv -Path $csvPath -NoTypeInformation
    } else {
        $newRows | Export-Csv -Path $csvPath -NoTypeInformation -Append
    }

    [System.Windows.Forms.MessageBox]::Show("Inventory successfully submitted!", "Success", 'OK', 'Information')
    $form.Close()
})

# Run the form
[void]$form.ShowDialog()

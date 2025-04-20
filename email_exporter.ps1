Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName Microsoft.Office.Interop.Outlook

function Show-FolderBrowserDialog {
    $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    if ($folderBrowser.ShowDialog() -eq "OK") {
        return $folderBrowser.SelectedPath
    }
    return $null
}

function Get-MailFoldersRecursive {
    param([System.__ComObject]$Folder)

    $folders = @()
    foreach ($f in $Folder.Folders) {
        $folders += $f
        $folders += Get-MailFoldersRecursive -Folder $f
    }
    return $folders
}

function CloseIEWindow {
    $oWindows = (New-Object -ComObject Shell.Application).Windows

    foreach ($oWindow in $oWindows.Invoke()) {

        if ($oWindow.Fullname -match "IEXPLORE.EXE") {

            Write-verbose "Closing tab $($oWindow.LocationURL)"
            $oWindow.Quit()
        }
    }
}

# Initialize Outlook
$outlook = New-Object -ComObject Outlook.Application
$namespace = $outlook.GetNamespace("MAPI")
$rootFolder = $namespace.Folders.Item(1) # First mailbox
$allFolders = @($rootFolder) + (Get-MailFoldersRecursive -Folder $rootFolder)

# Create Form
$form = New-Object System.Windows.Forms.Form
$form.Text = "Export Outlook Emails"
$form.Size = New-Object System.Drawing.Size(450, 320)
$form.StartPosition = "CenterScreen"

# Folder selection button
$btnSelectFolder = New-Object System.Windows.Forms.Button
$btnSelectFolder.Text = "Select Export Folder"
$btnSelectFolder.Location = New-Object System.Drawing.Point(30,20)
$btnSelectFolder.Size = New-Object System.Drawing.Size(150,30)
$form.Controls.Add($btnSelectFolder)

# Label for selected folder
$lblFolderPath = New-Object System.Windows.Forms.Label
$lblFolderPath.Text = "No folder selected"
$lblFolderPath.AutoSize = $true
$lblFolderPath.Location = New-Object System.Drawing.Point(30,60)
$form.Controls.Add($lblFolderPath)

# Number of emails label and textbox
$lblCount = New-Object System.Windows.Forms.Label
$lblCount.Text = "Number of Emails to Export:"
$lblCount.Location = New-Object System.Drawing.Point(30,90)
$form.Controls.Add($lblCount)

$txtCount = New-Object System.Windows.Forms.TextBox
$txtCount.Location = New-Object System.Drawing.Point(230,90)
$txtCount.Size = New-Object System.Drawing.Size(50,20)
$txtCount.Text = "10"
$form.Controls.Add($txtCount)

# Format selection
$lblFormat = New-Object System.Windows.Forms.Label
$lblFormat.Text = "Export Format:"
$lblFormat.Location = New-Object System.Drawing.Point(30,120)
$form.Controls.Add($lblFormat)

$comboFormat = New-Object System.Windows.Forms.ComboBox
$comboFormat.Items.Add("MHT")
$comboFormat.Items.Add("PDF")
$comboFormat.SelectedIndex = 0
$comboFormat.Location = New-Object System.Drawing.Point(230,120)
$form.Controls.Add($comboFormat)

# Mail folder dropdown
$lblMailFolder = New-Object System.Windows.Forms.Label
$lblMailFolder.Text = "Outlook Folder to Export From:"
$lblMailFolder.Location = New-Object System.Drawing.Point(30,150)
$form.Controls.Add($lblMailFolder)

$comboMailFolder = New-Object System.Windows.Forms.ComboBox
$comboMailFolder.Location = New-Object System.Drawing.Point(30,170)
$comboMailFolder.Size = New-Object System.Drawing.Size(370,21)
$comboMailFolder.DropDownStyle = 'DropDownList'
$form.Controls.Add($comboMailFolder)

# Populate the folder list
$folderMap = @{}
foreach ($folder in $allFolders) {
    $path = $folder.FolderPath
    $comboMailFolder.Items.Add($path)
    $folderMap[$path] = $folder
}
$comboMailFolder.SelectedIndex = 0

# Start button
$btnStart = New-Object System.Windows.Forms.Button
$btnStart.Text = "Start Export"
$btnStart.Location = New-Object System.Drawing.Point(30,210)
$form.Controls.Add($btnStart)

$exportFolder = ""

$btnSelectFolder.Add_Click({
    $exportFolder = Show-FolderBrowserDialog
    if ($exportFolder) {
        $lblFolderPath.Text = $exportFolder
    }
})

$btnStart.Add_Click({
    $count = [int]$txtCount.Text
    $format = $comboFormat.SelectedItem
    $selectedPath = $comboMailFolder.SelectedItem
    $exportFolder = $lblFolderPath.Text

    if (-not $exportFolder) {
        [System.Windows.Forms.MessageBox]::Show($exportFolder)
        [System.Windows.Forms.MessageBox]::Show("Please select an export folder.")
        return
    }

    if (-not $folderMap.ContainsKey($selectedPath)) {
        [System.Windows.Forms.MessageBox]::Show("Invalid Outlook folder selected.")
        return
    }

    $mailFolder = $folderMap[$selectedPath]
    $allEmails = $mailFolder.Items | Sort-Object ReceivedTime -Descending
    $exported = 0

    foreach ($mail in $allEmails) {
        if ($exported -ge $count) {
            break
        }

        try {
            $subjectSafe = ($mail.Subject -replace '[\\/:*?"<>|]', '_')
            $timestamp = $mail.ReceivedTime.ToString("yyyyMMdd_HHmmss")
            $filename = "${subjectSafe}_$timestamp"
            $fullPath = Join-Path $exportFolder "$filename.$($format.ToLower())"

            if (Test-Path $fullPath) {
                Write-Host "Skipping existing: $filename"
                continue
            }

            if ($format -eq "MHT") {
                $mail.SaveAs($fullPath, 9) # 9 = olRFC822
            } elseif ($format -eq "PDF") {
                $tempMht = Join-Path $env:TEMP "$filename.mht"
                $pdfPath = Join-Path $exportFolder "$filename.pdf"
            
                # Save as MHT
                $mail.SaveAs($tempMht, 10)
            
                # Launch Internet Explorer to open MHT
                $ie = New-Object -ComObject InternetExplorer.Application
                $ie.Visible = $true
                $ie.Navigate($tempMht)
            
                # while ($ie.Busy -eq $true -or $ie.ReadyState -ne 4) {
                #     Start-Sleep -Milliseconds 500
                # }
            
                # Simulate CTRL+P and print to PDF using SendKeys
                $wshell = New-Object -ComObject wscript.shell
            
                # Start-Sleep -Milliseconds 1000
                # $ie.Visible = $true
                $wshell.AppActivate("Internet Explorer")
                Start-Sleep -Milliseconds 1000
                $wshell.SendKeys("^(p)")  # Ctrl + P to print
                Start-Sleep -Milliseconds 1000
                $wshell.SendKeys("%(n)")    # Alt+N for printer name
                Start-Sleep -Milliseconds 1000
                $wshell.SendKeys("Microsoft Print to PDF")
                Start-Sleep -Milliseconds 1000
                $wshell.SendKeys("%(p)") # Navigate to Print button
                Set-Clipboard -Value $pdfPath
                Start-Sleep -Milliseconds 1000
                $wshell.SendKeys("%(n)")
                # Save dialog: send path and enter
                Start-Sleep -Milliseconds 2000
                $wshell.SendKeys("^(v)")
                Start-Sleep -Milliseconds 1000
                $wshell.SendKeys("%(s)")
            
                # Cleanup
                Start-Sleep -Seconds 5
                CloseIEWindow
                Remove-Item $tempMht -ErrorAction SilentlyContinue
            }

            Write-Host "Exported: $filename"
            $exported++
        } catch {
            Write-Warning "Failed to export email: $_"
        }
    }

    if ($exported -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No new emails were exported.")
    } else {
        [System.Windows.Forms.MessageBox]::Show("Export completed. $exported emails exported.")
    }
})

$form.Topmost = $false
$form.Add_Shown({ $form.Activate() })
[void]$form.ShowDialog()

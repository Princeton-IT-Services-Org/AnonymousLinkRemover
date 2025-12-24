# Requires:
# 1. PowerShell 7 or later
# 2. Microsoft.Graph PowerShell module
# 3. Azure AD App with: Sites.Read.All, Sites.ReadWrite.All, Files.ReadWrite.All, Directory.Read.All, User.Read.All

# --- GLOBAL STATE ---
$script:FoundLinks = @()      
$script:SelectedLinks = @()   
$script:LogFilePath = Join-Path $PSScriptRoot "AnonymousLinkScan_$(Get-Date -Format 'yyyyMMdd_HHmm').log"

# --- 1. GUI DEFINITIONS ---

function Get-GlobalConfiguration {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing

    $Form = New-Object System.Windows.Forms.Form
    $Form.Text = "Princeton IT Services - Anonymous Link Manager"
    $Form.Size = [System.Drawing.Size]::new(550, 480)
    $Form.StartPosition = "CenterScreen"
    $Form.FormBorderStyle = "FixedSingle"
    $Form.MaximizeBox = $false

    $TitleLabel = New-Object System.Windows.Forms.Label
    $TitleLabel.Location = "20, 20"; $TitleLabel.Size = "500, 30"
    $TitleLabel.Text = "Tenant-Wide Anonymous Sharing Scan"
    $TitleLabel.Font = [System.Drawing.Font]::new("Arial", 14, [System.Drawing.FontStyle]::Bold)

    $Y = 70
    $LabelTenant = New-Object System.Windows.Forms.Label
    $LabelTenant.Location = "20, $Y"; $LabelTenant.Size = "500, 20"; $LabelTenant.Text = "Azure Tenant ID:"
    $TxtTenant = New-Object System.Windows.Forms.TextBox
    $TxtTenant.Location = "20, $($Y+20)"; $TxtTenant.Size = "480, 25"; $TxtTenant.PlaceholderText = "Enter Tenant ID (e.g. 00000000-0000...)"
    $TxtTenant.Text = "" 

    $Y += 60
    $LabelApp = New-Object System.Windows.Forms.Label
    $LabelApp.Location = "20, $Y"; $LabelApp.Size = "500, 20"; $LabelApp.Text = "Application (Client) ID:"
    $TxtApp = New-Object System.Windows.Forms.TextBox
    $TxtApp.Location = "20, $($Y+20)"; $TxtApp.Size = "480, 25"; $TxtApp.PlaceholderText = "Enter Application (Client) ID"
    $TxtApp.Text = "" 

    $Y += 60
    $LabelSecret = New-Object System.Windows.Forms.Label
    $LabelSecret.Location = "20, $Y"; $LabelSecret.Size = "500, 20"; $LabelSecret.Text = "Application Secret:"
    $TxtSecret = New-Object System.Windows.Forms.TextBox
    $TxtSecret.Location = "20, $($Y+20)"; $TxtSecret.Size = "480, 25"; $TxtSecret.PlaceholderText = "Enter Application Secret"
    $TxtSecret.UseSystemPasswordChar = $true
    $TxtSecret.Text = "" 

    $Y += 70
    $InfoLabel = New-Object System.Windows.Forms.Label
    $InfoLabel.Location = "20, $Y"; $InfoLabel.Size = "480, 40"
    $InfoLabel.Text = "The scan will explore SharePoint and OneDrives. You will be prompted to remove links once the scan is complete."
    $InfoLabel.ForeColor = [System.Drawing.Color]::DarkSlateGray

    $BtnRun = New-Object System.Windows.Forms.Button
    $BtnRun.Location = "20, $($Y+50)"; $BtnRun.Size = "480, 45"
    $BtnRun.Text = "Connect & Start Scan"; $BtnRun.BackColor = "#0078D4"; $BtnRun.ForeColor = "White"
    $BtnRun.Font = [System.Drawing.Font]::new("Arial", 10, [System.Drawing.FontStyle]::Bold)

    $BtnRun.Add_Click({
        if ([string]::IsNullOrWhiteSpace($TxtTenant.Text) -or [string]::IsNullOrWhiteSpace($TxtApp.Text) -or [string]::IsNullOrWhiteSpace($TxtSecret.Text)) {
            [System.Windows.Forms.MessageBox]::Show("Please fill in all credentials before continuing.", "Missing Information")
            return
        }
        $script:TenantId = $TxtTenant.Text
        $script:AppId = $TxtApp.Text
        $script:AppSecret = $TxtSecret.Text
        $Form.Close()
    })

    $Form.Controls.AddRange(@($TitleLabel, $LabelTenant, $TxtTenant, $LabelApp, $TxtApp, $LabelSecret, $TxtSecret, $InfoLabel, $BtnRun))
    $Form.ShowDialog() | Out-Null
}

function Show-InteractiveRemovalPrompt {
    $PromptForm = New-Object System.Windows.Forms.Form
    $PromptForm.Text = "Phase 2: Managed Link Removal"
    $PromptForm.Size = [System.Drawing.Size]::new(950, 950)
    $PromptForm.StartPosition = "CenterScreen"
    $PromptForm.FormBorderStyle = "FixedDialog"

    $HeaderFound = New-Object System.Windows.Forms.Label
    $HeaderFound.Text = "1. Found Anonymous Links (Apply Filters & Select Items)"; $HeaderFound.Location = "20, 15"; $HeaderFound.Size = "840, 25"
    $HeaderFound.Font = [System.Drawing.Font]::new("Arial", 10, [System.Drawing.FontStyle]::Bold)

    $Y = 45
    $LblFilter = New-Object System.Windows.Forms.Label
    $LblFilter.Text = "File Name:"; $LblFilter.Location = "20, $Y"; $LblFilter.Size = "70, 20"
    $TxtFilterFile = New-Object System.Windows.Forms.TextBox
    $TxtFilterFile.Location = "95, $Y"; $TxtFilterFile.Size = "180, 25"

    $LblFilterSite = New-Object System.Windows.Forms.Label
    $LblFilterSite.Text = "Site/UPN:"; $LblFilterSite.Location = "290, $Y"; $LblFilterSite.Size = "70, 20"
    $TxtFilterSite = New-Object System.Windows.Forms.TextBox
    $TxtFilterSite.Location = "365, $Y"; $TxtFilterSite.Size = "180, 25"

    $BtnFilter = New-Object System.Windows.Forms.Button
    $BtnFilter.Text = "Filter"; $BtnFilter.Location = "560, $Y"; $BtnFilter.Size = "80, 25"

    $BtnSelectVisible = New-Object System.Windows.Forms.Button
    $BtnSelectVisible.Text = "Select All Visible"; $BtnSelectVisible.Location = "650, $Y"; $BtnSelectVisible.Size = "120, 25"
    
    $BtnClearFilter = New-Object System.Windows.Forms.Button
    $BtnClearFilter.Text = "Reset Filters"; $BtnClearFilter.Location = "780, $Y"; $BtnClearFilter.Size = "100, 25"

    $Y += 40
    $GridFound = New-Object System.Windows.Forms.DataGridView
    $GridFound.Location = "20, $Y"; $GridFound.Size = "890, 250"
    $GridFound.AllowUserToAddRows = $false; $GridFound.RowHeadersVisible = $false
    $GridFound.AutoSizeColumnsMode = "Fill"; $GridFound.BackgroundColor = "White"

    $colCheckF = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
    $colCheckF.HeaderText = "Select"; $colCheckF.Name = "Select"; $colCheckF.Width = 50
    $GridFound.Columns.Add($colCheckF) | Out-Null
    $GridFound.Columns.Add((New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{HeaderText="File Name"; Name="FileName"; ReadOnly=$true})) | Out-Null
    $GridFound.Columns.Add((New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{HeaderText="Site / UPN"; Name="OwnerSite"; ReadOnly=$true})) | Out-Null

    $Y += 270
    $HeaderSelected = New-Object System.Windows.Forms.Label
    $HeaderSelected.Text = "2. Selected for Revocation (Staging Area)"; $HeaderSelected.Location = "20, $Y"; $HeaderSelected.Size = "840, 25"
    $HeaderSelected.Font = [System.Drawing.Font]::new("Arial", 10, [System.Drawing.FontStyle]::Bold)
    $HeaderSelected.ForeColor = [System.Drawing.Color]::DarkRed

    $Y += 30
    $GridSelected = New-Object System.Windows.Forms.DataGridView
    $GridSelected.Location = "20, $Y"; $GridSelected.Size = "890, 200"
    $GridSelected.AllowUserToAddRows = $false; $GridSelected.RowHeadersVisible = $false
    $GridSelected.AutoSizeColumnsMode = "Fill"; $GridSelected.BackgroundColor = "#FFF5F5"

    $colCheckS = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
    $colCheckS.HeaderText = "Keep"; $colCheckS.Name = "Select"; $colCheckS.Width = 50
    $GridSelected.Columns.Add($colCheckS) | Out-Null
    $GridSelected.Columns.Add((New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{HeaderText="File Name"; Name="FileName"; ReadOnly=$true})) | Out-Null
    $GridSelected.Columns.Add((New-Object System.Windows.Forms.DataGridViewTextBoxColumn -Property @{HeaderText="Site / UPN"; Name="OwnerSite"; ReadOnly=$true})) | Out-Null

    function Update-Grids {
        $foundVisible = $script:FoundLinks | Where-Object { 
            $_.FileName -ilike "*$($TxtFilterFile.Text)*" -and $_.OwnerSite -ilike "*$($TxtFilterSite.Text)*" 
        }

        $GridFound.Rows.Clear()
        foreach ($item in $foundVisible) {
            $index = $GridFound.Rows.Add()
            $GridFound.Rows[$index].Cells["Select"].Value = $false
            $GridFound.Rows[$index].Cells["FileName"].Value = $item.FileName
            $GridFound.Rows[$index].Cells["OwnerSite"].Value = $item.OwnerSite
            $GridFound.Rows[$index].Tag = $item 
        }

        $GridSelected.Rows.Clear()
        foreach ($item in $script:SelectedLinks) {
            $index = $GridSelected.Rows.Add()
            $GridSelected.Rows[$index].Cells["Select"].Value = $true 
            $GridSelected.Rows[$index].Cells["FileName"].Value = $item.FileName
            $GridSelected.Rows[$index].Cells["OwnerSite"].Value = $item.OwnerSite
            $GridSelected.Rows[$index].Tag = $item 
        }
        $HeaderFound.Text = "1. Found Anonymous Links ($($script:FoundLinks.Count) Remaining)"
        $HeaderSelected.Text = "2. Selected for Revocation: ($($script:SelectedLinks.Count))"
    }

    $GridFound.Add_CellValueChanged({
        param($s, $e)
        if ($e.ColumnIndex -eq 0 -and $e.RowIndex -ge 0) {
            $row = $GridFound.Rows[$e.RowIndex]
            if ($row.Cells[0].Value -eq $true) {
                $item = $row.Tag
                $script:SelectedLinks += $item
                $script:FoundLinks = $script:FoundLinks | Where-Object { $_.PermissionId -ne $item.PermissionId }
                $PromptForm.BeginInvoke({ Update-Grids }) | Out-Null
            }
        }
    })

    $GridSelected.Add_CellValueChanged({
        param($s, $e)
        if ($e.ColumnIndex -eq 0 -and $e.RowIndex -ge 0) {
            $row = $GridSelected.Rows[$e.RowIndex]
            if ($row.Cells[0].Value -eq $false) {
                $item = $row.Tag
                $script:FoundLinks += $item
                $script:SelectedLinks = $script:SelectedLinks | Where-Object { $_.PermissionId -ne $item.PermissionId }
                $PromptForm.BeginInvoke({ Update-Grids }) | Out-Null
            }
        }
    })

    $GridFound.Add_CurrentCellDirtyStateChanged({ if ($GridFound.IsCurrentCellDirty) { $GridFound.CommitEdit([System.Windows.Forms.DataGridViewDataErrorContexts]::Commit) } })
    $GridSelected.Add_CurrentCellDirtyStateChanged({ if ($GridSelected.IsCurrentCellDirty) { $GridSelected.CommitEdit([System.Windows.Forms.DataGridViewDataErrorContexts]::Commit) } })

    $BtnFilter.Add_Click({ Update-Grids })
    $BtnClearFilter.Add_Click({ $TxtFilterFile.Text = ""; $TxtFilterSite.Text = ""; Update-Grids })
    $BtnSelectVisible.Add_Click({
        $visibleItems = @()
        foreach($row in $GridFound.Rows) { $visibleItems += $row.Tag }
        if ($visibleItems.Count -gt 0) {
            $script:SelectedLinks += $visibleItems
            $visibleIds = $visibleItems.PermissionId
            $script:FoundLinks = $script:FoundLinks | Where-Object { $_.PermissionId -notin $visibleIds }
            Update-Grids
        }
    })

    $Y += 220
    $TxtConsole = New-Object System.Windows.Forms.TextBox
    $TxtConsole.Location = "20, $Y"; $TxtConsole.Size = "890, 180"; $TxtConsole.Multiline = $true; $TxtConsole.ReadOnly = $true; $TxtConsole.ScrollBars = "Vertical"
    $TxtConsole.BackColor = "Black"; $TxtConsole.ForeColor = "LightGreen"; $TxtConsole.Font = [System.Drawing.Font]::new("Consolas", 9)

    $Y += 190
    $BtnRevoke = New-Object System.Windows.Forms.Button
    $BtnRevoke.Text = "Revoke All Selected Links"; $BtnRevoke.Location = "20, $Y"; $BtnRevoke.Size = "435, 45"
    $BtnRevoke.BackColor = "#D83B01"; $BtnRevoke.ForeColor = "White"; $BtnRevoke.Font = [System.Drawing.Font]::new("Arial", 10, [System.Drawing.FontStyle]::Bold)

    $BtnClose = New-Object System.Windows.Forms.Button
    $BtnClose.Text = "Exit"; $BtnClose.Location = "475, $Y"; $BtnClose.Size = "435, 45"

    $BtnRevoke.Add_Click({
        if ($script:SelectedLinks.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Staging area is empty.", "No Selection")
            return
        }

        $msg = "Are you sure you want to revoke anonymous access for all $($script:SelectedLinks.Count) selected files?"
        if ([System.Windows.Forms.MessageBox]::Show($msg, "Confirm Bulk Revocation", "YesNo") -eq "Yes") {
            $BtnRevoke.Enabled = $false
            $TxtConsole.AppendText("Starting bulk revocation...`r`n")
            
            $toProcess = $script:SelectedLinks
            $script:SelectedLinks = @() 

            foreach ($item in $toProcess) {
                try {
                    $TxtConsole.AppendText("Processing: $($item.FileName)... ")
                    Remove-MgDriveItemPermission -DriveId $item.DriveId -DriveItemId $item.ItemId -PermissionId $item.PermissionId -Confirm:$false -ErrorAction Stop
                    $TxtConsole.AppendText("SUCCESS`r`n")
                } catch {
                    $TxtConsole.AppendText("FAILED: $($_.Exception.Message)`r`n")
                    $script:SelectedLinks += $item 
                }
            }
            
            $TxtConsole.AppendText("Bulk operation complete.`r`n")
            Update-Grids
            $BtnRevoke.Enabled = $true
        }
    })

    $BtnClose.Add_Click({ $PromptForm.Close() })

    $PromptForm.Controls.AddRange(@($HeaderFound, $LblFilter, $TxtFilterFile, $LblFilterSite, $TxtFilterSite, $BtnFilter, $BtnSelectVisible, $BtnClearFilter, $GridFound, $HeaderSelected, $GridSelected, $TxtConsole, $BtnRevoke, $BtnClose))
    
    Update-Grids
    $PromptForm.ShowDialog() | Out-Null
}

# --- 2. LOGIC FUNCTIONS ---

function Invoke-AnonymousSharingScan {
    param ([string]$DriveId, [string]$DriveName, [string]$SiteName, [string]$ItemId = "root", [string]$ItemPath = "")
    
    if ($ItemId -ne "root") {
        try {
            $permissions = Get-MgDriveItemPermission -DriveId $DriveId -DriveItemId $ItemId -ErrorAction Stop
            foreach ($perm in $permissions) {
                if ($perm.Link -ne $null -and $perm.Link.Scope -eq "anonymous") {
                    $script:FoundLinks += [PSCustomObject]@{
                        FileName     = ($ItemPath -split '/' | Select-Object -Last 1)
                        FileFullPath = $ItemPath
                        OwnerSite    = $SiteName
                        DriveName    = $DriveName
                        LinkType     = $perm.Link.Type
                        LinkUrl      = $perm.Link.WebUrl
                        DriveId      = $DriveId
                        ItemId       = $ItemId
                        PermissionId = $perm.Id
                    }
                    Write-Host "    [FOUND] Anonymous link on: $($ItemPath)" -ForegroundColor Cyan
                }
            }
        } catch {}
    }

    try {
        $children = Get-MgDriveItemChild -DriveId $DriveId -DriveItemId $ItemId -All -ErrorAction SilentlyContinue
        foreach ($child in $children) {
            $path = if ($ItemPath -eq "") { $child.Name } else { "$ItemPath/$($child.Name)" }
            Invoke-AnonymousSharingScan -DriveId $DriveId -DriveName $DriveName -SiteName $SiteName -ItemId $child.Id -ItemPath $path
        }
    } catch {}
}

# --- 3. MAIN EXECUTION ---

Get-GlobalConfiguration

if (-not $script:AppSecret) { exit }

Start-Transcript -Path $script:LogFilePath

try {
    Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
    $SecureSecret = ConvertTo-SecureString -String $script:AppSecret -AsPlainText -Force
    $AppCredential = New-Object System.Management.Automation.PSCredential($script:AppId, $SecureSecret)
    Connect-MgGraph -TenantId $script:TenantId -Credential $AppCredential -NoWelcome

    Write-Host "`n--- Scanning SharePoint Sites ---" -ForegroundColor Yellow
    $sites = Get-MgSite -All -ErrorAction Stop
    foreach ($site in $sites) {
        if ($site.WebUrl -ilike "*personal*") { continue }
        Write-Host "Scanning Site: $($site.DisplayName) ($($site.WebUrl))" -ForegroundColor Gray
        try {
            $drives = Get-MgSiteDrive -SiteId $site.Id -All -ErrorAction Stop
            foreach ($drive in $drives) {
                Invoke-AnonymousSharingScan -DriveId $drive.Id -DriveName $drive.Name -SiteName $site.DisplayName
            }
        } catch { }
    }

    Write-Host "`n--- Scanning User OneDrives ---" -ForegroundColor Yellow
    $allUsers = Get-MgUser -All -ErrorAction Stop
    foreach ($user in $allUsers) {
        $upn = $user.UserPrincipalName
        Write-Host "Scanning OneDrive: $upn" -ForegroundColor Gray
        try {
            $drives = Get-MgUserDrive -UserId $upn -All -ErrorAction Stop | Where-Object { $_.DriveType -eq "business" }
            foreach ($drive in $drives) {
                Invoke-AnonymousSharingScan -DriveId $drive.Id -DriveName $drive.Name -SiteName $upn
            }
        } catch { }
    }

    $script:FoundLinks = $script:FoundLinks | Group-Object PermissionId | ForEach-Object { $_.Group[0] }

    if ($script:FoundLinks.Count -gt 0) {
        Write-Host "`nScan Complete. Found $($script:FoundLinks.Count) unique anonymous links. Launching removal GUI..." -ForegroundColor Green
        Show-InteractiveRemovalPrompt
    } else {
        [System.Windows.Forms.MessageBox]::Show("No anonymous links found in the tenant.", "Scan Complete")
        Write-Host "`nScan Complete. No anonymous links found." -ForegroundColor Gray
    }

} catch {
    Write-Error $_.Exception.Message
} finally {
    Stop-Transcript
}
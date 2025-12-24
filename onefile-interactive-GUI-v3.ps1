# Requires:
# 1. PowerShell 7 or later
# 2. Microsoft.Graph PowerShell module
# 3. Azure AD App with: Sites.Read.All, Sites.ReadWrite.All, Files.ReadWrite.All, Directory.Read.All, User.Read.All

# --- GLOBAL STATE ---
$script:AnonymousLinksFound = @()
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
    $TxtTenant.Location = "20, $($Y+20)"; $TxtTenant.Size = "480, 25"; $TxtTenant.Text = "63892a6f-84bf-4eab-99f3-f49e0cf0d4f8"

    $Y += 60
    $LabelApp = New-Object System.Windows.Forms.Label
    $LabelApp.Location = "20, $Y"; $LabelApp.Size = "500, 20"; $LabelApp.Text = "Application (Client) ID:"
    $TxtApp = New-Object System.Windows.Forms.TextBox
    $TxtApp.Location = "20, $($Y+20)"; $TxtApp.Size = "480, 25"; $TxtApp.Text = "77d06073-9de0-4fe2-be5f-64befd327653"

    $Y += 60
    $LabelSecret = New-Object System.Windows.Forms.Label
    $LabelSecret.Location = "20, $Y"; $LabelSecret.Size = "500, 20"; $LabelSecret.Text = "Application Secret:"
    $TxtSecret = New-Object System.Windows.Forms.TextBox
    $TxtSecret.Location = "20, $($Y+20)"; $TxtSecret.Size = "480, 25"; $TxtSecret.UseSystemPasswordChar = $true
    $TxtSecret.Text = "5wl8Q~AOCGdEKT63hntOKbEZ8c~QEKGv1SS2WaaL"

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
    $PromptForm.Text = "Phase 2: Multi-Select Link Removal"
    $PromptForm.Size = [System.Drawing.Size]::new(900, 800)
    $PromptForm.StartPosition = "CenterScreen"
    $PromptForm.FormBorderStyle = "FixedDialog"

    $Header = New-Object System.Windows.Forms.Label
    $Header.Text = "Step 1: Found $($script:AnonymousLinksFound.Count) Anonymous Links. Select items to revoke."; $Header.Location = "20, 20"; $Header.Size = "840, 25"
    $Header.Font = [System.Drawing.Font]::new("Arial", 11, [System.Drawing.FontStyle]::Bold)

    # --- Data Grid for Multi-Select ---
    $Grid = New-Object System.Windows.Forms.DataGridView
    $Grid.Location = "20, 60"
    $Grid.Size = "840, 300"
    $Grid.AllowUserToAddRows = $false
    $Grid.RowHeadersVisible = $false
    $Grid.SelectionMode = "FullRowSelect"
    $Grid.AutoSizeColumnsMode = "Fill"

    # Add Columns
    $colCheck = New-Object System.Windows.Forms.DataGridViewCheckBoxColumn
    $colCheck.HeaderText = "Select"; $colCheck.Name = "Select"; $colCheck.Width = 50
    $Grid.Columns.Add($colCheck) | Out-Null

    $colFile = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colFile.HeaderText = "File Name"; $colFile.Name = "FileName"; $colFile.ReadOnly = $true
    $Grid.Columns.Add($colFile) | Out-Null

    $colOwner = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $colOwner.HeaderText = "Site / UPN"; $colOwner.Name = "OwnerSite"; $colOwner.ReadOnly = $true
    $Grid.Columns.Add($colOwner) | Out-Null

    # Populate Initial Data
    function Update-Grid {
        param($Data)
        $Grid.Rows.Clear()
        foreach ($item in $Data) {
            $index = $Grid.Rows.Add()
            $Grid.Rows[$index].Cells["Select"].Value = $false
            $Grid.Rows[$index].Cells["FileName"].Value = $item.FileName
            $Grid.Rows[$index].Cells["OwnerSite"].Value = $item.OwnerSite
            $Grid.Rows[$index].Tag = $item # Store full object in tag for retrieval
        }
    }
    Update-Grid -Data $script:AnonymousLinksFound

    # --- Search / Filter Section ---
    $Y = 380
    $LblFilter = New-Object System.Windows.Forms.Label
    $LblFilter.Text = "Filter by Name:"; $LblFilter.Location = "20, $Y"; $LblFilter.Size = "100, 20"
    $TxtFilterFile = New-Object System.Windows.Forms.TextBox
    $TxtFilterFile.Location = "130, $Y"; $TxtFilterFile.Size = "250, 25"

    $LblFilterSite = New-Object System.Windows.Forms.Label
    $LblFilterSite.Text = "Filter by Site:"; $LblFilterSite.Location = "400, $Y"; $LblFilterSite.Size = "100, 20"
    $TxtFilterSite = New-Object System.Windows.Forms.TextBox
    $TxtFilterSite.Location = "510, $Y"; $TxtFilterSite.Size = "250, 25"

    $BtnFilter = New-Object System.Windows.Forms.Button
    $BtnFilter.Text = "Apply Filters"; $BtnFilter.Location = "770, $Y"; $BtnFilter.Size = "90, 25"

    $BtnFilter.Add_Click({
        $filtered = $script:AnonymousLinksFound | Where-Object { 
            $_.FileName -ilike "*$($TxtFilterFile.Text)*" -and $_.OwnerSite -ilike "*$($TxtFilterSite.Text)*" 
        }
        Update-Grid -Data $filtered
    })

    # --- Actions ---
    $Y = 430
    $BtnSelectAll = New-Object System.Windows.Forms.Button
    $BtnSelectAll.Text = "Select All Visible"; $BtnSelectAll.Location = "20, $Y"; $BtnSelectAll.Size = "150, 30"
    $BtnSelectAll.Add_Click({ foreach($row in $Grid.Rows) { $row.Cells["Select"].Value = $true } })

    $BtnDeselectAll = New-Object System.Windows.Forms.Button
    $BtnDeselectAll.Text = "Deselect All"; $BtnDeselectAll.Location = "180, $Y"; $BtnDeselectAll.Size = "150, 30"
    $BtnDeselectAll.Add_Click({ foreach($row in $Grid.Rows) { $row.Cells["Select"].Value = $false } })

    # Console-like output for progress
    $Y = 480
    $TxtConsole = New-Object System.Windows.Forms.TextBox
    $TxtConsole.Location = "20, $Y"; $TxtConsole.Size = "840, 200"; $TxtConsole.Multiline = $true; $TxtConsole.ReadOnly = $true; $TxtConsole.ScrollBars = "Vertical"
    $TxtConsole.BackColor = "Black"; $TxtConsole.ForeColor = "LightGreen"
    $TxtConsole.Font = [System.Drawing.Font]::new("Consolas", 9, [System.Drawing.FontStyle]::Regular)

    $Y = 700
    $BtnRevoke = New-Object System.Windows.Forms.Button
    $BtnRevoke.Text = "Revoke Selected Links"; $BtnRevoke.Location = "20, $Y"; $BtnRevoke.Size = "410, 45"
    $BtnRevoke.BackColor = "#D83B01"; $BtnRevoke.ForeColor = "White"; $BtnRevoke.Font = [System.Drawing.Font]::new("Arial", 10, [System.Drawing.FontStyle]::Bold)

    $BtnClose = New-Object System.Windows.Forms.Button
    $BtnClose.Text = "Exit Application"; $BtnClose.Location = "450, $Y"; $BtnClose.Size = "410, 45"

    $BtnRevoke.Add_Click({
        $selectedItems = @()
        foreach ($row in $Grid.Rows) {
            if ($row.Cells["Select"].Value -eq $true) { $selectedItems += $row.Tag }
        }

        if ($selectedItems.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("Please select at least one file to revoke.", "No Selection")
            return
        }

        $msg = "Are you sure you want to revoke anonymous access for $($selectedItems.Count) selected files?"
        if ([System.Windows.Forms.MessageBox]::Show($msg, "Confirm Bulk Revocation", "YesNo") -eq "Yes") {
            $BtnRevoke.Enabled = $false
            $TxtConsole.AppendText("Starting bulk revocation...`r`n")
            
            foreach ($item in $selectedItems) {
                try {
                    $TxtConsole.AppendText("Processing: $($item.FileName)... ")
                    Remove-MgDriveItemPermission -DriveId $item.DriveId -DriveItemId $item.ItemId -PermissionId $item.PermissionId -Confirm:$false -ErrorAction Stop
                    $TxtConsole.AppendText("SUCCESS`r`n")
                    # Remove from master list
                    $script:AnonymousLinksFound = $script:AnonymousLinksFound | Where-Object { $_.PermissionId -ne $item.PermissionId }
                } catch {
                    $TxtConsole.AppendText("FAILED: $($_.Exception.Message)`r`n")
                }
            }
            
            $TxtConsole.AppendText("Bulk operation complete.`r`n")
            Update-Grid -Data ($script:AnonymousLinksFound | Where-Object { 
                $_.FileName -ilike "*$($TxtFilterFile.Text)*" -and $_.OwnerSite -ilike "*$($TxtFilterSite.Text)*" 
            })
            $Header.Text = "Remaining Unique Links Found: $($script:AnonymousLinksFound.Count)"
            $BtnRevoke.Enabled = $true
        }
    })

    $BtnClose.Add_Click({ $PromptForm.Close() })

    $PromptForm.Controls.AddRange(@($Header, $Grid, $LblFilter, $TxtFilterFile, $LblFilterSite, $TxtFilterSite, $BtnFilter, $BtnSelectAll, $BtnDeselectAll, $TxtConsole, $BtnRevoke, $BtnClose))
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
                    $script:AnonymousLinksFound += [PSCustomObject]@{
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

    # Phase A: SharePoint
    Write-Host "PHASE 1A: Scanning SharePoint..." -ForegroundColor Yellow
    $sites = Get-MgSite -All -ErrorAction Stop
    foreach ($site in $sites) {
        if ($site.WebUrl -ilike "*personal*") { continue }
        try {
            $drives = Get-MgSiteDrive -SiteId $site.Id -All -ErrorAction Stop
            foreach ($drive in $drives) {
                Invoke-AnonymousSharingScan -DriveId $drive.Id -DriveName $drive.Name -SiteName $site.DisplayName
            }
        } catch { Write-Host "Skipping Drive scan for $($site.DisplayName)" -ForegroundColor Gray }
    }

    # Phase B: OneDrives
    Write-Host "PHASE 1B: Scanning OneDrives..." -ForegroundColor Yellow
    $allUsers = Get-MgUser -All -ErrorAction Stop
    foreach ($user in $allUsers) {
        $upn = $user.UserPrincipalName
        try {
            $drives = Get-MgUserDrive -UserId $upn -All -ErrorAction Stop | Where-Object { $_.DriveType -eq "business" }
            foreach ($drive in $drives) {
                Invoke-AnonymousSharingScan -DriveId $drive.Id -DriveName $drive.Name -SiteName $upn
            }
        } catch {
            if ($_.Exception.Message -ilike "*ResourceNotFound*") {
                Write-Host "Skipping User $upn (No MySite/OneDrive found)" -ForegroundColor DarkGray
            } else {
                Write-Host "Error accessing OneDrive for ${upn}: $($_.Exception.Message)" -ForegroundColor Red
            }
        }
    }

    # De-duplicate
    $script:AnonymousLinksFound = $script:AnonymousLinksFound | Group-Object PermissionId | ForEach-Object { $_.Group[0] }

    # Phase 2: Interactive Removal
    if ($script:AnonymousLinksFound.Count -gt 0) {
        Show-InteractiveRemovalPrompt
    } else {
        [System.Windows.Forms.MessageBox]::Show("No anonymous links found in the tenant.", "Scan Complete")
    }

} catch {
    Write-Error $_.Exception.Message
} finally {
    Stop-Transcript
}
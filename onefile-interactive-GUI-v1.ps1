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
    $PromptForm.Text = "Phase 2: Interactive Link Removal"
    $PromptForm.Size = [System.Drawing.Size]::new(600, 500)
    $PromptForm.StartPosition = "CenterScreen"
    $PromptForm.FormBorderStyle = "FixedDialog"

    $Header = New-Object System.Windows.Forms.Label
    $Header.Text = "Total Unique Links Found: $($script:AnonymousLinksFound.Count)"; $Header.Location = "20, 20"; $Header.Size = "540, 25"
    $Header.Font = [System.Drawing.Font]::new("Arial", 11, [System.Drawing.FontStyle]::Bold)

    $Y = 60
    $LblFile = New-Object System.Windows.Forms.Label
    $LblFile.Text = "[STEP 1] Enter File Name (Partial Match Allowed):"; $LblFile.Location = "20, $Y"; $LblFile.Size = "500, 20"
    $TxtFile = New-Object System.Windows.Forms.TextBox
    $TxtFile.Location = "20, $($Y+20)"; $TxtFile.Size = "540, 25"

    $Y += 60
    $LblSite = New-Object System.Windows.Forms.Label
    $LblSite.Text = "[STEP 2] Enter Site Name or User UPN:"; $LblSite.Location = "20, $Y"; $LblSite.Size = "500, 20"
    $TxtSite = New-Object System.Windows.Forms.TextBox
    $TxtSite.Location = "20, $($Y+20)"; $TxtSite.Size = "540, 25"

    $Y += 60
    $BtnSearch = New-Object System.Windows.Forms.Button
    $BtnSearch.Text = "Search & Review Link"; $BtnSearch.Location = "20, $Y"; $BtnSearch.Size = "540, 40"
    $BtnSearch.BackColor = "#F3F2F1"; $BtnSearch.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Bold)

    $Y += 60
    $TxtResult = New-Object System.Windows.Forms.TextBox
    $TxtResult.Location = "20, $Y"; $TxtResult.Size = "540, 100"; $TxtResult.Multiline = $true; $TxtResult.ReadOnly = $true
    $TxtResult.ScrollBars = "Vertical"

    $Y += 110
    $BtnRevoke = New-Object System.Windows.Forms.Button
    $BtnRevoke.Text = "Revoke Identified Link"; $BtnRevoke.Location = "20, $Y"; $BtnRevoke.Size = "260, 40"
    $BtnRevoke.Enabled = $false; $BtnRevoke.BackColor = "#D83B01"; $BtnRevoke.ForeColor = "White"; $BtnRevoke.Font = [System.Drawing.Font]::new("Arial", 9, [System.Drawing.FontStyle]::Bold)

    $BtnClose = New-Object System.Windows.Forms.Button
    $BtnClose.Text = "Exit"; $BtnClose.Location = "300, $Y"; $BtnClose.Size = "260, 40"

    # Search Logic
    $BtnSearch.Add_Click({
        $FileNameInput = $TxtFile.Text
        $SiteOwnerInput = $TxtSite.Text
        $script:TargetedLinks = $script:AnonymousLinksFound | Where-Object { 
            $_.FileName -ilike "*$FileNameInput*" -and $_.OwnerSite -ilike "*$SiteOwnerInput*" 
        }

        if ($script:TargetedLinks.Count -eq 0) {
            $TxtResult.Text = "ERROR: No matching links found."
            $BtnRevoke.Enabled = $false
        } elseif ($script:TargetedLinks.Count -gt 1) {
            $TxtResult.Text = "WARNING: Multiple matches found ($($script:TargetedLinks.Count)). Please refine your search.`r`n`r`n"
            foreach($l in $script:TargetedLinks) { $TxtResult.Text += " - $($l.FileName) in $($l.OwnerSite)`r`n" }
            $BtnRevoke.Enabled = $false
        } else {
            $target = $script:TargetedLinks[0]
            $TxtResult.Text = "MATCH FOUND:`r`nFile: $($target.FileName)`r`nSite: $($target.OwnerSite)`r`nType: $($target.LinkType)`r`nURL: $($target.LinkUrl)"
            $BtnRevoke.Enabled = $true
        }
    })

    # Revoke Logic
    $BtnRevoke.Add_Click({
        $target = $script:TargetedLinks[0]
        $msg = "Are you sure you want to revoke the anonymous link for '$($target.FileName)'?"
        if ([System.Windows.Forms.MessageBox]::Show($msg, "Confirm", "YesNo") -eq "Yes") {
            try {
                Remove-MgDriveItemPermission -DriveId $target.DriveId -DriveItemId $target.ItemId -PermissionId $target.PermissionId -Confirm:$false -ErrorAction Stop
                [System.Windows.Forms.MessageBox]::Show("Success: Link Revoked.", "Done")
                $script:AnonymousLinksFound = $script:AnonymousLinksFound | Where-Object { $_.PermissionId -ne $target.PermissionId }
                $Header.Text = "Total Unique Links Found: $($script:AnonymousLinksFound.Count)"
                $TxtResult.Text = ""; $BtnRevoke.Enabled = $false
            } catch {
                [System.Windows.Forms.MessageBox]::Show("Error: $($_.Exception.Message)", "Error")
            }
        }
    })

    $BtnClose.Add_Click({ $PromptForm.Close() })

    $PromptForm.Controls.AddRange(@($Header, $LblFile, $TxtFile, $LblSite, $TxtSite, $BtnSearch, $TxtResult, $BtnRevoke, $BtnClose))
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
                # FIXED: Wrapped upn in curly braces to fix ParserError on the colon
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
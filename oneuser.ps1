# --- CONFIGURATION (Application Credentials for Tenant-Wide Access) ---

$TenantId = "<replace w ur tenant id>"
$AppId = "<replace w ur app id>"
$AppSecret = "<replace w ur app secret>"

$Scopes = @(
    "Files.ReadWrite.All",
    "Directory.Read.All",
    "User.Read.All"
)

# --- CONNECT AND INITIALIZE ---

# Check if Application Credentials are provided
if ($TenantId -and $AppId -and $AppSecret) {
    Write-Host "Connecting to Microsoft Graph using Application Permissions (Client Secret)..." -ForegroundColor Cyan
    
    try {
        # FIX 1: Convert the App Secret string to a SecureString.
        $SecureSecret = ConvertTo-SecureString -String $AppSecret -AsPlainText -Force

        # FIX 2: Construct a PSCredential object. The App ID acts as the username.
        $AppCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $AppId, $SecureSecret

        # FIX 3: Use the -Credential parameter alone with -TenantId. This is the official MSFT application auth pattern
        # that avoids parameter set conflicts with -ClientId.
        Connect-MgGraph -TenantId $TenantId -Credential $AppCredential -ErrorAction Stop
        
    } catch {
        Write-Error "Failed to connect using Application Permissions. Check your Tenant ID, App ID, and App Secret. Error: $($_.Exception.Message)"
        exit 1
    }
} else {
    Write-Host "Connecting to Microsoft Graph with required Delegated scopes: $($Scopes -join ', ')" -ForegroundColor Cyan
    Connect-MgGraph -Scopes $Scopes -ErrorAction Stop
}


# Define the recursive cleanup function
function Remove-AnonymousSharing {
    param (
        [string]$DriveId,
        [string]$DriveName,
        [string]$ItemId,
        [string]$ItemPath = ""
    )
    
    # 1. Get permissions for the current item
    if ($ItemId -ne "root") {
        try {
            # Check for permissions on the current item
            $permissions = Get-MgDriveItemPermission -DriveId $DriveId -DriveItemId $ItemId -ErrorAction Stop
            
            foreach ($perm in $permissions) {
                # Identify anonymous sharing links: must have a 'link' property and the scope must be 'anonymous'
                if ($perm.Link -ne $null -and $perm.Link.Scope -eq "anonymous") {
                    Write-Host "    [REVOKING] Anonymous link found on: $ItemPath\$($perm.Link.Type) - ID: $($perm.Id)" -ForegroundColor Yellow
                    
                    # Revoke the anonymous link
                    Remove-MgDriveItemPermission -DriveId $DriveId -DriveItemId $ItemId -PermissionId $perm.Id -Confirm:$false
                    
                    Write-Host "    [SUCCESS] Revoked anonymous link for item '$ItemPath'" -ForegroundColor Green
                }
            }
        } catch {
            Write-Warning "    Failed to process permissions for item '$ItemPath' (Error: $($_.Exception.Message)). Skipping."
        }
    }

    # 2. Get child items and recurse
    try {
        # Check if the item is a folder by attempting to list children.
        $children = Get-MgDriveItemChild -DriveId $DriveId -DriveItemId $ItemId -All -ErrorAction Stop
        
        foreach ($child in $children) {
            $childPath = if ($ItemPath -eq "") { $child.Name } else { "$ItemPath/$($child.Name)" }
            
            # Recurse for all items (files and folders)
            Remove-AnonymousSharing -DriveId $DriveId -DriveName $DriveName -ItemId $child.Id -ItemPath $childPath
        }
    } catch {
        # Suppress common expected errors when trying to list children of a file,
        # or when the SDK fails on a non-folder item.
        $errorMessage = $_.Exception.Message
        if (-not ($errorMessage -like "*Item is a file*" -or $errorMessage -like "*Object reference not set to an instance of an object*")) {
            Write-Warning "    Failed to list children in '$DriveName/$ItemPath' (Error: $errorMessage). Skipping folder/file children."
        }
    }
}

# --- 2. PROCESS GIVEN USER'S ONEDRIVE ---

Write-Host "`n=== Starting User's OneDrive Cleanup ===`n" -ForegroundColor Cyan


    $UserUPN = "<replace with specific user upn>"
    Write-Host "`nProcessing OneDrive for User: $UserUPN" -ForegroundColor Yellow
    
    try {
        # Get the business drive for the user (their OneDrive for Business)
        # Using Application Permissions is the ONLY way to reliably bypass the 403 error here.
        $drives = Get-MgUserDrive -UserId $UserUPN -All -ErrorAction Stop | Where-Object { $_.DriveType -eq "business" }
    } catch {
        # IMPORTANT: A 403 (Access Denied) or 404 (ResourceNotFound) error here is common and often expected.
        # 403: Indicates the current admin, even with Delegated Permissions, lacks site collection admin rights to that user's drive.
        # 404: Indicates the user is an external user (#EXT#) or their OneDrive/MySite has not been provisioned yet.
        # We catch the error and continue to the next user gracefully.
        Write-Warning "Failed to fetch drives for user $UserUPN (Error: $($_.Exception.Message)). Skipping this user's OneDrive."
        continue
    }

    if ($drives.Count -eq 0) {
        Write-Host "No OneDrive for Business drive found for $UserUPN." -ForegroundColor DarkGray
    }

    foreach ($drive in $drives) {
        Write-Host "  -> Processing OneDrive Drive: $($drive.Name) (ID: $($drive.Id))" -ForegroundColor Green
        # Start recursion from the root of the OneDrive drive
        Remove-AnonymousSharing -DriveId $drive.Id -DriveName $drive.Name -ItemId "root" -ItemPath ""
    }

# --- FINAL MESSAGE ---

Write-Host "`n========================================" -ForegroundColor Cyan
Write-Host "Anonymous link removal operation complete." -ForegroundColor Cyan
Write-Host "========================================" -ForegroundColor Cyan
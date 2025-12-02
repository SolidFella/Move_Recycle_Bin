# --- Configuration ---
$serverPath = "\\10.1.10.207\Recycle_Bin_Desk_Manager"
$todayDate = Get-Date -Format "yyyy-MM-dd"
$destinationPath = Join-Path -Path $serverPath -ChildPath $todayDate

# --- Execution ---
$shell = New-Object -ComObject Shell.Application
$recycleBin = $shell.Namespace(0xA) # 0xA is Recycle Bin

# Only proceed if there are actually items to move
if ($recycleBin.Items().Count -gt 0) {
    
    # 1. Create the dated folder if it doesn't exist
    if (-not (Test-Path -Path $destinationPath)) {
        try {
            # "| Out-Null" prevents the script from printing the folder details on success
            New-Item -ItemType Directory -Path $destinationPath -Force | Out-Null
        }
        catch {
            Write-Error "CRITICAL: Failed to create daily folder on TrueNAS: $($_.Exception.Message)"
            exit
        }
    }

    # 2. Set the destination
    $destFolder = $shell.Namespace($destinationPath)

    if ($destFolder) {
        $items = $recycleBin.Items()
        
        # 3. Move the items
        # Flag 20 (16 + 4) = "Yes to All" + "Do not show progress dialog"
        # We wrap this in a try/catch just in case the Shell object throws a rare COM error
        try {
            $destFolder.MoveHere($items, 20)
        }
        catch {
            Write-Error "CRITICAL: Failed to move items to TrueNAS: $($_.Exception.Message)"
        }
    } else {
        Write-Error "CRITICAL: Could not bind to destination folder '$destinationPath'. Check network connection."
    }
}
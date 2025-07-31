# AVEVA P&ID Upgrade and DWG File Copy Script
# Run as Administrator

param(
    [string]$ZipSourcePath = "\\adm026\P&ID-WORK\02_P&ID_Software\2_Source\AVEVA_P&ID_12.4.zip",
    [string]$ExtractPath = "D:\AVEVA_Install\",
    [string]$InstallerPath = "D:\AVEVA_Install\PID12.4.exe",
    [string]$SourcePath1 = "\\adm026\P&ID-WORK\02_P&ID_Software\2_Source\RTR.DWG",
    [string]$SourcePath2 = "\\adm026\P&ID-WORK\02_P&ID_Software\2_Source\ARROW2.DWG",
    [string]$DestPath = "C:\Program Files\AVEVA\P&ID 12.4\Install\AutoCad\MetSym\Miscellaneous\",
    [string]$LogPath = "D:\upgrade_log.txt"
)

# Function to write logs
function Write-Log {
    param([string]$Message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] $Message"
    Write-Host $logEntry
    Add-Content -Path $LogPath -Value $logEntry
}

# Check if running as administrator
if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Log "ERROR: Script must be run as Administrator"
    exit 1
}

Write-Log "Starting AVEVA P&ID upgrade process..."

try {
    # Step 1: Check if zip file exists on network share
    if (-not (Test-Path $ZipSourcePath)) {
        Write-Log "ERROR: Zip file not found at $ZipSourcePath"
        exit 1
    }
    Write-Log "Zip file found: $ZipSourcePath"

    # Step 2: Create extraction directory
    if (Test-Path $ExtractPath) {
        Write-Log "Cleaning existing extraction directory: $ExtractPath"
        Remove-Item -Path $ExtractPath -Recurse -Force
    }
    New-Item -Path $ExtractPath -ItemType Directory -Force | Out-Null
    Write-Log "Created extraction directory: $ExtractPath"

    # Step 3: Extract zip file
    Write-Log "Extracting zip file to $ExtractPath..."
    try {
        Expand-Archive -Path $ZipSourcePath -DestinationPath $ExtractPath -Force
        Write-Log "Zip extraction completed successfully"
    } catch {
        Write-Log "ERROR: Failed to extract zip file: $($_.Exception.Message)"
        exit 1
    }

    # Step 4: Find PID12.4.exe installer
    $targetInstaller = Get-ChildItem -Path $ExtractPath -Filter "PID12.4.exe" -Recurse
    if ($targetInstaller.Count -eq 0) {
        Write-Log "ERROR: PID12.4.exe not found in extracted content"
        # List available exe files for troubleshooting
        $allExeFiles = Get-ChildItem -Path $ExtractPath -Filter "*.exe" -Recurse
        if ($allExeFiles.Count -gt 0) {
            Write-Log "Available exe files found:"
            foreach ($exe in $allExeFiles) {
                Write-Log "  - $($exe.FullName)"
            }
        }
        exit 1
    } else {
        $InstallerPath = $targetInstaller[0].FullName
        Write-Log "Found PID12.4.exe installer: $InstallerPath"
    }

    # Step 5: Install the program silently
    Write-Log "Starting installation..."
    $installProcess = Start-Process -FilePath $InstallerPath -ArgumentList "/S" -Wait -PassThru
    
    if ($installProcess.ExitCode -eq 0) {
        Write-Log "Installation completed successfully"
    } else {
        Write-Log "WARNING: Installation returned exit code: $($installProcess.ExitCode)"
    }

    # Step 6: Wait a moment for installation to settle
    Start-Sleep -Seconds 5

    # Step 7: Copy DWG files to AVEVA P&ID installation
    Write-Log "Starting DWG file copy to $DestPath"
    
    # Ensure destination directory exists
    if (-not (Test-Path $DestPath)) {
        Write-Log "Creating destination directory: $DestPath"
        New-Item -Path $DestPath -ItemType Directory -Force
    }
    
    # Array of source files to copy
    $sourceFiles = @($SourcePath1, $SourcePath2)
    $copyCount = 0
    
    foreach ($sourceFile in $sourceFiles) {
        $fileName = Split-Path $sourceFile -Leaf
        $destFile = Join-Path $DestPath $fileName
        
        try {
            # Check if source file exists
            if (Test-Path $sourceFile) {
                Copy-Item -Path $sourceFile -Destination $destFile -Force
                Write-Log "Successfully copied: $fileName"
                $copyCount++
            } else {
                Write-Log "ERROR: Source file not found: $sourceFile"
            }
        } catch {
            Write-Log "ERROR copying $fileName : $($_.Exception.Message)"
        }
    }
    
    Write-Log "DWG copy process completed. $copyCount of $($sourceFiles.Count) files copied successfully."

    # Step 8: Verify copied files
    Write-Log "Verifying copied DWG files..."
    $verifyFiles = @("RTR.DWG", "ARROW2.DWG")
    foreach ($file in $verifyFiles) {
        $fullPath = Join-Path $DestPath $file
        if (Test-Path $fullPath) {
            $fileInfo = Get-Item $fullPath
            Write-Log "Verified: $file (Size: $($fileInfo.Length) bytes, Modified: $($fileInfo.LastWriteTime))"
        } else {
            Write-Log "WARNING: $file not found in destination"
        }
    }

    Write-Log "Upgrade process completed successfully!"

    # Step 9: Clean up extracted files (optional)
    # Uncomment if you want to delete extracted files after installation
    # Write-Log "Cleaning up extracted files..."
    # Remove-Item -Path $ExtractPath -Recurse -Force

} catch {
    Write-Log "ERROR: $($_.Exception.Message)"
    exit 1
}

# Optional: Restart explorer to ensure changes are reflected
# Uncomment if needed
# Write-Log "Restarting Windows Explorer..."
# Stop-Process -Name explorer -Force
# Start-Process explorer

Write-Log "Script execution finished."
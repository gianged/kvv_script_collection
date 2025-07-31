# Run as Administrator

param(
    # Version Configuration - UPDATE THESE FOR NEW VERSIONS
    [string]$CurrentVersion = "12.4",
    [string]$ZipFileName = "AVEVA_P&ID_12.4.zip",
    [string]$InstallerFileName = "PID12.4.exe",
    
    # Network Paths - UPDATE IF PATHS CHANGE
    [string]$NetworkBasePath = "\\adm026\P&ID-WORK\02_P&ID_Software\2_Source",
    [string]$DWGFile1 = "RTR.DWG",
    [string]$DWGFile2 = "ARROW2.DWG",
    
    # Local Paths - USUALLY DON'T NEED TO CHANGE
    [string]$LocalExtractPath = "D:\AVEVA_Install\",
    [string]$LogPath = "D:\upgrade_log.txt",
    
    # Installation Paths - UPDATE IF AVEVA CHANGES STRUCTURE
    [string]$InstallBasePath = "C:\Program Files\AVEVA\P&ID",
    [string]$DWGDestinationSubPath = "Install\AutoCad\MetSym\Miscellaneous\"
)

# Build full paths from parameters
$ZipSourcePath = Join-Path $NetworkBasePath $ZipFileName
$InstallerPath = Join-Path $LocalExtractPath $InstallerFileName
$SourcePath1 = Join-Path $NetworkBasePath $DWGFile1
$SourcePath2 = Join-Path $NetworkBasePath $DWGFile2
$DestPath = Join-Path "$InstallBasePath $CurrentVersion" $DWGDestinationSubPath

# Function to write logs
function Write-Log {
    param([string]$Message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] $Message"
    Write-Host $logEntry
    Add-Content -Path $LogPath -Value $logEntry -ErrorAction SilentlyContinue
}

# Check if running as administrator
if (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Log "ERROR: Script must be run as Administrator"
    exit 1
}

Write-Log "========================================="
Write-Log "AVEVA P&ID Upgrade Script v$CurrentVersion"
Write-Log "========================================="
Write-Log "Configuration:"
Write-Log "  Version: $CurrentVersion"
Write-Log "  Zip Source: $ZipSourcePath"
Write-Log "  Installer: $InstallerFileName"
Write-Log "  Destination: $DestPath"
Write-Log "========================================="

try {
    # Step 1: Search for and uninstall existing AVEVA P&ID versions
    Write-Log "Searching for existing AVEVA P&ID installations..."
    
    # Search using WMI (slower but more reliable for uninstall)
    $existingPID = Get-WmiObject -Class Win32_Product | Where-Object {$_.Name -like "*AVEVA*P&ID*"}
    
    if ($existingPID) {
        foreach ($program in $existingPID) {
            Write-Log "Found existing installation: $($program.Name) (Version: $($program.Version))"
            
            # Skip if it's the same version we're installing
            if ($program.Name -like "*$CurrentVersion*") {
                Write-Log "Skipping uninstall - same version as target ($CurrentVersion)"
                continue
            }
            
            Write-Log "Uninstalling $($program.Name)..."
            try {
                $uninstallResult = $program.Uninstall()
                if ($uninstallResult.ReturnValue -eq 0) {
                    Write-Log "Successfully uninstalled: $($program.Name)"
                } else {
                    Write-Log "WARNING: Uninstall returned code $($uninstallResult.ReturnValue) for $($program.Name)"
                }
            } catch {
                Write-Log "ERROR: Failed to uninstall $($program.Name): $($_.Exception.Message)"
                Write-Log "Continuing with installation anyway..."
            }
        }
    } else {
        Write-Log "No existing AVEVA P&ID installations found"
    }

    # Step 2: Check if zip file exists on network share
    if (-not (Test-Path $ZipSourcePath)) {
        Write-Log "ERROR: Zip file not found at $ZipSourcePath"
        Write-Log "Please verify:"
        Write-Log "  1. Network path is accessible: $NetworkBasePath"
        Write-Log "  2. Zip file exists: $ZipFileName"
        exit 1
    }
    Write-Log "Zip file found: $ZipSourcePath"

    # Step 3: Create extraction directory
    if (Test-Path $LocalExtractPath) {
        Write-Log "Cleaning existing extraction directory: $LocalExtractPath"
        Remove-Item -Path $LocalExtractPath -Recurse -Force
    }
    New-Item -Path $LocalExtractPath -ItemType Directory -Force | Out-Null
    Write-Log "Created extraction directory: $LocalExtractPath"

    # Step 4: Extract zip file
    Write-Log "Extracting zip file to $LocalExtractPath..."
    try {
        Expand-Archive -Path $ZipSourcePath -DestinationPath $LocalExtractPath -Force
        Write-Log "Zip extraction completed successfully"
    } catch {
        Write-Log "ERROR: Failed to extract zip file: $($_.Exception.Message)"
        exit 1
    }

    # Step 5: Find installer executable
    $targetInstaller = Get-ChildItem -Path $LocalExtractPath -Filter $InstallerFileName -Recurse
    if ($targetInstaller.Count -eq 0) {
        Write-Log "ERROR: $InstallerFileName not found in extracted content"
        # List available exe files for troubleshooting
        $allExeFiles = Get-ChildItem -Path $LocalExtractPath -Filter "*.exe" -Recurse
        if ($allExeFiles.Count -gt 0) {
            Write-Log "Available exe files found:"
            foreach ($exe in $allExeFiles) {
                Write-Log "  - $($exe.FullName)"
            }
        }
        exit 1
    } else {
        $InstallerPath = $targetInstaller[0].FullName
        Write-Log "Found $InstallerFileName installer: $InstallerPath"
    }

    # Step 6: Start the installation (interactive)
    Write-Log "Starting interactive installation..."
    Write-Log "IMPORTANT: Please complete the installation manually when the installer appears"
    
    $installProcess = Start-Process -FilePath $InstallerPath -Wait -PassThru
    
    if ($installProcess.ExitCode -eq 0) {
        Write-Log "Installation completed successfully (Exit Code: 0)"
    } elseif ($installProcess.ExitCode -eq $null) {
        Write-Log "Installation window closed - assuming successful completion"
    } else {
        Write-Log "WARNING: Installation returned exit code: $($installProcess.ExitCode)"
        Write-Log "Do you want to continue with DWG file copy? (Press Ctrl+C to abort)"
        Read-Host "Press Enter to continue or Ctrl+C to abort"
    }

    # Step 7: Wait a moment for installation to settle
    Start-Sleep -Seconds 5

    # Step 8: Copy DWG files to AVEVA P&ID installation
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

    # Step 9: Verify copied files
    Write-Log "Verifying copied DWG files..."
    $verifyFiles = @($DWGFile1, $DWGFile2)
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

    # Step 10: Clean up extracted files
    Write-Log "Cleaning up extracted files..."
    try {
        Remove-Item -Path $LocalExtractPath -Recurse -Force
        Write-Log "Successfully deleted temporary extraction folder: $LocalExtractPath"
    } catch {
        Write-Log "WARNING: Could not delete extraction folder: $($_.Exception.Message)"
        Write-Log "You may manually delete: $LocalExtractPath"
    }

} catch {
    Write-Log "ERROR: $($_.Exception.Message)"
    exit 1
}

# Optional: Restart explorer to ensure changes are reflected
# Uncomment if needed
# Write-Log "Restarting Windows Explorer..."
# Stop-Process -Name explorer -Force
# Start-Process explorer

Write-Log "========================================="
Write-Log "Script execution finished."
Write-Log "Log saved to: $LogPath"
Write-Log "========================================="
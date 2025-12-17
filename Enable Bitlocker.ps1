#Requires -Version 4

<#
.SYNOPSIS
    Automates BitLocker enablement on the C: drive with TPM and Recovery Key, including AD backup.
    This script handles various BitLocker states and ensures proper key protector setup.

.DESCRIPTION
    This script is designed to:
    1.  Initialize the TPM (if not already initialized).
    2.  Check the current BitLocker status of the C: drive.
    3.  If BitLocker is 'Off':
        a.  Attempt to turn on BitLocker using 'manage-bde -on C:', which may implicitly add a TPM protector.
        b.  Explicitly add a TPM protector if one is not detected after the initial 'manage-bde -on' command.
        c.  Add a Recovery Password protector.
        d.  Resume BitLocker protection (if it's in a suspended state).
    4.  If -ForceDecryptAndEnable is set to $true, the script will first attempt to
    decrypt the C: drive (if encrypted) and remove all protectors to start the process from a clean slate.
    5.  Backup the Recovery Password protector to Active Directory.

    The script handles scenarios where a restart is required by BitLocker enablement and provides clear instructions.
    All script output will also be saved to a log file in C:\.

.NOTES
    OS: Windows 10+, Server 2012+
    Requires PowerShell 4.0 or higher.
    Requires administrative privileges to run.
#>

[CmdletBinding()]
param(
    [Parameter()]
    [Switch]$ForceDecryptAndEnable = [System.Convert]::ToBoolean($env:forceDecryptAndEnable)
)

begin {
    # --- Helper Function: Get-NinjaProperty (Included for completeness if needed, but not used in this specific BitLocker script section) ---
    function Get-NinjaProperty {
        [CmdletBinding()]
        Param(
            [Parameter(Mandatory = $True, ValueFromPipeline = $True)]
            [String]$Name,
            [Parameter()]
            [String]$Type,
            [Parameter()]
            [String]$DocumentName
        )
        # Placeholder for Get-NinjaProperty logic, removed content for brevity as it's not directly used here.
        # This function should be correctly defined if accessing Ninja custom fields.
        throw "Get-NinjaProperty function not fully implemented in this standalone script. If needed, re-integrate the full function definition."
    }
}

process {
    # Define log file path with a timestamp for uniqueness
    $logFilePath = "C:\BitLocker_Log_$(Get-Date -Format 'yyyyMMdd_HHmmss').log"

    # Check if Start-Transcript cmdlet is available
    $transcriptAvailable = $false
    if (Get-Command -Name Start-Transcript -ErrorAction SilentlyContinue) {
        $transcriptAvailable = $true
    }

    # Flag to track if transcript was successfully started
    $transcriptActive = $false

    # Start transcript to capture all output to the log file
    if ($transcriptAvailable) {
        try {
            Start-Transcript -Path $logFilePath -Append -Force
            $transcriptActive = $true
            Write-Host "Script output is being logged to: $logFilePath"
        } catch {
            Write-Warning "Could not start transcript logging to ${logFilePath}: $($_.Exception.Message)"
            $transcriptActive = $false # Ensure flag is false if start failed
        }
    } else {
        Write-Warning "Start-Transcript cmdlet not available. Script output will not be logged to a file."
    }

    try {
        Write-Host "`n--- BitLocker Enablement Process ---"

        # Ensure the TPM is initialized and ready
        Write-Host "Initializing TPM (if not already initialized)..."
        try {
            Initialize-Tpm -AllowClear -ErrorAction Stop
            Write-Host "TPM initialized successfully."
            Sleep 5 # Give TPM a moment to settle
        } catch {
            Write-Error "Failed to initialize TPM: $($_.Exception.Message)"
            exit 1
        }

        # --- Force Decrypt and Enable Logic ---
        if ($ForceDecryptAndEnable) {
            Write-Host "`n--- Force Decrypt and Re-Enable BitLocker ---"
            Write-Host "Attempting to decrypt C: drive and remove all key protectors to start fresh."

            # Get current BitLocker status for decryption check
            $currentStatusForDecrypt = Get-BitLockerVolume -MountPoint "C:"

            if ($currentStatusForDecrypt.ProtectionStatus -ne "Off" -or $currentStatusForDecrypt.KeyProtector.Count -gt 0) {
                Write-Host "BitLocker is currently '$($currentStatusForDecrypt.ProtectionStatus)' or has protectors. Proceeding with decryption."

                try {
                    # Remove all key protectors first
                    Write-Host "Removing all existing BitLocker key protectors..."
                    # Use manage-bde -protectors -get C: to get raw output and parse IDs
                    $rawProtectorOutput = manage-bde -protectors -get C: | Out-String
                    $protectorIds = $rawProtectorOutput | Select-String "ID: {(.*)}" | ForEach-Object {$_.Matches[0].Groups[1].Value.Trim()}

                    if ($protectorIds.Count -eq 0) {
                        Write-Host "No key protectors found to remove."
                    } else {
                        foreach ($id in $protectorIds) {
                            Write-Host "Removing protector with ID: ${id}"
                            $commandRemove = "manage-bde -protectors -delete C: -id $id"
                            $tempOutputFileRemove = [System.IO.Path]::GetTempFileName()
                            $processRemove = Start-Process -FilePath "cmd.exe" -ArgumentList "/c $commandRemove > `"$tempOutputFileRemove`" 2>&1" -NoNewWindow -Wait -PassThru -ErrorAction Stop
                            $manageBdeOutputRemove = Get-Content $tempOutputFileRemove | Out-String
                            Remove-Item $tempOutputFileRemove -ErrorAction SilentlyContinue

                            if ($processRemove.ExitCode -ne 0) {
                                Write-Warning "Failed to remove protector ${id}: $($manageBdeOutputRemove)"
                            } else {
                                Write-Host "Successfully removed protector ${id}."
                            }
                            Sleep 2
                        }
                        Write-Host "All key protectors removed."
                    }

                    # Turn off BitLocker (decrypt)
                    Write-Host "Attempting to turn off BitLocker (decrypt C: drive)..."
                    $commandOff = "manage-bde -off C:"
                    $tempOutputFileOff = [System.IO.Path]::GetTempFileName()
                    $processOff = Start-Process -FilePath "cmd.exe" -ArgumentList "/c $commandOff > `"$tempOutputFileOff`" 2>&1" -NoNewWindow -Wait -PassThru -ErrorAction Stop
                    $manageBdeOutputOff = Get-Content $tempOutputFileOff | Out-String
                    Remove-Item $tempOutputFileOff -ErrorAction SilentlyContinue

                    if ($processOff.ExitCode -ne 0) {
                        # Check for restart required after -off (less common but possible)
                        if ($manageBdeOutputOff -match "You must restart your computer") {
                            Write-Error "BitLocker decryption requires a system restart. Please restart the computer and re-run this script."
                            exit 1
                        } else {
                            throw "manage-bde -off command failed with exit code $($processOff.ExitCode). Output: $($manageBdeOutputOff)"
                        }
                    }
                    Write-Host "manage-bde -off output: $($manageBdeOutputOff)"

                    # Wait for decryption to complete
                    Write-Host "Waiting for decryption to complete. This may take a long time..."
                    do {
                        Sleep 30 # Check every 30 seconds
                        $decryptStatus = Get-BitLockerVolume -MountPoint "C:"
                        Write-Host "Decryption Status: $($decryptStatus.ConversionStatus) - $($decryptStatus.PercentageEncrypted)%"
                    } while ($decryptStatus.ConversionStatus -ne "Fully Decrypted")

                    Write-Host "C: drive is now fully decrypted."
                    Sleep 5 # Give system a moment

                } catch {
                    Write-Error "Failed during BitLocker decryption process: $($_.Exception.Message)"
                    exit 1
                }
            } else {
                Write-Host "C: drive is already fully decrypted and has no protectors. Starting fresh."
            }
            Write-Host "Decryption and cleanup complete. Proceeding with BitLocker enablement."
        }
        # --- End Force Decrypt and Enable Logic ---


        # Get current BitLocker status using manage-bde -status for direct parsing
        Write-Host "Checking BitLocker status for C: drive using manage-bde -status..."
        $commandStatus = "manage-bde -status C:"
        $tempOutputFileStatus = [System.IO.Path]::GetTempFileName()
        $processStatus = Start-Process -FilePath "cmd.exe" -ArgumentList "/c $commandStatus > `"$tempOutputFileStatus`" 2>&1" -NoNewWindow -Wait -PassThru -ErrorAction Stop
        $manageBdeStatusOutput = Get-Content $tempOutputFileStatus | Out-String
        Remove-Item $tempOutputFileStatus -ErrorAction SilentlyContinue

        Write-Host "manage-bde -status output:`n$manageBdeStatusOutput"

        # Parse current BitLocker state from manage-bde -status output
        $currentProtectionStatus = ($manageBdeStatusOutput | Select-String "Protection Status *: *(.*)" | ForEach-Object {$_.Matches[0].Groups[1].Value.Trim()})
        $currentConversionStatus = ($manageBdeStatusOutput | Select-String "Conversion Status *: *(.*)" | ForEach-Object {$_.Matches[0].Groups[1].Value.Trim()})
        $currentPercentageEncrypted = ($manageBdeStatusOutput | Select-String "Percentage Encrypted *: *(.*)" | ForEach-Object {[double]($_.Matches[0].Groups[1].Value -replace '%','').Trim()})
        $currentKeyProtectorsRaw = ($manageBdeStatusOutput | Select-String "Key Protectors *: *(.*)" | ForEach-Object {$_.Matches[0].Groups[1].Value.Trim()})

        Write-Host "Parsed Status: Protection=$currentProtectionStatus, Conversion=$currentConversionStatus, Encrypted=$currentPercentageEncrypted%, Protectors=$currentKeyProtectorsRaw"

        $bitlockerNeedsInitialOnCommand = $true # Assume it needs manage-bde -on unless proven otherwise

        # Scenario 1: Already On or Encrypting (skip initial -on)
        if ($currentProtectionStatus -eq "On" -or $currentProtectionStatus -eq "Encrypting") {
            Write-Host "BitLocker protection is already '$currentProtectionStatus'. Skipping initial 'manage-bde -on'."
            $bitlockerNeedsInitialOnCommand = $false
        }
        # Scenario 2: Protection Off but 100% Encrypted (skip initial -on, go straight to adding protectors)
        elseif ($currentProtectionStatus -eq "Protection Off" -and $currentPercentageEncrypted -eq 100.0) {
            Write-Host "Detected drive is 100% encrypted but protection is 'Off'. Skipping 'manage-bde -on'."
            $bitlockerNeedsInitialOnCommand = $false
        }
        # Scenario 3: Truly Off and not encrypted (proceed with initial -on)
        else {
            Write-Host "BitLocker is truly 'Off' and not encrypted. Attempting 'manage-bde -on'."
        }

        # Only attempt manage-bde -on if the logic determines it's necessary
        if ($bitlockerNeedsInitialOnCommand) {
            Write-Host "Attempting to turn on BitLocker for C: drive using manage-bde -on..."
            try {
                $commandOn = "manage-bde -on C: -used"
                $tempOutputFileOn = [System.IO.Path]::GetTempFileName()
                $processOn = Start-Process -FilePath "cmd.exe" -ArgumentList "/c $commandOn > `"$tempOutputFileOn`" 2>&1" -NoNewWindow -Wait -PassThru -ErrorAction Stop
                $manageBdeOutputOn = Get-Content $tempOutputFileOn | Out-String
                Remove-Item $tempOutputFileOn -ErrorAction SilentlyContinue

                if ($processOn.ExitCode -ne 0) {
                    if ($manageBdeOutputOn -match "0x8031004e") {
                        Write-Error "BitLocker enablement requires a system restart. Please restart the computer and re-run this script."
                        exit 1
                    } else {
                        throw "manage-bde -on command failed with exit code $($processOn.ExitCode). Output: $($manageBdeOutputOn)"
                    }
                }
                Write-Host "manage-bde -on output: $($manageBdeOutputOn)"
                Sleep 5
            } catch {
                Write-Error "Failed during initial BitLocker enablement with manage-bde -on: $($_.Exception.Message)"
                exit 1
            }
        }

        # --- IMPORTANT: Ensure Protectors are added, regardless of initial 'manage-bde -on' result ---
        # This section ensures that TPM and Recovery Password protectors are always added
        # if they are missing, allowing the process to continue even if 'manage-bde -on' was skipped or
        # put the drive in a pending state.

        Write-Host "Proceeding to ensure TPM and Recovery Password protectors are in place."

        # Re-get KEY after initial attempts/skips to ensure we have the latest status
        $KEY = Get-BitLockerVolume -MountPoint "C:"

        # This check is crucial: if a restart is required after manage-bde -on, exit now.
        # This prevents trying to add protectors to a volume that's not ready.
        $manageBdeOutputOnChecked = if ($bitlockerNeedsInitialOnCommand -and (Test-Path Variable:\manageBdeOutputOn)) { $manageBdeOutputOn } else { "" }

        if ($bitlockerNeedsInitialOnCommand -and ($manageBdeOutputOnChecked -match "Restart the computer to run a hardware test" -or $manageBdeOutputOnChecked -match "Encryption will begin after the hardware test succeeds")) {
            Write-Host "Initial BitLocker enablement requires a system restart for hardware test. Please restart the computer and re-run this script."
            exit 0 # Exit successfully, indicating a restart is the next step
        }


        Write-Host "BitLocker protection status before adding missing protectors: $($KEY.ProtectionStatus)."

        $TpmProtector = $KEY.KeyProtector | Where-Object {$_.KeyProtectorType -eq "Tpm"}
        $RecoveryKeyProtector = $KEY.KeyProtector | Where-Object {$_.KeyProtectorType -eq "RecoveryPassword"}

        # Add TPM protector if missing
        if (-not $TpmProtector) {
            Write-Host "TPM protector not found. Attempting to add it now..."
            try {
                $commandAddTpm = "manage-bde -protectors -add C: -TPM"
                $tempOutputFileTpm = [System.IO.Path]::GetTempFileName()
                $processTpm = Start-Process -FilePath "cmd.exe" -ArgumentList "/c $commandAddTpm > `"$tempOutputFileTpm`" 2>&1" -NoNewWindow -Wait -PassThru -ErrorAction Stop
                $manageBdeOutputTpm = Get-Content $tempOutputFileTpm | Out-String
                Remove-Item $tempOutputFileTpm -ErrorAction SilentlyContinue

                if ($processTpm.ExitCode -ne 0) {
                    throw "manage-bde -protectors -add -TPM command failed with exit code $($processTpm.ExitCode). Output: $($manageBdeOutputTpm)"
                }
                Write-Host "manage-bde -protectors -add -TPM output: $($manageBdeOutputTpm)"
                Sleep 5 # Give it a moment
                # Re-get KEY after adding protector to get the new ID
                $KEY = Get-BitLockerVolume -MountPoint "C:"
                $TpmProtector = $KEY.KeyProtector | Where-Object {$_.KeyProtectorType -eq "Tpm"}
                if (-not $TpmProtector) {
                    Write-Error "Failed to add TPM protector despite command success. Manual intervention may be required."
                    exit 1
                }
            } catch {
                Write-Error "Failed to add TPM protector: $($_.Exception.Message)"
                exit 1
            }
        } else {
            Write-Host "TPM protector already exists: ID = $($TpmProtector.KeyProtectorId)"
        }


        # Add Recovery Password protector if missing
        if (-not $RecoveryKeyProtector) {
            Write-Host "RecoveryPassword protector not found. Adding it now..."
            try {
                $commandAddRp = "manage-bde -protectors -add C: -RecoveryPassword"
                $tempOutputFileRp = [System.IO.Path]::GetTempFileName()
                $processRp = Start-Process -FilePath "cmd.exe" -ArgumentList "/c $commandAddRp > `"$tempOutputFileRp`" 2>&1" -NoNewWindow -Wait -PassThru -ErrorAction Stop
                $manageBdeOutputRp = Get-Content $tempOutputFileRp | Out-String
                Remove-Item $tempOutputFileRp -ErrorAction SilentlyContinue

                if ($processRp.ExitCode -ne 0) {
                    throw "manage-bde -protectors -add -RecoveryPassword command failed with exit code $($processRp.ExitCode). Output: $($manageBdeOutputRp)"
                }
                Write-Host "manage-bde -protectors -add -RecoveryPassword output: $($manageBdeOutputRp)"
                Sleep 5 # Give it a moment
                # Re-get KEY after adding protector to get the new ID
                $KEY = Get-BitLockerVolume -MountPoint "C:"
                $RecoveryKeyProtector = $KEY.KeyProtector | Where-Object {$_.KeyProtectorType -eq "RecoveryPassword"}
                if (-not $RecoveryKeyProtector) {
                    Write-Error "Failed to add RecoveryPassword protector despite command success. Manual intervention may be required."
                    exit 1
                }
            } catch {
                Write-Error "Failed to add Recovery Password protector: $($_.Exception.Message)"
                exit 1
            }
        } else {
            Write-Host "RecoveryPassword protector already exists: ID = $($RecoveryKeyProtector.KeyProtectorId)"
        }


        # --- IMPORTANT: Attempt to resume BitLocker here ---
        # This is the crucial step to transition from "Off" to "On" or "Encrypting" after protectors are added.
        # Perform a fresh Get-BitLockerVolume after all protector additions
        $KEY = Get-BitLockerVolume -MountPoint "C:" # Refresh $KEY after protector additions

        Write-Host "BitLocker protection status after adding all missing protectors: $($KEY.ProtectionStatus)."

        if ($KEY.ProtectionStatus -eq "Off" -or $KEY.ProtectionStatus -eq "Suspended") {
            Write-Host "BitLocker protection is still 'Off' or 'Suspended'. Attempting to resume protection..."
            try {
                Resume-BitLocker -MountPoint "C:" -Confirm:$false -ErrorAction Stop
                Write-Host "BitLocker protection resumed successfully."
                Sleep 15 # Give BitLocker some time to fully activate
                $KEY = Get-BitLockerVolume -MountPoint "C:" # Re-check status
            } catch {
                Write-Error "Failed to resume BitLocker protection: $($_.Exception.Message)"
                exit 1
            }
        }


        # Attempt to backup Recovery Password protector to Active Directory
        if ($KEY.ProtectionStatus -eq "On" -and $RecoveryKeyProtector) {
            Write-Host "BitLocker protection is 'On'. Attempting to backup Recovery Password protector to Active Directory..."
            try {
                Backup-BitLockerKeyProtector -MountPoint "C:" -KeyProtectorId $RecoveryKeyProtector.KeyProtectorId -Confirm:$false -ErrorAction Stop
                Write-Host "BitLocker recovery key backed up to Active Directory successfully."
            } catch {
                Write-Error "Failed to backup BitLocker recovery key to Active Directory: $($_.Exception.Message)"
            }
        } else {
            if ($KEY.ProtectionStatus -ne "On") {
                Write-Error "BitLocker protection is not 'On' (current status: $($KEY.ProtectionStatus)). Cannot backup key."
            }
            if (-not $RecoveryKeyProtector) {
                Write-Error "No RecoveryPassword protector found for backup. Manual intervention may be required."
            }
            Write-Error "BitLocker protection or key protector issue prevents backup. Manual intervention may be required."
        }

        Write-Host "`n--- BitLocker Enablement Process Completed ---"

    } finally {
        # Stop transcript to finalize the log file
        if ($transcriptActive) { # Check if transcript was actually started
            try {
                Stop-Transcript
                Write-Host "Transcript logging stopped. Log file saved to: $logFilePath"
            } catch {
                Write-Warning "Could not stop transcript: $($_.Exception.Message). Log file might be incomplete."
            }
        } else {
            Write-Warning "No active transcript to stop (Start-Transcript was not available or failed)."
        }
    }
}

end {
    # Script end block

}

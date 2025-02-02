# Virtual Machine Action Methods
class VMDKActions {
	VMDKActions() { }

	# Discard The Saved State Of The Selected Devices Virtual Machine
	DiscardSavedState( `
			[string]$vBoxManagePath, `
			[string]$vmName) `
	{
		Write-Host "Discarding $($vmName)'s Saved State..."

		# 2>&1 -> (2) 'stderr' stream (>) 'redirect', using an explicit (&) 'referrence', to (1) 'stdout' stream
		$commandOutput = & $vBoxManagePath discardstate $vmName 2>&1
		if ($commandOutput.Length) {
			$discardError = [string]::Concat(
				"Discarding State Failed. ",
				"The vm is either not shut down or there is no state to discard.")
			Write-Host $discardError -ForegroundColor Red
		}
		else {
			Write-Host "State Discarded."
		}
	}

	# Remove The Device A User Selected From A Virtual Machine
	RemoveAttachedDevice( `
			[string]$vBoxManagePath, `
			[string]$vmName, `
			[string]$controllerName, `
			[string]$port, `
			[string]$deviceGuid) `
	{
		Write-Host "Removing Attached Storage Device..."
		$commandOutput = & $vBoxManagePath storageattach $vmName `
			--storagectl $controllerName `
			--port $port `
			--device 0 `
			--type hdd `
			--medium none 2>&1
		if ($commandOutput.Length) {

			$removeError = [string]::Concat(
				"Removing Device Failed. ",
				"Please ensure the script and virtualbox are running as admin, ",
				"the vm is not in use, and a device has been selected or created.")
			Write-Host $removeError -ForegroundColor Red
		}
		else {
			$removeSuccess = [string]::Concat(
				"Device Removed. ",
				"If you plan to delete the vmdk please run the close disk option immediately.")
			Write-Host $removeSuccess
		}
	}

	# Close The Device That Was Attached To A Virtual Machine
	CloseDisk([string]$vBoxManagePath, [string]$deviceLocation) {
		Write-Host "Closing Disk $($deviceLocation)..."
		$commandOutput = & $vBoxManagePath closemedium disk $deviceLocation 2>&1
		if ($commandOutput.Length) {
			$closeError = [string]::Concat(
				"Closing Disk Failed. ",
				"Please remove the device from the controller then try again.")
			Write-Host $closeError -ForegroundColor Red
		}
		else {
			Write-Host "Disk Closed."
		}
	}

	# Delete A VMDK File That Was Attached To A Virtual Machine
	DeleteRelatedFiles( `
			[string]$devicePath, `
			[string]$vmdk, `
			[string]$dashPartitionVmdk, `
			[string]$dotPartitionVmdk) `
	{
		if ($devicePath.Length) {
			Write-Host "Deleting VMDK File: $devicePath..."
			Remove-Item -Path $devicePath -Force
			$commandOutput = Test-Path $devicePath
			if ($commandOutput) {
				$pathError = [string]::Concat(
					"Deleting file failed. Please ensure the script ",
					"is running as admin and try again.")
				Write-Host $pathError -ForegroundColor Red
			}
			else {
				Write-Host "File Deleted."
			}

			# Check for partition file and delete it if it exists
			$dashPartitionPath = $devicePath -replace $vmdk, $dashPartitionVmdk
			$dotPartitionPath = $devicePath -replace $vmdk, $dotPartitionVmdk
			$dashPartitionExists = Test-Path $dashPartitionPath
			$dotPartitionExists = Test-Path $dotPartitionPath
			if ($dashPartitionExists) {
				Write-Host "Deleting VMDK Partition File: $dashPartitionPath"
				$dashOutput = Remove-Item -Path $dashPartitionPath -Force 2>&1
				if ($dashOutput.Length) {
					$dashError = [string]::Concat(
						"Deleting partition file failed. Please ensure the script ",
						"is running as admin and try again.")
					Write-Host $dashError -ForegroundColor Red
				}
				else {
					Write-Host "File Deleted."
				}

			}
			elseif ($dotPartitionExists) {
				Write-Host "Deleting VMDK Partition File: $dotPartitionPath"
				$dotOutput = Remove-Item -Path $dotPartitionPath -Force 2>&1
				if ($dotOutput.Length) {
					$dotError = [string]::Concat(
						"Deleting partition file failed. Please ensure the script ",
						"is running as admin and try again.")
					Write-Host $dotError -ForegroundColor Red
				}
				else {
					Write-Host "File Deleted."
				}
			}
			else {
				Write-Host "Did not find any additional partition files to delete. :)"
			}
		}
		else {
			$noPathError = [string]::Concat(
				"Device path was not provided. ",
				"Please ensure a device has been selected or created and try again.")
			Write-Host $noPathError -ForegroundColor Red
		}
	}

	# Create A New VMDK File(s)
	CreateRelatedFiles( `
			[string]$vBoxManagePath, `
			[string]$vmLocation, `
			[string]$physicalDriveRegex, `
			[string]$fileNameRegex, `
			[string]$vmdk, `
			[string]$dev, `
			[string]$dri, `
			[int]$deviceOption, `
			[hashtable]$controllerAndDevice, `
			[hashtable]$userSelections, `
			[object]$helpers) `
	{
		$physicalDriveInfo = [string]::Empty
		$physicalDrives = @()
		$driveOptions = @()
		$option = 1
		powershell Get-WmiObject Win32_DiskDrive | ForEach-Object {
			$physicalDriveInfo += $_
		}
		$physicalDrives = $physicalDriveInfo -split [regex]::Escape("Partitions") | ForEach-Object {
			return "Partitions$($_)"
		}
		$drives = [System.Collections.ArrayList]::new($physicalDrives)
		$drives.RemoveAt(0)
		$drives | ForEach-Object {
			if ($_ -match $physicalDriveRegex) {
				$driveOptions += New-Object PSObject -Property ([ordered]@{
						"Option"     = $option.ToString()
						"Partitions" = $matches[1]
						"DeviceID"   = $matches[2]
						"Model"      = $matches[3]
						"Size"       = $matches[4]
						"Caption"    = $matches[5]
					})
				$option++
			}
		}
		$helpers.BuildTable($driveOptions)
		$helpers.Select($userSelections, $driveOptions, $dri)
		$option = [int]$userSelections[$dri] - 1
		$partitions = [string]::Empty
		switch ($driveOptions[$option].Partitions) {
			"0" {
				Write-Warning "The device has no partitions."
			}
			default {
				$result = Read-Host "Please enter a partition"
				if ([int]$result -gt [int]$driveOptions[$option].Partitions) {
					Write-Host "The desired partition is out of range." -ForegroundColor Red
					return
				}
				else {
					$partitions = $result.ToString()
				}
			}
		}

		$fileName = Read-Host "Please enter a file name ( without extension )"
		$folderPath = [string]::Empty
		if ((Test-Path -Path $vmLocation -PathType Container) -eq $false) {
			$folderPath = Read-Host "Please provide a destination folder path"
		}
		else {
			$folderPath = $vmLocation
			Write-Host "Creating $fileName$vmdk in vm folder: $folderPath"
			if ($folderPath[-1] -eq '\') {
				$folderPath = $folderPath.Substring(0, $folderPath.Length - 1)
			}
			$filePath = "$($folderPath)\$($fileName)$($vmdk)"
			$commandOutput = & $vBoxManagePath internalcommands createrawvmdk `
				-filename  $filePath `
				-rawdisk $driveOptions[$option].DeviceID `
				-partitions $partitions 2>&1
			if (Test-Path -Path $filePath) {
				Write-Host "VMDK File Created."
			}
			else {
				$createError = [string]::Concat("Creating the VMDK failed:`n`n$commandOutput`n`n",
					"Please ensure you ran the script as admin and try again.")
				Write-Host $createError -ForegroundColor Red
				return
			}

			# If no device has been selected add it to the dictionary with port 0
			# otherwise update it and increment the port by 1
			if (-not $controllerAndDevice[$dev] -or $controllerAndDevice[$dev]::IsNullOrEmpty) {
				$controllerAndDevice[$dev] = New-Object PSObject -Property ([ordered]@{
						"Option"     = $deviceOption
						"Port"       = 0
						"DeviceGuid" = [string]::Empty
						"Location"   = $filePath
					})
			}
			else {
				$port = [string]([int]$controllerAndDevice[$dev].Port + 1)
				$controllerAndDevice[$dev] = New-Object PSObject -Property ([ordered]@{
						"Option"     = $deviceOption
						"Port"       = $port
						"DeviceGuid" = [string]::Empty
						"Location"   = $filePath
					})
			}
		}
	}

	# Attach New or Currently Selected VMDK File To The Currently Selected Controller
	AttachDevice( `
			[string]$vBoxManagePath, `
			[hashtable]$controllerAndDevice, `
			[string]$fileNameAndExtRegex, `
			[string]$vmName, `
			[string]$con, `
			[string]$dev) `
	{
		$empty = "Empty"
		$devicePath = if ($controllerAndDevice[$dev].Location) {
			$controllerAndDevice[$dev].Location
		}
		else {
			$empty
		}
		$devicePort = $controllerAndDevice[$dev].Port
		$deviceName = if ($controllerAndDevice[$dev].Location -match $fileNameAndExtRegex) {
			$matches[1]
		}
		else {
			$empty
		}
		$controllerName = $controllerAndDevice[$con].ControllerName
		$usePath = $false

		while (!$usePath) {
			if ($devicePath -eq "q") {
				return
			}
			elseif ((-not [string]::IsNullOrEmpty($devicePath)) -and (Test-Path $devicePath)) {
				Write-Host "Device Path Found: $devicePath"
				$usePath = $true
			}
			else {
				$incorrectPath = if ([string]::IsNullOrEmpty($devicePath)) { $empty } else { $devicePath }
				Write-Host "Device Path ( $incorrectPath ) doesn't exist.`n" -ForegroundColor Red
				$devicePath = Read-Host "Please enter a device path or press 'q' to return to action menu"
			}
		}

		$commandOutput = & $vBoxManagePath storageattach $vmName `
			--storagectl  $controllerName `
			--port $devicePort `
			--device 0 `
			--type hdd `
			--medium $devicePath 2>&1

		if ($commandOutput.Length) {
			Write-Host "Attaching vmdk to storage failed: $_"
		}
		else {
			$attachSuccess = [string]::Concat(
				"Attached $deviceName to the $vmName ",
				"$controllerName controller on port $devicePort.")
			Write-Host $attachSuccess
		}
	}
}

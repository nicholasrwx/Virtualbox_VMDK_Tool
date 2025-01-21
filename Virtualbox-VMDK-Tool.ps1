# menu arrays
$virtualMachines = @()
$devices = @()
$controllers = @()
$actions = @(
	New-Object PSObject -Property ([ordered]@{ "Option" = "1"; "Action" = "Discard Machine State" })
	New-Object PSObject -Property ([ordered]@{ "Option" = "2"; "Action" = "Remove Device From Controller" })
	New-Object PSObject -Property ([ordered]@{ "Option" = "3"; "Action" = "Close Disk" })
	New-Object PSObject -Property ([ordered]@{ "Option" = "4"; "Action" = "Delete vmdk file(s)" })
	New-Object PSObject -Property ([ordered]@{ "Option" = "5"; "Action" = "Create vmdk file(s)" })
	New-Object PSObject -Property ([ordered]@{ "Option" = "6"; "Action" = "Attach Device To Controller" })
	New-Object PSObject -Property ([ordered]@{ "Option" = "Q"; "Action" = "Quit" }))

# user selected vm information
$virtualMachine = New-Object PSObject -Property (
	[ordered]@{ "Option" = [string]::Empty; "VmName" = [string]::Empty; "VmGuid" = [string]::Empty })

# user selected option numbers for each menu
$userSelections = @{}

# parsed 'vm specific' info based on user selection
$controllersAndDevices = @{}

# user selected controller and device
$controllerAndDevice = @{}

# strings
$vm = "VirtualMachine"
$con = "Controller"
$dev = "Device"
$act = "Actions"
$startParsingValue = "Storage Controllers:"
$endParsingValue = "NIC 1:"
$vmdk = ".vmdk"
$partitionedvmdk = "-pt.vmdk"
$attachPath = [string]::Empty
$virtualMachineInfo = [string]::Empty

# space used to differentiate from 'Ports' when splitting string
$port = 'Port '

# booleans
$startParsing = $false
$endParsing = $false
$quit = $false

# option variables for menu option values
$vmOption = 1
$controllerOption = 1
$deviceOption = 1

# VBoxManage executable file path
$vBoxManagePath = Get-ChildItem -Path C:\ -Recurse -File -Filter VBoxManage.exe -ErrorAction SilentlyContinue `
| Select-Object -First 1 | ForEach-Object { $_.FullName }

# regular expressions
$vBoxManageRegex = "VBoxManage\.exe"
$vmRegex = '"([^"]+)"\s*{([^}]+)}'
$deviceRegex = "(?:Port\s([0-9]+),\sUnit\s[0-9]+:\sUUID:\s([a-zA-Z0-9-]+)\s*.*Location:\s`"([^`"]+)`")"
$controllerRegex = "(?:([0-9]+):\s'([a-zA-Z0-9\s]+)',\sType:\s([a-zA-Z0-9]+))"
$physicalDriveRegex = "(?=Partitions\s:\s([0-9]*)DeviceID\s*:\s([\\\.a-zA-Z0-9]*)Model\s*:\s([a-zA-Z0-9\s\._-]*)" +
"Size\s*:\s([0-9]*)Caption\s*:\s([\w\s\._-]*(?=Partitions|$)))"

# instance containing action methods
$vmActions = [VmActions]::new()

# Constructs a formatted table of options and displays it
function BuildTable {
	param ([array]$tableOptions)
	$properties = $tableOptions | Get-Member -MemberType Property | Select-Object -ExpandProperty Name
	$tableOptions | Select-Object $properties | Format-Table -AutoSize | Out-Host
}

# Ensure VBoxManage Executable Exists On Users System
try {
	if ($vBoxManagePath -match $vBoxManageRegex) {
		Write-Host "VBoxManage Located: $($vBoxManagePath)"
	}
	else {
		throw "Please ensure VBoxManage exists on your system and run the script as admin."
	}
}
catch {
	Write-Host "VBoxManage Not Located: $($_.Exception.Message)"
	exit 1
}

# Create A List Of VMs and Display The Information
& $vBoxManagePath list vms | ForEach-Object {
	if ($_ -match $vmRegex) {
		$vmName = $matches[1]
		$vmGuid = $matches[2]
		$option = $vmOption.ToString()
		$virtualMachines += New-Object PSObject -Property ([ordered]@{
				"Option" = $option
				"VmName" = $vmName
				"VmGuid" = $vmGuid
			})
		$vmOption++
	}
}
BuildTable($virtualMachines)
$userSelections[$vm] = Read-Host "Please Select An Option"

# Capture The Selected VM Controller and Device Information
try {
	if ($userSelections[$vm] -gt $virtualMachines.Count) {
		throw "Option number ( $userSelections[$vm] ) is not available. Please re-enter an option number."
	}
	else {
		$virtualMachines | ForEach-Object {
			if ($_.Option -eq [int]$userSelections[$vm].ToString()) {
				$virtualMachine.Option = $_.Option
				$virtualMachine.VmName = $_.VmName
				$virtualMachine.VmGuid = $_.VmGuid
			}
		}
		& $vBoxManagePath showvminfo $virtualMachine.VmGuid | ForEach-Object {
			if (!$startParsing) {
				if ($_.Contains($startParsingValue)) {
					$startParsing = $true
				}
			}
			elseif (!$endParsing) {
				if ($_.Contains($endParsingValue)) {
					$endParsing = $true
				}
				else {
					$virtualMachineInfo += $_
				}
			}
		}

		# Create List Of Controller Information and a Dictionary of Associated Controller Device Information
		$controllerInformation = $virtualMachineInfo.Split('#').Where({ -not [string]::IsNullOrEmpty($_) })
		
		if ($controllerInformation.Count -gt 0) {
			$controllerInformation | ForEach-Object {
				$controllerOrDevice = [string]::Empty
				$previousKey = [string]::Empty
				$controllerKey = [string]::Empty
				$_ -split [regex]::Escape($port) | ForEach-Object {
					$controllerOrDevice = $port + $_
					if ($controllerOrDevice -match $controllerRegex) {
						$controllerKey = $matches[2]
						
						if ([string]::IsNullOrEmpty($previousKey)) {
							$previousKey = $controllerKey
						}

						if (!$controllersAndDevices.ContainsKey($controllerKey)) {
							$controllersAndDevices[$controllerKey] = @()
						}
						
						$controllers += New-Object PSObject -Property ([ordered]@{
								"Option"         = $controllerOption.ToString()
								"ControllerName" = $matches[2]
								"ControllerType" = $matches[3]
							})
						$controllerOption++
					} 
					elseif (($controllerOrDevice.Contains($vmdk) -or $controllerOrDevice.Contains($partitionedvmdk)) -and $controllerOrDevice -match $deviceRegex) {
						if (-not $previousKey.Equals($controllerKey)) {							
							$previousKey = $controllerKey
							$deviceOption = 1
						}

						$controllersAndDevices[$controllerKey] +=
						New-Object PSObject -Property ([ordered]@{
								"Option"     = $deviceOption
								"Port"       = $matches[1]
								"DeviceGuid" = $matches[2]
								"Location"   = $matches[3]
							})
						$deviceOption++
					}
				}
			}
			BuildTable($controllers)
			$userSelections[$con] = Read-Host "Please Select An Option"
		}
		else {
			throw "No Controllers were found for the vm: $($virtualMachine.VmName)"
		}
	}
}
catch {
	Write-Host $_
	exit
}

# Get The Device Information Selected By The User
try {
	if ($userSelections[$con] -gt $controllers.Count) {
		throw "Option number ( $userSelections[$con] ) is not available. Please re-enter an option number."
	}
	else {
		$controller = $controllers[[int]$userSelections[$con] - 1]
		$controllerAndDevice[$con] = $controller
		$key = $controller.ControllerName
		$devices = $controllersAndDevices[$key]
		if ($devices.Count -gt 0) {
			BuildTable($devices)

			$userSelections[$dev] = Read-Host "Please Select An Option"

			# Get The Device Information Selected By The User
			$controllerAndDevice[$dev] = $devices | ForEach-Object {
				if ($_.Option -eq [int]$userSelections[$dev].ToString()) {
					return $_
				}
			}
		}
		else {
			Write-Host "Skipping Devices Menu. No Attached Devices Found."
		}
	}
}
catch {
 Write-Host "Error Accessing Controller Devices: $_"
}

# Present a List Of Actions To Perform On The Selected Device Until The User Quits
While (!$quit) {
	BuildTable($actions)
	$userSelections[$act] = Read-Host "Please Select An Option"
	switch ($userSelections[$act]) {
		"1" {
			$vmActions.DiscardSavedState( `
					$vBoxManagePath, `
					$virtualMachine.VmName)
		}
		"2" {
			$vmActions.RemoveAttachedDevice( `
					$vBoxManagePath, `
					$virtualMachine.VmName, `
					$controllerAndDevice[$con].ControllerName, `
					$controllerAndDevice[$dev].Port, `
					$controllerAndDevice[$dev].DeviceGuid)
		}
		"3" {
			$vmActions.CloseDisk( `
					$vBoxManagePath, `
					$controllerAndDevice[$dev].Location)
		}
		"4" {
			$vmActions.DeleteRelatedFiles( `
					$controllerAndDevice[$dev].Location)
		}
		"5" {
			$attachPath = $vmActions.CreateRelatedFiles( `
					$vboxManagePath, `
					$controllerAndDevice[$dev].Location, `
					$physicalDriveRegex)
		}
		"6" {
			$vmActions.AttachDevice( `
					$vboxManagePath, `
					$controllerAndDevice[$con].ControllerName, `
					$virtualMachine.VmName, `
					$controllerAndDevice[$dev].Port)
		}
		"q" {
			$userConfirmation = Read-Host "Are you sure you would like to quit? (y or n)"
			if ($userConfirmation -eq "y") {
				$quit = $true
			}
		}
		default {
			Write-Host "Option ( $($userSelections[$act]) ) is out of bounds. Please Try Again, or press q to quit."
		}
	}
}

# Virtual Machine Action Methods
class VmActions {
	VmActions() { }

	# Discard The Saved State Of The Selected Devices Virtual Machine
	DiscardSavedState( `
			[string]$vBoxManagePath, `
			[string]$vmName) `
	{
		try {
			$errorOutput = & $vBoxManagePath discardstate $vmName 2>&1
			Write-Host "Discarding $($vmName)'s Saved State"
		}
		catch {
			Write-Error "Discarding State Failed: $_"
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
		try {
			& $vBoxManagePath storageattach $vmName `
				--storagectl $controllerName `
				--port $port `
				--device 0 `
				--type hdd `
				--medium none
			Write-Host "Removing Attached Storage Device $($deviceGuid)"`
				"From $($controllerName) Controller - Port $($port)"
		}
		catch {
			Write-Error "Removing Attached Storage Device Failed: $_"
		}
	}

	# Close The Device That Was Attached To A Virtual Machine
	CloseDisk([string]$vBoxManagePath, [string]$deviceLocation) {
		try {
			$errorOutput = & $vBoxManagePath closemedium disk $deviceLocation 2>&1
			Write-Host "Closing Medium $($deviceLocation)"
		}
		catch {
			Write-Error "Closing Medium Failed: $_"
		}
	}

	# Delete A VMDK File That Was Attached To A Virtual Machine
	DeleteRelatedFiles( `
			[string]$devicePath) `
	{
		try {
			Write-Host $devicePath
			Remove-Item -Path $devicePath -Force
		}
		catch {
			Write-Host "Deleting File Failed: "
		}
	}

	# Create A New VMDK File(s)
	CreateRelatedFiles( `
			[string]$vBoxManagePath, `
			[string]$oldDevicePath, `
			[string]$physicalDriveRegex) `
	{
		$physicalDriveInfo = [string]::Empty
		$physicalDrives = @()
		$driveOptions = @()
		$option = 0
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
		BuildTable($driveOptions)
		$option = Read-Host "Please Select An Option"
		$partitions = [string]::Empty
		Write-Host $option.GetType()
		Write-Host $driveOptions[$option].Partitions
		switch ($driveOptions[$option].Partitions) {
			"0" {
				Write-Host "The device has no partitions."
			}
			default {
				$result = Read-Host "Please enter a partition"
				if ([int]$result -gt [int]$driveOptions[$option].Partitions) {
					Write-Host "The desired partition is out of range."
				}
				else {
					$partitions = $result.ToString()
				}
			}
		}
		$usePath = [string]::Empty
		if ($oldDevicePath.Length -gt 0) {
			$usePath = Read-Host "Would you like to use ( $oldDevicePath ) for the filename and path? ( y or n )"
		}
		if ($usePath -eq "y") {
			& $vBoxManagePath internalcommands createrawvmdk `
				-filename $oldDevicePath `
				-rawdisk $driveOptions[$option].DeviceID `
				-partitions $partitions
			$oldDevicePath
		}
		else {
			$fileName = Read-Host "Please Enter a File Name"
			$folderPath = Read-Host "Please Enter a Folder Path"
			if ($fileName -match "(?:([a-zA-Z0-9\s\w_-]*)\..*)") {
				Write-Host "Filename is Valid: $fileName"
				if (Test-Path -Path $folderPath -PathType Container) {
					Write-Host "Folder Is Valid: $folderPath"
					if ($folderPath[-1] -eq '\') {
						$folderPath = $folderPath.Substring(0, $folderPath.Length - 1)
					}
					& $vBoxManagePath internalcommands createrawvmdk `
						-filename "$($folderPath)\$($fileName)" `
						-rawdisk $driveOptions[$option].DeviceID `
						-partitions $partitions
					$folderPath
				}
			}
		}
	}

	# Attach A New VMDK File To A New Or Existing SATA Controller On The Virtual Machine
	AttachDevice( `
			[string]$vBoxManagePath, `
			[string]$currentControllerName, `
			[string]$vmName, `
			[int]$currentDevicePort) `
	{
		try {
			$createController = Read-Host "Would you like to create a new controller ( y or n )"
			$pathExists = $false
			$attachPath = Read-Host "Please enter the vmdk file path"
			while (!$pathExists) {
				if (Test-Path $attachPath) {
					Write-Host "The Path Exists"
					$pathExists = $true
				}
				else {
					$attachPath = Read-Host "File Path doesn't exist, please re-enter vmdk file path"
				}
			}
			if ($createController -eq "y") {
				$newControllerName = Read-Host "Please enter a name for the controller"
				& $vBoxManagePath storagectl $vmName `
					--name $newControllerName `
					--add sata `
					--controller IntelAhci
				& $vBoxManagePath storageattach $vmName `
					--storagectl $newControllerName `
					--port 0 `
					--device 0 `
					--type hdd `
					--medium $attachPath
			}
			else {
				$devicePort = $currentDevicePort + 1
				Write-Host "PORT: $($devicePort)"
				& $vBoxManagePath storageattach $vmName `
					--storagectl $currentControllerName `
					--port $devicePort `
					--device 0 `
					--type hdd `
					--medium $attachPath
			}
		}
		catch {
			Write-Host "Attaching vmdk to storage failed: $_"
		}
	}
}

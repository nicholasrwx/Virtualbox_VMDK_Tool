$virtualMachines = @()
$deviceInfo = @()
$finalDeviceInfo = @()
$subDeviceInfo = @()
$deviceOptions = @()
# $controllers = @()		 # Split the second menu into two
# $controllerDevices = @()   # Attached and NoN-Attached Devices
$userSelections = @{ }
$virtualMachine = New-Object PSObject -Property ([ordered]@{ "Option" = [string]::Empty; "VmName" = [string]::Empty; "VmGuid" = [string]::Empty })
$virtualMachineInfo = [string]::Empty
$attachPath = [string]::Empty
$vm = "VirtualMachine"
$dev = "Device"
$act = "Actions"
$vBoxManagePath = Get-ChildItem -Path C:\ -Recurse -File -Filter VBoxManage.exe -ErrorAction SilentlyContinue `
| Select-Object -First 1 | ForEach-Object { $_.FullName }
$vmRegex = '"([^"]+)"\s*{([^}]+)}' # Improve
$deviceRegex = "(\d+):\s'([^']\w+)'(?:.*Port\s([0-9]+))(?:.*UUID:\s([a-zA-Z0-9-]+)?)?(?:\s*.*Location:\s`"([^`"]+)`"?)?"
$physicalDriveRegex = "(?=Partitions\s:\s([0-9]*)DeviceID\s*:\s([\\\.a-zA-Z0-9]*)Model\s*:\s([a-zA-Z0-9\s\._-]*)" +
"Size\s*:\s([0-9]*)Caption\s*:\s([\w\s\._-]*(?=Partitions|$)))"
$controllerRegex = "(?:[0-9]+(:.*)Bootable)"
$vBoxManageRegex = "VBoxManage\.exe"
$startParsingValue = "Storage Controllers:"
$endParsingValue = "NIC 1:"
$port = 'Port '
$count = 1
$startParsing = $false
$endParsing = $false
$quit = $false
$vmActions = [VmActions]::new()
$actions = @(
	New-Object PSObject -Property ([ordered]@{ "Option" = "1"; "Action" = "Discard Machine State" })
	New-Object PSObject -Property ([ordered]@{ "Option" = "2"; "Action" = "Remove Device From Controller" })
	New-Object PSObject -Property ([ordered]@{ "Option" = "3"; "Action" = "Close Disk" })
	New-Object PSObject -Property ([ordered]@{ "Option" = "4"; "Action" = "Delete vmdk file(s)" })
	New-Object PSObject -Property ([ordered]@{ "Option" = "5"; "Action" = "Create vmdk file(s)" })
	New-Object PSObject -Property ([ordered]@{ "Option" = "6"; "Action" = "Attach Device To Controller" })
	New-Object PSObject -Property ([ordered]@{ "Option" = "Q"; "Action" = "Quit" }))

# Constructs a formatted table of options and displays it
function BuildTable
{
	param ([array]$tableOptions)
	$properties = $tableOptions | Get-Member -MemberType Property | Select-Object -ExpandProperty Name
	$tableOptions | Select-Object $properties | Format-Table -AutoSize | Out-Host
}

# Ensure VBoxManage Executable Exists On Users System
try
{
	if ($vBoxManagePath -match $vBoxManageRegex)
	{
		Write-Host "VBoxManage Located: $($vBoxManagePath)"
	}
	else
	{
		throw "Please ensure VBoxManage exists on your system and run the script as admin."
	}
}
catch
{
	Write-Host "VBoxManage Not Located: $($_.Exception.Message)"
	exit 1
}

# Create A List Of VMs and Display The Information
& $vBoxManagePath list vms | ForEach-Object {
	if ($_ -match $vmRegex)
	{
		$vmName = $matches[1]
		$vmGuid = $matches[2]
		$option = $count.ToString()
		$virtualMachines += New-Object PSObject -Property ([ordered]@{
				"Option" = $option
				"VmName" = $vmName
				"VmGuid" = $vmGuid
			})
		$count++
	}
}
BuildTable($virtualMachines)
$userSelections[$vm] = Read-Host "Please Select An Option"

# Capture The Selected VM Controller and Device Information
try
{
	if ($userSelections[$vm] -gt $virtualMachines.Count)
	{
		throw "Option number ( $userSelections[$vm] ) is not available. Please re-enter an option number."
	}
	else
	{
		$virtualMachines | ForEach-Object {
			if ($_.Option -eq [int]$userSelections[$vm].ToString())
			{
				$virtualMachine.Option = $_.Option
				$virtualMachine.VmName = $_.VmName
				$virtualMachine.VmGuid = $_.VmGuid
			}
		}
		Write-Host $virtualMachine
		& $vBoxManagePath showvminfo $virtualMachine.VmGuid | ForEach-Object {
			if (!$startParsing)
			{
				if ($_.Contains($startParsingValue))
				{
					$startParsing = $true
				}
			}
			elseif (!$endParsing)
			{
				if ($_.Contains($endParsingValue))
				{
					$endParsing = $true
				}
				else
				{
					$virtualMachineInfo += $_
				}
			}
		}
		
		# Create List Of Controller and Device Information
		$controllerAndDeviceInfo = $virtualMachineInfo.Split('#').Where({ -not [string]::IsNullOrEmpty($_) })
		$optionNumber = 0
		$controllerAndDeviceInfo | ForEach-Object {
			$controllersAndDevices = @()
			$_ -split [regex]::Escape($port) | ForEach-Object {
				$controllersAndDevices += $port + $_
			}
			
			# Remove Controller Info and Create Devices Array
			$devices = [System.Collections.ArrayList]::new($controllersAndDevices)
			$controllerInfo = [string]::Empty
			if ($controllersAndDevices[0] -match $controllerRegex)
			{
				$controllerInfo = $matches[1]
				######	Write-Host $controllerInfo ( I will have to create another regex like the device regex below )
				$devices.RemoveAt(0)
			}
			
			######
			# create a list of controllers, with a sublist of devices
			# Each VM has one or more controllers
			# Each Controller has one or more devices
			# Each Device has one or more properties
			
			# display the controllers
			
			# displays the devices of the selected controllers
			######
			
			$devices | ForEach-Object {
				$deviceInfo = $_
				$deviceOption = $optionNumber.ToString() + $controllerInfo + $deviceInfo
				if ($deviceOption -match $deviceRegex)
				{
					$deviceOptions += New-Object PSObject -Property ([ordered]@{
							"Option"		 = $matches[1]
							"ControllerName" = $matches[2]
							"Port"		     = $matches[3]
							"DeviceGuid"	 = $matches[4]
							"Location"	     = $matches[5]
						})
					$optionNumber++
				}
			}
		}
	}
}
catch
{
	Write-Host "Error While Parsing: $_"
}
finally
{
	BuildTable($deviceOptions)
}
$userSelections[$dev] = Read-Host "Please Select An Option"

# Get The Device Information Selected By The User
$selectedDevice = $deviceOptions | ForEach-Object {
	if ($_.Option -eq [int]$userSelections[$dev].ToString())
	{
		return $_
	}
}

# Present a List Of Actions To Perform On The Selected Device Until The User Quits  
While (!$quit)
{
	BuildTable($actions)
	$userSelections[$act] = Read-Host "Please Select An Option"
	switch ($userSelections[$act])
	{
		"1" {
			$vmActions.DiscardSavedState( `
				$vBoxManagePath, `
				$virtualMachine.VmName)
		}
		"2" {
			$vmActions.RemoveAttachedDevice( `
				$vBoxManagePath, `
				$virtualMachine.VmName, `
				$selectedDevice.ControllerName, `
				$selectedDevice.Port, `
				$selectedDevice.DeviceGuid)
		}
		"3" {
			$vmActions.CloseDisk( `
				$vBoxManagePath, `
				$selectedDevice.Location)
		}
		"4" {
			$vmActions.DeleteRelatedFiles( `
				$selectedDevice.Location)
		}
		"5" {
			$attachPath = $vmActions.CreateRelatedFiles( `
				$vboxManagePath, `
				$selectedDevice.Location, `
				$physicalDriveRegex)
		}
		"6" {
			$vmActions.AttachDevice( `
				$vboxManagePath, `
				$selectedDevice.ControllerName, `
				$virtualMachine.VmName, `
				$selectedDevice.Port)
		}
		"q" {
			$userConfirmation = Read-Host "Are you sure you would like to quit? (y or n)"
			if ($userConfirmation = "y")
			{
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
		try
		{
			$errorOutput = & $vBoxManagePath discardstate $vmName 2>&1
			Write-Host "Discarding $($vmName)'s Saved State"
		}
		catch
		{
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
		try
		{
			& $vBoxManagePath storageattach $vmName `
							  --storagectl $controllerName `
							  --port $port `
							  --device 0 `
							  --type hdd `
							  --medium none
			Write-Host "Removing Attached Storage Device $($deviceGuid)"`
					   "From $($controllerName) Controller - Port $($port)"
		}
		catch
		{
			Write-Error "Removing Attached Storage Device Failed: $_"
		}
	}
	
	# Close The Device That Was Attached To A Virtual Machine
	CloseDisk([string]$vBoxManagePath, [string]$deviceLocation)
	{
		try
		{
			$errorOutput = & $vBoxManagePath closemedium disk $deviceLocation 2>&1
			Write-Host "Closing Medium $($deviceLocation)"
		}
		catch
		{
			Write-Error "Closing Medium Failed: $_"
		}
	}
	
	# Delete A VMDK File That Was Attached To A Virtual Machine
	DeleteRelatedFiles( `
		[string]$devicePath) `
	{
		try
		{
			Write-Host $devicePath
			Remove-Item -Path $devicePath -Force
		}
		catch
		{
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
			if ($_ -match $physicalDriveRegex)
			{
				$driveOptions += New-Object PSObject -Property ([ordered]@{
						"Option"	 = $option.ToString()
						"Partitions" = $matches[1]
						"DeviceID"   = $matches[2]
						"Model"	     = $matches[3]
						"Size"	     = $matches[4]
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
		switch ($driveOptions[$option].Partitions)
		{
			"0" {
				Write-Host "The device has no partitions."
			}
			default {
				$result = Read-Host "Please enter a partition"
				if ([int]$result -gt [int]$driveOptions[$option].Partitions)
				{
					Write-Host "The desired partition is out of range."
				}
				else
				{
					$partitions = $result.ToString()
				}
			}
		}
		$usePath = [string]::Empty
		if ($oldDevicePath.Length -gt 0)
		{
			$usePath = Read-Host "Would you like to use ( $oldDevicePath ) for the filename and path? ( y or n )"
		}
		if ($usePath -eq "y")
		{
			& $vBoxManagePath internalcommands createrawvmdk `
							  -filename $oldDevicePath `
							  -rawdisk $driveOptions[$option].DeviceID `
							  -partitions $partitions
			$oldDevicePath
		}
		else
		{
			$fileName = Read-Host "Please Enter a File Name"
			$folderPath = Read-Host "Please Enter a Folder Path"
			if ($fileName -match "(?:([a-zA-Z0-9\s\w_-]*)\..*)")
			{
				Write-Host "Filename is Valid: $fileName"
				if (Test-Path -Path $folderPath -PathType Container)
				{
					Write-Host "Folder Is Valid: $folderPath"
					if ($folderPath[-1] -eq '\')
					{
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
		try
		{
			$createController = Read-Host "Would you like to create a new controller ( y or n )"
			$pathExists = $false
			$attachPath = Read-Host "Please enter the vmdk file path"
			while (!$pathExists)
			{
				if (Test-Path $attachPath)
				{
					Write-Host "The Path Exists"
					$pathExists = $true
				}
				else
				{
					$attachPath = Read-Host "File Path doesn't exist, please re-enter vmdk file path"
				}
			}
			if ($createController -eq "y")
			{
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
			else
			{
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
		catch
		{
			Write-Host "Attaching vmdk to storage failed: $_"
		}
	}
}

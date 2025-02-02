. ".\VMDKActions.ps1"
. ".\VMDKHelpers.ps1"

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
	[ordered]@{
		"Option"   = [string]::Empty;
		"VmName"   = [string]::Empty;
		"VmGuid"   = [string]::Empty;
		"Location" = [string]::Empty
 })

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
$dri = "Drive"
$startParsingValue = "Storage Controllers:"
$endParsingValue = "NIC 1:"
$configFile = "Config File:"
$vmdk = ".vmdk"
$dashPartitionVmdk = "-pt.vmdk"
$dotPartitionVmdk = ".pt.vmdk"
$virtualMachineInfo = [string]::Empty
$actionDescriptions = @"
`n`n`
Discard State `t- returns the vm back to its original state ( optional )
Remove Device `t- removes a vmdk from the selected vm controller
Close Disk `t- closes the vmdk as it is still in use by the vm itself ( REQUIRED before deleting vmdk )
Delete Files `t- deletes the selected or newly created vmdk along with its partition file ( -pt.vmdk )
Create Files `t- creates a new vmdk in the currently selected vm root folder
Attach Device `t- attaches the vmdk to the currently selected vm controller
`n`n`
"@

# ending space used to differentiate from 'Ports' when splitting string
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
$physicalDriveRegex = [string]::Concat("(?=Partitions\s:\s([0-9]*)DeviceID\s*:\s([\\\.a-zA-Z0-9]*)",
	"Model\s*:\s([a-zA-Z0-9\s\._-]*)Size\s*:\s([0-9]*)Caption\s*:\s([\w\s\._-]*(?=Partitions|$)))")
$fileNameRegex = "(?:([a-zA-Z0-9\s\w_-]*))"
$fileNameAndExtRegex = "(?:([^\\]*.vmdk))"

# instance containing action methods
$VmActions = [VMDKActions]::new()

# instance containing select method
$helpers = [VMDKHelpers]::new()

# Ensure VBoxManage Executable Exists On Users System
try {
	if ($vBoxManagePath -match $vBoxManageRegex) {
		Write-Output "`nVBoxManage Located: $($vBoxManagePath)`n"
	}
	else {
		throw "Please ensure VBoxManage exists on your system and that you run this script as admin.`n"
	}
}
catch {
	Write-Output "VBoxManage Not Located: $($_.Exception.Message)"
	exit 1
}

# Create A List Of VMs and Display The Information
& $vBoxManagePath list vms | ForEach-Object {

	# parse out a vm path and remove everything after the last slash
	# store it so it is globally accessible
	if ($_ -match $vmRegex) {
		$vmName = $matches[1]
		$vmGuid = $matches[2]
		$option = $vmOption.ToString()
		$location = & $vBoxManagePath showvminfo $vmGuid | Select-String -Pattern $configFile |
		ForEach-Object { ($_ -replace "$configFile\s*", "").Trim(' ', '"') | Split-Path }
		$virtualMachines += New-Object PSObject -Property ([ordered]@{
				"Option"   = $option
				"VmName"   = $vmName
				"VmGuid"   = $vmGuid
				"Location" = $location
			})
		$vmOption++
	}
}
$helpers.BuildTable($virtualMachines)
$helpers.Select($userSelections, $virtualMachines, $vm)

# Capture The Selected VM Controller and Device Information
try {
	$virtualMachines | ForEach-Object {
		if ($_.Option -eq [int]$userSelections[$vm].ToString()) {
			$virtualMachine.Option = $_.Option
			$virtualMachine.VmName = $_.VmName
			$virtualMachine.VmGuid = $_.VmGuid
			$virtualMachine.Location = $_.Location
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
				elseif ($controllerOrDevice.Contains($vmdk) -and $controllerOrDevice -match $deviceRegex) {
					if (-not $previousKey.Equals($controllerKey)) {
						$previousKey = $controllerKey
						$deviceOption = 1
					}

					$controllersAndDevices[$controllerKey] +=
					New-Object PSObject -Property ([ordered]@{
							"Option"     = $deviceOption.ToString()
							"Port"       = $matches[1]
							"DeviceGuid" = $matches[2]
							"Location"   = $matches[3]
						})
					$deviceOption++
				}
			}
		}
		$helpers.BuildTable($controllers)
		$helpers.Select($userSelections, $controllers, $con)
	}
	else {
		$noControllersError = [string]::Concat(
			"No Controllers were found for the vm ( $($virtualMachine.VmName) ).`n",
			"Please add a Controller to the vm and run the script again.")
		throw $noControllersError
	}
}
catch {
	Write-Output $_
	exit
}

# Get The Device Information Selected By The User
try {
	$controller = $controllers[[int]$userSelections[$con] - 1]
	$controllerAndDevice[$con] = $controller
	$key = $controller.ControllerName
	$devices = $controllersAndDevices[$key]
	if ($devices.Count -gt 0) {
		$helpers.BuildTable($devices)
		$helpers.Select($userSelections, $devices, $dev)

		# Get The Device Information Selected By The User
		$controllerAndDevice[$dev] = $devices | ForEach-Object {
			if ($_.Option -eq [int]$userSelections[$dev].ToString()) {
				return $_
			}
		}
	}
	else {
		Write-Output "Skipping Devices Menu. No Attached Devices Found."
	}
}
catch {
 Write-Output "Error Accessing Controller Devices: $_"
}

# Present a List Of Actions To Perform On The Selected Device Until The User Quits
While (!$quit) {
	Write-Output $actionDescriptions
	$helpers.BuildTable($actions)
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
					$controllerAndDevice[$dev].Location, `
					$vmdk, `
					$dashPartitionVmdk, `
					$dotPartitionVmdk)
		}
		"5" {
			$vmActions.CreateRelatedFiles( `
					$vboxManagePath, `
					$virtualMachine.Location, `
					$physicalDriveRegex, `
					$fileNameRegex, `
					$vmdk, `
					$dev, `
					$dri, `
					$deviceOption, `
					$controllerAndDevice, `
					$userSelections, `
					$helpers)
		}
		"6" {
			$vmActions.AttachDevice( `
					$vboxManagePath, `
					$controllerAndDevice, `
					$fileNameAndExtRegex, `
					$virtualMachine.VmName, `
					$con, `
					$dev)
		}
		"q" {
			$userConfirmation = Read-Host "Are you sure you would like to quit? (y or n)"
			if ($userConfirmation -eq "y") {
				$quit = $true
			}
		}
		default {
			Write-Output "Option number ( $($userSelections[$act]) ) is not available."
		}
	}
}

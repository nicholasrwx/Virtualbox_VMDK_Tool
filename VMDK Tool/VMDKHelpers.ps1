class VMDKHelpers {
	VMDKHelpers() { }

	# Used for selecting an option under most menus and storing the value in the userSelections dictionary
	Select( `
			[hashtable]$userSelections, `
			[PSObject[]]$options, `
			[string]$key) {
		$selection = $false
		while (!$selection) {
			$userSelections[$key] = Read-Host "Please Select An Option"
			if (([int]$userSelections[$key] -gt $options.Count) -or ([int]$userSelections[$key] -eq 0)) {
				Write-Host "Option number ( $($userSelections[$key]) ) is not available."
			}
			else {
				$selection = $true
			}
		}
	}

	# Constructs a formatted table of options and displays it
	BuildTable ([PSObject[]]$tableOptions) {
		$properties = $tableOptions | Get-Member -MemberType Property | Select-Object -ExpandProperty Name
		$tableOptions | Select-Object $properties | Format-Table -AutoSize | Out-Host
	}
}

$AXSERVERHOSTNAME = "AXSERVERHOSTNAME"

Try {
	#Get list of VMs
	$VMList = Get-VM -ErrorAction Stop
	
	Do {
		$Index = 1
		Foreach ($VM in $VMList) {
			Write-Host "[$Index] $($VM.Name)"
			$Index++
		}
		
		$Selection = Read-Host "Choose the VM to Migrate"
	} 
	Until ($VMSelection = $VMList[$Selection-1])		
	
	#Check if the VM has any snapshots - Can improve this later to merge/delete maybe?
	If (Get-VMSnapshot $VMSelection) {
		Throw "$($VMSelection.Name) currently has one or more snapshots. Please remove any snapshots prior to migrating this VM"
		#Remove-VMSnapshot maybe?
	}	
	#Select the volume the VM will be exported to
	$AXVVolumes = Get-ChildItem -Path \$AXSERVERHOSTNAME\C$\ClusterStorage\ -Directory -Force -ErrorAction Stop

	Do {
		$Index = 1
		Foreach($Volume in $AXVVolumes) {
			$CurrentVMs = (Get-ChildItem -Path "$($Volume.FullName)\VMs")
			Write-Host "[$Index] $($Volume) currently has $CurrentVMs VMs stored"
			$Index++
		}
		
		$Selection = Read-Host "Choose the Volume to store the exported VM"
	} 
	Until ($VolumeSelection = $AXVVolumes[$Selection-1])

	#Confirm everything is correct
	$Title = "Confirm Migration"
	$Confirm = "Ready to migrate $($VMSelection.Name) to $($VolumeSelection.FullName)\VMs"
	$Choices  = '&Yes', '&No'
	$ConfirmMigration = $Host.UI.PromptForChoice($Title, $Confirm, $Choices, 0)
	If ($ConfirmMigration -eq 0) {
		#Check if the VM is running and attempt to shut it down gracefully
		If ($VMSelection.State -eq 'Running') {
			Write-Host "$($VMSelection.Name) is currently running. Attempting to shut the VM down"
			Stop-VM $VMSelection -ErrorAction Stop
			Write-Host "$($VMSelection.Name) has been shut down"
		}
		#Remove the DVD drive if the VM has one
		Get-VMDvdDrive $VMSelection | Remove-VMDvdDrive
		
		#Attempt to export the VM to the selected volume
		Write-Host "Exporting the VM to the new host"
		Export-VM $VMSelection -Path "$($VolumeSelection.FullName)\VMs" -ErrorAction Stop
		Write-Host "The VM has been exported successfully"
		Write-Host "Run the Import-VM script on axnode1 to complete importing"
		
	}
	Else {
		Throw "The VM migration has been cancelled"
	}
}
Catch {
	Write-Host -ForegroundColor Red $_.Exception.Message
	Return
}

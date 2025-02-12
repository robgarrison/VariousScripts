Try {
	$AXVVolumes = Get-ChildItem -Path "C:\ClusterStorage" -Directory -Force -ErrorAction Stop
	
	#Get the list of VMs stored within the AXVolumes
	Foreach($Volume in $AXVVolumes) {
		$VMFolders += (Get-ChildItem -Path "$($Volume.FullName)\VMs")
	}
	#Get the list of active VMs
	$ClusterNodes = Get-ClusterNode
	$ActiveVMs = @()
	$InactiveVMs = @()
	Foreach ($Cluster in $ClusterNodes) {
		$ActiveVMs += (Get-VM -ComputerName $Cluster.Name)
	}
	Foreach($Folder in $VMFolders){
	If ($Folder.Name -notin $ActiveVMs.Name) {
		$InactiveVMs += $Folder
	}
}
	#Select the VM to Import
	Do {
		$Index = 1
		Foreach($VM in $InactiveVMs){
			Write-Host "[$Index] $($VM.Name)"
			$Index++
		}
		$Selection = Read-Host "Choose the VM to import"
	}
	Until ($VMSelection = $InactiveVMs[$Selection-1])	

	#Import the VM onto the node
	Write-Host "Importing $($VMSelection.Name) to the node"
	$VMConfig = Get-ChildItem -Path "$($VMSelection.Fullname)\Virtual Machines\" -Recurse -Include *.vmcx
	
	Import-VM -Path $VMConfig.FullName -ErrorAction Stop | Out-Null
	Write-Host "$($VMSelection.Name) has been imported"
	
	#Add the VM to the failover cluster
	Write-Host "Adding $($VMSelection.Name) to the Failover Cluster Manager"
	Add-ClusterVirtualMachineRole -VirtualMachine $VMSelection.Name -ErrorAction Stop | Out-Null
	Write-Host "Adding $($VMSelection.Name) has been added to the Failover Cluster Manager"
	
	#Start the VM
	Write-Host "Starting the VM $($VMSelection.Name)"
	Start-Vm -Name $VMSelection.Name -ErrorAction Stop | Out-Null
	Write-Host "$($VMSelection.Name) has been started"
}
Catch {
	Write-Host -ForegroundColor Red $_.Exception.Message
	Return
}

## This script is useful for retrieving information information from Dayforce's API and taking action within AD. 
## The Script can look up Employee ID's for those objects that did not have an employee ID listed. It can also look up the employee's employment status and move the object to another OU if the employee is no longer active.
## Additionally we can assign departmental group membership based on departments that had been assigned within Dayforce.
## We were also using a solution for communication to our retail teams and that access was managed by AD group membership which was managed by this script as well.

$UserName = 'DAYFORCEUSERNAME'
$SecureSecret = 'C:\Scripts\Dayforce-AD-Integration\SecureSecret.txt'
$DFCredentials =New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $UserName, (Get-Content $SecureSecret | ConvertTo-SecureString)
$logFile = "C:\Scripts\Dayforce-AD-Integration\log\Dayforce-AD-Integration_$((Get-Date).ToString("yyyyMMdd")).log"
$ADServer = 'SERVERNAME'

Function Write-Log {
    Param (
		[string] $LogString
	)

    "$(Get-Date): $logstring" | Out-File $logfile -Append -Force
}

#Function used to retrieve the list of active AD user objects
Function GetADUsers {
	$ADUsers = @()
	
	Try {

	   
    $ADUsers = Get-ADUser -SearchBase "OU=FOLDER,EXAMPLE,DC=ad,DC=DOMAIN,DC=com" -Filter 'enabled -eq $true' -Properties * -Server $ADServer
		$ADUsers += Get-ADUser -SearchBase "OU=AnotherFolder,EXAMPLE,DC=ad,DC=DOMAIN,DC=com" -Filter 'enabled -eq $true' -Properties * -Server $ADServer
	}
	Catch {
		Write-Log "Error while retrieving AD Users. Verify OUs are reachable: $($_.Exception.Message)"
		Exit
	}
	Return $ADUsers
}

#Function used to retrieve info from DF
Function QueryDF  {
	
    Param
    (
         [string] $UserFirstName,	
         [string] $UserLastName,
		 [int] $EID
    )	
	
	If (!$EID) {
		$DFMatch =@()
		
		Try {
			$DFUsers = Invoke-RestMethod -Method Get -Uri $DAYFORCERESTENDPOINTURL/V1/Employees?displayName=$UserLastName -Credential $DFCredentials
		}
		Catch {
			Write-Log "Error while retrieving DF Users: $($_.Exception.Message)"
		}
		ForEach ($DFUser in $DFUsers.data) {
			$DFUserRecord = Invoke-RestMethod -Method Get -Uri $DAYFORCERESTENDPOINTURLV1/Employees/$($DFUser.XRefCode)?expand=WorkAssignments%2CEmployeeManagers%2CEmploymentStatuses -Credential $DFCredentials
			If ($DFUserRecord.Data.EmploymentStatuses.Items.EmploymentStatus.ShortName -eq 'Active' -and $DFUserRecord.Data.WorkAssignments.Items.Position.Department.ShortName -ne 'Retail' -and $DFUserRecord.Data.FirstName -eq $UserFirstName -and $DFUserRecord.Data.LastName -eq $UserLastName) {
				 $DFMatch += $DFUserRecord.Data				 
			}
		}
		#Multiple possible matches and set EID to NO_MATCH
		If ($DFMatch.count -gt 1) {
			Write-Log "Multiple matches found for $($ADUser.SamAccountName)"			
			ForEach ($Match in $DFMatch) {							
				Write-Log "Possible match - $($Match.FirstName) $($Match.LastName) $($Match.EmployeeNumber)"
			}			
			If ($ADUser.EmployeeID -ne 'NO_MATCH') {
				Set-ADUser -Identity $ADUser -EmployeeID 'NO_MATCH' -Server $ADServer | Out-Null
			}
			Return
		}
		#Assume this is the correct match and assign EID
		Elseif ($DFMatch.count -eq 1) {
			Write-Log "Match found for $($ADUser.SamAccountName). Assigned Employee ID: $($DFMatch.EmployeeNumber)"
			Set-ADUser -Identity $ADUser -EmployeeID $DFMatch.EmployeeNumber -Server $ADServer | Out-Null
			SetADUserDetails
		}
		#No Matches found set EID to NO_MATCH
		Else {
			Write-Log "No matches were found for $($ADUser.SamAccountName)"
			If ($ADUser.EmployeeID -ne 'NO_MATCH') {
				Set-ADUser -Identity $ADUser -EmployeeID 'NO_MATCH' -Server $ADServer | Out-Null
			}
			Return		
		}
	}
	Elseif ($EID) {
		Try {
			$DFMatch = (Invoke-RestMethod -Method Get -Uri $DAYFORCERESTENDPOINTURL/V1/Employees/$($EID)?expand=WorkAssignments%2CEmployeeManagers%2CEmploymentStatuses -Credential $DFCredentials).Data
		}
		Catch {
			Write-Log "Error while retrieving DF Users: $($_.Exception.Message)"
		}		
		SetADUserDetails
	}
}
Function SetADUserDetails {
	#If the department exist in DF attempt to match in AD
	If ($DFMatch.WorkAssignments.Items.Position.Department.ShortName -and $ADUser.Department -ne $DFMatch.WorkAssignments.Items.Position.Department.ShortName) {
		Write-Log "$($ADUser.SamAccountName) assigned Department: $($DFMatch.WorkAssignments.Items.Position.Department.ShortName)"
		Set-ADUser -Identity $ADUser -Department $DFMatch.WorkAssignments.Items.Position.Department.ShortName -Server $ADServer | Out-Null		
	}
	#If the title is provided in DF attempt to match in AD
	If ($DFMatch.WorkAssignments.Items.Position.Job.ShortName -and $ADUser.Title -ne $DFMatch.WorkAssignments.Items.Position.Job.ShortName) {
		Write-Log "$($ADUser.SamAccountName) assigned Title: $($DFMatch.WorkAssignments.Items.Position.Job.ShortName)"
		Set-ADUser -Identity $ADUser -Title $DFMatch.WorkAssignments.Items.Position.Job.ShortName -Server $ADServer | Out-Null
	}
	#If the manager is provided in DF attempt to match in AD
	If ($DFMatch.EmployeeManagers.Items.ManagerXrefCode) {
		Try {
			$ADUserManager = Get-ADUser -Filter * -Properties samaccountname,employeeID -Server $ADServer | Where {$_.EmployeeID -eq $DFMatch.EmployeeManagers.Items.ManagerXrefCode}
			If ($ADUserManager -and $ADUser.Manager -ne $ADUserManager) {
				Write-Log "$($ADUser.SamAccountName) assigned Manager: $($ADUserManager.SamAccountName)"
				Set-ADUser -Identity $ADUser -Manager $ADUserManager.SamAccountName -Server $ADServer | Out-Null
			}
		}
		Catch {
			Write-Log "Error while retrieving manager information for $($ADUser.SamAccountName): $($_.Exception.Message)"
		}
	}
	#If The account is not active in DF disable the account and move to recently disabled folder in AD
	If ($DFMatch.EmploymentStatuses.items.EmploymentStatus.XRefCode -and $DFMatch.EmploymentStatuses.items.EmploymentStatus.XRefCode -ne 'ACTIVE') {		
		Write-Log "$($ADUser.SamAccountName) was not active in DF and has been disabled in AD"
		Disable-ADAccount $ADUser -Server $ADServer
		
		Switch ($ADUser.DistinguishedName.split(',')[1]) {
			'OU=Users-DC1' {
				Write-Log "$($ADUser.SamAccountName) object moved to Recently-Disabled-Users-DC1"
				Move-ADObject -Identity $ADUser -TargetPath 'OU=Disabled-Users,DC=ad,DC=EXAMPLE,DC=com' -Server $ADServer | Out-Null
			}
			'OU=Users-DC4' {
				Write-Log "$($ADUser.SamAccountName) object moved to Recently-Disabled-Users-DC4"
				Move-ADObject -Identity $ADUser -TargetPath 'OU=Recently-ANOTHERDISABLEDUSERSFOLDER,OU=Disabled Users,DC=ad,DC=EXAMPLE,DC=com' -Server $ADServer | Out-Null
			}
			'OU=Utility-Users-DC4' {
				Write-Log "$($ADUser.SamAccountName) object moved to Recently-Disabled-Utility-Users-DC4"
				Move-ADObject -Identity $ADUser -TargetPath 'OU=ATHIRDDISABLEDUSERSFOLDER,OU=Disabled Users,DC=ad,DC=EXAMPLE,DC=com' -Server $ADServer | Out-Null
			}
		}
        #If the user was disabled return a true value ( for skipping group membership asignment )
        $UserDisabled = $True
        Return $UserDisabled
	}
}
Function SetADUserGroupMembership {
	$GG_DF_Groups = 'GG_DF_Accounting','GG_DF_Apparel','GG_DF_Bulk','GG_DF_Facilities','GG_DF_HR','GG_DF_Inbound','GG_DF_Inventory_Control','GG_DF_IT','GG_DF_Loss_Prevention','GG_DF_Marketing','GG_DF_Merchandising_Wholesale','GG_DF_NIC','GG_DF_Online','GG_DF_Outbound','GG_DF_Real_Estate_Support','GG_DF_Retail_Support','GG_DF_Reverse_Logistics','GG_DF_SGA','GG_DF_Shipping_Receiving','GG_DF_Supply_Chain','GG_DF_Supply_Chain_Support','GG_DF_Support_Services','GG_DF_Transportation'
	
	If ($ADUser.Department) {
		Switch ($ADUser.Department) {
			'Accounting' {
				If (!(Get-ADGroupMember -Identity 'GG_DF_Accounting' | Where-Object {$_.SamAccountName -eq $ADUser.SamAccountName})) {
					Try {
						Add-ADGroupMember -Identity GG_DF_Accounting -Members $ADUser -Server $ADServer | Out-Null
						Write-Log "$($ADUser.SamAccountName) has been added to GG_DF_Accounting"						
					}
					Catch {
						Write-Log "Error while adding user $($ADUser.SamAccountName) to GG_DF_Accounting: $($_.Exception.Message)"
					}
				}
				$GG_DF_Groups = $GG_DF_Groups -ne 'GG_DF_Accounting'
				SetZipLineHQMembership
			 }
			'Apparel' {
				If (!(Get-ADGroupMember -Identity 'GG_DF_Apparel' | Where-Object {$_.SamAccountName -eq $ADUser.SamAccountName})) {
					Try {
						Add-ADGroupMember -Identity GG_DF_Apparel -Members $ADUser -Server $ADServer | Out-Null
						Write-Log "$($ADUser.SamAccountName) has been added to GG_DF_Apparel"
					}
					Catch {
						Write-Log "Error while adding user $($ADUser.SamAccountName) to GG_DF_Apparel: $($_.Exception.Message)"
					}
				}
				$GG_DF_Groups = $GG_DF_Groups -ne 'GG_DF_Apparel'
				SetZipLineHQMembership
			 }
			'Bulk' {
				If (!(Get-ADGroupMember -Identity 'GG_DF_Bulk' | Where-Object {$_.SamAccountName -eq $ADUser.SamAccountName})) {
					Try {
						Add-ADGroupMember -Identity GG_DF_Bulk -Members $ADUser -Server $ADServer | Out-Null
						Write-Log "$($ADUser.SamAccountName) has been added to GG_DF_Bulk"
					}
					Catch {
						Write-Log "Error while adding user $($ADUser.SamAccountName) to GG_DF_Bulk: $($_.Exception.Message)"
					}
				}
				$GG_DF_Groups = $GG_DF_Groups -ne 'GG_DF_Bulk'
				SetZipLineHQMembership
			 }
			'Facilities' {
				If (!(Get-ADGroupMember -Identity 'GG_DF_Facilities' | Where-Object {$_.SamAccountName -eq $ADUser.SamAccountName})) {
					Try {
						Add-ADGroupMember -Identity GG_DF_Facilities -Members $ADUser -Server $ADServer | Out-Null
						Write-Log "$($ADUser.SamAccountName) has been added to GG_DF_Facilities"
					}
					Catch {
						Write-Log "Error while adding user $($ADUser.SamAccountName) to GG_DF_Facilities: $($_.Exception.Message)"
					}
				}
				$GG_DF_Groups = $GG_DF_Groups -ne 'GG_DF_Facilities'
				SetZipLineHQMembership
			 }
			'Human Resources' {
				If (!(Get-ADGroupMember -Identity 'GG_DF_HR' | Where-Object {$_.SamAccountName -eq $ADUser.SamAccountName})) {
					Try {
						Add-ADGroupMember -Identity GG_DF_HR -Members $ADUser -Server $ADServer | Out-Null
						Write-Log "$($ADUser.SamAccountName) has been added to GG_DF_HR"
					}
					Catch {
						Write-Log "Error while adding user $($ADUser.SamAccountName) to GG_DF_HR: $($_.Exception.Message)"
					}
				}
				$GG_DF_Groups = $GG_DF_Groups -ne 'GG_DF_HR'
				SetZipLineHQMembership
			 }
			'Inbound' {
				If (!(Get-ADGroupMember -Identity 'GG_DF_Inbound' | Where-Object {$_.SamAccountName -eq $ADUser.SamAccountName})) {
					Try {
						Add-ADGroupMember -Identity GG_DF_Inbound -Members $ADUser -Server $ADServer | Out-Null
						Write-Log "$($ADUser.SamAccountName) has been added to GG_DF_Inbound"
					}
					Catch {
						Write-Log "Error while adding user $($ADUser.SamAccountName) to GG_DF_Inbound: $($_.Exception.Message)"
					}
				}
				$GG_DF_Groups = $GG_DF_Groups -ne 'GG_DF_Inbound'
				SetZipLineHQMembership
			 }
			'Inventory Control' {
				If (!(Get-ADGroupMember -Identity 'GG_DF_Inventory_Control' | Where-Object {$_.SamAccountName -eq $ADUser.SamAccountName})) {
					Try {
						Add-ADGroupMember -Identity GG_DF_Inventory_Control -Members $ADUser -Server $ADServer | Out-Null
						Write-Log "$($ADUser.SamAccountName) has been added to GG_DF_Inventory_Control"
					}
					Catch {
						Write-Log "Error while adding user $($ADUser.SamAccountName) to GG_DF_Inventory_Control: $($_.Exception.Message)"
					}
				}
				$GG_DF_Groups = $GG_DF_Groups -ne 'GG_DF_Inventory_Control'
				SetZipLineHQMembership
			 }
			'IT' {
				If (!(Get-ADGroupMember -Identity 'GG_DF_IT' | Where-Object {$_.SamAccountName -eq $ADUser.SamAccountName})) {
					Try {
						Add-ADGroupMember -Identity GG_DF_IT -Members $ADUser -Server $ADServer | Out-Null
						Write-Log "$($ADUser.SamAccountName) has been added to GG_DF_IT"						
					}
					Catch {
						Write-Log "Error while adding user $($ADUser.SamAccountName) to GG_DF_IT: $($_.Exception.Message)"
					}
				}
				$GG_DF_Groups = $GG_DF_Groups -ne 'GG_DF_IT'
				SetZipLineHQMembership
			 }
			'Loss Prevention' {
				If (!(Get-ADGroupMember -Identity 'GG_DF_Loss_Prevention' | Where-Object {$_.SamAccountName -eq $ADUser.SamAccountName})) {
					Try {
						Add-ADGroupMember -Identity GG_DF_Loss_Prevention -Members $ADUser -Server $ADServer | Out-Null
						Write-Log "$($ADUser.SamAccountName) has been added to GG_DF_Loss_Prevention"
					}
					Catch {
						Write-Log "Error while adding user $($ADUser.SamAccountName) to GG_DF_Loss_Prevention: $($_.Exception.Message)"
					}
				}
				$GG_DF_Groups = $GG_DF_Groups -ne 'GG_DF_Loss_Prevention'
				SetZipLineHQMembership
			 }
			'Marketing' {
				If (!(Get-ADGroupMember -Identity 'GG_DF_Marketing' | Where-Object {$_.SamAccountName -eq $ADUser.SamAccountName})) {
					Try {
						Add-ADGroupMember -Identity GG_DF_Marketing -Members $ADUser -Server $ADServer | Out-Null
						Write-Log "$($ADUser.SamAccountName) has been added to GG_DF_Marketing"
					}
					Catch {
						Write-Log "Error while adding user $($ADUser.SamAccountName) to GG_DF_Marketing: $($_.Exception.Message)"
					}
				}
				$GG_DF_Groups = $GG_DF_Groups -ne 'GG_DF_Marketing'
				SetZipLineHQMembership
			 }
			'Merchandising and Wholesale' {
				If (!(Get-ADGroupMember -Identity 'GG_DF_Merchandising_Wholesale' | Where-Object {$_.SamAccountName -eq $ADUser.SamAccountName})) {
					Try {
						Add-ADGroupMember -Identity GG_DF_Merchandising_Wholesale -Members $ADUser -Server $ADServer | Out-Null
						Write-Log "$($ADUser.SamAccountName) has been added to GG_DF_Merchandising_Wholesale"
					}
					Catch {
						Write-Log "Error while adding user $($ADUser.SamAccountName) to GG_DF_Merchandising_Wholesale: $($_.Exception.Message)"
					}
				}
				$GG_DF_Groups = $GG_DF_Groups -ne 'GG_DF_Merchandising_Wholesale'
				SetZipLineHQMembership
			 }
			'NIC' {
				If (!(Get-ADGroupMember -Identity 'GG_DF_NIC' | Where-Object {$_.SamAccountName -eq $ADUser.SamAccountName})) {
					Try {
						Add-ADGroupMember -Identity GG_DF_NIC -Members $ADUser -Server $ADServer | Out-Null
						Write-Log "$($ADUser.SamAccountName) has been added to GG_DF_NIC"
					}
					Catch {
						Write-Log "Error while adding user $($ADUser.SamAccountName) to GG_DF_NIC: $($_.Exception.Message)"
					}
				}
				$GG_DF_Groups = $GG_DF_Groups -ne 'GG_DF_NIC'
				SetZipLineHQMembership
			 }
			'Online' {
				If (!(Get-ADGroupMember -Identity 'GG_DF_Online' | Where-Object {$_.SamAccountName -eq $ADUser.SamAccountName})) {
					Try {
						Add-ADGroupMember -Identity GG_DF_Online -Members $ADUser -Server $ADServer | Out-Null
						Write-Log "$($ADUser.SamAccountName) has been added to GG_DF_Online"
					}
					Catch {
						Write-Log "Error while adding user $($ADUser.SamAccountName) to GG_DF_Online: $($_.Exception.Message)"
					}
				}
				$GG_DF_Groups = $GG_DF_Groups -ne 'GG_DF_Online'
				SetZipLineHQMembership
			 }
			'Outbound' {
				If (!(Get-ADGroupMember -Identity 'GG_DF_Outbound' | Where-Object {$_.SamAccountName -eq $ADUser.SamAccountName})) {
					Try {
						Add-ADGroupMember -Identity GG_DF_Outbound -Members $ADUser -Server $ADServer | Out-Null
						Write-Log "$($ADUser.SamAccountName) has been added to GG_DF_Outbound"
					}
					Catch {
						Write-Log "Error while adding user $($ADUser.SamAccountName) to GG_DF_Outbound: $($_.Exception.Message)"
					}
				}
				$GG_DF_Groups = $GG_DF_Groups -ne 'GG_DF_Outbound'
				SetZipLineHQMembership
			 }
			'Real Estate Support' {
				If (!(Get-ADGroupMember -Identity 'GG_DF_Real_Estate_Support' | Where-Object {$_.SamAccountName -eq $ADUser.SamAccountName})) {
					Try {
						Add-ADGroupMember -Identity GG_DF_Real_Estate_Support -Members $ADUser -Server $ADServer | Out-Null
						Write-Log "$($ADUser.SamAccountName) has been added to GG_DF_Real_Estate_Support"
					}
					Catch {
						Write-Log "Error while adding user $($ADUser.SamAccountName) to GG_DF_Real_Estate_Support: $($_.Exception.Message)"
					}
				}
				$GG_DF_Groups = $GG_DF_Groups -ne 'GG_DF_Real_Estate_Support'
				SetZipLineHQMembership
			 }
			'Retail Support' {
				If (!(Get-ADGroupMember -Identity 'GG_DF_Retail_Support' | Where-Object {$_.SamAccountName -eq $ADUser.SamAccountName})) {
					Try {
						Add-ADGroupMember -Identity GG_DF_Retail_Support -Members $ADUser -Server $ADServer | Out-Null
						Write-Log "$($ADUser.SamAccountName) has been added to GG_DF_Retail_Support"
					}
					Catch {
						Write-Log "Error while adding user $($ADUser.SamAccountName) to GG_DF_Retail_Support: $($_.Exception.Message)"
					}
				}
				$GG_DF_Groups = $GG_DF_Groups -ne 'GG_DF_Retail_Support'
				SetZipLineHQMembership
			 }
			'Reverse Logistics' {
				If (!(Get-ADGroupMember -Identity 'GG_DF_Reverse_Logistics' | Where-Object {$_.SamAccountName -eq $ADUser.SamAccountName})) {
					Try {
						Add-ADGroupMember -Identity GG_DF_Reverse_Logistics -Members $ADUser -Server $ADServer | Out-Null
						Write-Log "$($ADUser.SamAccountName) has been added to GG_DF_Reverse_Logistics"
					}
					Catch {
						Write-Log "Error while adding user $($ADUser.SamAccountName) to GG_DF_Reverse_Logistics: $($_.Exception.Message)"
					}
				}
				$GG_DF_Groups = $GG_DF_Groups -ne 'GG_DF_Reverse_Logistics'
				SetZipLineHQMembership
			 }
			'SGA' {
				If (!(Get-ADGroupMember -Identity 'GG_DF_SGA' | Where-Object {$_.SamAccountName -eq $ADUser.SamAccountName})) {
					Try {
						Add-ADGroupMember -Identity GG_DF_SGA -Members $ADUser -Server $ADServer | Out-Null
						Write-Log "$($ADUser.SamAccountName) has been added to GG_DF_SGA"
					}
					Catch {
						Write-Log "Error while adding user $($ADUser.SamAccountName) to GG_DF_SGA: $($_.Exception.Message)"
					}
				}
				$GG_DF_Groups = $GG_DF_Groups -ne 'GG_DF_SGA'
				SetZipLineHQMembership
			 }
			'Shipping and Receiving' {
				If (!(Get-ADGroupMember -Identity 'GG_DF_Shipping_Receiving' | Where-Object {$_.SamAccountName -eq $ADUser.SamAccountName})) {
					Try {
						Add-ADGroupMember -Identity GG_DF_Shipping_Receiving -Members $ADUser -Server $ADServer | Out-Null
						Write-Log "$($ADUser.SamAccountName) has been added to GG_DF_Shipping_Receiving"
					}
					Catch {
						Write-Log "Error while adding user $($ADUser.SamAccountName) to GG_DF_Shipping_Receiving: $($_.Exception.Message)"
					}
				}
				$GG_DF_Groups = $GG_DF_Groups -ne 'GG_DF_Shipping_Receiving'
				SetZipLineHQMembership
			 }
			'Supply Chain' {
				If (!(Get-ADGroupMember -Identity 'GG_DF_Supply_Chain' | Where-Object {$_.SamAccountName -eq $ADUser.SamAccountName})) {
					Try {
						Add-ADGroupMember -Identity GG_DF_Supply_Chain -Members $ADUser -Server $ADServer | Out-Null
						Write-Log "$($ADUser.SamAccountName) has been added to GG_DF_Supply_Chain"
					}
					Catch {
						Write-Log "Error while adding user $($ADUser.SamAccountName) to GG_DF_Supply_Chain: $($_.Exception.Message)"
					}
				}
				$GG_DF_Groups = $GG_DF_Groups -ne 'GG_DF_Supply_Chain'
				SetZipLineHQMembership
			 }
			'SUPPLY CHAIN SUPPORT' {
				If (!(Get-ADGroupMember -Identity 'GG_DF_Supply_Chain_Support' | Where-Object {$_.SamAccountName -eq $ADUser.SamAccountName})) {
					Try {
						Add-ADGroupMember -Identity GG_DF_Supply_Chain_Support -Members $ADUser -Server $ADServer | Out-Null
						Write-Log "$($ADUser.SamAccountName) has been added to GG_DF_Supply_Chain_Support"
					}
					Catch {
						Write-Log "Error while adding user $($ADUser.SamAccountName) to GG_DF_Supply_Chain_Support: $($_.Exception.Message)"
					}
				}
				$GG_DF_Groups = $GG_DF_Groups -ne 'GG_DF_Supply_Chain_Support'
				SetZipLineHQMembership
			 }
			'Support Services' {
				If (!(Get-ADGroupMember -Identity 'GG_DF_Support_Services' | Where-Object {$_.SamAccountName -eq $ADUser.SamAccountName})) {
					Try {
						Add-ADGroupMember -Identity GG_DF_Support_Services -Members $ADUser -Server $ADServer | Out-Null
						Write-Log "$($ADUser.SamAccountName) has been added to GG_DF_Support_Services"
					}
					Catch {
						Write-Log "Error while adding user $($ADUser.SamAccountName) to GG_DF_Support_Services: $($_.Exception.Message)"
					}
				}
				$GG_DF_Groups = $GG_DF_Groups -ne 'GG_DF_Support_Services'
				SetZipLineHQMembership
			 }
			'Transportation' {
				If (!(Get-ADGroupMember -Identity 'GG_DF_Transportation' | Where-Object {$_.SamAccountName -eq $ADUser.SamAccountName})) {
					Try {
						Add-ADGroupMember -Identity GG_DF_Transportation -Members $ADUser -Server $ADServer | Out-Null
						Write-Log "$($ADUser.SamAccountName) has been added to GG_DF_Transportation"
					}
					Catch {
						Write-Log "Error while adding user $($ADUser.SamAccountName) to GG_DF_Transportation: $($_.Exception.Message)"
					}
				}
				$GG_DF_Groups = $GG_DF_Groups -ne 'GG_DF_Transportation'
				SetZipLineHQMembership
			}
		}
		#Remove the user from any of the other GG_DF_Groups
		ForEach ($Group in (Get-ADPrincipalGroupMembership -Identity $ADUser -Server $ADServer | Where {$_.GroupCategory -eq 'Security'})) {
			If ($Group.Name -in $GG_DF_Groups) {
				Try {
					Remove-ADGroupMember -Identity $Group -Members $ADUser -Server $ADServer -Confirm:$false | Out-Null
					Write-Log "$($ADUser.SamAccountName) has been removed from $($Group.Name)"
				}
				Catch {
					Write-Log "Error while removing user $($ADUser.SamAccountName) from $($Group.Name): $($_.Exception.Message)"
				}
			}
		}
	}
}
Function SetZipLineHQMembership {
	If (!(Get-ADGroupMember -Identity 'AZ_ZIPLINE_HQ' | Where-Object {$_.SamAccountName -eq $ADUser.SamAccountName})) {
		Try {
			Add-ADGroupMember -Identity AZ_ZIPLINE_HQ -Members $ADUser -Server $ADServer | Out-Null
			Write-Log "$($ADUser.SamAccountName) has been added to AZ_ZIPLINE_HQ"
		}
		Catch {
			Write-Log "Error while adding user $($ADUser.SamAccountName) to AZ_ZIPLINE_HQ: $($_.Exception.Message)"
		}
	}
}


#Retrieve initial list of all AD users

Write-Log "Retrieving AD Users."

$ADUsers = GetADUsers

If (([array]$ADUsers).Count -ge 1) {
	Write-Log "Retrieved $(([array]$ADUsers).Count) Users"
	Write-Log "Starting processing of DF information and group membership"
	#Iterate over AD users and find corresponding DF Info
	ForEach ($ADUser in $ADUsers) {
        $QueryDFResult = $Null

		If (($ADUser.EmployeeID -eq $Null) -or ($ADUser.EmployeeID -eq 'NO_MATCH')) {
			If (($ADUser.GivenName -eq $Null) -or ($ADUser.SurName -eq $Null)) {
				Write-Log "Verify first and last name are set for: $($ADUser.SamAccountName)"
				If ($ADUser.EmployeeID -ne 'NO_MATCH') {
					Set-ADUser -Identity $ADUser -EmployeeID 'NO_MATCH' -Server $ADServer | Out-Null
				}
			}
			Else {
				$QueryDFResult = QueryDF -UserFirstName $ADUser.GivenName -UserLastName $ADUser.SurName
                #If the user was not disabled during the query process group membership
                If ($QueryDFResult -ne $True) {
                    SetADUserGroupMembership
                }
			}
		}
		Elseif ($ADUser.EmployeeID -ne $null -and $ADUser.EmployeeID -ne '37013' -and $ADUser.EmployeeID -ne 'temp' -and $ADUser.EmployeeID -ne 'NO_MATCH') {
			$QueryDFResult = (QueryDF -EID $ADUser.EmployeeID)
			#If the user was not disabled during the query process group membership
            If ($QueryDFResult -ne $True) {
                SetADUserGroupMembership
            }
		}
	}
	Write-Log "Finished processing of DF information and group membership"
}



#Retrieve the list of accounts missing Employee IDs
Write-Log "Retrieving AD Users missing Employee IDs"

$NoMatchADUsers = GetADUsers | Where {$_.employeeid -eq 'NO_MATCH'}
#If users are found with missing Employee IDs add them to the log
If ($NoMatchADUsers) {
	Write-Log "The following users do not currently have employee IDs:"
	ForEach ($NoMatchADUser in $NoMatchADUsers) {
		Write-Log "$($NoMatchADUser.SamAccountName)"
	}
}
Else {
	Write-Log "No Users Missing Employee IDs found."
}

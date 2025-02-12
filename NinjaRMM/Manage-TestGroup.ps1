# Create the SecureSecret file using the following command
# $ClientSecret = "SECRETGOESHERE"
# $ClientSecret | ConvertTo-SecureString -AsPlainText -Force | ConvertFrom-SecureString | Out-File "C:\Scripts\Ninja\SecureSecret.txt"

$ClientID = 'SECRETIDGOESHERE'
$ClientSecret = Get-Content "C:\Scripts\Manage-TestGroup\SecureSecret.txt" | ConvertTo-SecureString
$LogFile = "C:\ScriptsManage-TestGroup\log\Manage-NinjaTestGroups_$((Get-Date).ToString("yyyyMMdd")).log"
$OUPaths = "OU=TESTOU,DC=ad,DC=EXAMPLE,DC=com","OU=ANOTHERTESTOU,DC=ad,DC=EXAMPLE,DC=com"
$TokenFile = "C:\Scripts\Manage-TestGroup\SecureToken.xml"

#Work around PS 5.1 limitation
$BSTR = [System.Runtime.InteropServices.Marshal]::SecureStringToBSTR($ClientSecret)
$CSPT = [System.Runtime.InteropServices.Marshal]::PtrToStringAuto($BSTR)
Function Write-Log {
    Param ($LogString)

    "$(Get-Date): $LogString" | Out-File $LogFile -Append -Force
}

If (Test-Path $TokenFile) {
    $MyToken = (Import-Clixml -Path $TokenFile).GetNetworkCredential().Password | ConvertFrom-Json
    Write-Log "Token found. Using existing credentials."
}
If (!$Headers){
    $Headers=@{}
    $Headers.Add("Content-Type", "application/x-www-form-urlencoded")
    $Headers.Add("Accept", "application/json")
}
Try {
    If (!$MyToken) {           
        #If we don't already have a token saved, let's generate a new one 
        $Body = "grant_type=client_credentials&client_id=$clientID&client_secret=$CSPT&scope=monitoring+management+control+offline_access"
        $MyToken = Invoke-RestMethod -Uri 'https://app.ninjarmm.com/ws/oauth/token' -Method POST -Headers $Headers -ContentType 'application/x-www-form-urlencoded' -Body $Body
        $MyToken | Add-Member -MemberType NoteProperty -Name "ExpireTime" -Value (get-date).AddSeconds($myToken.expires_in) -Force
        $NewTokenJSON = $MyToken | ConvertTo-Json
        $SecureJSON = ConvertTo-SecureString $NewTokenJSON -AsPlainText -Force
        $CredObject = [pscredential]::new('foo', $SecureJson)
        $CredObject | Export-Clixml -Path $TokenFile    
        Write-Log "The token Has been generated and saved to file."
    }
    Elseif ((Get-Date) -gt $MyToken.ExpireTime) {
        #If the token is present and has expired, let's grab a new one
        $Body = "grant_type=refresh_token&client_id=$clientID&client_secret=$CSPT&refresh_token="+$myToken.refresh_token
        $MyToken = Invoke-RestMethod -Uri 'https://app.ninjarmm.com/ws/oauth/token' -Method POST -Headers $Headers -ContentType 'application/x-www-form-urlencoded' -Body $Body
        $MyToken | Add-Member -MemberType NoteProperty -Name "ExpireTime" -Value (get-date).AddSeconds($myToken.expires_in) -Force
        $NewTokenJSON = $MyToken | ConvertTo-Json
        $SecureJSON = ConvertTo-SecureString $NewTokenJSON -AsPlainText -Force
        $CredObject = [pscredential]::new('foo', $SecureJson)
        $CredObject | Export-Clixml -Path $TokenFile
        Write-Log "The token has been updated and saved to file."
    }
}
Catch {
    Write-Log "Could not obtain token from NinjaAPI."
    Write-Log "Error: $($_.Exception.Message)"
    Break
}
#Future proofing token expiring
If (!$Headers.Authorization) {
    $Headers.Add("Authorization", "Bearer "+$MyToken.access_token)
}
Else {
    $Headers.Authorization = "Bearer "+$MyToken.access_token
}

#Let's get all the policies with their corresponding ID info
Try{
    $Policies = Invoke-RestMethod -Uri 'https://app.ninjarmm.com/v2/policies' -Method GET -Headers $Headers
    #Pull out the policies we are looking for
    $TESTGROUP1PolicyID = ($Policies | Where-Object {$_.name -eq "TESTGROUP1"}).id
    $TESTGROUP2TestGroupPolicyID = ($Policies | Where-Object {$_.name -eq "TESTGROUP2"}).id
    Write-Log "Obtained Policy IDs from NinjaAPI."
}
Catch {
    Write-Log "Could not obtain list of policies from NinjaAPI. Check if token is valid."
    Write-Log "Error: $($_.Exception.Message)"
    Break
}

#Get All Devices
Try {
    $AllDevices = Invoke-RestMethod -Uri "https://app.ninjarmm.com/v2/devices/" -Method GET -Headers $Headers
}
Catch {
    Write-Log "Could not obtain list of all devices from NinjaAPI. Check if token is valid."
    Write-Log "Error: $($_.Exception.Message)"
    Break
}
#Find Duplicates - leaving this here in case needed later
#$AllDevices.systemName | Group | Where {$_.Count -gt 1}

#Get all the register1 devices
$Register1Devices = $AllDevices | Where-Object {$_.systemName -like "HOSTNAME"}

#Find any Register 1 devices aren't in the test group policy and add them
ForEach ($Register in $Register1Devices) {
    If ($Register.policyId -ne $EXAMPLETestGroupPolicyID) {
        $URI = "https://app.ninjarmm.com/v2/device/"+$Register.id
        Invoke-RestMethod -Uri $URI -Method PATCH -Headers $Headers -ContentType 'application/json' -Body (@{policyId=$EXAMPLETestGroupPolicyID} | ConvertTo-Json)
        Write-Log "$($Register.systemName) has been assigned to the EXAMPLE test group policy"
    }
}

#Get all IT computer devices from the IT-Computers OU's
ForEach ($Path in $OUPaths) {
    $EXAMPLEs += Get-ADComputer -Filter * -SearchBase $Path
}
#Find any Register 1 devices aren't in the test group policy and add them
ForEach ($EXAMPLE in $EXAMPLEs) {    
    If ($EXAMPLE.Name -in $AllDevices.systemName) {
        $CurrentDevice = $AllDevices |  Where-Object {$_.systemName -like $EXAMPLE.Name}
        If (!$CurrentDevice.policyId -or $CurrentDevice.policyId -ne $CorporateTestGroupPolicyID) {
            $URI = "https://app.ninjarmm.com/v2/device/"+$CurrentDevice.id
            Invoke-RestMethod -Uri $URI -Method PATCH -Headers $Headers -ContentType 'application/json' -Body (@{policyId=$CorporateTestGroupPolicyID} | ConvertTo-Json)
            Write-Log "$($CurrentDevice.systemName) has been assigned to the TESTGROUP1 test group policy."
        }
    }
}
#Get all devices assigned to EXAMPLE organization
$AllEXAMPLElDevices = $AllDevices | Where {$_.organizationId -eq "2"}
#Get list of all EXAMPLE locations
$AllEXAMPLELocations = Invoke-RestMethod -Uri 'https://app.ninjarmm.com/v2/organization/2/locations' -Method GET -Headers $headers
#Assign EXAMPLE Device to location if not already assigned
ForEach ($EXAMPLEDevice in $AllEXAMPLEDevices) {
	If ($EXAMPLEDevice.systemName -match "register|manager|apc|rmo|rec|plt") {
		$DeviceStoreCode = $EXAMPLEDevice.systemName.split("-")[1]		
		$EXAMPLELocation = $AllEXAMPLELocations | Where-Object {$_ -like "* - $DeviceStoreCode -*"}
		
		If ($EXAMPLELocation) {
			If ($EXAMPLEDevice.locationId -ne $EXAMPLELocation.id) {
				$URI = "https://app.ninjarmm.com/v2/device/"+$EXAMPLEDevice.id
				Invoke-RestMethod -Uri $URI -Method PATCH -Headers $Headers -ContentType 'application/json' -Body (@{locationId=$EXAMPLELocation.id} | ConvertTo-Json)
				Write-Log "$($EXAMPLEDevice.systemName) has been assigned to the $($EXAMPLELocation.name) location"
			}
		}
		Else {
			Write-Log Error: Location ID Not found for $EXAMPLEDevice.systemName
		}
	}
	Else {
		Write-Log Error: Verify $EXAMPLEDevice.systemName is in the correct organization
	}
}

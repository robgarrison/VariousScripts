Function InstallO365(){
	$CurrentUser = $Null
	$ErrorActionPreference = "Stop"	
	$global:LogMsg = @()
	$MaxAttempts = 3   
  $SleepTime = 5
	$UserSID = $Null
  $InstallPath = 'INSTALLPATHHERE'
 
	
	$sb = {		
		#Grab the user that is currently logged in by checking the owner of the explorer process
		$CurrentUser = ((Get-WmiObject -Class Win32_Process -Filter 'Name="explorer.exe"' | % getowner).User.ToString()).ToLower()
		
		If (!($CurrentUser)) {
			Throw 'Problem obtaining the current user. The CurrentUser variable is empty'
		}
		Else {
			$global:LogMsg += 'The current user is ' + $CurrentUser
		}
		#Get the user's SID that is logged in
		$UserSID = (Get-WmiObject -Class win32_UserAccount -Filter "Domain = 'DOMAINHERE' and Name = '$CurrentUser'").SID

		If (!($UserSID)) {
			Throw 'Problem obtaining User SID. The UserSID Variable is empty'
		}
		If ($UserSID.Count -ne 1) {
			Throw 'Problem obtaining User SID. The UserSID variable returned multiple results'
		}
		Else {
			$global:LogMsg += 'The user SID is ' + $UserSID
		}

		#Create a PSDrive so we can access the users registry
		New-PSDrive HKU Registry HKEY_USERS | Out-Null

		#Create the new profile
		If (!(Test-Path -Path HKU:\$UserSID\SOFTWARE\Microsoft\Office\16.0\Outlook\Profiles\ExchangeOnline)) {
			$global:LogMsg += 'Creating ExchangeOnline profile'
			New-Item -Path HKU:\$UserSID\SOFTWARE\Microsoft\Office\16.0\Outlook\Profiles\ExchangeOnline -Force | Out-Null			
		}

		#Set the new profile as default	
		$global:LogMsg += 'Setting ExchangeOnline profile as default'
		Set-ItemProperty -Path HKU:\$UserSID\SOFTWARE\Microsoft\Office\16.0\Outlook -Name DefaultProfile -Value 'ExchangeOnline' -Type 'String' | Out-Null
		

		#Create autodiscover path
		If (!(Test-Path -Path HKU:\$UserSID\SOFTWARE\Microsoft\Office\16.0\Outlook\AutoDiscover)) {
			$global:LogMsg += 'Creating AutoDiscover path'		
			New-Item -Path HKU:\$UserSID\SOFTWARE\Microsoft\Office\16.0\Outlook\AutoDiscover -Force | Out-Null
		}
		
		#Create the ZCE property
		If (!(Get-ItemProperty -Path HKU:\$UserSID\Software\Microsoft\Office\16.0\Outlook\AutoDiscover -Name 'ZeroConfigExchange' -ErrorAction SilentlyContinue)) {
			$global:LogMsg += 'Setting ZeroConfigExchange property'
			Set-ItemProperty HKU:\$UserSID\Software\Microsoft\Office\16.0\Outlook\AutoDiscover -Name 'ZeroConfigExchange' -Value '1' -Type 'DWORD' | Out-Null
		}

		#Create general path
		If (!(Test-Path -Path HKU:\$UserSID\SOFTWARE\Microsoft\Office\16.0\Outlook\Options\General)) {
			$global:LogMsg += 'Creating General path'
			New-Item -Path HKU:\$UserSID\SOFTWARE\Microsoft\Office\16.0\Outlook\Options\General -Force | Out-Null
		}

		#Disable mobile app setup
		If (!(Get-ItemProperty -Path HKU:\$UserSID\Software\Microsoft\Office\16.0\Outlook\Options\General -Name  'DisableOutlookMobileHyperlink' -ErrorAction SilentlyContinue)) {
			$global:LogMsg += 'Creating DisableOutlookMobileHyperlink property'
			Set-ItemProperty HKU:\$UserSID\Software\Microsoft\Office\16.0\Outlook\Options\General -Name  'DisableOutlookMobileHyperlink' -Value '1' -Type 'DWORD' | Out-Null
		}

		#Close the PS Drive
		Remove-PSDrive HKU
	}	
    Do {
		Try {
			Invoke-Command -ScriptBlock $sb

			If (Test-path $InstallPath) {
				$global:LogMsg += 'Installing O365 Apps'
				$InstallPath\setup.exe /configure $InstallPath\BH-MEC-x64.xml
			}
			Else {
				$global:LogMsg += 'Could not reach O365 installation files'
			}			
			$global:LogMsg += 'Registry changes and App installation have been completed.'
			Write-Log
	        break;
        }
        Catch {
			
			$global:LogMsg += 'Error has occured ' + $MaxAttempts + ' attempts remaining. Trying again in ' + $SleepTime + ' Seconds...'
            $global:LogMsg +=  $_.Exception.Message
			Write-Log
        }
        $MaxAttempts--
        if ($MaxAttempts -gt 0) { 
			Start-Sleep -s $SleepTime 
		}
    } While ($MaxAttempts -gt 0)    
}
Function Write-Log() {
	$TimeStamp = get-date -format 'yyyy-MM-dd'
	$LogFile = 'C:\windows\temp\O365Install-'+$TimeStamp+'.log'

	ForEach ($Msg in $global:LogMsg) {
		$Msg = (get-date -format 'HH:mm ') + $Msg
		Add-content $Logfile -value $Msg
	}
	$global:LogMsg = @()
}

InstallO365

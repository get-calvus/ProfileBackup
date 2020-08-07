$destination = "\\server\ProfileBackups"

$folder = "Desktop",
"Documents",
"Downloads",
"Favorites",
"Music",
"Pictures",
"Videos",
"AppData\Local\Mozilla",
"AppData\Roaming\Mozilla",
"AppData\Local\Google\Chrome",
"AppData\Local\Microsoft\Edge",
"AppData\Local\Microsoft\Outlook\*.pst",
"AppData\Roaming\Microsoft\Signatures",
"AppData\Roaming\Microsoft\UProof",
"AppData\Roaming\Microsoft\Sticky Notes"



###############################################################################################################

$username = Read-Host -Prompt 'Enter the user profile to backup'
$userprofile = "c:\Users\$username"
$appData = "C:\Users\$username\AppData\Local"
# $userprofile = gc env:userprofile
# $appData = gc env:localAPPDATA



###### Restore data section ######
if ([IO.Directory]::Exists($destination + "\" + $username + "\")) 
{ 

	$caption = "Choose Action";
	$message = "A backup folder for $username already exists, would you like to restore the data to the local machine?";
	$Yes = new-Object System.Management.Automation.Host.ChoiceDescription "&Yes","Yes";
	$No = new-Object System.Management.Automation.Host.ChoiceDescription "&No","No";
	$choices = [System.Management.Automation.Host.ChoiceDescription[]]($Yes,$No);
	$answer = $host.ui.PromptForChoice($caption,$message,$choices,0)

	if ($answer -eq 0) 
	{
		
		write-host -ForegroundColor green "Restoring data to local machine for $username"
		foreach ($f in $folder)
		{	
			$currentLocalFolder = $userprofile + "\" + $f
			$currentRemoteFolder = $destination + "\" + $username + "\" + $f
			write-host -ForegroundColor cyan "  $f..."
			Copy-Item -ErrorAction silentlyContinue -recurse $currentRemoteFolder $userprofile
						
			if ($f -eq "AppData\Local\Mozilla") { rename-item $currentLocalFolder "$currentLocalFolder.old" }
			if ($f -eq "AppData\Roaming\Mozilla") { rename-item $currentLocalFolder "$currentLocalFolder.old" }
			if ($f -eq "AppData\Local\Google\Chrome") { rename-item $currentLocalFolder "$currentLocalFolder.old" }
			if ($f -eq "AppData\Local\Microsoft\Edge") { rename-item $currentLocalFolder "$currentLocalFolder.old" }
			if ($f -eq "AppData\Roaming\Microsoft\Signatures") { rename-item $currentLocalFolder "$currentLocalFolder.old" }
			if ($f -eq "AppData\Roaming\Microsoft\UProof") { rename-item $currentLocalFolder "$currentLocalFolder.old" }
			if ($f -eq "AppData\Roaming\Microsoft\Sticky Notes") { rename-item $currentLocalFolder "$currentLocalFolder.old" }
		
		}
		
		
		rename-item "$destination\$username" "$destination\$username.restored" -ErrorAction SilentlyContinue
		Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | sort DisplayName | Select-Object DisplayName, DisplayVersion, InstallSource | Export-Csv -Path $userprofile\Desktop\NewInstalledPrograms-RegistryExport.csv -NoTypeInformation
		get-ciminstance win32_product | sort Name | Select-Object Name,Caption,Version | Export-Csv -Path $userprofile\Desktop\NewInstalledPrograms-WMIExport.csv -NoTypeInformation
		
		Compare-Object $userprofile\Desktop\OldInstalledPrograms-RegistryExport.csv $userprofile\Desktop\NewInstalledPrograms-RegistryExport.csv -property DisplayName | Export-Csv -Path $userprofile\Desktop\InstallThesePrograms1.csv -NoTypeInformation
		Compare-Object $userprofile\Desktop\OldInstalledPrograms-WMIExport.csv $userprofile\Desktop\NewInstalledPrograms-WMIExport.csv -property Name | Export-Csv -Path $userprofile\Desktop\InstallThesePrograms2.csv -NoTypeInformation
		
		Remove-Item -Path $userprofile\Desktop\OldInstalledPrograms-RegistryExport.csv
		Remove-Item -Path $userprofile\Desktop\NewInstalledPrograms-RegistryExport.csv
		Remove-Item -Path $userprofile\Desktop\OldInstalledPrograms-WMIExport.csv
		Remove-Item -Path $userprofile\Desktop\NewInstalledPrograms-WMIExport.csv
		write-host -ForegroundColor green "Restore Complete!"
	}
	
	else
	{
		write-host -ForegroundColor yellow "Aborting backup"
		exit
	}
	
	
}

###### Backup Data section ########
else 
{ 
		
	Write-Host -ForegroundColor red "Outlook is about to close! `nSave any unsaved emails! `nPress any key to continue ..."

	$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
	
	Get-Process | Where { $_.Name -Eq "OUTLOOK" } | Kill

	write-host -ForegroundColor green "Generating installed printer list for $username"
	# Determine OS version
	# Win7 - 6.1.7601
	# Win10 LTSB - 10.0.14393
	# Win10 LTSC - 10.0.17763
	$name=(Get-CimInstance Win32_OperatingSystem).Caption
	$version=(Get-CimInstance Win32_OperatingSystem).Version
	
	# Now determine which method to use for printer backup
	If ($version -eq "6.1.7601") 
	 {Write-host -ForegroundColor green OS is $name and build is $version
	 get-WmiObject -class Win32_printer | ft name, location, systemName, shareName >> $userprofile\Desktop\InstalledPrinters.txt} 
	Else 
	 {Write-host -ForegroundColor green OS is $name and build is $version
	 Get-Printer | Select-Object Name, Type, DriverName, PortName | Export-Csv -Path $userprofile\Desktop\InstalledPrinters.csv -NoTypeInformation}  
	
	
	write-host -ForegroundColor green "Generating installed programs list from local machine"
	$Arch=(Get-WmiObject -Class Win32_operatingsystem).Osarchitecture
	If ($Arch -eq "32-bit") 
		{Get-ItemProperty HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\* | Select-Object DisplayName, DisplayVersion, InstallSource | Export-Csv -Path $userprofile\Desktop\OldInstalledPrograms-RegistryExport.csv -NoTypeInformation
		get-ciminstance win32_product | sort Name | Select-Object Name,Caption,Version | Export-Csv -Path $userprofile\Desktop\OldInstalledPrograms-WMIExport.csv -NoTypeInformation}
	Else 
		{Get-ItemProperty HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\* | Select-Object DisplayName, DisplayVersion, InstallSource | Export-Csv -Path $userprofile\Desktop\OldInstalledPrograms-RegistryExport.csv -NoTypeInformation
		get-ciminstance win32_product | sort Name | Select-Object Name,Caption,Version | Export-Csv -Path $userprofile\Desktop\OldInstalledPrograms-WMIExport.csv -NoTypeInformation}
	
	
	write-host -ForegroundColor green "Backing up data from local machine for $username"
	foreach ($f in $folder)
	{	
		$currentLocalFolder = $userprofile + "\" + $f
		$currentRemoteFolder = $destination + "\" + $username + "\" + $f
		$currentFolderSize = (Get-ChildItem -ErrorAction silentlyContinue $currentLocalFolder -Recurse -Force | Measure-Object -ErrorAction silentlyContinue -Property Length -Sum ).Sum / 1MB
		$currentFolderSizeRounded = [System.Math]::Round($currentFolderSize)
		write-host -ForegroundColor cyan "  $f... ($currentFolderSizeRounded MB)"
		Copy-Item -ErrorAction silentlyContinue -recurse $currentLocalFolder $currentRemoteFolder
	}
	
	
	
	$oldStylePST = [IO.Directory]::GetFiles($appData + "\Microsoft\Outlook", "*.pst") 
	foreach($pst in $oldStylePST)	
	{ 
		if ((test-path -path ($destination + "\" + $username + "\Documents\Outlook Files\oldstyle")) -eq 0){new-item -type directory -path ($destination + "\" + $username + "\Documents\Outlook Files\oldstyle") | out-null}
		write-host -ForegroundColor yellow "  $pst..."
		Copy-Item $pst ($destination + "\" + $username + "\Documents\Outlook Files\oldstyle")
	}    
	
	write-host -ForegroundColor green "Backup complete!"
	
} 

Write-Host "Press any key to continue..."
$x = $host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")

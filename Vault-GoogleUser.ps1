<#	
	.NOTES
	===========================================================================
	 Created with: 	SAPIEN Technologies, Inc., PowerShell Studio 2020 v5.7.173
	 Created on:   	10/10/20 11:10 AM
	 Created by:   	jsavage
	 Organization: 	
	 Filename:     	Vault-GoogleUser.ps1
	===========================================================================
	.SYNOPSIS	
		Archives a previous employee's GSuite data. 
		
	.DESCRIPTION
		Completes the vaulting/archiving process using GAM by creating a matter, exporting the content, downloading to a share, and deleting the users.

	.PARAMETER $GAMCSVFilePath
		The file path of the csv file containing the users to be vaulted.

	.PARAMETER $ExportDataLocation
		The save location of the users who are being vaulted. 

	.NOTES
		- The save location must be within <NETWORK SHARE> for security purposes.

		- This function uses GAM CSV bulk operations to create each matter (More info here: https://github.com/jay0lee/GAM/wiki/BulkOperations).

		- GAM needs to be located in the users' folder running the script.

	.EXAMPLE
		This example shows what might be entered as a csv file and save location.
			
		PS C:\> Vault-GoogleUser -GAMCSVFilePath C:\Users\<USERNAME>\gam\users09.12.2020.csv -ExportDataLocation <NETWORK SHARE>

		
#>
Function Vault-GoogleUser
{
	
	param (
		
		[parameter(Mandatory, Position = 0)]
		[ValidateScript({
				
				#Check if it actually exists
				if (-Not ($_ | Test-Path))
				{
					throw "Could not find File or Folder."
				}
				
				#Check that it is a file not a folder
				if (-Not ($_ | Test-Path -PathType Leaf))
				{
					throw "The Path argument must be a file. Folder paths are not allowed."
				}
				
				#Check that the file is a .csv file.
				if ($_ -notmatch "(\.csv)")
				{
					throw "The file specified must be a csv file."
				}
				
				else { return $true }
				
			})]
		
		[System.IO.FileInfo]$GAMCSVFilePath,
		
		[parameter(Mandatory, Position = 1)]
		[ValidateScript({
				
				
				#Check that it is a file not a folder
				if ($_ | Test-Path -PathType Leaf)
				{
					throw "The Path argument must be a Folder, not a file."
				}
				
				if (-Not ($_ | Test-Path))
				{
					Write-Error $_
					throw "Cannot find path specified."
				}
				else
				{
					return $true
				}

			})]
		[System.String]$ExportDataLocation,

		[parameter(Mandatory, Position = 2)]
		[System.String]logginglocation,

		
		[parameter(Mandatory, Position = 3)]
		[System.String]$ErrorLogFolder,
				
		
		[parameter(Mandatory = $false, Position = 4)]
		[int32]$GAM_THREADS = 1,
	)
	
	
		#Ensure this is running from a VM before doing anything, otherwise stop the script.
		if (!((Get-ComputerInfo).BiosManufacturer -match "VMware"))
		{
			Throw "Error: Vault-GoogleUser script not running from a Virtual Machine."
		}
	

		Start-Transcript -Path $logginglocation -Append
		
		#set GAM location
		$GAMLocation = "C:\users\$env:username\gam\"
	
	
		#Set error logs location.
		$ErrorLogLocation = $ErrorLogFolder+"\errorlogs.txt"
	
		#Make sure it exists.
		if (Test-Path $ErrorLogLocation)
		{

			#set Log function just for errors.
			function write-log ($status, $msg)
			{
				
				$ErrorLogLocation = Get-ChildItem $ErrorLogLocation
				
				$date = Get-Date
				$date = $date.Tostring('yyyy-MM-dd-HH:mm')
				
				$newlog += "`n" + $date + " " + $status + $msg
			
				$newlog | Out-file $ErrorLogLocation -Append 
				
			}

		}
	
		else
		{
			throw "Could not find $ErrorLogLocation. Ensure this exists before proceeding"
		}
	
	
		#verify GAM is available before proceeding
		if ((Test-Path -Path $GAMLocation) -and (Get-ChildItem -Path $GAMLocation -Filter "gam.exe"))
		{
			#set GAM thread amount.
			set GAM_THREADS=$GAM_THREADS
			
			Set-Location $GAMLocation
			
			#Import the CSV file so we can work with it
			$GAMCSVFile = Import-Csv -Path $GAMCSVFilePath
			
				#region Create Matters
			
			.\gam.exe csv $GAMCSVFilePath gam create vaultmatter name ~name
			
			#Export mail file
			.\gam.exe csv $GAMCSVFilePath gam create export matter ~name name ~mailname corpus mail accounts ~Email
			
			#Export hangouts file
			.\gam.exe csv $GAMCSVFilePath gam create export matter ~name name ~hangoutsname corpus hangouts_chat accounts ~Email
			
			#Export Drive file
			.\gam.exe csv $GAMCSVFilePath gam create export matter ~name name ~drivename corpus drive accounts ~Email
			
			#endregion Create Matters
			
			foreach ($user in $GAMCSVFile)
			{
				
				#region Create Export Folders
				
				#Create a folder for each corresponding export to go into.
				
				try
				{
					$mailfolder = New-Item -Path $ExportDataLocation -Name $user.mailname -ItemType Directory -ErrorAction Stop
					Write-Host "Creating Mail folder for $($user.email) in $ExportDataLocation" -ForegroundColor Yellow
				}
				
				Catch
				{
					write-log ("Error:", "$_")
					$_ | Tee-Object -FilePath $ExportDataLocation\$($user.name)_errorlogs.txt -Append
				}
				
				try
				{
					$hangoutsfolder = New-Item -Path $ExportDataLocation -Name $user.hangoutsname -ItemType Directory -ErrorAction Stop
					Write-Host "Creating Hangouts folder for $($user.email) in $ExportDataLocation" -ForegroundColor Yellow
				}
				
				Catch
				{
					write-log ("Error:", "$_")
					$_ | Tee-Object -FilePath $ExportDataLocation\$($user.name)_errorlogs.txt -Append
				}
				
				try
				{
					$drivefolder = New-Item -Path $ExportDataLocation -Name $user.drivename -ItemType Directory -ErrorAction Stop
					Write-Host "Creating Drive folder for $($user.email) in $ExportDataLocation" -ForegroundColor Yellow
				}
				
				Catch
				{
					write-log ("Error:", "$_")
					$_ | Tee-Object -FilePath $ExportDataLocation\$($user.name)_errorlogs.txt -Append
				}
			#endregion Create Export Folders
				
				
				#region Download Vault data
				
				#Check that the export status is "status: COMPLETED" before attempting export.
				#Begin with Hangouts, then Drive, as this gives time for Mail to finish which normally takes the longest.
				
				do
				{
					
					$hangouts_status = .\gam.exe info export $user.name $user.hangoutsname
					if ($hangoutsflag = $hangouts_status -match "status: COMPLETED")
					{
						write-host "Hangouts export done" -ForegroundColor Yellow
					}
					
					else
					{
						Write-Host "Hangouts export not yet complete... Checking again in 60 seconds."
						Start-Sleep -Seconds 60
						
					}
				}
				Until ($hangoutsflag)
				
				
				
				Write-Host "Downloading to $($hangoutsfolder.fullname)" -ForegroundColor Yellow
				
				#now that we've verified the export is complete, let's download it to the share location.
				.\gam.exe download export $user.name $user.hangoutsname noextract targetfolder $hangoutsfolder.FullName
				
				
				
				
				
				#Repeat the same for DRIVE data
				do
				{
					
					$Drive_status = .\gam.exe info export $user.name $user.drivename
					if ($driveflag = $Drive_status -match "status: COMPLETED")
					{
						write-host "Drive export done" -ForegroundColor Yellow
					}
					
					else
					{
						Write-Host "Drive export not yet complete... Checking again in 60 seconds."
						Start-Sleep -Seconds 60
					}
				}
				Until ($driveflag)
				
				
				Write-Host "Downloading to $($drivefolder.fullname)" -ForegroundColor Yellow
				
				#now that we've verified the export is complete, let's download it to the share location.
				.\gam.exe download export $user.name $user.drivename noextract targetfolder $drivefolder.FullName
				
				
				
				
				
				#Repeat the same for MAIL data
				
				do
				{
					
					$mail_status = .\gam.exe info export $user.name $user.mailname
					if ($mailflag = $mail_status -match "status: COMPLETED")
					{
						
						write-host "Mail Export Done" -ForegroundColor Yellow
					}
					else
					{
						Write-Host "Mail export not yet complete... Checking again in 60 seconds."
						Start-Sleep -Seconds 60
					}
				}
				Until ($mailflag)
				
				
				
				Write-Host "Downloading to $($mailfolder.fullname)" -ForegroundColor Yellow
				
				#now that we've verified the export is complete, let's download it to the share location.
				.\gam.exe download export $user.name $user.mailname noextract targetfolder $mailfolder.FullName
				
				
				
				Write-Host "All of $($user.email)'s data has been successfully vaulted and content moved to $ExportDataLocation" -ForegroundColor Yellow
				
				
				#endregion Download Vault data
				
				
				#region Delete Gsuite account
				
				#THIS SECTION COMAPARES ALL ZIP FILES IN THE VAULT AND NETWORK SHARE TO ENSURE THEY ARE EXACT IN NAME AND AMOUNT BEFORE DELETING.
			
			
				#region Compare Hangouts zip Files
				
				 
			
				#Get information regarding the export
				$hangouts_content = .\gam.exe info export $user.name $user.hangoutsname
				
				
				#Grab the line of string that contains the .zip files from $hangouts_content
				$object1 = $hangouts_content | Select-String -Pattern "objectName:"
				
				
				#Filter out ONLY the zip files, and edit it to look exactly as it would in the share location.
				$Vault_Hangouts_files = $object1 | where { $_ -match ".zip" } | ForEach-Object { ($_ -replace ("objectName: ", "") -replace ("/", "-")).trim(" ") }
				
				
				Write-Host "`nHere are the zip files for Hangouts in the Vault:" -ForegroundColor Green
				$Vault_Hangouts_files
			
				try
				{
					#Grab the equivalent zip files from the share location.
					$Hangouts_Share_zip = Get-ChildItem $hangoutsfolder.FullName -Filter "*.zip" -ErrorAction Stop
					
					Write-Host "`nHere are the zip files in $($hangoutsfolder.FullName):" -ForegroundColor Green
					$Hangouts_Share_zip.name
				}
			
				catch
				{
				  	Write-log ("Error:", "$_")
					$_ | Tee-Object -FilePath $ExportDataLocation\$($user.name)_errorlogs.txt -Append
				}
			
			
				try
				{
					#compare the two objects and their differences
					$Hangouts_comparison = Compare-Object -ReferenceObject $Vault_Hangouts_files -DifferenceObject $Hangouts_Share_zip.name -IncludeEqual -ErrorAction Stop
					
					
					#if there are any side indicators, then they are not exact, and therefore DO NOT DELETE
					if (($Hangouts_comparison.sideindicator -contains "<=") -or ($Hangouts_comparison.sideindicator -contains "=>"))
					{
						Write-Output "Hangouts zip files do not match in both locations for $($user.Email)" | Tee-Object -FilePath $ExportDataLocation\$($user.name)_errorlogs.txt -Append
						write-log ("Error", "Hangouts zip files do not match in both locations for $($user.Email)")
						
					}
					
					else
					{
						Write-host "`nZIP FILES MATCH in both locations." -ForegroundColor Yellow
						$Hangouts_files_match = $true
					}
				}
				
				Catch
				{
					write-log ("Error:", "$_")
					$_ | Tee-Object -FilePath $ExportDataLocation\$($user.name)_errorlogs.txt -Append
				}
			
				#endregion
			
			
				#region Compare Drive zip files
				
				
				
				
				#Get information regarding the export
				$Drive_content = .\gam.exe info export $user.name $user.drivename
				
				
				#Grab the line of string that contains the .zip files from $Drive_content
				$object2 = $drive_content | Select-String -Pattern "objectName:"
				
				
				#Filter out ONLY the zip files, and edit it to look exactly as it would in the share location.
				$Vault_Drive_files = $object2 | where { $_ -match ".zip" } | ForEach-Object { ($_ -replace ("objectName: ", "") -replace ("/", "-")).trim(" ") }
				
				
				Write-Host "`nHere are the zip files for Drive in the Vault:" -ForegroundColor Green
				$Vault_Drive_files
			
				try
				{
					#Grab the equivalent zip files from the share location.
					$Drive_Share_zip = Get-ChildItem $drivefolder.FullName -Filter "*.zip" -ErrorAction Stop
					
					Write-Host "`nHere are the zip files in $($drivefolder.FullName):" -ForegroundColor Green
					$Drive_Share_zip.name
				}
			
			
				catch
				{
					Write-log ("Error:", "$_")
					$_ | Tee-Object -FilePath $ExportDataLocation\$($user.name)_errorlogs.txt -Append
				}
			
			
				try
				{
					#compare the two objects and their differences
					$Drive_comparison = Compare-Object -ReferenceObject $Vault_Drive_files -DifferenceObject $Drive_Share_zip.name -IncludeEqual -ErrorAction Stop
					
					
					#if there are any side indicators, then they are not exact, and therefore DO NOT DELETE
					if (($Drive_comparison.sideindicator -contains "<=") -or ($Drive_comparison.sideindicator -contains "=>"))
					{
						Write-Output "Drive zip files do not match in both locations for $($user.Email)" | Tee-Object -FilePath $ExportDataLocation\$($user.name)_errorlogs.txt -Append
						write-log ("Error", "Drive zip files do not match in both locations for $($user.Email)")
						
					}
					
					else
					{
						Write-Host "`nZIP FILES MATCH in both locations." -ForegroundColor Yellow
						$Drive_files_match = $true
					}
				}
				catch
				{
					write-log ("Error:", "$_")
					$_ | Tee-Object -FilePath $ExportDataLocation\$($user.name)_errorlogs.txt -Append
				}
			
				#endregion
			
			
				#region Compare Mail zip files
				
				
				
				#Get information regarding the export
				$Mail_content = .\gam.exe info export $user.name $user.mailname
				
				
				#Grab the line of string that contains the .zip files from $Mail_content
				$object3 = $Mail_content | Select-String -Pattern "objectName:"
				
				
				#Filter out ONLY the zip files, and edit it to look exactly as it would in the share location.
				$Vault_Mail_files = $object3 | where { $_ -match ".zip" } | ForEach-Object { ($_ -replace ("objectName: ", "") -replace ("/", "-")).trim(" ") }
				
				
				Write-Host "`nHere are the zip files for Mail in the Vault:" -ForegroundColor Green
				$Vault_Mail_files
			
				try
				{
					#Grab the equivalent zip files from the share location.
					$Mail_Share_zip = Get-ChildItem $mailfolder.FullName -Filter "*.zip" -ErrorAction Stop
					
					Write-Host "`nHere are the zip files in $($mailfolder.FullName):" -ForegroundColor Green
					$Mail_Share_zip.name
				}
			
				catch
				{
					Write-log ("Error:", "$_")
					$_ | Tee-Object -FilePath $ExportDataLocation\$($user.name)_errorlogs.txt -Append
				}
				
			try
				{
					#compare the two objects and there differences
					$Mail_comparison = Compare-Object -ReferenceObject $Vault_Mail_files -DifferenceObject $Mail_Share_zip.name -IncludeEqual -ErrorAction Stop
					
					
					#if there are any side indicators, then they are not exact, and therefore DO NOT DELETE
					if (($Mail_comparison.sideindicator -contains "<=") -or ($Mail_comparison.sideindicator -contains "=>"))
					{
						Write-Output "Mail zip files do not match in both locations for $($user.Email)" | Tee-Object -FilePath $ExportDataLocation\$($user.name)_errorlogs.txt -Append
						write-log ("Error", "Mail zip files do not match in both locations for $($user.Email)")
					}
					
					else
					{
						Write-Host "`nZIP FILES MATCH in both locations." -ForegroundColor Yellow
						$Mail_files_match = $true
					}
				}
				
				Catch
				{
					write-log ("Error:", "$_")
					$_ | Tee-Object -FilePath $ExportDataLocation\$($user.name)_errorlogs.txt -Append
				}
			
				#endregion
				
				
				#If all three export locations contain the same data as shown in the vault, Delete the user.
				
				if (($Mail_files_match) -and ($Drive_files_match) -and ($Hangouts_files_match))
				{
					Write-host "`nDeleting $($user.email) from Google..." -ForegroundColor Yellow
					
					.\gam.exe delete user $user.Email
					
					Write-Host "`n$($user.email) has been vaulted and deleted successfully" -ForegroundColor Green
				}
				
				else
				{
					Write-Output "`nCould not delete $($user.email) in Google. Not all data has been vaulted. Please check this before deleting manually." | Tee-Object -FilePath $ExportDataLocation\$($user.name)_errorlogs.txt -Append
					write-log ("Error", "Could not delete $($user.email) in Google. Not all data has been vaulted. Please check this before deleting manually.")
				}
				
				#endregion Delete Gsuite account
				
			} #end first Foreach loop
			
		} #end IF for test GAM path
		
		else
		{
			Write-Error "Could not locate gam.exe in $GAMLocation. Please ensure GAM is configured here before proceeding"
			write-log ("Error:","Could not locate gam.exe in $GAMLocation. Please ensure GAM is configured here before proceeding")
		}
	
		Stop-Transcript
		
}
	
	
<#	
	.NOTES
	===========================================================================
	 Created with: 	SAPIEN Technologies, Inc., PowerShell Studio 2019 v5.6.167
	 Created on:   	9/09/2020 9:50 AM
	 Created by:   	jsavage
	 Organization: 	
	 Filename:     	Add-GoogleGroupMembers.ps1
	===========================================================================
	.DESCRIPTION
		Takes a CSV file and adds all users listed in the "email" row to to the specified Google group.

	.NOTES
		Uses GAM's native CSV function to iterate over multiple objects (more information on this can be found here: https://gamcheatsheet.com/GAM%20Cheat%20Sheet%20A3.pdf)

		The first row in the CSV file needs to be labelled as "email" (without quotes), as this is what the GAM command is looking for in the CSV file. 

	.EXAMPLE
		Add-GoogleGroupMembers -EmailGroup TestGroup@domain1.com.au -$FullCSVFilePath C:\support\test.csv

		This example adds all members listed in C:\Support\test.csv into the 'TestGroup@domain1.com.au' Google group.
#>



function Add-GoogleGroupMembers
{
	param (
		
		[CmdletBinding()]
		
		[parameter(Mandatory, Position = 0)]
		[ValidateScript({
				#Validate the Google Group exists. If true, then proceed.
				if (C:\gam\gam.exe info group $_)
				{ return $true }
				else
				{
					throw "Error finding '$_' in GSuite"
				}
				
			})]
		[mailaddress]$EmailGroup,
		
		
		[parameter(Mandatory, Position = 1)]
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
				
				#Check if the column name is 'email' in the csv file.
				$columnname = Import-Csv -Path $_ | Get-Member -MemberType NoteProperty | select Name
				if (($columnname).Name -eq "email")
					{
						return $true 
					}
				else
					{
						throw "The column in the excel sheet listed as '$($columnname.name)' needs to be 'email'"
					}
				
			})]
		[System.IO.FileInfo]
		$FullCSVFilePath.

		[parameter(Mandatory, Position = 2)]
		$loggingLocation,

		[parameter(Mandatory = $false, Position = 3)]
		$GAMLocation = "C:\gam\gam.exe",
	)
	
	#Set a log location to record everytime the script is run
	$logginglocation = $loggingLocation+"\Add-GoogleGroupMembersLOG.txt"
	
	#Check logging location first before starting transcript.
	if ($logpath = Test-Path $logginglocation -ErrorAction SilentlyContinue)
		{
			#Keep transcript silent
			Start-Transcript -Path $logginglocation -Append | Out-Null
		}
	
	else
		{
			#Knowledge of transcript is only visible when location is not available.
			Write-Output "Transcript not logging to $logginglocation"
		}
	
	
	#Ensure GAM is available
	if (Test-Path -Path $GAMLocation)
		{
			C:\gam\gam.exe csv $FullCSVFilePath gam update group $EmailGroup add ~email
		}
	
	else
		{
			Write-Error "Could not locate GAM on the local machine."
		}
	
	#Only attempt to stop the transcript if the path is valid and transcript starts.
	if ($logpath)
	{
		Stop-Transcript -ErrorAction Ignore | Out-Null
	}
	
}

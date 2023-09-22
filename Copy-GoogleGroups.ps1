
<#  
    .NOTES

    ===========================================================================

     Created with:  SAPIEN Technologies, Inc., PowerShell Studio 2019 v5.6.167

     Created on:    3/08/2020 8:48 AM

     Created by:    jsavage

     Organization:  

     Filename: Copy-GoogleGroups.ps1

    ===========================================================================

    .DESCRIPTION
        Mirrors the Google group membership for one account to another
        
#>

function Copy-GoogleGroups
{
    param
    (
        [parameter(Mandatory = $true)]
        [ValidateNotNullorEmpty()]
        [mailaddress]$SourceEmail,

        [parameter(Mandatory = $true)]
        [ValidateNotNullorEmpty()]
        [mailaddress]$DestinationEmail
    )

    #Ensure both email addresses exist, if not bail out

    if ((Get-GSUser -User $SourceEmail) -and (Get-GSuser -User $DestinationEmail))
    {
        #retrieve the google groups of the person to be copied
        $EquivUserGroups = Get-GSGroup -Where_IsAMember $SourceEmail

        #Ensure there are actually groups to add
        if (!(($EquivUserGroups).count -eq 0))
        {
            try
            {
                foreach ($Group in $EquivUserGroups.Email)
                {
                    Add-GSGroupMember $Group -Member $DestinationEmail | out-null
                    write-output "Successfully added $DestinationEmail to $Group"
                }
            }
            Catch
            {
                Write-Error "Failed to add $DestinationEmail to $group"
            }
        } #end second if    

        else
        {
            Write-output "Could not add $DestinationEmail to any Google groups, as $SourceEmail is not a part of any groups."
        }
    } #end first if

    else
    {
        Write-Error "Could not find either $SourceEmail or $DestinationEmail in GSuite"
    }
} 




##############################################################
# PowerShell get-MsolUserByLicenseSKU Ver 0.1 02/11/2016     #
# Powershell Version 4.0                                     #
# If this script works it was written by Nathan Kewley       #
# If this script does not work is was written by Michael Cho #
#                                                            #
# This cmdlet will return a list of users that have a given  #
# license SKU assigned in MSOL.                              #
##############################################################

#gets users by license SKUID
function global:get-MsolUserByLicenseSku {
    <#
        .SYNOPSIS
            This cmdlet will query MSOl for users who has a specified license assigned
        .DESCRIPTION
            This cmdlet will query MSOl for users who has a specified license assigned. Ver 0.1 02/11/2016 Created by Nathan Kewley
        .EXAMPLE
            get-MsolUserByLicenseSku -SKU penrithcitycouncil:VISIOCLIENT
        .EXAMPLE
            get-MsolUserByLicenseSku -SKU penrithcitycouncil:ENTERPRISEPACK
        .EXAMPLE
            get-MsolUserByLicenseSku -SKU penrithcitycouncil:PROJECTCLIENT
        .LINK
            nathankewley.info
    #>


    #setup cmdlet binding
    [CmdletBinding(
        DefaultParameterSetName=”SKU”
    )]

    #set up cmdlet parameters
    PARAM(
        [parameter(
            HelpMessage='Enter the license SKU that you want to query MSOL for license allocation.', 
            Mandatory=$True, 
            Position=0, 
            ValueFromPipeline=$True
        )]
        [STRING[]]$SKU
    )
    
    #only query msol for users once
    Begin{
        #get users
        $users = Get-MsolUser -All
    }

    #run for every value passed through the pipeline
    Process{
        #create return object so this cmdlet is pipeline friendly
        $userUPN = @()

        #check user licenses
        foreach($user in $users){
            $licensedetails = $user.Licenses
            if ($licensedetails.Count -gt 0){
                foreach ($i in $licensedetails){
                    if($i.AccountSkuId -eq $SKU){
                        #add the user to the return object if they have the license in question
                        $tmp = New-Object -TypeName PSObject
                        Add-Member -InputObject $tmp -MemberType NoteProperty -Name UserPrincipalName -Value $user.UserPrincipalName
                        $userUPN += $tmp
                    }
                }
            }
        }
        return $userUPN
    }

    #only run once
    End{}
}
#get devices in a device collection
function get-CMDeviceMembership {
    #define the parameters that the cmdlet will take
    param(
        [parameter(HelpMessage='Enter Name Of Colleciton to Query')]
        [STRING]$Name
    )

    #Import SCCM and Connect To Site
    import-module "C:\Program Files (x86)\Microsoft Configuration Manager\AdminConsole\bin\ConfigurationManager.psd1"
    CD PR1:

    #query SCCM for all devices in the specified collection
    $members = Get-CMDevice -CollectionName $Name | select name

    #return the members in the collection
    return $members
}get
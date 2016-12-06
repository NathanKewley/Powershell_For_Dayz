##############################################################
# PowerShell NUC start IE Script v1.01             19/01/2016#
# Powershell Version 4.0                                     #
# If this script works it was written by Nathan Kewley       #
# If this script does not work is was written by Michael Cho #
#                                                            #
# This script is called from a .bat file at login. When run  #
# this script will create an instance of Internet Explorer   #
# which it will then make fullscreen.                        #
#                                                            #
# The script then navigates Internet Explorer to the         #
# specified webpage (in this case the zecom test pages).     #
#                                                            #
# v1.01                                                      #  
#  - checks hostname to determine page to load so same script#
#    can be used on every standard setup NUC                 #
############################################################## 

#start Internet Explorer
$ie = new-object -com "InternetExplorer.Application";
$ie.visible = $true;

#wether default function should be run
$default = $true;

#get hostname of the device the script is running on
$hostname = $env:COMPUTERNAME;

#Wait till Internet Exploere is running
while ($ie.busy) {sleep -milliseconds 500;}

#Set Internet Explorer to fullscreen
$ie.fullscreen = $true;

#Wait till Internet Exploere is running
while ($ie.busy) {sleep -milliseconds 500;}

#use a switch statment to determine which webpage should be loaded
switch($hostname){
    #find the nuc and set the webpage that should be loaded
    #NUC-1516-001{$webpage = "http://sr-vm-zcc-01/snapshot/Default.aspx?template=WasteTest%20-%20Glenn.xml";}       #NOT SET UP
    #NUC-1516-002{$webpage = "http://sr-vm-zcc-01/snapshot/Default.aspx?template=WasteTest%20-%20Glenn.xml";}       #NOT SET UP
    #NUC-1516-003{$webpage = "http://sr-vm-zcc-01/snapshot/Default.aspx?template=WasteTest%20-%20Glenn.xml";}       #NOT SET UP
    #NUC-1516-004{WUGLogin $ie; $default = $false;}                                                                      #ICT - Right TV
    NUC-1516-005{$webpage = "http://sr-vm-zcc-01/snapshot/Default.aspx?template=ICTTest%20-%20Glenn.xml";}          #ICT - center TV
    NUC-1516-006{$webpage = "http://sr-vm-zcc-01/snapshot/Default.aspx?template=ICT%20Service%20Desk.xml";}         #ICT - Left TV
    NUC-1516-007{$webpage = "http://sr-vm-zcc-01/snapshot/Default.aspx?template=Operator AA.xml";}                     #da operator
    #NUC-1516-008{$webpage = "http://sr-vm-zcc-01/snapshot/Default.aspx?template=WasteTest%20-%20Glenn.xml";}       #NOT SET UP
    #NUC-1516-009{$webpage = "http://sr-vm-zcc-01/snapshot/Default.aspx?template=WasteTest%20-%20Glenn.xml";}       #NOT SET UP
    NUC-1516-010{$webpage = "http://sr-vm-zcc-01/snapshot/Default.aspx?template=DSTest%20-%20Glenn.xml";}           #DEVELOPMENT SERVICES
    #NUC-1516-011{$webpage = "http://sr-vm-zcc-01/snapshot/Default.aspx?template=WasteTest%20-%20Glenn.xml";}       #NOT SET UP
    NUC-1516-012{$webpage = "http://sr-vm-zcc-01/snapshot/Default.aspx?template=RatesTest%20-%20Glenn.xml";}        #RATES
    NUC-1516-013{$webpage = "http://sr-vm-zcc-01/snapshot/Default.aspx?template=CSTEST%20-%20Glenn.xml";}           #CHILDREN SERVICES
    NUC-1516-014{$webpage = "http://sr-vm-zcc-01/snapshot/Default.aspx?template=DSTest%20-%20Glenn.xml";}           #Development Services... slideshow
    NUC-1516-015{$webpage = "http://sr-vm-zcc-01/snapshot/Default.aspx?template=WasteRangersNFTest%20-%20Glenn.xml";} #WASTE
    NUC-1516-016{$webpage = "http://sr-vm-zcc-01/snapshot/Default.aspx?template=CSTEST%20-%20Glenn.xml";}           #children services (office)
    
    #MY Sufrace for testing
    #TB-CC-1415-001{WUGLogin; $default = $false;}      #MY PC FOR TESTING :)
    #TB-CC-1415-001{WUGLogin $ie; $default = $false;}      #MY PC FOR TESTING :)
    TB-CC-1415-001{$webpage = "http://sr-vm-zcc-01/snapshot/Default.aspx?template=Operator AA.xml";}
    
    #if the nuc/device is not recognised display the "All" page by default
    default {$webpage = "http://sr-vm-zcc-01/snapshot/Default.aspx?template=All%20Queues.xml";}
}

#Navigate Internet Explorer to the webpage required for the NUC
if($default -eq $true){
    $ie.navigate($webpage);
    sleep -Seconds 5;
    $ie.navigate($webpage); #just in case....
}
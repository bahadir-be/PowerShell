Set objNetwork = CreateObject("WScript.Network")
objNetwork.MapNetworkDrive "M:" , "\\Ad\netlogon"
objNetwork.MapNetworkDrive "N:" ,"\\Ad\sysvol"
objNetwork.MapNetworkDrive "o:" ,"\\Sccm\paylasmacilar"
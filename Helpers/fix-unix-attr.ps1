import-module ActiveDirectory
$Username = "praiyani"

#UNIX ATTRIBUTES
$NIS = Get-ADObject "CN=int,CN=ypservers,CN=ypServ30,CN=RpcServices,CN=System,DC=int,DC=tt,DC=local" -Properties:*
$maxUid = $NIS.msSFU30MaxUidNumber + 1
Set-ADObject $NIS -Replace @{msSFU30MaxUidNumber = "$($maxUid)"}

   Set-ADUser -Identity $Username -Replace @{mssfu30nisdomain = "int"} #Enable NIS
   Set-ADUser -Identity $Username -Replace @{gidnumber="10000"} #Set Group ID
   $maxUid++ #Raise the User ID number
   Set-ADUser -Identity $Username -Replace @{uidnumber=$maxUid} #Set User ID
   Set-ADUser -Identity $Username -Replace @{loginshell="/bin/bash"} #Set user login shell
   Set-ADUser -Identity $Username -Replace @{msSFU30Name="$($Username)"}
   Set-ADUser -Identity $Username -Replace @{unixHomeDirectory="/home/$($Username)"}
   Write-Host -Backgroundcolor Green -Foregroundcolor Black $usr.SamAccountName changed #Write Changed Username to console	

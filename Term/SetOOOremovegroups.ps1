if($args.length -ne 1)
{
	write-host "usage:  SetOOO.ps1 <username>"
	exit
}
$name = $args[0]

Add-PSSnapin Quest.ActiveRoles.ADManagement | out-null

Write-Host -Foreground Gray "------------------------------------------------------------------------------------------------------------------------"
Write-Host -Foreground Cyan "Termed Employee Information"
Write-Host -Foreground Gray "------------------------------------------------------------------------------------------------------------------------"
#first check to see if the employee is a tradeco employee
connect-qadservice -service tradeco.int.tt.local | out-null 
$user = Get-QADUser $name
if($user -eq $null)
{
$tradeco = "yes"
disconnect-qadservice | out-null
connect-qadservice | out-null
$user = Get-QADUser $name
}


$user | out-default
Write-Host -Foreground Gray "------------------------------------------------------------------------------------------------------------------------"
Write-Host ""
Write-Host ""



$DispName = ""
$Location = ""
$suffix = ""

$DispName = $user.Name
$split = $DispName.split(" ")
$suffix = $split[2] 
try{$suffix = $suffix.substring(1,$suffix.length-2)}
catch{}
try{$location = $suffix.split("-")}
catch{}

$location = $location[1]

$Gmail = $user.Email


Write-Host ""	
	$answer=Read-Host "For"$Gmail",do you wish to use a standard OOO reply (1),set a custom OOO reply (2), or do not set any OOO reply (3)?"
	if ($answer -eq 1)
	{
	$oooreply=“This person is no longer an employee of Trading Technologies. For all TT related inquiries please contact us at +1(312)476-1000.”
	}
	elseif ($answer -eq 2)
	{
	$oooreply=Read-Host "Please type the custom OOO message and then press enter"
	}
	
Write-Host ""
Write-Host ""	
Write-Host -Foreground Gray "-----------------"
Write-Host -Foreground Cyan "OOO reply set to:" $oooreply
Write-Host -Foreground Gray "-----------------"	
	
	
	#Set status of a users AutoReply and make sure it is set
if ($answer -eq 3)
	{
	write-host "No OOO reply will be set for this user"
	}
else
	{
c:\gam\gam.exe user $Gmail vacation on subject "This email is no longer valid" message $oooreply enddate 2154-1-18
}

Write-Host "Removing $Dispname from all Google Groups" -ForegroundColor Red
$purge_usr = $Gmail
$purge= c:\gam\gam.exe info user $purge_usr
$purge_chunk= $purge | Select-String "Groups:" -context 0,100
$purge_grps=$purge_chunk.tostring().split(")")
foreach ($line in $purge_grps)
{
$grpaddresspurge=$line.tostring().split("<")[-1].Trim(">")
    if ($grpaddresspurge.contains("(direct member")) 
    {  
    $purgegrp=$grpaddresspurge.replace("> (direct member","") 
    
    c:\gam\gam.exe update group $purgegrp remove owner $purge_usr
    c:\gam\gam.exe update group $purgegrp remove member $purge_usr
    }
}

$purge_usr = $Gmail
$purge= c:\gam\gam.exe info user $purge_usr
$purge_chunk= $purge | Select-String "Groups:" -context 0,100
$purge_grps=$purge_chunk.tostring().split(")")
foreach ($line in $purge_grps)
{
$grpaddresspurge=$line.tostring().split("<")[-1].Trim(">")
    if ($grpaddresspurge.contains("(direct member")) 
    {  
    $purgegrp=$grpaddresspurge.replace("> (direct member","") 
    
    c:\gam\gam.exe update group $purgegrp remove owner $purge_usr
    c:\gam\gam.exe update group $purgegrp remove member $purge_usr
    }
}
Write-Host ""
Write-Host "User Removed from all Google Groups"
Write-Host ""
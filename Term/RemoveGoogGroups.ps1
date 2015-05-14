$Gmail = $args[0]

Write-Host ""
Write-Host ""
Write-Host ""
Write-Host ""	
	$answer=Read-Host "Do you wish to use a standard OOO reply (1),set a custom OOO reply (2), or do not set any OOO reply (3)?"
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


    #Remove user from all Google Groups

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
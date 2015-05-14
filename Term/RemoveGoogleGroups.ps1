if($args.length -ne 1)
{
	write-host "usage:  RemoveGoogleGroups.ps1 <username>"
	exit
}

$purge_usr = $args[0]
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